#!/usr/bin/env python3
"""Loopback-only adversarial HTTP fixtures. Development tool, never app runtime."""
from __future__ import annotations

import argparse
import json
import re
import socket
import threading
import time
from html import escape
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Dict, Optional, Tuple
from urllib.parse import urlsplit

FIXTURES = Path(__file__).resolve().parent / 'fixtures'
FEEDS = {
    '/feeds/single.xml': 'rss-single.xml', '/feeds/empty.xml': 'rss-empty.xml',
    '/feeds/collisions.xml': 'rss-collisions.xml', '/feeds/atom.xml': 'atom-dates.xml',
    '/feeds/dates.xml': 'rss-date-cases.xml', '/feeds/media.xml': 'rss-media.xml',
    '/feeds/page-1.xml': 'rss-page-1.xml', '/feeds/page-2.xml': 'rss-page-2.xml',
    '/feeds/dtd.xml': 'xml-with-dtd.xml', '/feeds/malformed.xml': 'xml-malformed.xml',
    '/feeds/single-alias.xml': 'rss-single.xml',
    '/show': 'show-multiple.html', '/show/not-feed': 'show-not-feed.html',
}


class FixtureServer(ThreadingHTTPServer):
    """Only binds IPv4 loopback, even when instantiated by a test."""
    daemon_threads = True
    allow_reuse_address = True

    def __init__(self, port: int = 0):
        self.counts: Dict[str, int] = {}
        self.count_lock = threading.Lock()
        self.audio = (FIXTURES / 'silence.mp3').read_bytes()
        super().__init__(('127.0.0.1', port), FixtureHandler)

    @property
    def base_url(self) -> str:
        return 'http://127.0.0.1:{}'.format(self.server_address[1])

    def count(self, path: str) -> int:
        with self.count_lock:
            self.counts[path] = self.counts.get(path, 0) + 1
            return self.counts[path]

    def stats(self) -> Dict[str, int]:
        with self.count_lock:
            return dict(self.counts)


class FixtureHandler(BaseHTTPRequestHandler):
    protocol_version = 'HTTP/1.0'
    server_version = 'UPDLocalFixture/1.0'

    def log_message(self, *_args: object) -> None:
        # Avoid recording even fake request query values or headers.
        return

    def do_GET(self) -> None:
        self._safe_dispatch(False)

    def do_HEAD(self) -> None:
        self._safe_dispatch(True)

    def do_POST(self) -> None:
        if urlsplit(self.path).path != '/__reset':
            self._send(404, b'Unknown fixture route', False, 'text/plain')
            return
        with self.server.count_lock:
            self.server.counts.clear()
        self._send(200, b'{"reset":true}', False, 'application/json')

    def _safe_dispatch(self, head: bool) -> None:
        try:
            self._dispatch(head)
        except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError):
            # Expected when a downloader cancels or times out against a fixture.
            self.close_connection = True

    def _headers(self, status: int, length: Optional[int], content_type: str,
                 extra: Optional[Dict[str, str]] = None) -> None:
        self.send_response(status)
        self.send_header('Content-Type', content_type)
        self.send_header('Connection', 'close')
        if length is not None:
            self.send_header('Content-Length', str(length))
        for key, value in (extra or {}).items():
            self.send_header(key, value)
        self.end_headers()
        self.close_connection = True

    def _send(self, status: int, body: bytes, head: bool,
              content_type: str = 'audio/mpeg',
              extra: Optional[Dict[str, str]] = None) -> None:
        self._headers(status, len(body), content_type, extra)
        if not head:
            self.wfile.write(body)

    def _dispatch(self, head: bool) -> None:
        path = urlsplit(self.path).path
        number = self.server.count(path)
        if path == '/__stats':
            self._send(200, json.dumps(self.server.stats()).encode(), head, 'application/json')
            return
        hostile_titles = {
            '/feeds/hostile-dot.xml': '.',
            '/feeds/hostile-dotdot.xml': '..',
            '/feeds/hostile-drive.xml': r'C:\UPD-Synthetic-DoNotCreate',
            '/feeds/hostile-unc.xml': r'\\127.0.0.1\upd-fixture-do-not-create',
        }
        if path in hostile_titles or path == '/feeds/long-names.xml':
            title = hostile_titles.get(path, 'Long podcast title ' * 25)
            episode_title = 'Long episode title ' * 30 if path == '/feeds/long-names.xml' else 'Synthetic episode'
            identifiers = ('long-one', 'long-two') if path == '/feeds/long-names.xml' else ('hostile',)
            items = ''.join(
                '<item><title>{}</title><guid isPermaLink="false">{}</guid>'
                '<pubDate>Tue, 01 Sep 2026 12:00:00 +0000</pubDate>'
                '<enclosure url="{}/media/ok.mp3?id={}" type="audio/mpeg"/></item>'.format(
                    escape(episode_title), identifier, self.server.base_url, identifier)
                for identifier in identifiers)
            body = '<rss version="2.0"><channel><title>{}</title>{}</channel></rss>'.format(escape(title), items)
            self._send(200, body.encode('utf-8'), head, 'application/xml; charset=utf-8')
            return
        if path in FEEDS:
            name = FEEDS[path]
            text = (FIXTURES / name).read_text(encoding='utf-8')
            text = text.replace('{{BASE_URL}}', self.server.base_url)
            text = text.replace('{{AUDIO_BYTES}}', str(len(self.server.audio)))
            mime = 'text/html; charset=utf-8' if name.endswith('.html') else 'application/xml; charset=utf-8'
            self._send(200, text.encode('utf-8'), head, mime)
            return
        if path in ('/redirect/show', '/redirect/loop'):
            target = '/show' if path == '/redirect/show' else '/redirect/loop'
            self._send(302, b'', head, 'text/plain', {'Location': target})
            return
        if path in ('/status/403', '/status/404', '/status/429', '/status/503'):
            status = int(path.rsplit('/', 1)[1])
            extra = {'Retry-After': '1'} if status in (429, 503) else None
            self._send(status, b'Synthetic status', head, 'text/plain', extra)
            return
        if path == '/retry/once.mp3' and number == 1:
            self._send(503, b'Synthetic first-attempt failure', head, 'text/plain', {'Retry-After': '1'})
            return
        known = {
            '/media/ok.mp3', '/media/ignore-range.mp3', '/media/changed.mp3',
            '/media/weak.mp3', '/media/no-validator.mp3', '/media/bad-range.mp3',
            '/media/always-416.mp3', '/media/truncated.mp3', '/media/empty.mp3',
            '/media/html.mp3', '/media/octet-stream', '/media/no-length.mp3',
            '/media/stall.mp3', '/media/stall-headers.mp3', '/retry/once.mp3',
        }
        if path not in known:
            self._send(404, b'Unknown fixture route', head, 'text/plain')
            return
        data = self.server.audio
        etag = '"fixture-v1"'
        if path == '/media/changed.mp3':
            # Still MP3 data, but a different representation and validator.
            data = data + b'\0' * 16
            etag = '"fixture-v2"'
        if path == '/media/weak.mp3':
            etag = 'W/"fixture-v1"'
        extra = {'Accept-Ranges': 'bytes'}
        if path != '/media/no-validator.mp3':
            extra['ETag'] = etag
        if path == '/media/empty.mp3':
            self._send(200, b'', head, extra=extra)
            return
        if path == '/media/html.mp3':
            self._send(200, b'<!doctype html><html><body>Synthetic error</body></html>', head)
            return
        if path == '/media/always-416.mp3':
            self._send(416, b'', head, extra={'Content-Range': 'bytes */{}'.format(len(data))})
            return
        if path == '/media/truncated.mp3':
            self._headers(200, len(data), 'audio/mpeg', extra)
            if not head:
                self.wfile.write(data[:len(data) // 3])
                self.wfile.flush()
                try:
                    self.connection.shutdown(socket.SHUT_WR)
                except OSError:
                    pass
            return
        if path == '/media/no-length.mp3':
            self._headers(200, None, 'audio/mpeg', extra)
            if not head:
                self.wfile.write(data)
            return
        if path == '/media/stall.mp3':
            self._headers(200, len(data), 'audio/mpeg', extra)
            if not head:
                self.wfile.write(data[:1])
                self.wfile.flush()
                time.sleep(0.6)
                self.wfile.write(data[1:])
            return
        if path == '/media/stall-headers.mp3':
            time.sleep(0.6)
        range_header = self.headers.get('Range')
        if_range = self.headers.get('If-Range')
        allow_range = path != '/media/ignore-range.mp3'
        if if_range is not None and (if_range != etag or etag.startswith('W/')):
            allow_range = False
        mime = 'application/octet-stream' if path == '/media/octet-stream' else 'audio/mpeg'
        if range_header and allow_range and not head:
            span = parse_single_range(range_header, len(data))
            if span is None:
                self._send(416, b'', head, extra={'Content-Range': 'bytes */{}'.format(len(data))})
                return
            start, end = span
            advertised_start = start + 1 if path == '/media/bad-range.mp3' else start
            extra['Content-Range'] = 'bytes {}-{}/{}'.format(advertised_start, end, len(data))
            self._send(206, data[start:end + 1], head, mime, extra)
            return
        self._send(200, data, head, mime, extra)


def parse_single_range(value: str, length: int) -> Optional[Tuple[int, int]]:
    """Minimal single-byte-range parser; unsupported/malformed ranges return None."""
    match = re.fullmatch(r'bytes=(\d*)-(\d*)', value)
    if not match or length <= 0:
        return None
    first, last = match.groups()
    if not first and not last:
        return None
    if not first:
        suffix = int(last)
        return (max(0, length - suffix), length - 1) if suffix > 0 else None
    start = int(first)
    end = int(last) if last else length - 1
    if start >= length or end < start:
        return None
    return start, min(end, length - 1)


def write_ready_file(path: Path, payload: str) -> None:
    """Ready file must be new; caller owns and pre-creates its temp parent."""
    with path.open('x', encoding='utf-8') as handle:
        handle.write(payload + '\n')


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--port', type=int, default=0, help='0 selects an ephemeral loopback port')
    parser.add_argument('--ready-file', type=Path, help='New file in caller-owned temporary directory')
    args = parser.parse_args()
    if not 0 <= args.port <= 65535:
        parser.error('port must be between 0 and 65535')
    with FixtureServer(args.port) as server:
        payload = json.dumps({'base_url': server.base_url, 'fixture_only': True})
        if args.ready_file:
            write_ready_file(args.ready_file, payload)
        print(payload, flush=True)
        try:
            server.serve_forever(poll_interval=0.1)
        except KeyboardInterrupt:
            pass
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
