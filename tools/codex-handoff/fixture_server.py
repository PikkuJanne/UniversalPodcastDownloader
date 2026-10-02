#!/usr/bin/env python3
"""Loopback-only adversarial HTTP fixtures. Development tool, never app runtime."""
from __future__ import annotations

import argparse
import json
import math
import re
import socket
import threading
import time
from email.utils import formatdate
from html import escape
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Dict, Optional, Tuple
from urllib.parse import parse_qs, urlsplit

FIXTURES = Path(__file__).resolve().parent / 'fixtures'
FEEDS = {
    '/feeds/single.xml': 'rss-single.xml', '/feeds/empty.xml': 'rss-empty.xml',
    '/feeds/collisions.xml': 'rss-collisions.xml', '/feeds/atom.xml': 'atom-dates.xml',
    '/feeds/dates.xml': 'rss-date-cases.xml', '/feeds/media.xml': 'rss-media.xml',
    '/feeds/page-1.xml': 'rss-page-1.xml', '/feeds/page-2.xml': 'rss-page-2.xml',
    '/feeds/dtd.xml': 'xml-with-dtd.xml', '/feeds/malformed.xml': 'xml-malformed.xml',
    '/feeds/single-alias.xml': 'rss-single.xml',
    '/show': 'show-multiple.html', '/show/not-feed': 'show-not-feed.html',
    '/discovery/final/show.html': 'show-relative.html',
    '/discovery/base.html': 'show-base.html',
    '/discovery/single.html': 'show-single.html',
    '/discovery/nonfeed-link.html': 'show-nonfeed-link.html',
    '/feeds/empty-atom.xml': 'atom-empty.xml',
    '/feeds/wrong-root.xml': 'xml-wrong-root.xml',
    '/feeds/wrong-atom-namespace.xml': 'atom-wrong-namespace.xml',
    '/feeds/publication-order.xml': 'rss-publication-order.xml',
    '/feeds/publication-midnight.xml': 'atom-publication-midnight.xml',
}


class FixtureServer(ThreadingHTTPServer):
    """Only binds IPv4 loopback, even when instantiated by a test."""
    daemon_threads = True
    allow_reuse_address = True

    def __init__(self, port: int = 0):
        self.counts: Dict[str, int] = {}
        self.transport_events = []
        self.resume_events = []
        self.count_lock = threading.Lock()
        self.recovered = threading.Event()
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
        if urlsplit(self.path).path == '/__recover':
            self.server.recovered.set()
            self._send(200, b'{"recovered":true}', False, 'application/json')
            return
        if urlsplit(self.path).path != '/__reset':
            self._send(404, b'Unknown fixture route', False, 'text/plain')
            return
        with self.server.count_lock:
            self.server.counts.clear()
            self.server.transport_events.clear()
            self.server.resume_events.clear()
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

    def _transport(self, path: str, number: int, head: bool) -> None:
        # Only these named, bounded scenarios exist. Never proxy or read a path
        # supplied by the caller. Events contain fixed route names, not queries.
        match = re.fullmatch(r'/transport/(feed|metadata|media)/([a-z0-9-]+)', path)
        if not match:
            self._send(404, b'Unknown transport fixture', head, 'text/plain')
            return
        kind, scenario = match.groups()
        scenarios = {'once-' + str(code) for code in (408, 429, 500, 502, 503, 504)}
        scenarios.update('permanent-' + str(code) for code in (400, 403, 404, 501))
        scenarios.update(('always-503', 'retry-delta', 'retry-date', 'defer-delta',
                          'defer-date', 'delay-headers', 'stall-body', 'progress',
                          'truncated-once', 'redirect-wait', 'redirect-defer', 'redirect-target'))
        if scenario not in scenarios:
            self._send(404, b'Unknown transport fixture', head, 'text/plain')
            return
        body = ('<rss version="2.0"><channel><title>Transport fixture</title>'
                '<item><title>Original synthetic audio</title><guid>transport-001</guid>'
                '<enclosure url="{}/transport/media/{}" type="audio/mpeg" '
                'length="{}"/></item></channel></rss>').format(
                    self.server.base_url, scenario, len(self.server.audio)).encode('utf-8')
        content_type = 'application/xml; charset=utf-8'
        if kind == 'feed':
            self._send(200, body, head, content_type)
            return
        if kind == 'media':
            body, content_type = self.server.audio, 'audio/mpeg'
        event = {'path': path, 'number': number, 'received': time.time()}
        with self.server.count_lock:
            self.server.transport_events.append(event)
        if scenario in ('redirect-wait', 'redirect-defer'):
            delay = 5 if scenario == 'redirect-defer' else 1
            with self.server.count_lock:
                event['not_before'] = time.time() + delay
            self._send(302, b'', head, 'text/plain', {
                'Location': '/transport/{}/redirect-target'.format(kind),
                'Retry-After': str(delay),
            })
            return
        if scenario.startswith('permanent-') or scenario == 'always-503' or (
                scenario.startswith('once-') and number == 1):
            status = int(scenario.rsplit('-', 1)[1])
            self._send(status, b'Synthetic response', head, 'text/plain')
            return
        if scenario.startswith(('retry-', 'defer-')) and number == 1:
            delay = 5 if scenario.startswith('defer-') else 1
            now = time.time()
            not_before = math.ceil(now) + delay if scenario.endswith('-date') else now + delay
            retry_after = formatdate(not_before, usegmt=True) if scenario.endswith('-date') else str(delay)
            with self.server.count_lock:
                event['not_before'] = not_before
            self._send(429, b'Synthetic response', head, 'text/plain', {'Retry-After': retry_after})
            return
        if scenario == 'delay-headers':
            time.sleep(1.0)
        if scenario == 'stall-body':
            self._headers(200, len(body), content_type)
            if not head:
                self.wfile.write(body[:10])
                self.wfile.flush()
                time.sleep(1.0)
                self.wfile.write(body[10:])
            return
        if scenario == 'progress':
            self._headers(200, len(body), content_type)
            if not head:
                width = max(1, len(body) // 20)
                for offset in range(0, len(body), width):
                    self.wfile.write(body[offset:offset + width])
                    self.wfile.flush()
                    time.sleep(0.075)
            return
        if scenario == 'truncated-once' and number == 1:
            self._headers(200, len(body), content_type)
            if not head:
                self.wfile.write(body[:len(body) // 2])
                self.wfile.flush()
                self.connection.shutdown(socket.SHUT_WR)
            return
        self._send(200, body, head, content_type)

    def _resume(self, path: str, number: int, head: bool) -> None:
        match = re.fullmatch(r'/resume/(feed|media)/([a-z0-9-]+)', path)
        scenarios = {'valid', 'ignore-range', 'changed', 'weak', 'absent',
                     'last-modified', 'no-length', 'bad-start', 'bad-end',
                     'bad-total', 'bad-length', 'bad-validator', 'missing-validator',
                     'weak-validator', 'bad-type', 'bad-encoding', 'multipart',
                     'range-416', 'range-416-local', 'always-416', 'signed',
                     'range-truncated-once', 'fresh-truncated-once'}
        if not match or match.group(2) not in scenarios:
            self._send(404, b'Unknown resume fixture', head, 'text/plain')
            return
        kind, scenario = match.groups()
        recovered = self.server.recovered.is_set()
        data = self.server.audio * 64
        etag = '"resume-v1"'
        if scenario == 'changed' and recovered:
            data += b'\0' * 32
            etag = '"resume-v2"'
        if kind == 'feed':
            suffix = '?signature={}'.format('renewed' if recovered else 'original') if scenario == 'signed' else ''
            body = ('<rss version="2.0"><channel><title>Resume fixture</title>'
                    '<item><title>Original resume audio</title><guid>resume-001</guid>'
                    '<enclosure url="{}/resume/media/{}{}" type="audio/mpeg" '
                    'length="{}"/></item></channel></rss>').format(
                        self.server.base_url, scenario, suffix, len(data)).encode('utf-8')
            self._send(200, body, head, 'application/xml; charset=utf-8')
            return
        query_variant = None
        if scenario == 'signed':
            expected = 'signature={}'.format('renewed' if recovered else 'original')
            query_variant = 'renewed' if recovered else 'original'
            if urlsplit(self.path).query != expected:
                self._send(403, b'Synthetic signature mismatch', head, 'text/plain')
                return
        range_header = self.headers.get('Range')
        if_range = self.headers.get('If-Range')
        # Record only bounded, syntactically safe synthetic range evidence.
        event = {'path': path, 'number': number, 'query_variant': query_variant,
                 'range': range_header if range_header and re.fullmatch(r'bytes=[0-9]+-', range_header) else None,
                 'if_range': if_range if if_range in ('"resume-v1"', '"resume-v2"') else None,
                 'has_range': range_header is not None, 'has_if_range': if_range is not None}
        with self.server.count_lock:
            self.server.resume_events.append(event)
        headers = {'Accept-Ranges': 'bytes', 'ETag': etag}
        if scenario == 'weak':
            headers['ETag'] = 'W/"resume-v1"'
        if scenario in ('absent', 'last-modified'):
            headers.pop('ETag')
        if scenario == 'last-modified':
            headers['Last-Modified'] = 'Tue, 01 Sep 2026 12:00:00 GMT'
        if scenario == 'fresh-truncated-once' and number == 1:
            self._headers(200, len(data), 'audio/mpeg', headers)
            if not head:
                self.wfile.write(data[:len(data) // 3])
                self.wfile.flush()
                self.connection.shutdown(socket.SHUT_WR)
            return
        if recovered and scenario == 'always-416':
            self._send(416, b'', head, extra={'Content-Range': 'bytes */{}'.format(len(data))})
            return
        if range_header and recovered:
            if scenario in ('range-416', 'range-416-local'):
                total = int(range_header[6:-1]) if scenario == 'range-416-local' and re.fullmatch(r'bytes=[0-9]+-', range_header) else len(data)
                self._send(416, b'', head, extra={'Content-Range': 'bytes */{}'.format(total)})
                return
            span = parse_single_range(range_header, len(data))
            if span is None:
                self._send(416, b'', head, extra={'Content-Range': 'bytes */{}'.format(len(data))})
                return
            if scenario != 'ignore-range' and if_range == etag:
                start, end = span
                range_start = start + 1 if scenario == 'bad-start' else start
                range_end = end - 1 if scenario == 'bad-end' else end
                total = len(data) + 1 if scenario == 'bad-total' else len(data)
                headers['Content-Range'] = 'bytes {}-{}/{}'.format(range_start, range_end, total)
                if scenario == 'bad-validator':
                    headers['ETag'] = '"resume-v2"'
                elif scenario == 'missing-validator':
                    headers.pop('ETag')
                elif scenario == 'weak-validator':
                    headers['ETag'] = 'W/"resume-v1"'
                elif scenario == 'bad-encoding':
                    headers['Content-Encoding'] = 'gzip'
                mime = 'audio/ogg' if scenario == 'bad-type' else 'audio/mpeg'
                if scenario == 'multipart':
                    mime = 'multipart/byteranges; boundary=synthetic'
                body = data[start:end + 1]
                length = len(body) + 1 if scenario == 'bad-length' else len(body)
                self._headers(206, length, mime, headers)
                if not head:
                    if scenario == 'range-truncated-once' and number == 2:
                        self.wfile.write(body[:len(body) // 2])
                        self.wfile.flush()
                        self.connection.shutdown(socket.SHUT_WR)
                        return
                    self.wfile.write(body)
                return
        self._headers(200, None if scenario == 'no-length' else len(data), 'audio/mpeg', headers)
        if not head:
            if not recovered:
                boundary = 65536
                self.wfile.write(data[:boundary])
                self.wfile.flush()
                # The owned test worker is terminated before the local test
                # releases this finite gate. Later requests complete normally.
                self.server.recovered.wait(timeout=15)
                self.wfile.write(data[boundary:])
            else:
                self.wfile.write(data)

    def _dispatch(self, head: bool) -> None:
        path = urlsplit(self.path).path
        number = self.server.count(path)
        if path == '/__stats':
            self._send(200, json.dumps(self.server.stats()).encode(), head, 'application/json')
            return
        if path == '/__transport':
            with self.server.count_lock:
                events = list(self.server.transport_events)
            self._send(200, json.dumps(events).encode(), head, 'application/json')
            return
        if path.startswith('/transport/'):
            self._transport(path, number, head)
            return
        if path == '/__resume':
            with self.server.count_lock:
                events = list(self.server.resume_events)
            self._send(200, json.dumps(events).encode(), head, 'application/json')
            return
        if path.startswith('/resume/'):
            self._resume(path, number, head)
            return
        if path == '/entity/never':
            self._send(200, b'ENTITY_MUST_NOT_BE_READ', head, 'text/plain')
            return
        boundary_xml = {
            '/feeds/external-http.xml', '/feeds/external-file.xml',
            '/feeds/internal-dtd.xml', '/feeds/oversized.xml',
            '/feeds/oversized-no-length.xml', '/feeds/deep.xml',
            '/feeds/many-nodes.xml', '/feeds/redirect-media.xml',
            '/feeds/redirect-file-media.xml', '/feeds/no-cookie.xml',
        }
        if path in boundary_xml:
            declaration = ''
            title = 'Input boundary fixture'
            extra_xml = ''
            media_path = '/media/ok.mp3'
            if path in ('/feeds/external-http.xml', '/feeds/external-file.xml'):
                target = self.server.base_url + '/entity/never'
                if path.endswith('external-file.xml'):
                    # Reflect a bounded URI into XML only. Never open the path.
                    target = parse_qs(urlsplit(self.path).query).get('file', ['file:///synthetic/never'])[0]
                    if len(target) > 4096 or not target.startswith('file:///'):
                        self._send(400, b'Expected a bounded synthetic file URI', head, 'text/plain')
                        return
                declaration = '<!DOCTYPE rss [<!ENTITY xxe SYSTEM "{}">]>'.format(escape(target, quote=True))
                title = '&xxe;'
            elif path.endswith('internal-dtd.xml'):
                declaration = '<!DOCTYPE rss [<!ENTITY x "synthetic"><!ENTITY y "&x;&x;&x;&x;">]>'
                title = '&y;'
            elif 'oversized' in path:
                extra_xml = '<description>' + ('x' * (8 * 1024 * 1024 + 64)) + '</description>'
            elif path.endswith('deep.xml'):
                extra_xml = '<n>' * 70 + 'bounded depth fixture' + '</n>' * 70
            elif path.endswith('many-nodes.xml'):
                extra_xml = '<n/>' * 100010
            elif path.endswith('redirect-media.xml'):
                media_path = '/media/redirect.mp3'
            elif path.endswith('redirect-file-media.xml'):
                media_path = '/media/redirect-file.mp3'
            elif path.endswith('no-cookie.xml'):
                if self.headers.get('Cookie') or self.headers.get('Authorization'):
                    self.server.count('/credential-received')
            body = ('{}<rss version="2.0"><channel><title>{}</title>{}'
                    '<item><title>Boundary episode</title><guid>boundary-001</guid>'
                    '<enclosure url="{}{}" type="audio/mpeg"/></item></channel></rss>').format(
                        declaration, title, extra_xml, self.server.base_url, media_path).encode('utf-8')
            if path.endswith('oversized-no-length.xml'):
                self._headers(200, None, 'application/xml; charset=utf-8')
                if not head:
                    self.wfile.write(body)
            else:
                self._send(200, body, head, 'application/xml; charset=utf-8')
            return
        if path in ('/feeds/publication-history-rss.xml', '/feeds/publication-history-atom.xml'):
            changed = self.server.recovered.is_set()
            if path.endswith('-rss.xml'):
                date = 'Fri, 01 Jan 2027 00:00:00 +0000' if changed else 'Tue, 01 Sep 2026 00:30:00 +1400'
                body = ('<rss version="2.0"><channel><title>Publication history</title>'
                        '<item><title>Publication history episode</title>'
                        '<guid isPermaLink="false">publication-history-001</guid>'
                        '<pubDate>{}</pubDate><enclosure url="{}/media/ok.mp3" '
                        'type="audio/mpeg" length="{}"/></item></channel></rss>').format(
                            date, self.server.base_url, len(self.server.audio))
            else:
                date = '2027-01-01T00:00:00Z' if changed else '2026-09-01T00:30:00+14:00'
                body = ('<feed xmlns="http://www.w3.org/2005/Atom"><title>Publication history</title>'
                        '<entry><title>Publication history episode</title>'
                        '<id>urn:fixture:publication-history-001</id>'
                        '<published>{}</published><updated>2028-01-01T00:00:00Z</updated>'
                        '<link rel="enclosure" href="{}/media/ok.mp3" '
                        'type="audio/mpeg" length="{}"/></entry></feed>').format(
                            date, self.server.base_url, len(self.server.audio))
            self._send(200, body.encode('utf-8'), head, 'application/xml; charset=utf-8')
            return
        if path in ('/feeds/history.xml', '/feeds/legacy-changing.xml'):
            changed = self.server.recovered.is_set()
            title = 'Renamed history show' if changed else 'Original history show'
            episode = 'Renamed episode' if changed else 'Original episode'
            token = 'renewed' if changed else 'original'
            media_name = 'legacy-changing' if path == '/feeds/legacy-changing.xml' else 'history'
            body = ('<rss version="2.0"><channel><title>{}</title>'
                    '<item><title>{}</title><guid isPermaLink="false">history-stable-001</guid>'
                    '<pubDate>Tue, 01 Sep 2026 12:00:00 +0000</pubDate>'
                    '<enclosure url="{}/media/{}.mp3?signature={}&amp;part=1" '
                    'type="audio/mpeg" length="{}"/></item></channel></rss>').format(
                        title, episode, self.server.base_url, media_name, token, len(self.server.audio))
            self._send(200, body.encode('utf-8'), head, 'application/xml; charset=utf-8')
            return
        transaction_media = {
            '/feeds/transaction-empty.xml': ('/media/empty.mp3', len(self.server.audio)),
            '/feeds/transaction-html.xml': ('/media/html.mp3', len(self.server.audio)),
            '/feeds/transaction-truncated.xml': ('/media/truncated.mp3', len(self.server.audio)),
            '/feeds/transaction-no-length.xml': ('/media/no-length.mp3', len(self.server.audio)),
            '/feeds/transaction-octet.xml': ('/media/octet-stream', len(self.server.audio)),
            '/feeds/transaction-enclosure-mismatch.xml': ('/media/ok.mp3', len(self.server.audio) + 12345),
            '/feeds/transaction-crash.xml': ('/media/ok.mp3', len(self.server.audio)),
            '/feeds/transaction-recover.xml': ('/media/recover.mp3', len(self.server.audio)),
            '/feeds/transaction-interrupt.xml': ('/media/interrupt.mp3', len(self.server.audio) * 64),
            '/feeds/transaction-json.xml': ('/media/json.mp3', len(self.server.audio)),
            '/feeds/transaction-xml.xml': ('/media/xml.mp3', len(self.server.audio)),
            '/feeds/transaction-partial.xml': ('/media/unsolicited-partial.mp3', len(self.server.audio)),
        }
        if path in transaction_media:
            media_path, estimate = transaction_media[path]
            body = ('<rss version="2.0"><channel><title>Transactional fixtures</title>'
                    '<item><title>Transactional episode</title><guid isPermaLink="false">transactional-001</guid>'
                    '<pubDate>Tue, 01 Sep 2026 12:00:00 +0000</pubDate>'
                    '<enclosure url="{}{}" type="audio/mpeg" length="{}"/></item></channel></rss>').format(
                        self.server.base_url, media_path, estimate)
            self._send(200, body.encode('utf-8'), head, 'application/xml; charset=utf-8')
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
        if path in ('/redirect/show', '/redirect/loop', '/discovery/redirect'):
            target = {
                '/redirect/show': '/show', '/redirect/loop': '/redirect/loop',
                '/discovery/redirect': '/discovery/final/show.html',
            }[path]
            self._send(302, b'', head, 'text/plain', {'Location': target})
            return
        redirects = {
            '/redirect/feed': '../feeds/single.xml',
            '/redirect/file': 'file:///synthetic/never',
            '/redirect/userinfo': self.server.base_url.replace('http://', 'http://fake:FAKE_TOKEN@') + '/feeds/single.xml',
            '/media/redirect.mp3': '/media/ok.mp3',
            '/media/redirect-file.mp3': 'file:///synthetic/never',
            '/redirect/cookie': 'http://localhost:{}/feeds/no-cookie.xml'.format(self.server.server_address[1]),
        }
        if path in redirects:
            headers = {'Location': redirects[path]}
            if path == '/redirect/cookie':
                headers['Set-Cookie'] = 'synthetic=FAKE_COOKIE; Path=/'
            self._send(302, b'', head, 'text/plain', headers)
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
            '/media/recover.mp3', '/media/interrupt.mp3', '/media/json.mp3',
            '/media/xml.mp3', '/media/unsolicited-partial.mp3', '/media/history.mp3',
            '/media/legacy-changing.mp3',
        }
        if path not in known:
            self._send(404, b'Unknown fixture route', head, 'text/plain')
            return
        if path in ('/media/history.mp3', '/media/legacy-changing.mp3'):
            token = 'renewed' if self.server.recovered.is_set() else 'original'
            if urlsplit(self.path).query != 'signature={}&part=1'.format(token):
                self._send(403, b'Synthetic signature mismatch', head, 'text/plain')
                return
        data = self.server.audio
        etag = '"fixture-v1"'
        if path == '/media/legacy-changing.mp3' and self.server.recovered.is_set():
            data = data + b'\0' * 32
            etag = '"legacy-changed-v2"'
        if path == '/media/changed.mp3':
            # Still MP3 data, but a different representation and validator.
            data = data + b'\0' * 16
            etag = '"fixture-v2"'
        if path == '/media/weak.mp3':
            etag = 'W/"fixture-v1"'
        extra = {'Accept-Ranges': 'bytes'}
        # Legacy fault cases exercise cleanup/restart without resume evidence.
        # Validator-bearing crash and recovery cases have explicit /resume routes.
        no_validator = {'/media/no-validator.mp3', '/media/truncated.mp3',
                        '/media/recover.mp3', '/media/interrupt.mp3'}
        if path not in no_validator:
            extra['ETag'] = etag
        if path == '/media/empty.mp3':
            self._send(200, b'', head, extra=extra)
            return
        if path == '/media/html.mp3':
            self._send(200, b'<!doctype html><html><body>Synthetic error</body></html>', head)
            return
        if path in ('/media/json.mp3', '/media/xml.mp3'):
            body = b'{"error":"synthetic denial"}' if path.endswith('json.mp3') else b'<?xml version="1.0"?><error>synthetic denial</error>'
            self._send(200, body, head, 'audio/mpeg')
            return
        if path == '/media/unsolicited-partial.mp3':
            self._send(206, data[:len(data) // 2], head, extra={'Content-Range': 'bytes 0-{}/{}'.format(len(data) // 2 - 1, len(data))})
            return
        if path == '/media/interrupt.mp3':
            data = data * 64
            self._headers(200, len(data), 'audio/mpeg', extra)
            if not head:
                boundary = min(65536, len(data) - 1)
                self.wfile.write(data[:boundary])
                self.wfile.flush()
                # Tests release this loopback-only gate after terminating their
                # owned downloader child. The bounded wait cannot hang cleanup.
                self.server.recovered.wait(timeout=15)
                self.wfile.write(data[boundary:])
            return
        if path == '/media/always-416.mp3':
            self._send(416, b'', head, extra={'Content-Range': 'bytes */{}'.format(len(data))})
            return
        if path == '/media/truncated.mp3' or (path == '/media/recover.mp3' and not self.server.recovered.is_set()):
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
