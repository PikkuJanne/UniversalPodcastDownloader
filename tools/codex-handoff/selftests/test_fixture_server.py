"""Helper-only tests. Passing these is NOT downloader acceptance."""
import http.client
import importlib.util
import json
from pathlib import Path
import tempfile
import subprocess
import sys
import time
import threading
import unittest
import xml.etree.ElementTree as ET

KIT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location('upd_fixture_server', KIT / 'fixture_server.py')
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


class HelperTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.server = module.FixtureServer()
        cls.thread = threading.Thread(target=cls.server.serve_forever,
                                      kwargs={'poll_interval': 0.02}, daemon=True)
        cls.thread.start()
        cls.port = cls.server.server_address[1]
        cls.audio = cls.server.audio

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        cls.thread.join(timeout=2)

    def request(self, path, headers=None, method='GET'):
        conn = http.client.HTTPConnection('127.0.0.1', self.port, timeout=3)
        try:
            conn.request(method, path, headers=headers or {})
            resp = conn.getresponse()
            return resp.status, dict(resp.getheaders()), resp.read()
        finally:
            conn.close()

    def test_loopback_only(self):
        self.assertEqual(self.server.server_address[0], '127.0.0.1')

    def test_single_feed_is_substituted_and_valid(self):
        status, _, data = self.request('/feeds/single.xml')
        self.assertEqual(status, 200)
        self.assertNotIn(b'{{', data)
        tree = ET.fromstring(data)
        enclosure = tree.find('./channel/item/enclosure')
        self.assertEqual(int(enclosure.attrib['length']), len(self.audio))
        self.assertTrue(enclosure.attrib['url'].startswith(self.server.base_url))

    def test_history_feed_changes_titles_and_requires_the_exact_renewed_query(self):
        self.server.recovered.clear()
        try:
            _, _, original = self.request('/feeds/history.xml')
            before = ET.fromstring(original)
            old_url = before.find('./channel/item/enclosure').attrib['url']
            old_path = old_url.removeprefix(self.server.base_url)
            self.assertEqual(self.request(old_path)[2], self.audio)
            self.assertEqual(self.request('/media/history.mp3')[0], 403)
            self.request('/__recover', method='POST')
            _, _, renewed = self.request('/feeds/history.xml')
            after = ET.fromstring(renewed)
            self.assertNotEqual(before.find('./channel/title').text, after.find('./channel/title').text)
            self.assertNotEqual(before.find('./channel/item/title').text, after.find('./channel/item/title').text)
            self.assertEqual(before.find('./channel/item/guid').text, after.find('./channel/item/guid').text)
            new_url = after.find('./channel/item/enclosure').attrib['url']
            self.assertEqual(self.request(new_url.removeprefix(self.server.base_url))[2], self.audio)
            self.assertEqual(self.request(old_path)[0], 403)
        finally:
            self.server.recovered.clear()

    def test_safe_xml_fixture_shapes(self):
        for name in ('rss-single.xml', 'rss-empty.xml', 'rss-collisions.xml',
                     'atom-dates.xml', 'rss-date-cases.xml', 'rss-media.xml',
                     'rss-page-1.xml', 'rss-page-2.xml'):
            with self.subTest(name=name):
                ET.fromstring((KIT / 'fixtures' / name).read_bytes())

    def test_legacy_remote_change_preserves_guid_but_changes_url_bytes_and_validator(self):
        self.server.recovered.clear()
        try:
            before = ET.fromstring(self.request('/feeds/legacy-changing.xml')[2])
            old_path = before.find('./channel/item/enclosure').attrib['url'].removeprefix(self.server.base_url)
            status, old_headers, old_body = self.request(old_path)
            self.assertEqual((status, old_body), (200, self.audio))
            self.request('/__recover', method='POST')
            after = ET.fromstring(self.request('/feeds/legacy-changing.xml')[2])
            new_path = after.find('./channel/item/enclosure').attrib['url'].removeprefix(self.server.base_url)
            self.assertNotEqual(old_path, new_path)
            self.assertEqual(before.find('./channel/item/guid').text, after.find('./channel/item/guid').text)
            status, new_headers, new_body = self.request(new_path)
            self.assertEqual(status, 200)
            self.assertNotEqual(old_body, new_body)
            self.assertNotEqual(old_headers['ETag'], new_headers['ETag'])
            self.assertEqual(int(new_headers['Content-Length']), len(new_body))
            self.assertEqual(self.request(old_path)[0], 403)
        finally:
            self.server.recovered.clear()

    def test_malformed_fixture_really_is_malformed(self):
        with self.assertRaises(ET.ParseError):
            ET.fromstring((KIT / 'fixtures/xml-malformed.xml').read_bytes())

    def test_dtd_fixture_is_present_without_resolving_it(self):
        data = (KIT / 'fixtures/xml-with-dtd.xml').read_text()
        self.assertIn('<!DOCTYPE', data)
        self.assertIn('entity-must-never-be-fetched.invalid', data)

    def test_normal_media_and_head(self):
        status, hdr, data = self.request('/media/ok.mp3')
        self.assertEqual((status, data), (200, self.audio))
        self.assertEqual(hdr['ETag'], '"fixture-v1"')
        _, hh, hd = self.request('/media/ok.mp3', method='HEAD')
        self.assertEqual(int(hh['Content-Length']), len(self.audio))
        self.assertEqual(hd, b'')

    def test_range_can_reassemble_exact_original(self):
        offset = 100
        status, hdr, body = self.request('/media/ok.mp3',
            {'Range': 'bytes=100-', 'If-Range': '"fixture-v1"'})
        self.assertEqual(status, 206)
        self.assertEqual(hdr['Content-Range'], 'bytes 100-{}/{}'.format(len(self.audio)-1, len(self.audio)))
        self.assertEqual(self.audio[:offset] + body, self.audio)

    def test_suffix_and_invalid_range(self):
        status, _, data = self.request('/media/ok.mp3', {'Range': 'bytes=-12'})
        self.assertEqual((status, data), (206, self.audio[-12:]))
        status, hdr, _ = self.request('/media/ok.mp3', {'Range': 'bytes=999999-'})
        self.assertEqual(status, 416)
        self.assertEqual(hdr['Content-Range'], 'bytes */{}'.format(len(self.audio)))

    def test_ignored_range_and_changed_entity(self):
        status, _, data = self.request('/media/ignore-range.mp3', {'Range': 'bytes=100-'})
        self.assertEqual((status, data), (200, self.audio))
        status, hdr, data = self.request('/media/changed.mp3',
            {'Range': 'bytes=100-', 'If-Range': '"fixture-v1"'})
        self.assertEqual(status, 200)
        self.assertEqual(hdr['ETag'], '"fixture-v2"')
        self.assertNotEqual(data, self.audio)

    def test_weak_and_absent_validator(self):
        status, hdr, data = self.request('/media/weak.mp3',
            {'Range': 'bytes=100-', 'If-Range': 'W/"fixture-v1"'})
        self.assertEqual((status, data), (200, self.audio))
        self.assertTrue(hdr['ETag'].startswith('W/'))
        _, hdr, _ = self.request('/media/no-validator.mp3')
        self.assertNotIn('ETag', hdr)

    def test_intentionally_bad_range(self):
        status, hdr, data = self.request('/media/bad-range.mp3', {'Range': 'bytes=100-'})
        self.assertEqual(status, 206)
        self.assertTrue(hdr['Content-Range'].startswith('bytes 101-'))
        self.assertEqual(data, self.audio[100:])

    def test_forced_416(self):
        self.assertEqual(self.request('/media/always-416.mp3')[0], 416)

    def test_truncation_is_detectable(self):
        conn = http.client.HTTPConnection('127.0.0.1', self.port, timeout=3)
        try:
            conn.request('GET', '/media/truncated.mp3')
            response = conn.getresponse()
            with self.assertRaises(http.client.IncompleteRead) as cm:
                response.read()
            self.assertEqual(cm.exception.partial, self.audio[:len(self.audio)//3])
        finally:
            conn.close()

    def test_wrong_or_missing_content_metadata(self):
        self.assertEqual(self.request('/media/empty.mp3')[2], b'')
        _, headers, data = self.request('/media/html.mp3')
        self.assertEqual(headers['Content-Type'], 'audio/mpeg')
        self.assertTrue(data.startswith(b'<!doctype'))
        _, headers, data = self.request('/media/octet-stream')
        self.assertEqual(headers['Content-Type'], 'application/octet-stream')
        self.assertEqual(data, self.audio)
        _, headers, data = self.request('/media/no-length.mp3')
        self.assertNotIn('Content-Length', headers)
        self.assertEqual(data, self.audio)

    def test_retry_once_and_statuses(self):
        self.request('/__reset', method='POST')
        status, headers, _ = self.request('/retry/once.mp3')
        self.assertEqual(status, 503)
        self.assertEqual(headers['Retry-After'], '1')
        self.assertEqual(self.request('/retry/once.mp3')[0], 200)
        for code in (403, 404, 429, 503):
            self.assertEqual(self.request('/status/{}'.format(code))[0], code)

    def test_redirect_and_query_redacted_stats(self):
        status, hdr, _ = self.request('/redirect/show')
        self.assertEqual((status, hdr['Location']), (302, '/show'))
        self.request('/media/ok.mp3?fake_token=DO_NOT_LOG_QUERY')
        _, _, data = self.request('/__stats')
        self.assertNotIn(b'DO_NOT_LOG_QUERY', data)
        self.assertIn('/media/ok.mp3', json.loads(data))

    def test_unknown_route_is_not_filesystem_access(self):
        self.assertEqual(self.request('/../../etc/passwd')[0], 404)
        self.assertEqual(self.request('/fixtures/../fixture_server.py')[0], 404)

    def test_ready_file_never_overwrites(self):
        with tempfile.TemporaryDirectory(prefix='upd-helper-') as root:
            path = Path(root)/'ready.json'
            module.write_ready_file(path, '{"first":true}')
            with self.assertRaises(FileExistsError):
                module.write_ready_file(path, '{"overwrite":true}')
            self.assertIn('first', path.read_text())

    def test_cli_readiness_and_no_overwrite(self):
        with tempfile.TemporaryDirectory(prefix='upd-cli-helper-') as directory:
            ready = Path(directory) / 'ready.json'
            proc = subprocess.Popen([sys.executable, '-B', str(KIT / 'fixture_server.py'),
                                     '--port', '0', '--ready-file', str(ready)],
                                    stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            try:
                deadline = time.monotonic() + 5
                while not ready.exists() and time.monotonic() < deadline:
                    if proc.poll() is not None:
                        self.fail('Fixture CLI exited before readiness')
                    time.sleep(0.02)
                self.assertTrue(ready.exists())
                payload = json.loads(ready.read_text())
                self.assertTrue(payload['base_url'].startswith('http://127.0.0.1:'))
                self.assertTrue(payload['fixture_only'])
            finally:
                proc.terminate()
                proc.communicate(timeout=5)
            original = ready.read_bytes()
            result = subprocess.run([sys.executable, '-B', str(KIT / 'fixture_server.py'),
                                     '--ready-file', str(ready)], capture_output=True, timeout=5)
            self.assertNotEqual(result.returncode, 0)
            self.assertEqual(ready.read_bytes(), original)

    def test_stall_routes_finish(self):
        self.assertEqual(self.request('/media/stall.mp3')[2], self.audio)
        self.assertEqual(self.request('/media/stall-headers.mp3')[2], self.audio)


if __name__ == '__main__':
    unittest.main()
