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
from email.utils import parsedate_to_datetime

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

    def test_discovery_redirect_and_relative_base_fixtures_are_named_and_local(self):
        status, headers, body = self.request('/discovery/redirect')
        self.assertEqual((status, headers['Location'], body),
                         (302, '/discovery/final/show.html', b''))
        status, _, body = self.request(headers['Location'])
        self.assertEqual(status, 200)
        self.assertIn(b'href="../../feeds/single.xml"', body)
        status, _, body = self.request('/discovery/base.html')
        self.assertEqual(status, 200)
        self.assertIn(b'<base href="base/">', body)
        self.assertIn(b'fixture=base&amp;part=1', body)

    def test_discovery_duplicate_entities_and_nonfeed_targets_are_preserved(self):
        status, _, body = self.request('/discovery/single.html')
        self.assertEqual(status, 200)
        self.assertIn(b'fixture=single&amp;part=1', body)
        self.assertIn(b'fixture=single&#38;part=1', body)
        status, _, body = self.request('/discovery/nonfeed-link.html')
        self.assertEqual(status, 200)
        self.assertIn(b'href="/show/not-feed"', body)

    def test_empty_atom_and_unsupported_xml_root_shapes_are_distinct(self):
        empty = ET.fromstring(self.request('/feeds/empty-atom.xml')[2])
        self.assertEqual(empty.tag, '{http://www.w3.org/2005/Atom}feed')
        self.assertEqual(len(empty.findall('{http://www.w3.org/2005/Atom}entry')), 0)
        nested = ET.fromstring(self.request('/feeds/wrong-root.xml')[2])
        self.assertEqual(nested.tag, 'document')
        self.assertIsNotNone(nested.find('rss'))
        unsupported = ET.fromstring(self.request('/feeds/wrong-atom-namespace.xml')[2])
        self.assertEqual(unsupported.tag, '{urn:fixture:unsupported}feed')

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

    def test_entity_routes_only_return_xml_and_do_not_read_external_resources(self):
        _, _, http_xml = self.request('/feeds/external-http.xml')
        self.assertIn(b'<!DOCTYPE rss', http_xml)
        self.assertIn((self.server.base_url + '/entity/never').encode(), http_xml)
        _, _, file_xml = self.request('/feeds/external-file.xml?file=file%3A%2F%2F%2Fsynthetic%2Fdoes-not-exist')
        self.assertIn(b'file:///synthetic/does-not-exist', file_xml)
        self.assertIn(b'&xxe;', file_xml)
        self.assertEqual(self.request('/feeds/external-file.xml?file=https%3A%2F%2Fexample.invalid')[0], 400)

    def test_internal_dtd_fixture_has_an_entity_reference(self):
        _, _, data = self.request('/feeds/internal-dtd.xml')
        self.assertIn(b'<!ENTITY', data)
        self.assertIn(b'<title>&y;</title>', data)

    def test_oversized_metadata_is_finite_with_both_framing_variants(self):
        _, headers, data = self.request('/feeds/oversized.xml')
        self.assertGreater(len(data), 8 * 1024 * 1024)
        self.assertLess(len(data), 8 * 1024 * 1024 + 4096)
        self.assertEqual(int(headers['Content-Length']), len(data))
        _, headers, unframed = self.request('/feeds/oversized-no-length.xml')
        self.assertNotIn('Content-Length', headers)
        self.assertEqual(unframed, data)

    def test_depth_and_node_fixtures_are_well_formed_but_exceed_policy_budgets(self):
        _, _, deep = self.request('/feeds/deep.xml')
        self.assertEqual(deep.count(b'<n>'), 70)
        ET.fromstring(deep)
        _, _, wide = self.request('/feeds/many-nodes.xml')
        self.assertEqual(wide.count(b'<n/>'), 100010)
        ET.fromstring(wide)

    def test_boundary_redirect_routes_expose_controlled_locations(self):
        cases = {
            '/redirect/feed': '../feeds/single.xml',
            '/redirect/file': 'file:///synthetic/never',
            '/media/redirect.mp3': '/media/ok.mp3',
            '/media/redirect-file.mp3': 'file:///synthetic/never',
        }
        for route, location in cases.items():
            with self.subTest(route=route):
                status, headers, body = self.request(route)
                self.assertEqual((status, headers['Location'], body), (302, location, b''))
        _, headers, _ = self.request('/redirect/userinfo')
        self.assertIn('fake:FAKE_TOKEN@127.0.0.1', headers['Location'])
        _, headers, _ = self.request('/redirect/cookie')
        self.assertIn('localhost:', headers['Location'])
        self.assertEqual(headers['Set-Cookie'], 'synthetic=FAKE_COOKIE; Path=/')

    def test_redirect_media_feed_points_to_its_controlled_redirect(self):
        for route, target in (
                ('/feeds/redirect-media.xml', '/media/redirect.mp3'),
                ('/feeds/redirect-file-media.xml', '/media/redirect-file.mp3')):
            tree = ET.fromstring(self.request(route)[2])
            self.assertEqual(tree.find('./channel/item/enclosure').attrib['url'], self.server.base_url + target)

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
        for path in ('/media/truncated.mp3', '/media/recover.mp3', '/media/interrupt.mp3'):
            _, headers, _ = self.request(path, method='HEAD')
            self.assertNotIn('ETag', headers)

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

    def test_transport_feed_only_points_to_named_synthetic_media(self):
        tree = ET.fromstring(self.request('/transport/feed/once-503')[2])
        self.assertEqual(tree.find('./channel/item/enclosure').attrib['url'],
                         self.server.base_url + '/transport/media/once-503')
        self.assertEqual(self.request('/transport/media/not-a-scenario')[0], 404)

    def test_transport_status_routes_recover_only_when_configured(self):
        self.request('/__reset', method='POST')
        for kind in ('metadata', 'media'):
            for code in (408, 429, 500, 502, 503, 504):
                route = '/transport/{}/once-{}'.format(kind, code)
                self.assertEqual(self.request(route)[0], code)
                self.assertEqual(self.request(route)[0], 200)
            for code in (400, 403, 404, 501):
                route = '/transport/{}/permanent-{}'.format(kind, code)
                self.assertEqual(self.request(route)[0], code)
                self.assertEqual(self.request(route)[0], code)

    def test_transport_retry_after_events_preserve_actual_server_deadline(self):
        self.request('/__reset', method='POST')
        for scenario in ('retry-delta', 'retry-date', 'defer-delta', 'defer-date'):
            route = '/transport/media/' + scenario
            status, headers, _ = self.request(route + '?fake_token=DO_NOT_RECORD')
            self.assertEqual(status, 429)
            events_raw = self.request('/__transport')[2]
            self.assertNotIn(b'DO_NOT_RECORD', events_raw)
            event = next(event for event in json.loads(events_raw) if event['path'] == route)
            if scenario.endswith('-date'):
                self.assertEqual(event['not_before'], parsedate_to_datetime(headers['Retry-After']).timestamp())
            else:
                self.assertGreaterEqual(event['not_before'] - event['received'], int(headers['Retry-After']))
            self.assertEqual(self.request(route)[2], self.audio)

    def test_transport_truncation_exposes_half_then_complete_original(self):
        self.request('/__reset', method='POST')
        with self.assertRaises(http.client.IncompleteRead) as caught:
            self.request('/transport/media/truncated-once')
        self.assertEqual(caught.exception.partial, self.audio[:len(self.audio) // 2])
        self.assertEqual(self.request('/transport/media/truncated-once')[2], self.audio)

    def test_transport_redirect_exposes_server_wait_and_owned_target(self):
        for scenario, delay in (('redirect-wait', '1'), ('redirect-defer', '5')):
            status, headers, body = self.request('/transport/media/' + scenario)
            self.assertEqual((status, body), (302, b''))
            self.assertEqual(headers['Retry-After'], delay)
            self.assertEqual(headers['Location'], '/transport/media/redirect-target')
        self.assertEqual(self.request('/transport/media/redirect-target')[2], self.audio)

    def test_transport_active_progress_delivers_the_original_body(self):
        started = time.monotonic()
        self.assertEqual(self.request('/transport/media/progress')[2], self.audio)
        self.assertGreater(time.monotonic() - started, 1.0)

    def test_resume_feed_has_stable_identity_and_rotates_only_synthetic_signature(self):
        self.server.recovered.clear()
        try:
            before = ET.fromstring(self.request('/resume/feed/signed')[2])
            self.server.recovered.set()
            after = ET.fromstring(self.request('/resume/feed/signed')[2])
            self.assertEqual(before.find('./channel/item/guid').text, after.find('./channel/item/guid').text)
            self.assertIn('signature=original', before.find('./channel/item/enclosure').attrib['url'])
            self.assertIn('signature=renewed', after.find('./channel/item/enclosure').attrib['url'])
            self.assertEqual(self.request('/resume/media/signed?signature=original')[0], 403)
            self.assertEqual(self.request('/resume/media/signed?signature=renewed')[2], self.audio * 64)
        finally:
            self.server.recovered.clear()

    def test_resume_partial_gate_and_matching_range_reconstruct_original_bytes(self):
        self.server.recovered.clear()
        conn = http.client.HTTPConnection('127.0.0.1', self.port, timeout=3)
        try:
            conn.request('GET', '/resume/media/valid')
            response = conn.getresponse()
            self.assertEqual(response.getheader('ETag'), '"resume-v1"')
            prefix = response.read(1024)
            self.assertEqual(prefix, (self.audio * 64)[:1024])
        finally:
            conn.close()
            self.server.recovered.set()
        try:
            status, headers, suffix = self.request('/resume/media/valid', {'Range': 'bytes=1024-', 'If-Range': '"resume-v1"'})
            self.assertEqual(status, 206)
            self.assertEqual(headers['Content-Range'], 'bytes 1024-{}/{}'.format(len(self.audio) * 64 - 1, len(self.audio) * 64))
            self.assertEqual(prefix + suffix, self.audio * 64)
        finally:
            self.server.recovered.clear()

    def test_resume_unknown_or_weak_representation_never_claims_a_strong_validator(self):
        for scenario in ('weak', 'absent', 'last-modified', 'no-length'):
            status, headers, _ = self.request('/resume/media/' + scenario, method='HEAD')
            self.assertEqual(status, 200)
            if scenario == 'weak':
                self.assertEqual(headers['ETag'], 'W/"resume-v1"')
            if scenario in ('absent', 'last-modified'):
                self.assertNotIn('ETag', headers)
            if scenario == 'last-modified':
                self.assertIn('Last-Modified', headers)
            if scenario == 'no-length':
                self.assertNotIn('Content-Length', headers)

    def test_resume_adversarial_range_headers_and_fresh_responses_are_distinct(self):
        self.server.recovered.set()
        try:
            cases = {
                'bad-start': ('Content-Range', 'bytes 1025-'),
                'bad-end': ('Content-Range', 'bytes 1024-{}'.format(len(self.audio) * 64 - 2)),
                'bad-total': ('Content-Range', 'bytes 1024-{}/{}'.format(len(self.audio) * 64 - 1, len(self.audio) * 64 + 1)),
                'bad-validator': ('ETag', '"resume-v2"'),
                'weak-validator': ('ETag', 'W/"resume-v1"'),
                'bad-type': ('Content-Type', 'audio/ogg'),
                'bad-encoding': ('Content-Encoding', 'gzip'),
                'multipart': ('Content-Type', 'multipart/byteranges'),
            }
            for scenario, (header, prefix) in cases.items():
                status, headers, _ = self.request('/resume/media/' + scenario, {'Range': 'bytes=1024-', 'If-Range': '"resume-v1"'})
                self.assertEqual(status, 206)
                self.assertTrue(headers[header].startswith(prefix), scenario)
                self.assertEqual(self.request('/resume/media/' + scenario)[2], self.audio * 64)
            with self.assertRaises(http.client.IncompleteRead):
                self.request('/resume/media/bad-length', {'Range': 'bytes=1024-', 'If-Range': '"resume-v1"'})
            status, headers, _ = self.request('/resume/media/missing-validator', {'Range': 'bytes=1024-', 'If-Range': '"resume-v1"'})
            self.assertEqual(status, 206)
            self.assertNotIn('ETag', headers)
        finally:
            self.server.recovered.clear()

    def test_resume_416_and_ignored_or_changed_ranges_are_explicit(self):
        self.server.recovered.set()
        try:
            request_headers = {'Range': 'bytes=1024-', 'If-Range': '"resume-v1"'}
            self.assertEqual(self.request('/resume/media/range-416', request_headers)[0], 416)
            self.assertEqual(self.request('/resume/media/range-416')[2], self.audio * 64)
            status, headers, _ = self.request('/resume/media/range-416-local', request_headers)
            self.assertEqual((status, headers['Content-Range']), (416, 'bytes */1024'))
            self.assertEqual(self.request('/resume/media/always-416')[0], 416)
            self.assertEqual(self.request('/resume/media/ignore-range', request_headers)[0], 200)
            status, headers, body = self.request('/resume/media/changed', request_headers)
            self.assertEqual((status, headers['ETag'], body), (200, '"resume-v2"', self.audio * 64 + b'\0' * 32))
        finally:
            self.server.recovered.clear()

    def test_resume_events_record_only_bounded_synthetic_range_values(self):
        self.request('/__reset', method='POST')
        self.server.recovered.set()
        try:
            self.request('/resume/media/valid?secret=DO_NOT_RECORD', {'Range': 'bytes=1024-', 'If-Range': '"resume-v1"'})
            self.request('/resume/media/valid?secret=DO_NOT_RECORD', {'If-Range': 'UNTRUSTED_DO_NOT_RECORD'})
            raw = self.request('/__resume')[2]
            self.assertNotIn(b'DO_NOT_RECORD', raw)
            events = json.loads(raw)
            self.assertEqual(events[0]['range'], 'bytes=1024-')
            self.assertEqual(events[0]['if_range'], '"resume-v1"')
            self.assertTrue(events[1]['has_if_range'])
            self.assertIsNone(events[1]['if_range'])
        finally:
            self.server.recovered.clear()

    def test_resume_caught_failure_fixtures_advance_from_the_observed_bytes(self):
        self.request('/__reset', method='POST')
        self.server.recovered.set()
        try:
            original = self.audio * 64
            with self.assertRaises(http.client.IncompleteRead) as fresh:
                self.request('/resume/media/fresh-truncated-once')
            prefix = fresh.exception.partial
            self.assertEqual(prefix, original[:len(original) // 3])
            status, _, tail = self.request('/resume/media/fresh-truncated-once', {
                'Range': 'bytes={}-'.format(len(prefix)), 'If-Range': '"resume-v1"'})
            self.assertEqual((status, prefix + tail), (206, original))
            self.request('/resume/media/range-truncated-once', method='HEAD')
            with self.assertRaises(http.client.IncompleteRead) as resumed:
                self.request('/resume/media/range-truncated-once', {'Range': 'bytes=1024-', 'If-Range': '"resume-v1"'})
            middle = resumed.exception.partial
            self.assertEqual(middle, original[1024:1024 + len(middle)])
            status, _, tail = self.request('/resume/media/range-truncated-once', {
                'Range': 'bytes={}-'.format(1024 + len(middle)), 'If-Range': '"resume-v1"'})
            self.assertEqual((status, original[:1024] + middle + tail), (206, original))
        finally:
            self.server.recovered.clear()


if __name__ == '__main__':
    unittest.main()
