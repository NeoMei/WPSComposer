"""Offline startup checks plus real task-only TLS on ephemeral loopback ports."""
import hashlib
import http.client
import importlib.util
import json
from pathlib import Path
import shutil
import socket
import ssl
import subprocess
import tempfile
import threading
import time
import unittest
from unittest.mock import patch
ROOT=Path(__file__).resolve().parents[1]
PACKAGE=ROOT.parent/'officejs-word-probe-prepared'
spec=importlib.util.spec_from_file_location('probe_static_server',ROOT/'server.py')
server=importlib.util.module_from_spec(spec);spec.loader.exec_module(server)

class ServerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp=tempfile.TemporaryDirectory(prefix='word-probe-server-tls-')
        cls.path=Path(cls.temp.name);cls.cert=cls.path/'test-cert.pem';cls.key=cls.path/'test-key.pem'
        subprocess.run(['openssl','req','-x509','-newkey','rsa:2048','-nodes','-keyout',str(cls.key),'-out',str(cls.cert),'-days','1','-subj','/CN=localhost','-addext','subjectAltName=DNS:localhost,IP:127.0.0.1'],check=True,capture_output=True,timeout=30)
        cls.key.chmod(0o600)
    @classmethod
    def tearDownClass(cls):
        cls.temp.cleanup()
        assert not cls.path.exists(), 'Task-only certificate directory was not removed'

    def setUp(self):
        self.work=tempfile.TemporaryDirectory(dir=self.path,prefix='assets-');self.package=Path(self.work.name)
        for name in (*server.ASSET_HASHES,'HASHES.json'):shutil.copy2(PACKAGE/name,self.package/name)
        self.httpd=None;self.thread=None
    def tearDown(self):
        if self.httpd:
            self.httpd.shutdown();self.httpd.server_close();self.thread.join(timeout=3)
            self.assertFalse(self.thread.is_alive())
            with self.assertRaises(OSError):socket.create_connection(('127.0.0.1',self.port),timeout=.2)
        self.work.cleanup()
    def start(self,timeout=.3):
        self.httpd=server.create_server(self.package,self.cert,self.key,port=0,socket_timeout=timeout)
        self.thread=threading.Thread(target=self.httpd.serve_forever,kwargs={'poll_interval':.02});self.thread.start()
        self.port=self.httpd.server_address[1]
    def request(self,path,method='GET',host=None):
        context=ssl.create_default_context(cafile=str(self.cert)) # process-local trust only
        conn=http.client.HTTPSConnection('127.0.0.1',self.port,context=context,timeout=2)
        conn.request(method,path,headers={'Host':host or f'localhost:{self.port}'})
        response=conn.getresponse();body=response.read();result=(response.status,dict(response.getheaders()),body);conn.close();return result
    def test_assets_match_frozen_manifest_and_pinned_hashes(self):
        assets=server.load_assets(self.package)
        self.assertEqual(set(assets),set('/'+n for n in server.ASSET_HASHES))
        for path,(_,data) in assets.items():self.assertEqual(hashlib.sha256(data).hexdigest(),server.ASSET_HASHES[path[1:]])
    def test_tamper_missing_symlink_and_forged_manifest_fail_before_bind(self):
        for mutation in ('asset','missing','symlink','manifest'):
            with self.subTest(mutation=mutation):
                for name in (*server.ASSET_HASHES,'HASHES.json'):
                    p=self.package/name
                    if p.exists() or p.is_symlink():p.unlink()
                    shutil.copy2(PACKAGE/name,p)
                if mutation=='asset':(self.package/'probe.js').write_text('modified')
                if mutation=='missing':(self.package/'taskpane.css').unlink()
                if mutation=='symlink':
                    (self.package/'probe.js').unlink();(self.package/'probe.js').symlink_to(PACKAGE/'probe.js')
                if mutation=='manifest':
                    p=self.package/'HASHES.json';data=json.loads(p.read_text());data['sha256']['probe.js']='0'*64;p.write_text(json.dumps(data))
                with patch.object(server.StaticServer,'server_bind',side_effect=AssertionError('Must not bind')),self.assertRaises(ValueError):
                    server.create_server(self.package,self.cert,self.key,port=0)
    def test_tls_serves_only_four_exact_assets_with_correct_types(self):
        self.start();self.assertEqual(self.httpd.server_address[0],'127.0.0.1')
        for name in server.ASSET_HASHES:
            status,headers,body=self.request('/'+name)
            self.assertEqual(status,200);self.assertEqual(headers['Content-Type'],server.CONTENT_TYPES[name]);self.assertEqual(body,(PACKAGE/name).read_bytes())
    def test_reject_paths_methods_and_foreign_hosts_without_leaking_files(self):
        self.start()
        for path in ('/','/HASHES.json','/manifest.word.xml','/test-key.pem','/../probe.js','/%2e%2e/probe.js','//probe.js','/probe.js#x','/probe.js/','https://localhost/probe.js'):
            with self.subTest(path=path):self.assertEqual(self.request(path)[0],404)
        for method in ('POST','PUT','DELETE','HEAD','OPTIONS','PATCH','TRACE'):
            with self.subTest(method=method):self.assertEqual(self.request('/probe.js',method)[0],405)
        self.assertEqual(self.request('/probe.js',host='foreign.example')[0],403)
    def test_opaque_host_query_is_ignored_on_exact_asset_path(self):
        self.start()
        for path in ('/taskpane.html?_host_Info=Word$Mac$opaque', '/taskpane.html?file=../test-key.pem', '/taskpane.html?url=https://foreign/secret', '/taskpane.html?%2f..%2f=opaque'):
            self.assertEqual(self.request(path)[2],(PACKAGE/'taskpane.html').read_bytes())
        self.assertEqual(self.request('/taskpane.html?'+'a'*2100)[0],414)
        self.assertEqual(self.request('/../test-key.pem?_host_Info=opaque')[0],404)
    def test_bytes_are_frozen_after_startup_not_reloaded_from_disk(self):
        self.start();(self.package/'probe.js').write_text('unreviewed change')
        self.assertEqual(self.request('/probe.js')[2],(PACKAGE/'probe.js').read_bytes())
    def test_tls_and_read_timeouts_do_not_block_next_client(self):
        self.start(timeout=.15)
        raw=socket.create_connection(('127.0.0.1',self.port),timeout=2)
        context=ssl.create_default_context(cafile=str(self.cert));idle=context.wrap_socket(socket.create_connection(('127.0.0.1',self.port),timeout=2),server_hostname='localhost')
        time.sleep(.35)
        self.assertEqual(raw.recv(1),b'');self.assertEqual(idle.recv(1),b'');raw.close();idle.close()
        self.assertEqual(self.request('/probe.js')[0],200)
    def test_missing_certificate_and_invalid_timeout_never_bind(self):
        for cert,timeout in ((self.path/'absent.pem',2),(self.cert,0),(self.cert,11)):
            with patch.object(server.StaticServer,'server_bind',side_effect=AssertionError('Must not bind')),self.assertRaises((ValueError,FileNotFoundError)):
                server.create_server(self.package,cert,self.key,port=0,socket_timeout=timeout)
if __name__=='__main__':unittest.main()
