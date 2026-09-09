"""Independent bounded raw-HTTP and worker-saturation TLS checks."""
import importlib.util
from pathlib import Path
import socket
import ssl
import time
import unittest

root=Path(__file__).resolve().parents[1]/'officejs-word-probe-server'
spec=importlib.util.spec_from_file_location('prepared_tls_tests',root/'tests/test_server.py')
existing=importlib.util.module_from_spec(spec);spec.loader.exec_module(existing)

class Review(existing.ServerTests):
    def test_raw_duplicate_and_missing_host_never_serve_asset(self):
        self.start()
        context=ssl.create_default_context(cafile=str(self.cert))
        for headers in [b'', f'Host: localhost:{self.port}\r\nHost: localhost:{self.port}\r\n'.encode()]:
            with socket.create_connection(('127.0.0.1',self.port),timeout=2) as raw:
                with context.wrap_socket(raw,server_hostname='localhost') as client:
                    client.sendall(b'GET /probe.js?_host_Info=opaque HTTP/1.1\r\n'+headers+b'\r\n')
                    data=b''
                    while True:
                        chunk=client.recv(65536)
                        if not chunk:break
                        data+=chunk
            self.assertIn(b'403',data.split(b'\r\n',1)[0])
            self.assertNotIn(b'WordCapabilityProbe',data)
    def test_eight_stalled_handshakes_release_workers_and_close(self):
        self.start(timeout=.2)
        clients=[]
        try:
            for _ in range(8):
                clients.append(socket.create_connection(('127.0.0.1',self.port),timeout=2))
            time.sleep(.35)
            for client in clients:self.assertEqual(client.recv(1),b'')
            self.assertEqual(self.request('/probe.js')[0],200)
        finally:
            for client in clients:client.close()

if __name__=='__main__':
    suite=unittest.TestSuite(Review(name) for name in (
        'test_raw_duplicate_and_missing_host_never_serve_asset',
        'test_eight_stalled_handshakes_release_workers_and_close'))
    result=unittest.TextTestRunner(verbosity=2).run(suite)
    raise SystemExit(not result.wasSuccessful())
