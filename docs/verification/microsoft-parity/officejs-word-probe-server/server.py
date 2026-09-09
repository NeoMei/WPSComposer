"""Serve only four pinned Office.js probe assets over supplied-certificate TLS.

No installation, certificate generation/trust, app access or command endpoint.
"""
from __future__ import annotations
import argparse
import hashlib
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
from pathlib import Path
import socket
import ssl
import threading
from types import MappingProxyType

ASSET_HASHES = MappingProxyType({
    'taskpane.html': '66d009da4426202fd357a10d6f9a9c6793ea703dd68c3f46f84541dbcd2458b4',
    'taskpane.css': 'c6ec516146a3e54e1b04104dde0091a986544c5fc2d62ec2d3b3ca64d5feadc2',
    'taskpane.js': '45e32946e5d74b496a44908613a1154c42cdf19b015ee46c958278c69a5170a9',
    'probe.js': '8da68b4f4e2420e66262e3b56086a71db3251c3269d54d85ccd7fd8ebdf0f046',
})
MANIFEST_HASH = '53bd0cf03da7a9dd885d69b885f17ef34c60e5441fd62569b371b26f69af0c5c'
CONTENT_TYPES = MappingProxyType({'taskpane.html':'text/html; charset=utf-8',
    'taskpane.css':'text/css; charset=utf-8','taskpane.js':'application/javascript; charset=utf-8',
    'probe.js':'application/javascript; charset=utf-8'})
MAX_FILE_BYTES = 1024 * 1024
MAX_TARGET = 2048


def _read_regular(root, name):
    path = root / name
    if path.is_symlink() or not path.is_file() or path.resolve().parent != root:
        raise ValueError('Missing or nonregular frozen file')
    with path.open('rb') as stream:
        value = stream.read(MAX_FILE_BYTES + 1)
    if len(value) > MAX_FILE_BYTES:
        raise ValueError('Frozen file exceeds size bound')
    return value


def load_assets(package):
    supplied = Path(package)
    if supplied.is_symlink():
        raise ValueError('Package symlink is not allowed')
    try:
        root = supplied.resolve(strict=True)
        manifest = _read_regular(root, 'HASHES.json')
        if hashlib.sha256(manifest).hexdigest() != MANIFEST_HASH:
            raise ValueError('Frozen manifest hash mismatch')
        declared = json.loads(manifest)['sha256']
        assets = {}
        for name, expected in ASSET_HASHES.items():
            if declared.get(name) != expected:
                raise ValueError('Frozen manifest asset mismatch')
            data = _read_regular(root, name)
            if hashlib.sha256(data).hexdigest() != expected:
                raise ValueError('Frozen asset hash mismatch')
            assets['/' + name] = (CONTENT_TYPES[name], data)
        return MappingProxyType(assets)
    except (OSError, KeyError, TypeError, json.JSONDecodeError) as exc:
        raise ValueError('Frozen assets could not be verified') from exc


class StaticServer(ThreadingHTTPServer):
    daemon_threads = False
    block_on_close = True
    allow_reuse_address = False
    request_queue_size = 8

    def __init__(self, assets, context, port, socket_timeout):
        self.assets = assets
        self.tls_context = context
        self.socket_timeout = socket_timeout
        self._slots = threading.BoundedSemaphore(8)
        super().__init__(('127.0.0.1', port), AssetHandler)

    def get_request(self):
        connection, address = super().get_request()
        connection.settimeout(self.socket_timeout)
        return connection, address

    def process_request(self, request, client_address):
        if not self._slots.acquire(blocking=False):
            self.shutdown_request(request)
            return
        try:
            super().process_request(request, client_address)
        except BaseException:
            self._slots.release()
            raise

    def process_request_thread(self, request, client_address):
        try:
            super().process_request_thread(request, client_address)
        finally:
            self._slots.release()

    def handle_error(self, request, client_address):
        # No request URL/query, TLS data, document metadata or paths are logged.
        pass


class AssetHandler(BaseHTTPRequestHandler):
    protocol_version = 'HTTP/1.0'
    server_version = 'FrozenWordProbe'
    sys_version = ''

    def setup(self):
        # TLS handshakes run in bounded workers, not the accept loop.
        self.request = self.server.tls_context.wrap_socket(self.request, server_side=True)
        self.request.settimeout(self.server.socket_timeout)
        super().setup()
        self._deadline = threading.Timer(self.server.socket_timeout, self._expire)
        self._deadline.daemon = True
        self._deadline.start()

    def _expire(self):
        try:
            self.connection.shutdown(socket.SHUT_RDWR)
        except OSError:
            pass

    def finish(self):
        self._deadline.cancel()
        try:
            super().finish()
        finally:
            self.connection.close()

    def log_message(self, *args):
        pass

    def _reply(self, status, body=b'', content_type='text/plain; charset=utf-8'):
        self.close_connection = True
        self.send_response(status)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Connection', 'close')
        self.send_header('Cache-Control', 'no-store')
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.end_headers()
        if self.command != 'HEAD':
            self.wfile.write(body)

    def send_error(self, code, message=None, explain=None):
        self._reply(code, b'Request rejected\n')

    def parse_request(self):
        if not super().parse_request():
            return False
        if self.command != 'GET':
            self._reply(405, b'Method rejected\n')
            return False
        return True

    def do_GET(self):
        # BaseHTTPRequestHandler normalizes leading // in self.path; use raw target.
        target = self.requestline.split()[1]
        if len(target) > MAX_TARGET:
            self._reply(414, b'Target too long\n')
            return
        if '#' in target:
            self._reply(404, b'Not found\n')
            return
        path = target.partition('?')[0]
        # Opaque Office host query is ignored, never parsed or used as identity.
        asset = self.server.assets.get(path)
        if asset is None:
            self._reply(404, b'Not found\n')
            return
        port = self.server.server_address[1]
        if self.headers.get_all('Host', []) not in ([f'localhost:{port}'], [f'127.0.0.1:{port}']):
            self._reply(403, b'Host rejected\n')
            return
        content_type, data = asset
        self._reply(200, data, content_type)


def create_server(package, cert, key, *, port=3443, socket_timeout=2.0):
    if type(port) is not int or not 0 <= port <= 65535:
        raise ValueError('Invalid port')
    if isinstance(socket_timeout, bool) or not isinstance(socket_timeout, (int, float)) or not 0 < socket_timeout <= 10:
        raise ValueError('Socket timeout must be within (0, 10] seconds')
    assets = load_assets(package)
    context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    context.minimum_version = ssl.TLSVersion.TLSv1_2
    # Empty password avoids an interactive prompt for encrypted private keys.
    context.load_cert_chain(certfile=str(cert), keyfile=str(key), password=b'')
    return StaticServer(assets, context, port, socket_timeout)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--package', required=True, type=Path)
    parser.add_argument('--cert', required=True, type=Path)
    parser.add_argument('--key', required=True, type=Path)
    parser.add_argument('--port', type=int, default=3443)
    args = parser.parse_args()
    if not 1 <= args.port <= 65535:
        parser.error('Port must be between 1 and 65535')
    with create_server(args.package, args.cert, args.key, port=args.port) as httpd:
        print(f'Frozen probe assets ready on https://127.0.0.1:{args.port}', flush=True)
        try:
            httpd.serve_forever(poll_interval=0.2)
        except KeyboardInterrupt:
            pass


if __name__ == '__main__':
    main()
