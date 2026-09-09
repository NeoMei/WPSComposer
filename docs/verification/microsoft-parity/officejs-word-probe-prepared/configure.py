"""Render only this package's manifest. No install, server, certificate or app access."""
from __future__ import annotations
import argparse
from pathlib import Path
from urllib.parse import urlsplit
from xml.sax.saxutils import escape
ROOT=Path(__file__).resolve().parent
OUTPUT=ROOT/'manifest.word.xml'

def render(origin):
    if not isinstance(origin,str) or any(c.isspace() for c in origin):raise ValueError('Expected HTTPS loopback origin')
    try:
        url=urlsplit(origin)
        valid=(url.scheme=='https' and url.hostname in ('localhost','127.0.0.1','::1')
               and url.username is None and url.password is None and not url.path and not url.query and not url.fragment
               and (url.port is None or 1<=url.port<=65535))
        # Reject alternate spellings and delimiters rather than normalize uncertain input.
        host='[::1]' if url.hostname=='::1' else url.hostname
        canonical='https://'+str(host)+(':'+str(url.port) if url.port is not None else '')
    except ValueError as exc:raise ValueError('Invalid origin') from exc
    if not valid or canonical!=origin:raise ValueError('Only an exact HTTPS loopback origin is allowed')
    return (ROOT/'manifest.word.xml.template').read_text().replace('__HTTPS_LOOPBACK_ORIGIN__',escape(origin,{'"':'&quot;'}))

def main():
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('--origin',required=True);args=parser.parse_args()
    value=render(args.origin)
    with OUTPUT.open('x',encoding='utf-8') as stream:stream.write(value)
    print(OUTPUT)
if __name__=='__main__':main()
