from __future__ import annotations

import socket
import struct
import time

import pytest

from skills.WPSComposer.scripts.msoffice.macos_excel_launcher import _transfer, _remaining, LauncherError


def test_protocol_roundtrip_uses_exact_frame_and_independent_clock():
    left,right=socket.socketpair()
    left.setblocking(False);right.setblocking(False)
    try:
        _transfer(left,{'version':1,'id':7},10,lambda:1)
        assert _transfer(right,None,10,lambda:1,receive=True)=={'version':1,'id':7}
    finally:left.close();right.close()


@pytest.mark.parametrize('raw',[b'[]',b'null',b'{invalid'])
def test_nonobject_or_malformed_json_is_rejected(raw):
    left,right=socket.socketpair()
    left.setblocking(False);right.setblocking(False)
    try:
        left.sendall(struct.pack('!I',len(raw))+raw)
        with pytest.raises(LauncherError,match='EXCEL_LAUNCHER_BAD_FRAME'):
            _transfer(right,None,time.monotonic()+1,time.monotonic,receive=True)
    finally:left.close();right.close()


def test_partial_frame_shares_original_deadline():
    left,right=socket.socketpair()
    left.setblocking(False);right.setblocking(False)
    try:
        left.sendall(struct.pack('!I',20)+b'{')
        started=time.monotonic()
        with pytest.raises(LauncherError,match='EXCEL_LAUNCHER_TIMEOUT'):
            _transfer(right,None,started+.03,time.monotonic,receive=True)
        assert time.monotonic()-started < .5
    finally:left.close();right.close()


@pytest.mark.parametrize('deadline',[0,float('nan'),float('inf')])
def test_invalid_remaining_budget_never_waits(deadline):
    with pytest.raises(LauncherError,match='EXCEL_LAUNCHER_TIMEOUT'):
        _remaining(deadline,lambda:1)
