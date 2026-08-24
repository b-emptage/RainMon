"""Single-instance guard tests, for both processes that hold one.

The field incident (first seen on the dome server, same template): Ctrl-C the
process, restart it straight away, and be refused because "another instance is
already running" -- the restart raced the old process, which was still
shutting down. Both guards now wait a bounded grace for their port instead of
failing on the first bind, and both shutdown paths restore default signal
handlers so a stuck stop can always be killed -- freeing the port -- with a
second Ctrl-C.
"""
import signal
import socket
import threading

import pytest

import app
from greenhill import broadcaster


@pytest.fixture
def app_guard(monkeypatch):
    """Fresh guard global, quick polling, no leaked port after the test."""
    monkeypatch.setattr(app, '_instance_guard', None)
    monkeypatch.setattr(app, '_GUARD_POLL_SECONDS', 0.05)
    yield
    if app._instance_guard is not None:
        app._instance_guard.close()
        app._instance_guard = None


@pytest.fixture
def bridge_guard(monkeypatch):
    monkeypatch.setattr(broadcaster, '_instance_guard', None)
    monkeypatch.setattr(broadcaster, 'GUARD_POLL_SECONDS', 0.05)
    yield
    if broadcaster._instance_guard is not None:
        broadcaster._instance_guard.close()
        broadcaster._instance_guard = None


def _hold(port):
    incumbent = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    incumbent.bind(('127.0.0.1', port))
    incumbent.listen(1)
    return incumbent


# --------------------------------------------------------------------------- #
# The weather device server (device/app.py, port 50817)                       #
# --------------------------------------------------------------------------- #

def test_app_acquires_a_free_port_without_waiting(app_guard, capsys):
    assert app.acquire_single_instance_lock() is True
    assert app._instance_guard is not None
    assert capsys.readouterr().err == ''


def test_app_refused_immediately_without_grace(app_guard, capsys):
    incumbent = _hold(app._SINGLE_INSTANCE_PORT)
    try:
        assert app.acquire_single_instance_lock(grace_seconds=0.0) is False
    finally:
        incumbent.close()
    # A single failed attempt announces nothing: there was no wait to explain.
    assert 'Waiting' not in capsys.readouterr().err


def test_app_waits_out_a_dying_incumbent(app_guard, capsys):
    """The incident, in miniature: the port is released a moment after the new
    server asks for it, exactly as when a restart races the old shutdown."""
    incumbent = _hold(app._SINGLE_INSTANCE_PORT)
    releaser = threading.Timer(0.3, incumbent.close)
    releaser.start()
    try:
        assert app.acquire_single_instance_lock(grace_seconds=5.0) is True
    finally:
        releaser.join()

    assert 'may still be shutting down' in capsys.readouterr().err


def test_app_gives_up_when_the_incumbent_stays(app_guard, capsys):
    incumbent = _hold(app._SINGLE_INSTANCE_PORT)
    try:
        assert app.acquire_single_instance_lock(grace_seconds=0.4) is False
    finally:
        incumbent.close()

    assert 'may still be shutting down' in capsys.readouterr().err


def test_app_restore_default_signal_handlers():
    """On the way into the shutdown path, every further Ctrl-C or kill must act
    at the C level (SIG_DFL): a Python-level handler cannot run while the main
    thread is blocked in weather_dev.stop(), and a stuck process squats the
    guard port, refusing every restart."""
    saved = {sig: signal.getsignal(sig) for sig in (signal.SIGINT, signal.SIGTERM)}
    try:
        signal.signal(signal.SIGINT, lambda *a: None)
        signal.signal(signal.SIGTERM, app._signal_shutdown)

        app._restore_default_signal_handlers()

        assert signal.getsignal(signal.SIGINT) is signal.SIG_DFL
        assert signal.getsignal(signal.SIGTERM) is signal.SIG_DFL
    finally:
        for sig, handler in saved.items():
            signal.signal(sig, handler)


# --------------------------------------------------------------------------- #
# The rain bridge (greenhill/broadcaster.py, port 50816)                      #
# --------------------------------------------------------------------------- #

def test_bridge_refused_immediately_without_grace(bridge_guard, capsys):
    incumbent = _hold(broadcaster.SINGLE_INSTANCE_PORT)
    try:
        assert broadcaster.acquire_single_instance_lock(grace_seconds=0.0) is False
    finally:
        incumbent.close()
    assert 'Waiting' not in capsys.readouterr().err


def test_bridge_waits_out_a_dying_incumbent(bridge_guard, capsys):
    incumbent = _hold(broadcaster.SINGLE_INSTANCE_PORT)
    releaser = threading.Timer(0.3, incumbent.close)
    releaser.start()
    try:
        assert broadcaster.acquire_single_instance_lock(grace_seconds=5.0) is True
    finally:
        releaser.join()

    assert 'may still be shutting down' in capsys.readouterr().err


def test_bridge_second_signal_is_left_to_kill_us(monkeypatch):
    """The first signal asks the loop to stop AND restores SIG_DFL, so a second
    Ctrl-C kills the process even when it is wedged inside a serial driver
    call, where the Python-level handler could never run. The kill frees the
    mutex port, so the next start is not refused."""
    monkeypatch.setattr(broadcaster, '_shutdown_requested', False)
    saved = {}
    for name in ('SIGINT', 'SIGTERM', 'SIGBREAK'):
        sig = getattr(signal, name, None)
        if sig is not None:
            saved[sig] = signal.getsignal(sig)
    try:
        broadcaster.install_signal_handlers()
        handler = signal.getsignal(signal.SIGTERM)
        assert callable(handler)

        handler(signal.SIGTERM, None)

        assert broadcaster._shutdown_requested is True
        assert signal.getsignal(signal.SIGTERM) is signal.SIG_DFL
        # Only the received signal is restored; Ctrl-C still asks nicely first.
        assert signal.getsignal(signal.SIGINT) is handler
    finally:
        for sig, restored in saved.items():
            signal.signal(sig, restored)
