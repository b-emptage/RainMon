# -*- coding: utf-8 -*-
"""The Greenhill rain-sensor bridge.

Runs on the Windows 7 box. Owns COM7, polls the rain detectors once a second,
and multicasts what they said. That is the whole job.

It decides nothing. There are no thresholds here, no latching, no dome
commands and no network listener -- every rule that might ever want changing
lives on the Windows 11 weather service, so this machine never has to be
touched again. RESET (*RnI) is deliberately absent too: a detector that has
locked up is worth walking up to.

The one behaviour worth stating outright: THIS PROCESS KEEPS EMITTING WHEN THE
SERIAL PORT IS BROKEN. A packet with `port_ok: false` says "the bridge is alive
and its sensors are not"; silence says "the bridge is gone". Both are unsafe
and the receiver treats them as such, but they send an engineer to different
places, so the difference is worth a datagram a second.

Python 3.8 compatible. Dependencies: pyserial, and nothing else.
"""

import argparse
import configparser
import datetime
import logging
import logging.handlers
import os
import signal
import socket
import sys
import time
from typing import Dict, List, Optional, Tuple

from greenhill import rain_protocol
from greenhill.rain_serial import PortFailure, RainSensorPort

CONFIG_FILENAME = 'rain_broadcaster.ini'
CONFIG_SECTION = 'rain'
LOG_FILENAME = 'rain_broadcaster.log'

# Loopback port held as a process mutex. 50815 is the dome server's; this is
# the next one along. Chosen over a lock file because the OS releases it even
# when we are killed outright, where a stale file would block every later start.
SINGLE_INSTANCE_PORT = 50816

# How long to wait between attempts to reopen a serial port that has failed.
# Backs off so a genuinely absent port does not fill the log at 1 Hz, but stays
# short enough that a cable pushed back in is picked up while the engineer is
# still standing there.
REOPEN_BACKOFF_SECONDS = (1, 2, 5, 10, 30)

# Emit a liveness line even when nothing changes, so a quiet log means a dead
# process rather than a calm night.
HEARTBEAT_SECONDS = 300

logger = logging.getLogger('rain_broadcaster')

_shutdown_requested = False
_instance_guard = None


DEFAULTS = {
    'com': 'COM7',
    'detectors': '0:H127, 1:H50, 2:ACC',
    'baud': '9600',
    'read_timeout': '0.9',
    'write_timeout': '0.4',
    'interval': '1.0',
    'group': rain_protocol.MULTICAST_GROUP,
    'port': str(rain_protocol.MULTICAST_PORT),
    'ttl': '1',
    'interface': '',
    'log_level': 'INFO',
    'log_max_mb': '5',
    'log_keep': '10',
}


def config_directory():
    # type: () -> str
    """Where the .ini and the log live.

    Beside the executable when frozen, NOT inside the PyInstaller bundle. A
    path resolved relative to this module lands in _MEIPASS in a frozen build:
    a read-only temporary directory that is deleted on exit, so the operator
    could neither edit the settings nor find the log. The legacy RainMon.ini is
    read from the executable's directory for exactly this reason.
    """
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def load_config(path=None):
    # type: (Optional[str]) -> Dict[str, str]
    """Read the .ini, falling back to defaults for anything absent.

    A missing file is not an error: the defaults describe the installation as
    it stands today, and a bridge that refuses to start because nobody wrote a
    config file is a bridge that is not watching the sky.
    """
    parser = configparser.ConfigParser()
    parser[CONFIG_SECTION] = dict(DEFAULTS)

    if path is None:
        path = os.path.join(config_directory(), CONFIG_FILENAME)
    if os.path.exists(path):
        try:
            parser.read(path)
        except configparser.Error as exc:
            print('==CONFIG== {} is malformed ({}); using defaults.'.format(path, exc),
                  file=sys.stderr)
    else:
        print('==CONFIG== no {} found; using defaults.'.format(path), file=sys.stderr)

    settings = dict(DEFAULTS)
    if parser.has_section(CONFIG_SECTION):
        settings.update({key: parser.get(CONFIG_SECTION, key)
                         for key in parser.options(CONFIG_SECTION)})
    settings['_path'] = path
    return settings


def parse_detectors(spec):
    # type: (str) -> List[Tuple[int, str]]
    """"0:H127, 1:H50, 2:ACC" -> [(0, 'H127'), (1, 'H50'), (2, 'ACC')].

    The identifiers travel on the wire and are what the operators call the
    detectors, so they are configuration rather than constants -- a fourth unit
    is a line in the .ini, not a rebuild.
    """
    detectors = []      # type: List[Tuple[int, str]]
    seen_index = set()
    seen_name = set()
    for chunk in spec.split(','):
        chunk = chunk.strip()
        if not chunk:
            continue
        if ':' not in chunk:
            raise ValueError('detector {!r} is not index:name'.format(chunk))
        raw_index, name = chunk.split(':', 1)
        try:
            index = int(raw_index.strip())
        except ValueError:
            raise ValueError('detector index {!r} is not a number'.format(raw_index))
        name = name.strip()
        if not name:
            raise ValueError('detector {} has no name'.format(index))
        if not 0 <= index <= 3:
            raise ValueError('detector index {} is outside 0-3'.format(index))
        if index in seen_index:
            raise ValueError('detector index {} appears twice'.format(index))
        if name in seen_name:
            raise ValueError('detector name {!r} appears twice'.format(name))
        seen_index.add(index)
        seen_name.add(name)
        detectors.append((index, name))
    if not detectors:
        raise ValueError('no detectors configured')
    return detectors


def init_logging(settings):
    # type: (Dict[str, str]) -> None
    """Rotating log beside the executable.

    Nothing about the log file may stop the bridge from running: a directory
    that cannot be written falls back to the console and says so. The dome
    server learned this the hard way -- a start that followed a crash used to
    die on a log rotation it could not perform, leaving the observatory with no
    server at all over a file nobody was reading.
    """
    level = getattr(logging, settings['log_level'].upper(), logging.INFO)
    logger.setLevel(level)
    formatter = logging.Formatter('%(asctime)s %(levelname)s %(message)s',
                                  datefmt='%Y-%m-%d %H:%M:%S')

    console = logging.StreamHandler()
    console.setFormatter(formatter)
    logger.addHandler(console)

    path = os.path.join(config_directory(), LOG_FILENAME)
    try:
        handler = logging.handlers.RotatingFileHandler(
            path,
            maxBytes=int(float(settings['log_max_mb']) * 1024 * 1024),
            backupCount=int(settings['log_keep']),
            delay=True)
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    except Exception as exc:
        print('==LOGGING== cannot write {} ({}); console only.'.format(path, exc),
              file=sys.stderr)


def acquire_single_instance_lock():
    # type: () -> bool
    global _instance_guard
    guard = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        guard.bind(('127.0.0.1', SINGLE_INSTANCE_PORT))
        guard.listen(1)
    except OSError:
        guard.close()
        return False
    _instance_guard = guard
    return True


def install_signal_handlers():
    # type: () -> None
    """Ask the loop to stop. Nothing here latches in hardware -- unlike the
    dome, there are no relays to de-energise -- so a clean exit only has to
    close the serial port."""
    def _request_shutdown(signum, frame):
        global _shutdown_requested
        _shutdown_requested = True

    for name in ('SIGINT', 'SIGTERM', 'SIGBREAK'):
        sig = getattr(signal, name, None)
        if sig is None:
            continue
        try:
            signal.signal(sig, _request_shutdown)
        except (ValueError, OSError):
            continue


def create_sender(ttl, interface):
    # type: (int, str) -> socket.socket
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM, socket.IPPROTO_UDP)
    # TTL 1 keeps this on the local segment, which is where both listeners live.
    # Raise it only if the stream is ever deliberately routed between VLANs.
    sock.setsockopt(socket.IPPROTO_IP, socket.IP_MULTICAST_TTL, ttl)
    if interface:
        # Without this, Windows picks the outgoing interface from the routing
        # table -- which on a multi-homed box is a coin toss, and a silent one:
        # the send succeeds and nobody receives it.
        sock.setsockopt(socket.IPPROTO_IP, socket.IP_MULTICAST_IF,
                        socket.inet_aton(interface))
    return sock


def utc_timestamp():
    # type: () -> str
    now = datetime.datetime.now(datetime.timezone.utc)
    return '{}.{:03d}Z'.format(now.strftime('%Y-%m-%dT%H:%M:%S'),
                               now.microsecond // 1000)


def _format_statuses(readings):
    return ' '.join('{}={}'.format(r.identifier, r.status) for r in readings)


def run(settings):
    # type: (Dict[str, str]) -> int
    detectors = parse_detectors(settings['detectors'])
    interval = float(settings['interval'])
    group = settings['group']
    port_number = int(settings['port'])

    sender = create_sender(int(settings['ttl']), settings['interface'].strip())
    sensors = RainSensorPort(
        settings['com'], detectors,
        baud=int(settings['baud']),
        read_timeout=float(settings['read_timeout']),
        write_timeout=float(settings['write_timeout']))

    logger.info('==STARTUP== rain bridge on %s -> %s:%d every %.1fs, detectors %s',
                settings['com'], group, port_number, interval,
                ', '.join('{}:{}'.format(i, n) for i, n in detectors))
    logger.info('config: %s', settings['_path'])

    sequence = 0
    reopen_attempt = 0
    next_reopen = 0.0
    last_statuses = None            # type: Optional[Dict[str, str]]
    last_heartbeat = time.monotonic()
    send_failures = 0
    deadline = time.monotonic()

    while not _shutdown_requested:
        cycle_start = time.monotonic()

        if not sensors.is_open and cycle_start >= next_reopen:
            try:
                sensors.open()
                logger.info('serial port %s open', settings['com'])
                reopen_attempt = 0
            except PortFailure as exc:
                backoff = REOPEN_BACKOFF_SECONDS[
                    min(reopen_attempt, len(REOPEN_BACKOFF_SECONDS) - 1)]
                # Only shout the first time; after that it is a known state and
                # the heartbeat carries the news.
                if reopen_attempt == 0:
                    logger.error('%s', exc)
                else:
                    logger.debug('reopen failed: %s', exc)
                reopen_attempt += 1
                next_reopen = cycle_start + backoff

        readings, port_ok = sensors.poll()
        poll_ms = int((time.monotonic() - cycle_start) * 1000)

        packet_detectors = [
            rain_protocol.Detector(r.identifier, r.status, r.temperature_c)
            for r in readings]

        datagram = None
        try:
            datagram = rain_protocol.build(
                sequence, utc_timestamp(), packet_detectors, port_ok, poll_ms)
        except rain_protocol.ProtocolError as exc:
            # Only reachable by adding detectors until the packet outgrows an
            # MTU. Log loudly and keep going: the next cycle will fail the same
            # way, but a bridge that exits helps nobody.
            logger.error('cannot build packet: %s', exc)

        if datagram is not None:
            try:
                sender.sendto(datagram, (group, port_number))
                if send_failures:
                    logger.info('multicast send recovered after %d failures',
                                send_failures)
                    send_failures = 0
            except OSError as exc:
                send_failures += 1
                if send_failures == 1 or send_failures % 60 == 0:
                    logger.error('multicast send failed (%d in a row): %s',
                                 send_failures, exc)

        statuses = {r.identifier: r.status for r in readings}
        if statuses != last_statuses:
            logger.info('%s%s', _format_statuses(readings),
                        '' if port_ok else '  [SERIAL PORT DOWN]')
            for reading in readings:
                if reading.reason:
                    logger.debug('  %s: %s', reading.identifier, reading.reason)
            last_statuses = statuses

        if time.monotonic() - last_heartbeat >= HEARTBEAT_SECONDS:
            logger.info('alive: seq=%d %s poll=%dms%s', sequence,
                        _format_statuses(readings), poll_ms,
                        '' if port_ok else '  [SERIAL PORT DOWN]')
            last_heartbeat = time.monotonic()

        sequence += 1

        deadline += interval
        remaining = deadline - time.monotonic()
        if remaining > 0:
            time.sleep(remaining)
        else:
            # The cycle overran, almost always because dead detectors are being
            # waited out. Resync rather than trying to catch up, which would
            # burst packets and never recover the cadence.
            deadline = time.monotonic()

    logger.info('==SHUTDOWN== stopping')
    sensors.close()
    sender.close()
    return 0


def main(argv=None):
    # type: (Optional[List[str]]) -> int
    parser = argparse.ArgumentParser(
        description='Multicast the Greenhill rain detectors to the local network.')
    parser.add_argument('--config', help='path to rain_broadcaster.ini')
    parser.add_argument('--verbose', action='store_true', help='debug logging')
    args = parser.parse_args(argv)

    if not acquire_single_instance_lock():
        print('==STARTUP FAILED== another rain bridge is already running. Two '
              'processes cannot share the serial port; refusing to start.',
              file=sys.stderr)
        return 1

    settings = load_config(args.config)
    if args.verbose:
        settings['log_level'] = 'DEBUG'
    init_logging(settings)
    install_signal_handlers()

    try:
        return run(settings)
    except ValueError as exc:
        logger.error('==STARTUP FAILED== bad configuration: %s', exc)
        return 2
    except Exception as exc:                        # noqa: BLE001
        logger.exception('==CRASH== %s', exc)
        return 3
