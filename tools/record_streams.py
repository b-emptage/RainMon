#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Capture the Greenhill weather multicast streams to a file.

Run this FIRST, before writing a line of the fusion logic. It records exactly
what the anemometer and the rain bridge put on the wire, so Phase 2 can be
built and tested against real traffic instead of against an assumption.

    python3 tools/record_streams.py --out capture.jsonl
    python3 tools/record_streams.py --out capture.jsonl --duration 3600

Ctrl-C stops it and prints a summary. One JSON object per datagram:

    {"stream": "wind", "t_wall": "...Z", "t_mono": 12.345,
     "src": "10.0.0.7:51322", "len": 47, "text": "$WIMWV,..."}

Datagrams that are not decodable text are stored as base64 under "b64" instead
of "text", so a capture never loses a packet just because it was garbled --
which is precisely the packet worth having.

There is a specific question this is meant to settle. The existing wind display
reads direction from `parts[2]` and speed from `parts[4]` of a comma-separated
sentence, by fixed index, behind a bare `except:` that silently yields no
reading. Nobody has written down what the anemometer actually emits. The exit
summary prints every field of a sample sentence with its index, so the parser
in Phase 2 can be written against the real format.

Runs anywhere -- no hardware, no Windows, no dependencies beyond the standard
library.
"""

import argparse
import base64
import datetime
import json
import os
import select
import signal
import socket
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from greenhill import rain_protocol  # noqa: E402
from greenhill.core.receiver import join_multicast_group  # noqa: E402

_stop = False


def _request_stop(signum, frame):
    global _stop
    _stop = True


class StreamStats(object):
    def __init__(self, name):
        self.name = name
        self.count = 0
        self.bytes = 0
        self.first_mono = None
        self.last_mono = None
        self.max_gap = 0.0
        self.sample = None
        self.undecodable = 0
        # Rain only: gaps in the sequence number reveal loss that arrival
        # timing alone cannot distinguish from a slow sender.
        self.sequence_gaps = 0
        self.last_sequence = None
        self.protocol_errors = 0

    def record(self, mono, payload, text):
        if self.first_mono is None:
            self.first_mono = mono
        elif mono - self.last_mono > self.max_gap:
            self.max_gap = mono - self.last_mono
        self.last_mono = mono
        self.count += 1
        self.bytes += len(payload)
        if text is None:
            self.undecodable += 1
        else:
            self.sample = text

    @property
    def duration(self):
        if self.first_mono is None or self.last_mono is None:
            return 0.0
        return self.last_mono - self.first_mono

    @property
    def rate(self):
        return self.count / self.duration if self.duration > 0 else 0.0


def summarise(stats, out=sys.stderr):
    print('', file=out)
    print('=' * 68, file=out)
    for stream in stats.values():
        print('{}: {} packets, {} bytes, {:.2f}/s over {:.0f}s'.format(
            stream.name, stream.count, stream.bytes, stream.rate,
            stream.duration), file=out)
        if not stream.count:
            print('    NOTHING RECEIVED. Check the group address, the '
                  'interface, and the firewall.', file=out)
            continue
        print('    largest gap between packets: {:.2f}s'.format(stream.max_gap),
              file=out)
        if stream.undecodable:
            print('    {} datagrams were not decodable text'.format(
                stream.undecodable), file=out)
        if stream.name == 'rain':
            print('    sequence gaps: {}   protocol errors: {}'.format(
                stream.sequence_gaps, stream.protocol_errors), file=out)
        if stream.sample:
            print('    sample: {!r}'.format(stream.sample), file=out)
        if stream.name == 'wind' and stream.sample:
            _dissect_wind(stream.sample, out)
    print('=' * 68, file=out)


def _dissect_wind(sample, out):
    """Print the anemometer sentence field by field.

    The existing display reads direction from index 2 and speed from index 4.
    This is how we find out whether that is right, and what the units are.
    """
    fields = sample.strip().split(',')
    print('    fields ({} total):'.format(len(fields)), file=out)
    for index, field in enumerate(fields):
        marker = ''
        if index == 2:
            marker = '   <- read as DIRECTION by wind_sensor.py'
        elif index == 4:
            marker = '   <- read as SPEED by wind_sensor.py'
        print('      [{}] {!r}{}'.format(index, field, marker), file=out)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.split('\n')[0])
    parser.add_argument('--out', default='capture.jsonl',
                        help='output file (JSON lines, appended)')
    parser.add_argument('--duration', type=float, default=None,
                        help='stop after this many seconds')
    parser.add_argument('--interface', default='0.0.0.0',
                        help='local interface IP to receive on')
    parser.add_argument('--wind-group', default=rain_protocol.WIND_MULTICAST_GROUP)
    parser.add_argument('--wind-port', type=int, default=rain_protocol.WIND_MULTICAST_PORT)
    parser.add_argument('--rain-group', default=rain_protocol.MULTICAST_GROUP)
    parser.add_argument('--rain-port', type=int, default=rain_protocol.MULTICAST_PORT)
    parser.add_argument('--quiet', action='store_true',
                        help='no live progress line')
    args = parser.parse_args(argv)

    signal.signal(signal.SIGINT, _request_stop)
    if hasattr(signal, 'SIGTERM'):
        signal.signal(signal.SIGTERM, _request_stop)

    sockets = {}
    stats = {}
    for name, group, port in (('wind', args.wind_group, args.wind_port),
                              ('rain', args.rain_group, args.rain_port)):
        try:
            sockets[join_multicast_group(group, port, args.interface)] = name
            stats[name] = StreamStats(name)
            print('listening: {} on {}:{}'.format(name, group, port),
                  file=sys.stderr)
        except OSError as exc:
            print('could not join {} on {}:{}: {}'.format(name, group, port, exc),
                  file=sys.stderr)

    if not sockets:
        print('no streams could be joined; nothing to record.', file=sys.stderr)
        return 1

    started = time.monotonic()
    last_progress = started
    handle = open(args.out, 'a', encoding='utf-8')
    try:
        while not _stop:
            if args.duration is not None and time.monotonic() - started >= args.duration:
                break

            readable, _, _ = select.select(list(sockets), [], [], 0.5)
            for sock in readable:
                name = sockets[sock]
                try:
                    payload, source = sock.recvfrom(65535)
                except OSError:
                    continue

                mono = time.monotonic() - started
                record = {
                    'stream': name,
                    't_wall': datetime.datetime.now(
                        datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%S.%fZ'),
                    't_mono': round(mono, 4),
                    'src': '{}:{}'.format(source[0], source[1]),
                    'len': len(payload),
                }
                try:
                    text = payload.decode('ascii').strip()
                    record['text'] = text
                except UnicodeDecodeError:
                    text = None
                    record['b64'] = base64.b64encode(payload).decode('ascii')

                stream = stats[name]
                stream.record(mono, payload, text)

                # Validate rain packets as they arrive. A capture that turns out
                # to be full of rejects is worth knowing about while the sender
                # is still running, not a week later during Phase 2.
                if name == 'rain':
                    try:
                        packet = rain_protocol.parse(payload)
                    except rain_protocol.ProtocolError as exc:
                        stream.protocol_errors += 1
                        record['protocol_error'] = str(exc)
                    else:
                        if (stream.last_sequence is not None
                                and packet.sequence != stream.last_sequence + 1):
                            stream.sequence_gaps += 1
                        stream.last_sequence = packet.sequence

                handle.write(json.dumps(record, sort_keys=True) + '\n')

            now = time.monotonic()
            if not args.quiet and now - last_progress >= 1.0:
                last_progress = now
                print('\r{:.0f}s  {}'.format(
                    now - started,
                    '  '.join('{}={}'.format(s.name, s.count)
                              for s in stats.values())),
                    end='', file=sys.stderr)
                sys.stderr.flush()
    finally:
        handle.close()
        for sock in sockets:
            sock.close()

    summarise(stats)
    print('written to {}'.format(args.out), file=sys.stderr)
    return 0


if __name__ == '__main__':
    sys.exit(main())
