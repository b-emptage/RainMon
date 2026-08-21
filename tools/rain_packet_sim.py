#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Emit synthetic rain packets, so the receiving side can be built without the
Windows 7 box or the detectors.

    python3 tools/rain_packet_sim.py --scenario raindrop
    python3 tools/rain_packet_sim.py --scenario rain --duration 120

Scenarios exist to exercise the cases the safety logic has to tell apart:

    dry        all detectors report D. The nothing-is-happening baseline.
    raindrop   one section goes wet and dries again within a few seconds.
               THIS MUST BE TREATED AS RAIN -- a drop evaporating off a heated
               sensor is the signature of real weather.
    fouling    one section goes wet and stays wet indefinitely. THIS MUST NOT
               BE TREATED AS RAIN -- it is bird droppings, an insect, or a
               failed sensor, and closing for it every night is how a safety
               system gets switched off.
    rain       two sections wet. The unambiguous close case.
    downpour   everything wet, then a long slow dry, so the latch and the
               2-5 minute drying time can be exercised.
    port-down  the bridge is alive but its serial port is not: every detector
               reports "e" and port_ok is false.
    silence    emits briefly and then stops without exiting, so the receiver's
               staleness watchdog can be tested. This is the case the original
               software could not distinguish from a clear night.

Standard library only. Runs anywhere.
"""

import argparse
import datetime
import os
import socket
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from greenhill import rain_protocol  # noqa: E402

DEFAULT_DETECTORS = ('H127', 'H50', 'ACC')


def timestamp():
    now = datetime.datetime.now(datetime.timezone.utc)
    return '{}.{:03d}Z'.format(now.strftime('%Y-%m-%dT%H:%M:%S'),
                               now.microsecond // 1000)


def statuses_for(scenario, elapsed, count):
    """(list of status letters, port_ok, still_emitting) at `elapsed` seconds."""
    dry = ['D'] * count

    if scenario == 'dry':
        return dry, True, True

    if scenario == 'raindrop':
        # Wet at t=5, dry again at t=8: inside the 10 s window that marks it as
        # a real drop rather than a persistent contaminant.
        if 5.0 <= elapsed < 8.0:
            return ['w'] + dry[1:], True, True
        return dry, True, True

    if scenario == 'fouling':
        if elapsed >= 5.0:
            return ['w'] + dry[1:], True, True
        return dry, True, True

    if scenario == 'rain':
        if elapsed >= 5.0:
            return ['w', 'w'] + dry[2:], True, True
        return dry, True, True

    if scenario == 'downpour':
        if elapsed < 5.0:
            return dry, True, True
        if elapsed < 65.0:
            return ['W'] * count, True, True
        if elapsed < 185.0:             # a couple of minutes drying out
            return ['w'] * count, True, True
        return dry, True, True

    if scenario == 'port-down':
        if elapsed >= 5.0:
            return ['e'] * count, False, True
        return dry, True, True

    if scenario == 'silence':
        return dry, True, elapsed < 10.0

    raise ValueError('unknown scenario {!r}'.format(scenario))


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.split('\n')[0])
    parser.add_argument('--scenario', default='dry',
                        choices=['dry', 'raindrop', 'fouling', 'rain',
                                 'downpour', 'port-down', 'silence'])
    parser.add_argument('--group', default=rain_protocol.MULTICAST_GROUP)
    parser.add_argument('--port', type=int, default=rain_protocol.MULTICAST_PORT)
    parser.add_argument('--interval', type=float, default=1.0)
    parser.add_argument('--duration', type=float, default=None,
                        help='stop after this many seconds')
    parser.add_argument('--detectors', default=','.join(DEFAULT_DETECTORS))
    parser.add_argument('--ttl', type=int, default=1)
    parser.add_argument('--interface', default='',
                        help='local interface IP to send from. Use 127.0.0.1 to '
                             'test against a listener on this same machine when '
                             'there is no usable network interface.')
    args = parser.parse_args(argv)

    names = [n.strip() for n in args.detectors.split(',') if n.strip()]
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM, socket.IPPROTO_UDP)
    sock.setsockopt(socket.IPPROTO_IP, socket.IP_MULTICAST_TTL, args.ttl)
    if args.interface:
        sock.setsockopt(socket.IPPROTO_IP, socket.IP_MULTICAST_IF,
                        socket.inet_aton(args.interface))

    print('scenario {!r} -> {}:{} every {:.1f}s, detectors {}'.format(
        args.scenario, args.group, args.port, args.interval, ', '.join(names)),
        file=sys.stderr)

    sequence = 0
    started = time.monotonic()
    announced_silence = False
    try:
        while True:
            elapsed = time.monotonic() - started
            if args.duration is not None and elapsed >= args.duration:
                break

            letters, port_ok, emitting = statuses_for(args.scenario, elapsed, len(names))
            if emitting:
                detectors = [
                    rain_protocol.Detector(
                        name, letter,
                        None if letter in ('e', 'E') else round(12.0 + 0.1 * index, 1))
                    for index, (name, letter) in enumerate(zip(names, letters))]
                sock.sendto(
                    rain_protocol.build(sequence, timestamp(), detectors,
                                        port_ok, poll_ms=60),
                    (args.group, args.port))
                print('\rseq={} {}{}'.format(
                    sequence, ' '.join(letters), '' if port_ok else '  PORT DOWN'),
                    end='', file=sys.stderr)
                sequence += 1
            elif not announced_silence:
                print('\n-- gone silent; the receiver should now go unsafe on '
                      'staleness --', file=sys.stderr)
                announced_silence = True

            time.sleep(args.interval)
    except KeyboardInterrupt:
        pass
    finally:
        sock.close()
        print('', file=sys.stderr)
    return 0


if __name__ == '__main__':
    sys.exit(main())
