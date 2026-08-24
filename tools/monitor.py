#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Live safety verdict from the two weather streams.

The Phase 2 core with a console attached: the same fusion the Alpaca devices
will publish, printed as it happens. Useful twice over -- as a commissioning
tool at the observatory, and as the way to check the logic against a real
night before any of it is allowed near the dome.

    python3 tools/monitor.py                        # live, from the network
    python3 tools/monitor.py --replay capture.jsonl  # from a recording

REPLAY IS THE IMPORTANT MODE. Point it at a file from tools/record_streams.py
and the whole safety core runs against real traffic, at speed, with no
hardware and no network -- so a night of weather can be re-examined as often as
needed, and any change to the rules can be checked against what actually
happened rather than against what someone remembers happening.

    python3 tools/monitor.py --replay night.jsonl --verbose

Nothing here commands anything. It watches.
"""

import argparse
import json
import os
import signal
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from greenhill import rain_protocol  # noqa: E402
from greenhill.core.config import WeatherConfig  # noqa: E402
from greenhill.core.receiver import MulticastReader  # noqa: E402
from greenhill.core.safety import SafetyEvaluator  # noqa: E402

_stop = False


def _request_stop(signum, frame):
    global _stop
    _stop = True


def format_state(now, state):
    verdict = 'SAFE  ' if state.is_safe else 'UNSAFE'
    wind = ('{:5.1f}'.format(state.wind_speed_ms)
            if state.wind_speed_ms is not None else '    -')
    gust = ('{:5.1f}'.format(state.wind_gust_ms)
            if state.wind_gust_ms is not None else '    -')
    detectors = ''.join(state.detector_states.get(name, '?')
                        for name in sorted(state.detector_states))
    return '{:7.0f}s  {}  rain[{}] rate={:.1f}  wind {} m/s gust {} @ {:3.0f}deg'.format(
        now, verdict, detectors or '---', state.rain_rate_mm_h, wind, gust,
        state.wind_direction_deg)


def run_replay(path, config, verbose):
    """Feed a capture file through the evaluator using its recorded timings."""
    evaluator = SafetyEvaluator(config)
    counts = {'rain': 0, 'wind': 0, 'rejected': 0, 'ignored': 0}
    now = 0.0
    last_report = -1e9
    transitions = []

    with open(path, encoding='utf-8') as handle:
        for line in handle:
            line = line.strip()
            if not line:
                continue
            try:
                record = json.loads(line)
            except ValueError:
                continue

            now = record.get('t_mono', now)
            stream = record.get('stream')

            if stream == 'rain' and 'text' in record:
                try:
                    packet = rain_protocol.parse(record['text'].encode('utf-8'))
                except rain_protocol.ProtocolError:
                    counts['rejected'] += 1
                else:
                    evaluator.rain.update(now, packet)
                    counts['rain'] += 1
            elif stream == 'wind' and 'text' in record:
                before = evaluator.wind.ignored_count
                if evaluator.wind.update(now, record['text']):
                    counts['wind'] += 1
                elif evaluator.wind.ignored_count > before:
                    # Not a wind sentence at all -- the instrument's periodic
                    # $WIVER identification. Normal traffic, not a fault.
                    counts['ignored'] += 1
                else:
                    counts['rejected'] += 1

            state = evaluator.update(now)
            if evaluator.changed(state):
                transitions.append((now, state))
                print(format_state(now, state))
                for reason in state.reasons:
                    print('           {}'.format(reason))
            elif verbose and now - last_report >= 30.0:
                last_report = now
                print(format_state(now, state))

    print('')
    print('replayed {:.0f}s: {} rain packets, {} wind sentences, '
          '{} non-wind datagrams ignored, {} rejected'.format(
              now, counts['rain'], counts['wind'], counts['ignored'],
              counts['rejected']))
    print('{} safety transitions'.format(len(transitions)))
    if not transitions:
        print('NOTE: the verdict never changed. On a capture that starts from '
              'nothing this means it never became safe -- check both streams '
              'were actually being received.')
    return 0


def run_live(args, config):
    evaluator = SafetyEvaluator(config)
    reader = MulticastReader()
    reader.add('wind', args.wind_group, args.wind_port, args.interface)
    reader.add('rain', args.rain_group, args.rain_port, args.interface)

    for name, error in reader.failures.items():
        print('could not join {}: {}'.format(name, error), file=sys.stderr)
    if not reader.stream_count:
        print('no streams joined; nothing to monitor.', file=sys.stderr)
        return 1

    print(config.describe_wind_thresholds(), file=sys.stderr)
    print('watching. the verdict starts UNSAFE and clears once both streams '
          'are healthy.', file=sys.stderr)

    started = time.monotonic()
    last_report = -1e9
    try:
        while not _stop:
            if args.duration is not None and time.monotonic() - started >= args.duration:
                break
            for stream, payload in reader.poll(0.5):
                now = time.monotonic() - started
                if stream == 'rain':
                    try:
                        evaluator.rain.update(now, rain_protocol.parse(payload))
                    except rain_protocol.ProtocolError as exc:
                        if args.verbose:
                            print('bad rain packet: {}'.format(exc), file=sys.stderr)
                else:
                    try:
                        evaluator.wind.update(now, payload.decode('ascii'))
                    except UnicodeDecodeError:
                        pass

            # Evaluated on a timer as well as on arrival: a verdict that only
            # changes when data arrives can never notice data NOT arriving,
            # which is the failure this whole design exists to catch.
            now = time.monotonic() - started
            state = evaluator.update(now)
            if evaluator.changed(state) or now - last_report >= 10.0:
                last_report = now
                print(format_state(now, state))
                if not state.is_safe:
                    for reason in state.reasons:
                        print('           {}'.format(reason))
    finally:
        reader.close()
    return 0


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.split('\n')[0])
    parser.add_argument('--replay', help='a capture from tools/record_streams.py')
    parser.add_argument('--interface', default='0.0.0.0')
    parser.add_argument('--wind-group', default=rain_protocol.WIND_MULTICAST_GROUP)
    parser.add_argument('--wind-port', type=int,
                        default=rain_protocol.WIND_MULTICAST_PORT)
    parser.add_argument('--rain-group', default=rain_protocol.MULTICAST_GROUP)
    parser.add_argument('--rain-port', type=int,
                        default=rain_protocol.MULTICAST_PORT)
    parser.add_argument('--duration', type=float, default=None,
                        help='stop after this many seconds (live mode)')
    parser.add_argument('--verbose', action='store_true')
    args = parser.parse_args(argv)

    signal.signal(signal.SIGINT, _request_stop)
    config = WeatherConfig()

    if args.replay:
        return run_replay(args.replay, config, args.verbose)
    return run_live(args, config)


if __name__ == '__main__':
    sys.exit(main())
