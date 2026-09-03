# -*- coding: utf-8 -*-
"""Wire format for the Greenhill rain-sensor multicast stream.

One datagram per poll cycle, emitted by the Windows 7 sensor bridge and
consumed by the weather service on the Windows 11 box. This module is the
single definition of that format and is imported by BOTH ends, so it must stay
compatible with the oldest interpreter in play -- Python 3.8, 32-bit.

    {"v": 1,
     "seq": 4213,
     "t": "2026-08-17T10:23:45.123Z",
     "dets": [{"id": "H127", "st": "D", "tC": 12.4},
              {"id": "H50",  "st": "w", "tC": 11.9},
              {"id": "ACC",  "st": "W", "tC": 13.1}],
     "port_ok": true,
     "poll_ms": 63}

The bridge publishes RAW OBSERVATIONS AND NOTHING ELSE. No thresholds, no
latching, no decisions -- those live on the Windows 11 side, where they can be
changed and tested without touching the old machine.

Two properties of this format are load-bearing for safety:

* `dets` always carries EVERY CONFIGURED DETECTOR, including ones that failed
  to answer (as `st = "e"`). The receiver therefore knows the expected
  population and can treat a short list as a malformed packet rather than
  quietly evaluating the weather on fewer sensors than the observatory has.

* `port_ok` distinguishes "the bridge is alive but its serial port is broken"
  from silence, which means the bridge itself is gone. Both are unsafe, but
  they are different faults and the operator should not have to guess which.

DO NOT USE `t` AS A STALENESS CLOCK. It is the sender's wall clock, and the
Windows 7 box has no guaranteed time sync; a machine whose clock is an hour out
would otherwise appear permanently stale, or -- far worse -- permanently fresh.
Staleness is measured by the RECEIVER against ARRIVAL TIME. `t` is for logs,
and for spotting a sender whose clock has stopped.
"""

import json
from typing import Any, Dict, List, Optional

PROTOCOL_VERSION = 1

# Adjacent to the anemometer's group, which the hardware already emits on and
# which nothing here is allowed to change.
MULTICAST_GROUP = '239.192.0.5'
MULTICAST_PORT = 60005
WIND_MULTICAST_GROUP = '239.192.0.4'
WIND_MULTICAST_PORT = 60004

# Comfortably inside one Ethernet MTU: three detectors come to roughly 200
# bytes, and a fourth adds about 35. A datagram that had to fragment would turn
# one lost fragment into one lost reading, so the format must never grow past
# this without being reconsidered.
MAX_DATAGRAM_BYTES = 1400


class ProtocolError(ValueError):
    """A datagram that is not a valid rain packet.

    Raised rather than returning a partial result on purpose. A malformed
    packet is not weak evidence about the weather, it is no evidence at all,
    and must not refresh the receiver's staleness clock.
    """


# Status letter -> (human description, wet sections, is a wetness observation).
#
# `wet_sections` is the count this detector contributes: each detector has two
# independently triggering sections, so a half-wet detector ("w") contributes
# one and a fully wet one ("W") contributes two. This is why the operational
# rule is "2 of 6 SECTIONS", not "2 of 3 detectors" -- a single detector
# reporting "W" reaches the threshold on its own.
#
# `is_observation` is the fail-closed hinge. Only D, w and W are statements
# about wetness. A detector that is parked, moving, initialising or in error is
# telling us NOTHING about the sky, and must never be counted as dry.
STATUS = {
    'P': ('parked', 0, False),
    'M': ('moving', 0, False),
    'I': ('initialising', 0, False),
    'D': ('dry', 0, True),
    'w': ('half wet', 1, True),
    'W': ('wet', 2, True),
    'E': ('device error', 0, False),   # reported by the detector itself
    'e': ('comms error', 0, False),    # synthesised by the bridge: no/bad reply
}

VALID_STATUS = frozenset(STATUS)


def wet_sections(status: str) -> int:
    """Sections this status letter reports wet (0, 1 or 2)."""
    try:
        return STATUS[status][1]
    except KeyError:
        raise ProtocolError('unknown status letter {!r}'.format(status))


def is_observation(status: str) -> bool:
    """Is this letter a trustworthy statement about wetness?

    False for P/M/I/E/e. A receiver deciding safety must treat a detector that
    fails this as absent, NOT as dry.
    """
    try:
        return STATUS[status][2]
    except KeyError:
        raise ProtocolError('unknown status letter {!r}'.format(status))


def describe(status: str) -> str:
    return STATUS[status][0] if status in STATUS else 'unknown({!r})'.format(status)


class Detector(object):
    """One detector's reading, as carried on the wire."""

    __slots__ = ('id', 'status', 'temperature_c')

    def __init__(self, id, status, temperature_c=None):
        # type: (str, str, Optional[float]) -> None
        self.id = id
        self.status = status
        self.temperature_c = temperature_c

    @property
    def wet_sections(self):
        # type: () -> int
        return wet_sections(self.status)

    @property
    def is_observation(self):
        # type: () -> bool
        return is_observation(self.status)

    def __repr__(self):
        return 'Detector({!r}, {!r}, {!r})'.format(
            self.id, self.status, self.temperature_c)

    def __eq__(self, other):
        return (isinstance(other, Detector)
                and self.id == other.id
                and self.status == other.status
                and self.temperature_c == other.temperature_c)


class RainPacket(object):
    """A validated rain datagram."""

    __slots__ = ('sequence', 'sent_at', 'detectors', 'port_ok', 'poll_ms')

    def __init__(self, sequence, sent_at, detectors, port_ok, poll_ms=None):
        # type: (int, str, List[Detector], bool, Optional[int]) -> None
        self.sequence = sequence
        self.sent_at = sent_at          # advisory only -- see the module note
        self.detectors = detectors
        self.port_ok = port_ok
        self.poll_ms = poll_ms

    @property
    def total_wet_sections(self):
        # type: () -> int
        """Wet sections across all detectors that are actually observing."""
        return sum(d.wet_sections for d in self.detectors if d.is_observation)

    @property
    def observing_detectors(self):
        # type: () -> List[Detector]
        return [d for d in self.detectors if d.is_observation]

    def __repr__(self):
        return 'RainPacket(seq={}, port_ok={}, dets={!r})'.format(
            self.sequence, self.port_ok, self.detectors)


def build(sequence, sent_at, detectors, port_ok, poll_ms=None):
    # type: (int, str, List[Detector], bool, Optional[int]) -> bytes
    """Serialise one packet. Raises ProtocolError if it would exceed one MTU."""
    payload = {
        'v': PROTOCOL_VERSION,
        'seq': int(sequence),
        't': sent_at,
        'dets': [{'id': d.id, 'st': d.status, 'tC': d.temperature_c}
                 for d in detectors],
        'port_ok': bool(port_ok),
    }
    if poll_ms is not None:
        payload['poll_ms'] = int(poll_ms)

    # separators= drops the whitespace json.dumps adds by default; sort_keys
    # makes captures diffable, which matters when comparing two nights of them.
    data = json.dumps(payload, separators=(',', ':'), sort_keys=True).encode('utf-8')
    if len(data) > MAX_DATAGRAM_BYTES:
        raise ProtocolError(
            'packet is {} bytes, over the {} byte limit'.format(
                len(data), MAX_DATAGRAM_BYTES))
    return data


def _require(condition, message):
    if not condition:
        raise ProtocolError(message)


def parse(data):
    # type: (bytes) -> RainPacket
    """Validate and decode one datagram.

    Strict by design: every failure raises rather than yielding a best-effort
    reading. See ProtocolError.
    """
    _require(isinstance(data, (bytes, bytearray)), 'expected bytes')
    _require(len(data) <= MAX_DATAGRAM_BYTES,
             'datagram is {} bytes, over the limit'.format(len(data)))

    try:
        text = bytes(data).decode('utf-8')
    except UnicodeDecodeError as exc:
        raise ProtocolError('not UTF-8: {}'.format(exc))

    try:
        payload = json.loads(text)
    except ValueError as exc:
        raise ProtocolError('not JSON: {}'.format(exc))

    _require(isinstance(payload, dict), 'top level is not an object')

    version = payload.get('v')
    _require(version == PROTOCOL_VERSION,
             'protocol version {!r}, expected {}'.format(version, PROTOCOL_VERSION))

    sequence = payload.get('seq')
    # bool is a subclass of int; True would otherwise pass as sequence 1.
    _require(isinstance(sequence, int) and not isinstance(sequence, bool)
             and sequence >= 0,
             'seq must be a non-negative integer, got {!r}'.format(sequence))

    sent_at = payload.get('t')
    _require(isinstance(sent_at, str) and sent_at, 't must be a non-empty string')

    port_ok = payload.get('port_ok')
    _require(isinstance(port_ok, bool), 'port_ok must be a boolean')

    poll_ms = payload.get('poll_ms')
    if poll_ms is not None:
        _require(isinstance(poll_ms, int) and not isinstance(poll_ms, bool)
                 and poll_ms >= 0,
                 'poll_ms must be a non-negative integer, got {!r}'.format(poll_ms))

    raw_detectors = payload.get('dets')
    _require(isinstance(raw_detectors, list), 'dets must be a list')
    _require(raw_detectors, 'dets is empty')

    detectors = []          # type: List[Detector]
    seen = set()            # type: set
    for position, entry in enumerate(raw_detectors):
        where = 'dets[{}]'.format(position)
        _require(isinstance(entry, dict), '{} is not an object'.format(where))

        identifier = entry.get('id')
        _require(isinstance(identifier, str) and identifier,
                 '{}.id must be a non-empty string'.format(where))
        _require(identifier not in seen,
                 '{}.id {!r} is duplicated'.format(where, identifier))
        seen.add(identifier)

        status = entry.get('st')
        _require(isinstance(status, str) and status in VALID_STATUS,
                 '{}.st {!r} is not a known status letter'.format(where, status))

        temperature = entry.get('tC')
        if temperature is not None:
            _require(isinstance(temperature, (int, float))
                     and not isinstance(temperature, bool),
                     '{}.tC must be a number or null, got {!r}'.format(
                         where, temperature))
            temperature = float(temperature)

        detectors.append(Detector(identifier, status, temperature))

    return RainPacket(sequence, sent_at, detectors, port_ok, poll_ms)
