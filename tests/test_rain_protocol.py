# -*- coding: utf-8 -*-
"""Wire format tests.

The parser is a safety boundary: everything downstream trusts what comes out of
it, so most of what is tested here is what it REFUSES. A malformed packet must
raise rather than yield a partial reading, because a partial reading would
refresh the receiver's staleness clock and count as evidence about the sky.
"""

import json
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from greenhill import rain_protocol as proto  # noqa: E402
from greenhill.rain_protocol import Detector, ProtocolError  # noqa: E402


def make(detectors=None, sequence=1, sent_at='2026-08-17T10:23:45.123Z',
         port_ok=True, poll_ms=63):
    if detectors is None:
        detectors = [Detector('H127', 'D', 12.4),
                     Detector('H50', 'w', 11.9),
                     Detector('ACC', 'W', 13.1)]
    return proto.build(sequence, sent_at, detectors, port_ok, poll_ms)


def mutate(**changes):
    """A valid packet with individual fields replaced, for the rejection tests."""
    payload = json.loads(make().decode('utf-8'))
    payload.update(changes)
    return json.dumps(payload).encode('utf-8')


class TestRoundTrip:
    def test_survives_a_round_trip(self):
        packet = proto.parse(make())
        assert packet.sequence == 1
        assert packet.sent_at == '2026-08-17T10:23:45.123Z'
        assert packet.port_ok is True
        assert packet.poll_ms == 63
        assert [d.id for d in packet.detectors] == ['H127', 'H50', 'ACC']
        assert [d.status for d in packet.detectors] == ['D', 'w', 'W']
        assert packet.detectors[0].temperature_c == pytest.approx(12.4)

    def test_temperature_may_be_absent(self):
        packet = proto.parse(make([Detector('H127', 'e', None)]))
        assert packet.detectors[0].temperature_c is None

    def test_poll_ms_is_optional(self):
        packet = proto.parse(
            proto.build(1, 'now', [Detector('H127', 'D', 1.0)], True, None))
        assert packet.poll_ms is None

    def test_a_fourth_detector_still_fits_one_datagram(self):
        # The installation is expandable to four units; the format must not be
        # the thing that stops it.
        detectors = [Detector(name, 'W', 12.345)
                     for name in ('H127', 'H50', 'ACC', 'NYI')]
        assert len(make(detectors)) <= proto.MAX_DATAGRAM_BYTES


class TestWetnessSemantics:
    @pytest.mark.parametrize('status,expected', [
        ('D', 0), ('w', 1), ('W', 2),
        ('P', 0), ('M', 0), ('I', 0), ('E', 0), ('e', 0),
    ])
    def test_wet_section_counts(self, status, expected):
        assert proto.wet_sections(status) == expected

    @pytest.mark.parametrize('status', ['D', 'w', 'W'])
    def test_only_dwW_are_observations(self, status):
        assert proto.is_observation(status) is True

    @pytest.mark.parametrize('status', ['P', 'M', 'I', 'E', 'e'])
    def test_everything_else_observes_nothing(self, status):
        # The fail-closed hinge. A detector that is parked, moving,
        # initialising or in error says NOTHING about the sky, and must never
        # be counted as evidence that it is dry.
        assert proto.is_observation(status) is False

    def test_a_single_fully_wet_detector_reaches_the_threshold(self):
        # "2 of 6 sensors" counts SECTIONS. One detector reporting W supplies
        # both on its own, which is the intended behaviour and the reason the
        # rule is not phrased in detectors.
        packet = proto.parse(make([Detector('H127', 'W', 12.0),
                                   Detector('H50', 'D', 12.0),
                                   Detector('ACC', 'D', 12.0)]))
        assert packet.total_wet_sections == 2

    def test_errored_detectors_do_not_dilute_the_count(self):
        packet = proto.parse(make([Detector('H127', 'w', 12.0),
                                   Detector('H50', 'e', None),
                                   Detector('ACC', 'E', None)]))
        assert packet.total_wet_sections == 1
        assert [d.id for d in packet.observing_detectors] == ['H127']

    def test_unknown_letters_raise_rather_than_counting_zero(self):
        with pytest.raises(ProtocolError):
            proto.wet_sections('X')
        with pytest.raises(ProtocolError):
            proto.is_observation('X')


class TestRejection:
    def test_rejects_non_utf8(self):
        with pytest.raises(ProtocolError, match='UTF-8'):
            proto.parse(b'\xff\xfe not text')

    def test_rejects_non_json(self):
        with pytest.raises(ProtocolError, match='JSON'):
            proto.parse(b'hello')

    def test_rejects_a_json_array(self):
        with pytest.raises(ProtocolError, match='not an object'):
            proto.parse(b'[1,2,3]')

    def test_rejects_a_future_protocol_version(self):
        with pytest.raises(ProtocolError, match='version'):
            proto.parse(mutate(v=2))

    def test_rejects_a_missing_version(self):
        with pytest.raises(ProtocolError, match='version'):
            proto.parse(b'{"seq":1}')

    @pytest.mark.parametrize('sequence', [-1, 'x', 1.5, None, True])
    def test_rejects_bad_sequence_numbers(self, sequence):
        # True is in that list on purpose: bool subclasses int, so a naive
        # isinstance check would accept it as sequence 1.
        with pytest.raises(ProtocolError, match='seq'):
            proto.parse(mutate(seq=sequence))

    @pytest.mark.parametrize('port_ok', ['yes', 1, 0, None])
    def test_rejects_non_boolean_port_ok(self, port_ok):
        # 1 and 0 matter: a sender that wrote integers here would otherwise
        # have its "serial port is broken" flag silently read as truthy.
        with pytest.raises(ProtocolError, match='port_ok'):
            proto.parse(mutate(port_ok=port_ok))

    def test_rejects_an_empty_detector_list(self):
        with pytest.raises(ProtocolError, match='empty'):
            proto.parse(mutate(dets=[]))

    def test_rejects_a_missing_detector_list(self):
        with pytest.raises(ProtocolError, match='dets'):
            proto.parse(mutate(dets=None))

    def test_rejects_an_unknown_status_letter(self):
        with pytest.raises(ProtocolError, match='status letter'):
            proto.parse(mutate(dets=[{'id': 'H127', 'st': 'X', 'tC': 1.0}]))

    def test_rejects_a_nameless_detector(self):
        with pytest.raises(ProtocolError, match='id'):
            proto.parse(mutate(dets=[{'id': '', 'st': 'D', 'tC': 1.0}]))

    def test_rejects_duplicate_detector_ids(self):
        # Two entries claiming the same detector means the sender is confused;
        # counting both would double one sensor's vote.
        with pytest.raises(ProtocolError, match='duplicated'):
            proto.parse(mutate(dets=[{'id': 'H127', 'st': 'W', 'tC': 1.0},
                                     {'id': 'H127', 'st': 'W', 'tC': 1.0}]))

    def test_rejects_a_non_numeric_temperature(self):
        with pytest.raises(ProtocolError, match='tC'):
            proto.parse(mutate(dets=[{'id': 'H127', 'st': 'D', 'tC': 'warm'}]))

    def test_rejects_an_oversized_datagram(self):
        with pytest.raises(ProtocolError, match='limit'):
            proto.parse(b'{"v":1}' + b' ' * proto.MAX_DATAGRAM_BYTES)

    def test_refuses_to_build_an_oversized_datagram(self):
        many = [Detector('detector-{:03d}'.format(i), 'D', 12.345)
                for i in range(200)]
        with pytest.raises(ProtocolError, match='over the'):
            proto.build(1, 'now', many, True)
