# -*- coding: utf-8 -*-
"""The rain rule.

These are the cases the observatory's whole rain policy comes down to, and the
ones the legacy software got wrong. Written against packets rather than serial
bytes, so they exercise exactly what the Windows 11 service will see.
"""

import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from greenhill import rain_protocol as proto  # noqa: E402
from greenhill.core import rain as R  # noqa: E402
from greenhill.core.config import WeatherConfig  # noqa: E402

NAMES = ('H127', 'H50', 'ACC')


class Bridge:
    """Builds the packet stream a real bridge would send."""

    def __init__(self, names=NAMES):
        self.names = names
        self.sequence = 0

    def packet(self, statuses, port_ok=True, temperatures=None):
        self.sequence += 1
        if temperatures is None:
            temperatures = [12.0] * len(statuses)
        return proto.RainPacket(
            self.sequence, 'ts',
            [proto.Detector(name, status, temperature)
             for name, status, temperature
             in zip(self.names, statuses, temperatures)],
            port_ok)


def config(**kwargs):
    base = dict(raindrop_window_s=10.0, rain_max_age_s=15.0,
                rain_min_observing_detectors=2)
    base.update(kwargs)
    return WeatherConfig(**base)


def run(monitor, bridge, steps, start=0.0, interval=1.0):
    """Feed (repeat_count, statuses) steps at one packet per interval."""
    now = start
    for repeats, statuses in steps:
        for _ in range(repeats):
            monitor.update(now, bridge.packet(statuses))
            now += interval
    return now


class TestThreshold:
    def test_dry_is_no_rain(self):
        monitor, bridge = R.RainMonitor(config()), Bridge()
        now = run(monitor, bridge, [(5, ['D', 'D', 'D'])])
        assert monitor.verdict(now)[0] == R.NO_RAIN

    def test_two_half_wet_sections_are_rain(self):
        monitor, bridge = R.RainMonitor(config()), Bridge()
        now = run(monitor, bridge, [(2, ['w', 'w', 'D'])])
        assert monitor.verdict(now)[0] == R.RAIN

    def test_one_fully_wet_detector_reaches_the_threshold_alone(self):
        # "2 of 6" counts SECTIONS, and a detector reporting W supplies both.
        # This is intended, and it is why the rule is not phrased in detectors.
        monitor, bridge = R.RainMonitor(config()), Bridge()
        now = run(monitor, bridge, [(2, ['W', 'D', 'D'])])
        assert monitor.verdict(now)[0] == R.RAIN

    def test_the_threshold_is_configurable(self):
        monitor = R.RainMonitor(config(rain_wet_sections_trigger=3))
        bridge = Bridge()
        now = run(monitor, bridge, [(2, ['w', 'w', 'D'])])
        assert monitor.verdict(now)[0] == R.NO_RAIN


class TestSingleWetSection:
    """The raindrop-versus-contamination rule."""

    def test_a_drop_that_evaporates_quickly_is_rain(self):
        # A heated sensor dries a real raindrop within seconds. That signature
        # IS the detection -- it is not a false positive to be filtered out.
        monitor, bridge = R.RainMonitor(config()), Bridge()
        now = run(monitor, bridge, [(3, ['D', 'D', 'D']),
                                    (3, ['w', 'D', 'D']),
                                    (1, ['D', 'D', 'D'])])
        verdict, reasons = monitor.verdict(now)
        assert verdict == R.RAIN
        assert any('raindrop' in r for r in reasons)

    def test_a_persistent_single_section_is_not_rain(self):
        # Bird droppings, an insect, or a failed sensor. Closing for this every
        # night is how a safety system ends up switched off.
        monitor, bridge = R.RainMonitor(config()), Bridge()
        now = run(monitor, bridge, [(3, ['D', 'D', 'D']),
                                    (60, ['w', 'D', 'D'])])
        assert monitor.verdict(now)[0] == R.NO_RAIN

    def test_contamination_that_finally_dries_is_still_not_rain(self):
        # THIS IS THE DELIBERATE CHANGE FROM THE LEGACY BEHAVIOUR. The old
        # software closed the dome whenever a single wet section eventually
        # cleared, however long it had been wet, because a rescheduling timer
        # meant the ten-second window was never really enforced.
        monitor, bridge = R.RainMonitor(config()), Bridge()
        now = run(monitor, bridge, [(3, ['D', 'D', 'D']),
                                    (60, ['w', 'D', 'D']),
                                    (2, ['D', 'D', 'D'])])
        assert monitor.verdict(now)[0] == R.NO_RAIN

    def test_the_window_boundary_is_configurable(self):
        monitor = R.RainMonitor(config(raindrop_window_s=30.0))
        bridge = Bridge()
        now = run(monitor, bridge, [(3, ['D', 'D', 'D']),
                                    (20, ['w', 'D', 'D']),
                                    (1, ['D', 'D', 'D'])])
        # Twenty seconds is inside a thirty-second window, so this is a drop.
        assert monitor.verdict(now)[0] == R.RAIN

    def test_a_drop_stays_asserted_between_packets(self):
        # The evidence lasts one packet. If the verdict were not held briefly,
        # a caller evaluating on its own timer could miss it entirely.
        monitor, bridge = R.RainMonitor(config()), Bridge()
        now = run(monitor, bridge, [(2, ['w', 'D', 'D']), (1, ['D', 'D', 'D'])])
        assert monitor.verdict(now)[0] == R.RAIN
        assert monitor.verdict(now + 5.0)[0] == R.RAIN
        assert monitor.verdict(now + 11.0)[0] == R.NO_RAIN

    def test_a_single_section_escalating_to_two_is_rain_at_once(self):
        monitor, bridge = R.RainMonitor(config()), Bridge()
        now = run(monitor, bridge, [(2, ['w', 'D', 'D']), (1, ['w', 'w', 'D'])])
        assert monitor.verdict(now)[0] == R.RAIN


class TestUnavailable:
    """Everything that must NOT be read as a clear night."""

    def test_no_data_is_unavailable(self):
        assert R.RainMonitor(config()).verdict(0.0)[0] == R.UNAVAILABLE

    def test_a_stopped_stream_is_unavailable(self):
        # The bridge has died. In the legacy software this was silent: the
        # errored detectors produced a wet count of zero and nothing closed.
        monitor, bridge = R.RainMonitor(config()), Bridge()
        now = run(monitor, bridge, [(5, ['D', 'D', 'D'])])
        assert monitor.verdict(now)[0] == R.NO_RAIN
        assert monitor.verdict(now + 20.0)[0] == R.UNAVAILABLE

    def test_a_failed_serial_port_is_unavailable(self):
        monitor, bridge = R.RainMonitor(config()), Bridge()
        now = 0.0
        monitor.update(now, bridge.packet(['e', 'e', 'e'], port_ok=False))
        verdict, reasons = monitor.verdict(now)
        assert verdict == R.UNAVAILABLE
        assert any('serial port' in r for r in reasons)

    def test_too_few_working_detectors_is_unavailable(self):
        # One survivor is not asked to speak for the whole site.
        monitor, bridge = R.RainMonitor(config()), Bridge()
        now = 0.0
        monitor.update(now, bridge.packet(['D', 'e', 'E']))
        verdict, reasons = monitor.verdict(now)
        assert verdict == R.UNAVAILABLE
        assert any('detector' in r for r in reasons)

    def test_one_dead_detector_does_not_blind_the_observatory(self):
        monitor, bridge = R.RainMonitor(config()), Bridge()
        now = 0.0
        monitor.update(now, bridge.packet(['D', 'D', 'e']))
        assert monitor.verdict(now)[0] == R.NO_RAIN

    def test_a_parked_detector_is_not_a_dry_one(self):
        # P, M, I, E and e all say nothing about the sky. Counting them as dry
        # is the single defect this design most needed to make impossible.
        monitor, bridge = R.RainMonitor(config()), Bridge()
        now = 0.0
        monitor.update(now, bridge.packet(['P', 'M', 'I']))
        assert monitor.verdict(now)[0] == R.UNAVAILABLE

    def test_rain_still_wins_when_a_detector_is_dead(self):
        monitor, bridge = R.RainMonitor(config()), Bridge()
        now = 0.0
        monitor.update(now, bridge.packet(['W', 'D', 'e']))
        assert monitor.verdict(now)[0] == R.RAIN

    def test_recovering_from_unavailable_returns_a_verdict(self):
        monitor, bridge = R.RainMonitor(config()), Bridge()
        now = 0.0
        monitor.update(now, bridge.packet(['e', 'e', 'e'], port_ok=False))
        assert monitor.verdict(now)[0] == R.UNAVAILABLE
        now += 1.0
        monitor.update(now, bridge.packet(['D', 'D', 'D']))
        assert monitor.verdict(now)[0] == R.NO_RAIN


class TestStreamRobustness:
    def test_a_reordered_datagram_is_dropped(self):
        # UDP reorders. An old "dry" arriving after a newer "wet" would
        # otherwise overwrite the newer reading.
        monitor, bridge = R.RainMonitor(config()), Bridge()
        monitor.update(0.0, bridge.packet(['W', 'W', 'D']))
        stale = proto.RainPacket(1, 'ts',
                                 [proto.Detector(n, 'D', 12.0) for n in NAMES],
                                 True)
        assert monitor.update(1.0, stale) is False
        assert monitor.out_of_order_count == 1
        assert monitor.verdict(1.0)[0] == R.RAIN

    def test_the_bridge_restarting_is_not_mistaken_for_reordering(self):
        monitor, bridge = R.RainMonitor(config()), Bridge()
        bridge.sequence = 5000
        monitor.update(0.0, bridge.packet(['D', 'D', 'D']))
        restarted = proto.RainPacket(0, 'ts',
                                     [proto.Detector(n, 'D', 12.0) for n in NAMES],
                                     True)
        assert monitor.update(1.0, restarted) is True
        assert monitor.restart_count == 1

    def test_temperature_is_the_mean_of_what_reported_one(self):
        monitor, bridge = R.RainMonitor(config()), Bridge()
        monitor.update(0.0, bridge.packet(['D', 'D', 'e'],
                                          temperatures=[10.0, 14.0, None]))
        assert monitor.temperature_c() == pytest.approx(12.0)

    def test_temperature_is_none_when_nobody_reported_one(self):
        monitor, bridge = R.RainMonitor(config()), Bridge()
        monitor.update(0.0, bridge.packet(['e', 'e', 'e'],
                                          temperatures=[None, None, None]))
        assert monitor.temperature_c() is None
