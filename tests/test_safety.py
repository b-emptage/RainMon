# -*- coding: utf-8 -*-
"""Fusion, latching and the fail-closed behaviour.

This is the module that decides whether the dome closes, so most of what is
tested here is that it says UNSAFE when it should -- including all the ways the
software this replaces said nothing at all.
"""

import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from greenhill import rain_protocol as proto  # noqa: E402
from greenhill.core.config import WeatherConfig  # noqa: E402
from greenhill.core.safety import LatchedCondition, SafetyEvaluator  # noqa: E402
from conftest import mwv_datagram  # noqa: E402

NAMES = ('H127', 'H50', 'ACC')


def config(**kwargs):
    base = dict(settle_s=30.0, rain_clear_s=600.0, wind_clear_s=120.0,
                wind_min_samples=5, rain_max_age_s=15.0, wind_max_age_s=15.0)
    base.update(kwargs)
    return WeatherConfig(**base)


class Site:
    """Drives an evaluator with both streams, like the real service."""

    def __init__(self, cfg=None):
        self.config = cfg or config()
        self.evaluator = SafetyEvaluator(self.config)
        self.now = 0.0
        self.sequence = 0

    def tick(self, seconds=1.0, rain=('D', 'D', 'D'), wind_speed=2.0,
             rain_ok=True, send_rain=True, send_wind=True):
        """Advance one second, delivering a packet from each stream."""
        steps = int(seconds)
        for _ in range(steps):
            if send_rain:
                self.sequence += 1
                self.evaluator.rain.update(self.now, proto.RainPacket(
                    self.sequence, 'ts',
                    [proto.Detector(n, s, 12.0) for n, s in zip(NAMES, rain)],
                    rain_ok))
            if send_wind:
                self.evaluator.wind.update(
                    self.now, mwv_datagram(angle=90.0, speed=wind_speed))
            state = self.evaluator.update(self.now)
            self.now += 1.0
        return state

    def settle(self):
        """Run long enough for a clean start to become safe.

        The settle timer cannot even begin until wind has enough samples to
        offer a verdict, so this is the settle period PLUS that warm-up.
        """
        return self.tick(self.config.settle_s + self.config.wind_min_samples + 5)


class TestLatchedCondition:
    def test_starts_tripped(self):
        # A service that came up believing all was well would be safe for the
        # first seconds of every restart, including a restart in the rain.
        assert LatchedCondition('x', 600.0, 30.0).active is True

    def test_the_first_clear_only_waits_the_settle_period(self):
        # The startup trip is the absence of good news, not the presence of bad
        # news -- so it must not cost the full rain delay on a clear night.
        condition = LatchedCondition('x', 600.0, 30.0)
        condition.update(0.0, False)
        condition.update(29.0, False)
        assert condition.active is True
        condition.update(31.0, False)
        assert condition.active is False

    def test_a_real_trip_waits_the_full_delay(self):
        condition = LatchedCondition('x', 600.0, 30.0)
        condition.update(0.0, False)
        condition.update(31.0, False)
        condition.update(40.0, True, ['because'])
        condition.update(41.0, False)
        condition.update(100.0, False)
        assert condition.active is True          # 30s would have been enough before
        condition.update(645.0, False)
        assert condition.active is False

    def test_the_countdown_restarts_if_it_trips_again(self):
        condition = LatchedCondition('x', 100.0, 10.0)
        condition.update(0.0, False)
        condition.update(11.0, False)
        condition.update(20.0, True)
        condition.update(30.0, False)
        condition.update(80.0, True)             # trips again mid-countdown
        condition.update(150.0, False)           # countdown restarts here
        assert condition.active is True
        condition.update(190.0, False)
        assert condition.active is True          # only 40s of the 100s served
        condition.update(251.0, False)
        assert condition.active is False

    def test_reports_how_long_is_left(self):
        condition = LatchedCondition('x', 100.0, 100.0)
        condition.update(0.0, False)
        condition.update(40.0, False)
        assert condition.clearing_in_s(40.0) == pytest.approx(60.0)


class TestStartup:
    def test_starts_unsafe(self):
        site = Site()
        state = site.tick(1)
        assert state.is_safe is False
        assert state.reasons

    def test_becomes_safe_once_both_streams_are_healthy(self):
        site = Site()
        assert site.settle().is_safe is True

    def test_stays_unsafe_if_only_one_stream_arrives(self):
        # Half a weather station is not a weather station.
        site = Site()
        state = site.tick(60, send_wind=False)
        assert state.is_safe is False
        assert state.conditions['wind_data'] is True
        assert state.conditions['rain_data'] is False

    def test_a_restart_in_the_rain_does_not_briefly_report_safe(self):
        site = Site()
        state = site.tick(5, rain=('W', 'W', 'D'))
        assert state.is_safe is False


class TestRainLatch:
    def test_rain_makes_it_unsafe(self):
        site = Site()
        site.settle()
        assert site.tick(2, rain=('W', 'W', 'D')).is_safe is False

    def test_the_latch_holds_after_the_sensors_dry(self):
        # The whole point: a heated sensor takes two to five minutes to dry, and
        # reopening onto sensors that are still evaporating defeats the purpose.
        site = Site()
        site.settle()
        site.tick(5, rain=('W', 'W', 'W'))
        state = site.tick(120, rain=('D', 'D', 'D'))
        assert state.is_safe is False
        assert state.conditions['rain'] is True

    def test_the_latch_releases_after_the_full_clear_period(self):
        site = Site()
        site.settle()
        site.tick(5, rain=('W', 'W', 'W'))
        state = site.tick(625, rain=('D', 'D', 'D'))
        assert state.is_safe is True

    def test_rain_during_the_countdown_restarts_it(self):
        site = Site(config(rain_clear_s=100.0))
        site.settle()
        site.tick(3, rain=('W', 'W', 'D'))
        site.tick(50, rain=('D', 'D', 'D'))
        site.tick(3, rain=('W', 'W', 'D'))          # it starts again
        state = site.tick(60, rain=('D', 'D', 'D'))
        assert state.is_safe is False
        state = site.tick(60, rain=('D', 'D', 'D'))
        assert state.is_safe is True


class TestWindLatch:
    def test_sustained_wind_makes_it_unsafe(self):
        site = Site()
        site.settle()
        assert site.tick(60, wind_speed=8.0).is_safe is False

    def test_wind_clears_faster_than_rain(self):
        # Nothing has to dry out. Holding the dome shut for ten minutes after
        # the wind drops would cost sky for no reason.
        #
        # Note the two-part delay, which is easy to overlook: the 60 s rolling
        # mean has to forget the gale BEFORE the 120 s latch even starts
        # counting. Total hold is roughly 140 s here, and up to ~240 s for a
        # gust, which ages out of a 120 s window. Still far short of rain's ten
        # minutes, which is the point.
        site = Site()
        site.settle()
        site.tick(60, wind_speed=8.0)
        assert site.tick(120, wind_speed=1.0).is_safe is False
        state = site.tick(40, wind_speed=1.0)
        assert state.is_safe is True

    def test_wind_and_rain_latch_independently(self):
        site = Site()
        site.settle()
        site.tick(60, wind_speed=8.0, rain=('W', 'W', 'D'))
        state = site.tick(200, wind_speed=1.0, rain=('D', 'D', 'D'))
        assert state.conditions['wind'] is False     # wind has cleared
        assert state.conditions['rain'] is True      # rain has not
        assert state.is_safe is False


class TestFailClosed:
    def test_the_rain_stream_going_silent_is_unsafe(self):
        # The Windows 7 box has died. The legacy software could not tell this
        # from a clear night.
        site = Site()
        site.settle()
        state = site.tick(30, send_rain=False)
        assert state.is_safe is False
        assert state.conditions['rain_data'] is True

    def test_the_wind_stream_going_silent_is_unsafe(self):
        site = Site()
        site.settle()
        assert site.tick(30, send_wind=False).is_safe is False

    def test_a_failed_serial_port_is_unsafe(self):
        site = Site()
        site.settle()
        state = site.tick(5, rain=('e', 'e', 'e'), rain_ok=False)
        assert state.is_safe is False

    def test_all_detectors_errored_is_unsafe_not_dry(self):
        site = Site()
        site.settle()
        assert site.tick(5, rain=('e', 'e', 'E')).is_safe is False

    def test_a_stream_that_comes_back_recovers_quickly(self):
        # A network blip must not cost ten minutes of sky.
        site = Site()
        site.settle()
        site.tick(30, send_rain=False)
        state = site.tick(35)
        assert state.is_safe is True


class TestPublishedValues:
    def test_rain_rate_is_zero_when_all_is_well(self):
        site = Site()
        assert site.settle().rain_rate_mm_h == 0.0

    def test_rain_rate_stays_up_while_the_latch_holds(self):
        # Arcsecond reopens on the transition back to GO. A rate that fell to
        # zero the moment the sensors dried would reopen the dome early.
        site = Site()
        site.settle()
        site.tick(5, rain=('W', 'W', 'D'))
        state = site.tick(120, rain=('D', 'D', 'D'))
        assert state.rain_rate_mm_h > 0.0

    def test_rain_rate_is_raised_when_the_sensors_cannot_be_read(self):
        # Deliberately defensive. Arcsecond's evaluator SKIPS a condition whose
        # value is missing and, with nothing else triggered, decides GO -- so a
        # blind sensor reads to it as good weather. Publishing the wet value
        # means a device that can be reached but cannot see still says NOGO.
        site = Site()
        site.settle()
        state = site.tick(5, rain=('e', 'e', 'e'), rain_ok=False)
        assert state.rain_rate_mm_h > 0.0

    def test_wind_values_are_published_in_metres_per_second(self):
        site = Site()
        state = site.tick(60, wind_speed=4.0)
        assert state.wind_speed_ms == pytest.approx(4.0, abs=0.1)
        assert state.wind_gust_ms == pytest.approx(4.0, abs=0.5)

    def test_detector_states_are_exposed_for_the_display(self):
        site = Site()
        state = site.tick(2, rain=('D', 'w', 'W'))
        assert state.detector_states == {'H127': 'D', 'H50': 'w', 'ACC': 'W'}

    def test_reasons_name_what_is_wrong(self):
        site = Site()
        site.settle()
        state = site.tick(60, wind_speed=9.0)
        assert any('wind' in r for r in state.reasons)


class TestTransitions:
    def test_change_is_reported_once(self):
        site = Site()
        state = site.tick(1)
        assert site.evaluator.changed(state) is True
        assert site.evaluator.changed(state) is False


class TestBlindnessDoesNotClearALatch:
    """A latch may only be released by evidence that the weather is good --
    never by the absence of evidence."""

    def test_a_dead_bridge_does_not_run_out_the_rain_latch(self):
        # It rains; then the Windows 7 box dies. If the ten-minute countdown
        # ran through the blackout, the observatory would be declared safe on
        # the strength of ten minutes in which nobody could see the sensors.
        site = Site()
        site.settle()
        site.tick(5, rain=('W', 'W', 'D'))
        site.tick(700, send_rain=False)             # blind, far longer than the latch
        state = site.tick(5, rain=('D', 'D', 'D'))  # bridge returns, reporting dry
        assert state.is_safe is False
        assert state.conditions['rain'] is True

    def test_and_then_clears_normally_once_it_can_see_again(self):
        site = Site()
        site.settle()
        site.tick(5, rain=('W', 'W', 'D'))
        site.tick(700, send_rain=False)
        state = site.tick(625, rain=('D', 'D', 'D'))
        assert state.is_safe is True

    def test_a_silent_anemometer_does_not_run_out_the_wind_latch(self):
        site = Site()
        site.settle()
        site.tick(60, wind_speed=9.0)
        site.tick(300, send_wind=False)
        state = site.tick(5, wind_speed=1.0)
        assert state.is_safe is False
