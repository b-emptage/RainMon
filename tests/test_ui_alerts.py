# -*- coding: utf-8 -*-
"""What the observatory says out loud, and when.

The failure mode of a spoken alert is not silence, it is nagging: one that
repeats every second gets muted, and a muted alert protects nobody. So most of
these tests are about restraint.
"""

import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from greenhill.ui.alerts import FAULTS, RAIN, SAFETY, AlertPolicy, Snapshot  # noqa: E402


def snap(wet=0, active=6, rain_ok=True, wind_ok=True, safe=True, reasons=None):
    return Snapshot(wet_sections=wet, active_sections=active, rain_ok=rain_ok,
                    wind_ok=wind_ok, is_safe=safe, reasons=reasons or [])


class Run:
    """Feeds snapshots and collects everything said."""

    def __init__(self, **kwargs):
        # Most tests are not about the startup grace, so it is short here and
        # exercised deliberately in TestStartupGrace.
        kwargs.setdefault('startup_grace_seconds', 0.0)
        self.policy = AlertPolicy(**kwargs)
        self.now = 0.0
        self.said = []

    def feed(self, snapshot, advance=1.0, times=1):
        for _ in range(times):
            self.said.extend(self.policy.update(self.now, snapshot))
            self.now += advance
        return self.said

    @property
    def last(self):
        return self.said[-1] if self.said else None


class TestRestraint:
    def test_says_nothing_on_the_first_snapshot(self):
        # Otherwise opening the window greets you with "rain detected" for rain
        # that stopped an hour ago.
        run = Run()
        assert run.feed(snap(wet=4, safe=False)) == []

    def test_says_nothing_while_nothing_changes(self):
        run = Run()
        run.feed(snap(), times=100)
        assert run.said == []

    def test_the_rain_announcement_repeats_but_does_not_nag(self):
        run = Run(repeat_seconds=30.0)
        run.feed(snap())                        # establish a baseline
        run.feed(snap(wet=2, safe=False), times=120)    # two minutes of rain
        rain_phrases = [s for s in run.said if 'Rain detected' in s]
        # One on the transition, then one every thirty seconds. Not 120.
        assert 4 <= len(rain_phrases) <= 5


class TestRain:
    def test_announces_rain_with_the_count(self):
        run = Run()
        run.feed(snap())
        run.feed(snap(wet=3, active=6, safe=False))
        assert 'Rain detected. 3 of 6 sections wet.' in run.said

    def test_announces_drying(self):
        run = Run()
        run.feed(snap())
        run.feed(snap(wet=2, safe=False))
        run.feed(snap(wet=0))
        assert 'Rain sensors dry.' in run.said

    def test_does_not_announce_dry_when_the_sensors_went_unreadable(self):
        # The most misleading sentence this could possibly utter. Wet then
        # unreadable is a fault, not good news.
        run = Run()
        run.feed(snap())
        run.feed(snap(wet=2, safe=False))
        run.feed(snap(wet=None, rain_ok=False, safe=False))
        assert 'Rain sensors dry.' not in run.said

    def test_a_second_rain_episode_is_announced_again(self):
        run = Run(repeat_seconds=30.0)
        run.feed(snap())
        run.feed(snap(wet=2, safe=False))
        run.feed(snap(wet=0))
        run.feed(snap(wet=2, safe=False))
        assert len([s for s in run.said if 'Rain detected' in s]) == 2


class TestStartupGrace:
    """Every source looks dead for the first fraction of a second."""

    def test_does_not_cry_fault_the_moment_the_window_opens(self):
        run = Run(startup_grace_seconds=20.0)
        run.feed(snap(rain_ok=False, wind_ok=False, wet=None, safe=False),
                 times=10)
        assert run.said == []

    def test_does_not_announce_a_recovery_that_was_never_a_failure(self):
        # Opening the window used to say "rain sensors responding again"
        # before a single packet had been missed -- a recovery from a failure
        # nobody was ever told about.
        run = Run(startup_grace_seconds=20.0)
        run.feed(snap(rain_ok=False, wind_ok=False, wet=None, safe=False))
        run.feed(snap(), times=5)
        assert not any('responding again' in s for s in run.said)

    def test_a_genuinely_dead_sensor_is_announced_once_settled(self):
        # Unlike rain, which would be stale news, a sensor that is dead right
        # now is worth saying as soon as we can be sure.
        run = Run(startup_grace_seconds=20.0)
        run.feed(snap(rain_ok=False, wet=None, safe=False), times=40)
        assert 'Rain sensors not responding.' in run.said


class TestFaults:
    def test_announces_a_stream_that_stops(self):
        # The announcement the legacy monitor could not make, because it could
        # not tell. An astronomer who hears nothing assumes it is dry.
        run = Run()
        run.feed(snap())
        run.feed(snap(rain_ok=False, wet=None, safe=False))
        assert 'Rain sensors not responding.' in run.said

    def test_announces_recovery(self):
        run = Run()
        run.feed(snap())
        run.feed(snap(rain_ok=False, wet=None, safe=False))
        run.feed(snap())
        assert 'Rain sensors responding again.' in run.said

    def test_the_wind_sensor_is_named_separately(self):
        run = Run()
        run.feed(snap())
        run.feed(snap(wind_ok=False, safe=False))
        assert 'Wind sensor not responding.' in run.said

    def test_a_fault_is_announced_once_not_repeatedly(self):
        run = Run()
        run.feed(snap())
        run.feed(snap(rain_ok=False, wet=None, safe=False), times=60)
        assert len([s for s in run.said if 'not responding' in s]) == 1


class TestSafety:
    def test_announces_going_unsafe_and_back(self):
        run = Run()
        run.feed(snap())
        run.feed(snap(safe=False, wet=2))
        assert 'Conditions unsafe.' in run.said
        run.feed(snap(safe=True))
        assert 'Conditions safe.' in run.said


class TestCategories:
    def test_rain_can_be_silenced_on_its_own(self):
        run = Run(categories=(SAFETY, FAULTS))
        run.feed(snap())
        run.feed(snap(wet=2, safe=False))
        assert not any('Rain detected' in s for s in run.said)
        assert 'Conditions unsafe.' in run.said

    def test_faults_can_be_kept_when_everything_else_is_silenced(self):
        # The one to keep if you keep only one.
        run = Run(categories=(FAULTS,))
        run.feed(snap())
        run.feed(snap(wet=2, rain_ok=False, safe=False))
        assert run.said == ['Rain sensors not responding.']
