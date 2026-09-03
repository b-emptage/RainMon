# -*- coding: utf-8 -*-
"""Fusion: rain plus wind plus the health of both streams, into one verdict.

This is the only thing in the system that decides anything. Both close routes
read it, `SafetyMonitor.IsSafe` publishes it, and `ObservingConditions` is a
view onto the same state.

Two principles run through it.

**Fail closed.** The service starts unsafe and every source that goes quiet,
stale or self-reports a fault takes it back to unsafe. In the software this
replaces, a dead serial link, a disconnected sensor and an errored detector all
produced a wet count of zero -- indistinguishable from a clear night.

**Latch per cause, with a clear delay that fits the cause.** Rain holds for ten
minutes because a heated sensor takes two to five to dry, and reopening onto
sensors that are merely still evaporating defeats the point. Wind holds for two
because nothing has to dry, it only has to stop gusting. A stream that dropped
out and came back holds for thirty seconds, because a network blip should not
cost the observatory ten minutes of sky. One delay for all three would be wrong
for two of them.
"""

from typing import Dict, List, Optional

from greenhill.core import rain as rain_module
from greenhill.core import wind as wind_module


class LatchedCondition(object):
    """One reason the observatory might be unsafe, with hysteresis.

    Starts TRIPPED. A safety service that came up believing all was well until
    told otherwise would be safe for the first few seconds of every restart,
    including a restart in the rain.
    """

    def __init__(self, name, clear_seconds, startup_clear_seconds):
        self.name = name
        self.clear_seconds = clear_seconds
        # The startup trip is not evidence of bad weather -- it is the absence
        # of evidence of good weather. Requiring the full ten-minute rain delay
        # after every restart on a clear night would be theatre, so the first
        # clear only has to wait out the settle period.
        self._effective_clear_seconds = startup_clear_seconds
        self._tripped = True
        self._clear_since = None    # type: Optional[float]
        self._reasons = []          # type: List[str]
        self._trip_count = 0

    def update(self, now, currently_bad, reasons=None, hold=False):
        # type: (float, bool, Optional[List[str]], bool) -> None
        """`hold` means "we cannot currently tell", and it FREEZES the clearing
        countdown rather than advancing it.

        This distinction is not cosmetic. Suppose it rains, and then the rain
        bridge dies. Without `hold`, the ten-minute countdown would run right
        through the blackout and release the latch -- reporting the observatory
        safe on the strength of ten minutes during which nobody could see the
        sensors at all. The clear period has to be ten minutes of OBSERVED
        dryness, so the countdown resets whenever the observation lapses.
        """
        if currently_bad:
            if not self._tripped:
                self._trip_count += 1
            self._tripped = True
            self._clear_since = None
            self._effective_clear_seconds = self.clear_seconds
            self._reasons = list(reasons or [])
            return

        if not self._tripped:
            self._reasons = []
            return

        if hold:
            self._clear_since = None
            return

        if self._clear_since is None:
            self._clear_since = now
        elif now - self._clear_since >= self._effective_clear_seconds:
            self._tripped = False
            self._clear_since = None
            self._reasons = []
            self._effective_clear_seconds = self.clear_seconds

    @property
    def active(self):
        # type: () -> bool
        return self._tripped

    @property
    def trip_count(self):
        return self._trip_count

    def clearing_in_s(self, now):
        # type: (float) -> Optional[float]
        """Seconds until this clears, or None if it is not counting down."""
        if not self._tripped or self._clear_since is None:
            return None
        return max(0.0, self._effective_clear_seconds - (now - self._clear_since))

    def describe(self, now):
        # type: (float) -> str
        if not self._tripped:
            return '{}: ok'.format(self.name)
        remaining = self.clearing_in_s(now)
        if remaining is not None:
            return '{}: clearing, {:.0f}s to go'.format(self.name, remaining)
        if self._reasons:
            return '{}: {}'.format(self.name, '; '.join(self._reasons))
        return '{}: tripped'.format(self.name)


class SafetyState(object):
    """The answer, at one moment."""

    def __init__(self, is_safe, reasons, rain_rate_mm_h, wind_speed_ms,
                 wind_gust_ms, wind_direction_deg, temperature_c,
                 conditions, rain_age_s, wind_age_s, detector_states):
        self.is_safe = is_safe
        self.reasons = reasons
        self.rain_rate_mm_h = rain_rate_mm_h
        self.wind_speed_ms = wind_speed_ms
        self.wind_gust_ms = wind_gust_ms
        self.wind_direction_deg = wind_direction_deg
        self.temperature_c = temperature_c
        self.conditions = conditions            # type: Dict[str, bool]
        self.rain_age_s = rain_age_s
        self.wind_age_s = wind_age_s
        self.detector_states = detector_states

    def __repr__(self):
        return 'SafetyState(is_safe={}, reasons={!r})'.format(
            self.is_safe, self.reasons)


class SafetyEvaluator(object):
    """Owns the two monitors and the latches.

    `update(now)` must be called after every arriving packet AND on a timer, so
    that a source going silent is noticed. A verdict that only changes when
    data arrives can never notice data not arriving, which is the failure this
    whole design exists to catch.
    """

    def __init__(self, config, rain_monitor=None, wind_monitor=None):
        self._config = config
        self.rain = rain_monitor or rain_module.RainMonitor(config)
        self.wind = wind_monitor or wind_module.WindMonitor(config)

        settle = config.settle_s
        self._conditions = {
            'rain': LatchedCondition('rain', config.rain_clear_s, settle),
            'wind': LatchedCondition('wind', config.wind_clear_s, settle),
            'rain_data': LatchedCondition('rain_data', settle, settle),
            'wind_data': LatchedCondition('wind_data', settle, settle),
        }
        self._last_is_safe = None       # type: Optional[bool]

    @property
    def conditions(self):
        return self._conditions

    def update(self, now):
        # type: (float) -> SafetyState
        rain_verdict, rain_reasons = self.rain.verdict(now)
        wind_verdict, wind_reasons = self.wind.verdict(now)

        # The weather conditions are HELD, not advanced, while their source
        # cannot be read: a latch may only be released by positive evidence
        # that the weather is good, never by the absence of evidence.
        self._conditions['rain_data'].update(
            now, rain_verdict == rain_module.UNAVAILABLE, rain_reasons)
        self._conditions['rain'].update(
            now, rain_verdict == rain_module.RAIN, rain_reasons,
            hold=rain_verdict == rain_module.UNAVAILABLE)
        self._conditions['wind_data'].update(
            now, wind_verdict == wind_module.UNKNOWN, wind_reasons)
        self._conditions['wind'].update(
            now, wind_verdict == wind_module.WINDY, wind_reasons,
            hold=wind_verdict == wind_module.UNKNOWN)

        active = [c for c in self._conditions.values() if c.active]
        is_safe = not active
        reasons = [c.describe(now) for c in
                   sorted(active, key=lambda c: c.name)]

        return SafetyState(
            is_safe=is_safe,
            reasons=reasons,
            rain_rate_mm_h=self._rain_rate(),
            wind_speed_ms=self.wind.mean_speed_ms(now),
            wind_gust_ms=self.wind.gust_ms(now),
            wind_direction_deg=self.wind.direction_deg(now),
            temperature_c=self.rain.temperature_c(),
            conditions={name: c.active for name, c in self._conditions.items()},
            rain_age_s=self.rain.age_s(now),
            wind_age_s=self.wind.age_s(now),
            detector_states=self.rain.detector_states)

    def _rain_rate(self):
        # type: () -> float
        """What ObservingConditions publishes as RainRate.

        Non-zero while the rain latch holds -- NOT while the sensors happen to
        be wet. Arcsecond fires its recovery procedure on the transition back
        to GO, so a rate that dropped to zero the moment the sensors dried
        would reopen the dome onto sensors that are still evaporating.

        Also non-zero when the rain sensors cannot be read at all. That is
        deliberate and it is defensive: Arcsecond's evaluator currently SKIPS a
        condition whose value is missing, and with nothing else triggered
        decides GO -- so a blind sensor would read to it as good weather.
        Publishing the wet value instead means a device that can be reached but
        cannot see still produces a NOGO.

        This does not cover the weather service itself being unreachable. Only
        the Arcsecond-side fix, or a dome-side deadman, covers that.
        """
        blind = (self._conditions['rain'].active
                 or self._conditions['rain_data'].active)
        return self._config.rain_rate_when_wet_mm_h if blind else 0.0

    def changed(self, state):
        # type: (SafetyState) -> bool
        """Whether IsSafe flipped since the last call. For logging, so a long
        quiet night does not fill the log with identical lines."""
        changed = self._last_is_safe is None or state.is_safe != self._last_is_safe
        self._last_is_safe = state.is_safe
        return changed
