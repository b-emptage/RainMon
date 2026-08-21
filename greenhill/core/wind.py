# -*- coding: utf-8 -*-
"""Wind from the anemometer's own multicast stream.

The instrument multicasts a comma-separated ASCII sentence to the LAN; nothing
in this project produces it, and `wind_sensor.py` in the WindSensor repo was
only ever a display client for it.

TWO THINGS HERE ARE PROVISIONAL AND BOTH ARE IN CONFIG.

The sentence format is undocumented. The legacy display reads direction from
field 2 and speed from field 4, by fixed index, behind a bare `except:` that
silently yields no reading -- so if those positions were ever wrong, nothing
would have said so. `tools/record_streams.py` prints a real sentence field by
field to settle it. The positions are configuration, so correcting them is a
config edit rather than a release.

The legacy display also adds 30 degrees to the raw direction, with no note
saying why; presumably a mounting correction. It is carried forward so
behaviour does not change silently, but WindDirection should not be trusted
until someone has checked it against a known reference.

Speed is metres per second -- confirmed with the observatory -- and stays that
way everywhere. ASCOM and Arcsecond both want m/s, and a km/h value anywhere in
this code would eventually become a threshold wrong by a factor of 3.6.
"""

import math
from typing import List, Optional, Tuple

UNKNOWN = 'unknown'
CALM = 'calm'
WINDY = 'windy'


class WindParseError(ValueError):
    """A sentence that is not a usable wind reading."""


def parse_sentence(text, direction_field, speed_field, max_plausible_ms):
    # type: (str, int, int, float) -> Tuple[float, float]
    """One sentence -> (raw direction degrees, speed m/s).

    Strict: anything questionable raises. The legacy parser returned "no
    reading" for every failure, which is indistinguishable from a calm night
    and left the display frozen on its last good values indefinitely.
    """
    if not isinstance(text, str):
        raise WindParseError('sentence is not text')

    # NMEA-style sentences end with a "*hh" checksum. Nothing here validates it
    # -- the format is not confirmed to be NMEA -- but it must not be parsed as
    # part of the last field.
    body = text.strip().split('*')[0]
    fields = body.split(',')

    for index, name in ((direction_field, 'direction'), (speed_field, 'speed')):
        if index >= len(fields):
            raise WindParseError(
                'sentence has {} fields, need index {} for {}'.format(
                    len(fields), index, name))

    try:
        direction = float(fields[direction_field])
    except ValueError:
        raise WindParseError(
            'direction field {} is {!r}, not a number'.format(
                direction_field, fields[direction_field]))
    try:
        speed = float(fields[speed_field])
    except ValueError:
        raise WindParseError(
            'speed field {} is {!r}, not a number'.format(
                speed_field, fields[speed_field]))

    if not 0.0 <= direction <= 360.0:
        raise WindParseError('direction {} is outside 0-360'.format(direction))
    if speed < 0.0:
        raise WindParseError('speed {} is negative'.format(speed))
    if speed > max_plausible_ms:
        # A reading this high is a misparsed field, not weather: 100 m/s is
        # 360 km/h. Rejecting it matters because it would otherwise trip the
        # gust threshold and close the dome on a formatting change.
        raise WindParseError(
            'speed {} m/s exceeds the plausible maximum {}'.format(
                speed, max_plausible_ms))

    return direction, speed


def circular_mean_and_scatter(degrees_list):
    # type: (List[float]) -> Tuple[float, float]
    """Mean bearing and angular scatter, both in degrees.

    Ordinary averaging cannot do this: the mean of 350 and 10 is 180, which
    points exactly the wrong way.
    """
    if not degrees_list:
        return 0.0, 0.0
    count = float(len(degrees_list))
    sin_mean = sum(math.sin(math.radians(d)) for d in degrees_list) / count
    cos_mean = sum(math.cos(math.radians(d)) for d in degrees_list) / count

    mean = math.degrees(math.atan2(sin_mean, cos_mean)) % 360.0
    resultant = math.hypot(sin_mean, cos_mean)
    if resultant <= 0.0:
        # Perfectly opposed samples: there is no mean direction to report.
        return mean, 180.0
    scatter = math.degrees(math.sqrt(-2.0 * math.log(min(resultant, 1.0))))
    return mean, scatter


class WindMonitor(object):
    """Rolling wind statistics from the anemometer stream.

    Times are supplied by the caller rather than read from a clock, so the
    tests can advance time deterministically -- and, more importantly, so that
    ARRIVAL TIME is what ages a sample. Nothing here trusts a timestamp from
    the network.
    """

    def __init__(self, config):
        self._config = config
        self._samples = []          # type: List[Tuple[float, float, float]]
        self._last_update = None    # type: Optional[float]
        self._rejected = 0
        self._accepted = 0

    # -- ingest ---------------------------------------------------------------

    def update(self, now, text):
        # type: (float, str) -> bool
        """Take one sentence. Returns whether it was usable.

        A rejected sentence does NOT refresh the staleness clock. A sender
        emitting garbage is a dead source, not a live one, and must age out
        into UNKNOWN exactly as silence does.
        """
        try:
            direction, speed = parse_sentence(
                text,
                self._config.wind_direction_field,
                self._config.wind_speed_field,
                self._config.wind_max_plausible_ms)
        except WindParseError:
            self._rejected += 1
            return False

        self._samples.append((now, direction, speed))
        self._last_update = now
        self._accepted += 1
        self._prune(now)
        return True

    def _prune(self, now):
        horizon = now - max(self._config.wind_mean_window_s,
                            self._config.wind_gust_window_s)
        while self._samples and self._samples[0][0] < horizon:
            self._samples.pop(0)

    # -- statistics -----------------------------------------------------------

    def _within(self, now, window):
        cutoff = now - window
        return [s for s in self._samples if s[0] >= cutoff]

    def mean_speed_ms(self, now):
        # type: (float) -> Optional[float]
        window = self._within(now, self._config.wind_mean_window_s)
        if not window:
            return None
        return sum(s[2] for s in window) / len(window)

    def gust_ms(self, now):
        # type: (float) -> Optional[float]
        """Strongest short-term average in the gust window.

        The ASCOM definition, and the physically meaningful one: a gust is a
        few seconds of strong wind, not one fast sample. An ultrasonic head
        produces the odd spiky reading, and closing the dome on one of those
        would teach the operators to distrust the system.
        """
        window = self._within(now, self._config.wind_gust_window_s)
        if not window:
            return None

        span = self._config.wind_gust_average_s
        best = 0.0
        start = 0
        running = 0.0
        for end in range(len(window)):
            running += window[end][2]
            while window[end][0] - window[start][0] > span:
                running -= window[start][2]
                start += 1
            best = max(best, running / (end - start + 1))
        return best

    def peak_ms(self, now):
        # type: (float) -> Optional[float]
        """Fastest single sample in the gust window. Diagnostic only -- the
        threshold uses gust_ms, so this is what tells an operator whether a
        borderline night was gusty or merely spiky."""
        window = self._within(now, self._config.wind_gust_window_s)
        return max(s[2] for s in window) if window else None

    def direction_deg(self, now):
        # type: (float) -> float
        """Mean bearing over the averaging window, north offset applied.

        Samples below the minimum speed are excluded: an anemometer pointing
        nowhere in particular reports a direction anyway, and averaging those
        in would swing the result around. Returns 0.0 when there is nothing to
        average, which is also what ASCOM asks for in a calm.
        """
        window = [s for s in self._within(now, self._config.wind_mean_window_s)
                  if s[2] >= self._config.wind_direction_min_speed_ms]
        if not window:
            return 0.0
        mean, _ = circular_mean_and_scatter([s[1] for s in window])
        return (mean + self._config.wind_north_offset_deg) % 360.0

    def direction_scatter_deg(self, now):
        # type: (float) -> float
        window = [s for s in self._within(now, self._config.wind_mean_window_s)
                  if s[2] >= self._config.wind_direction_min_speed_ms]
        if not window:
            return 0.0
        return circular_mean_and_scatter([s[1] for s in window])[1]

    # -- health ---------------------------------------------------------------

    @property
    def sample_count(self):
        return len(self._samples)

    @property
    def rejected_count(self):
        return self._rejected

    @property
    def accepted_count(self):
        return self._accepted

    def age_s(self, now):
        # type: (float) -> Optional[float]
        if self._last_update is None:
            return None
        return now - self._last_update

    def is_stale(self, now):
        # type: (float) -> bool
        age = self.age_s(now)
        return age is None or age > self._config.wind_max_age_s

    def verdict(self, now):
        # type: (float) -> Tuple[str, List[str]]
        """(UNKNOWN | CALM | WINDY, reasons).

        UNKNOWN is not a mild version of CALM -- the caller treats it as
        unsafe. A stream that has stopped, or that has not yet said enough to
        average, is exactly as unhelpful as one reporting a gale.
        """
        if self._last_update is None:
            return UNKNOWN, ['no wind data received']
        if self.is_stale(now):
            return UNKNOWN, ['wind data is {:.0f}s old'.format(self.age_s(now))]
        if self.sample_count < self._config.wind_min_samples:
            return UNKNOWN, ['only {} wind samples, need {}'.format(
                self.sample_count, self._config.wind_min_samples)]

        reasons = []
        mean = self.mean_speed_ms(now)
        gust = self.gust_ms(now)
        if mean is not None and mean > self._config.wind_sustained_max_ms:
            reasons.append('sustained wind {:.1f} m/s over {:.0f} m/s'.format(
                mean, self._config.wind_sustained_max_ms))
        if gust is not None and gust > self._config.wind_gust_max_ms:
            reasons.append('gust {:.1f} m/s over {:.0f} m/s'.format(
                gust, self._config.wind_gust_max_ms))
        return (WINDY, reasons) if reasons else (CALM, [])
