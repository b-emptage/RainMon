# -*- coding: utf-8 -*-
r"""Wind from the anemometer's own multicast stream.

The instrument multicasts to the LAN; nothing in this project produces that
traffic, and `wind_sensor.py` in the WindSensor repo was only ever a display
client for it.

WHAT IS ACTUALLY ON THE WIRE, from a 15 minute capture at Greenhill:

    UdPbC\x00\s:D383P1,s:WI4383*23\$IIMWV,273,R,007.51,M,A*15
    |_______________________________||_________________________|
     proprietary UDP wrapper and an   the NMEA 0183 sentence
     NMEA 0183 v4 TAG block

The instrument is an Observator OMC-140 -- it says so itself in a `$WIVER`
sentence it emits every few minutes among the wind ones. The wind sentence is
MWV (Wind Speed and Angle):

    $IIMWV,<angle>,<reference>,<speed>,<units>,<status>*<checksum>

so the reading carries its own units and its own validity flag, and neither has
to be assumed.

THIS IS PARSED AS NMEA, NOT BY FIELD INDEX. The legacy display splits the whole
datagram on commas and reads fields 2 and 4, which happens to land on angle and
speed -- but only because the TAG block contributes exactly one comma. Any
change to that header silently shifts both fields, and the reading would still
look like a number. It also means the `$WIVER` sentence is read as wind: angle
1, speed "WI", saved only by the float() raising into a bare except.

Speed is converted from whatever `units` says, rather than assumed to be m/s.
Everything downstream is m/s, because ASCOM and Arcsecond both require it and a
km/h value anywhere would become a threshold wrong by a factor of 3.6.

The `reference` field is why the north offset exists: the OMC-140 is reporting
`R`, an angle relative to its own zero mark, so the offset rotates it to true
north. The offset is applied ONLY to relative readings -- if the instrument is
ever configured to emit `T`, adding it would introduce the very error it exists
to remove.
"""


import math
import re
from typing import List, Optional, Tuple

UNKNOWN = 'unknown'
CALM = 'calm'
WINDY = 'windy'


class WindParseError(ValueError):
    """A wind sentence that is present but not usable."""


class NoWindSentence(ValueError):
    """This datagram carries no wind reading at all.

    Not an error. The instrument also emits `$WIVER` identification sentences,
    and a datagram that is simply not about wind must not be counted against
    the stream's health -- while equally not being allowed to pass for a
    reading.
    """


# One MWV sentence anywhere inside the datagram, past whatever wrapper and TAG
# block precede it.
MWV = re.compile(
    r'\$(?P<talker>[A-Z]{2})MWV,'
    r'(?P<angle>[^,]*),(?P<reference>[^,]*),'
    r'(?P<speed>[^,]*),(?P<units>[^,]*),(?P<status>[^,*]*)'
    r'(?:\*(?P<checksum>[0-9A-Fa-f]{2}))?')

# NMEA speed units -> metres per second.
UNIT_TO_MS = {
    'M': 1.0,               # metres per second
    'N': 0.514444,          # knots
    'K': 1.0 / 3.6,         # kilometres per hour
    'S': 0.44704,           # miles per hour
}

REFERENCE_RELATIVE = 'R'
REFERENCE_TRUE = 'T'
STATUS_VALID = 'A'


def nmea_checksum(sentence_body):
    # type: (str) -> int
    """XOR of every character between the '$' and the '*'."""
    value = 0
    for char in sentence_body:
        value ^= ord(char)
    return value


def parse_sentence(text, max_plausible_ms, verify_checksum=True):
    # type: (str, float, bool) -> Tuple[float, float, str]
    """One datagram -> (angle degrees, speed m/s, reference).

    Raises NoWindSentence when the datagram is not about wind, and
    WindParseError when it is but cannot be trusted. The distinction matters:
    the first is normal traffic, the second is a sick instrument.
    """
    if not isinstance(text, str):
        raise NoWindSentence('datagram is not text')

    match = MWV.search(text)
    if match is None:
        raise NoWindSentence('no MWV sentence in this datagram')

    if verify_checksum and match.group('checksum'):
        body = match.group(0)
        body = body[1:body.rindex('*')]         # between '$' and '*'
        if nmea_checksum(body) != int(match.group('checksum'), 16):
            raise WindParseError('MWV checksum does not match')

    status = match.group('status').strip()
    if status and status != STATUS_VALID:
        # 'V' means the instrument is telling us its own reading is void.
        raise WindParseError('MWV status is {!r}, not valid'.format(status))

    units = match.group('units').strip().upper()
    if units not in UNIT_TO_MS:
        raise WindParseError('unknown speed units {!r}'.format(units))

    try:
        angle = float(match.group('angle'))
    except ValueError:
        raise WindParseError(
            'angle {!r} is not a number'.format(match.group('angle')))
    try:
        speed = float(match.group('speed')) * UNIT_TO_MS[units]
    except ValueError:
        raise WindParseError(
            'speed {!r} is not a number'.format(match.group('speed')))

    if not 0.0 <= angle <= 360.0:
        raise WindParseError('angle {} is outside 0-360'.format(angle))
    if speed < 0.0:
        raise WindParseError('speed {} is negative'.format(speed))
    if speed > max_plausible_ms:
        # A reading this high is a misparse, not weather: 100 m/s is 360 km/h.
        # It matters because it would trip the gust threshold and close the
        # dome on nothing worse than a change of sentence format.
        raise WindParseError(
            'speed {} m/s exceeds the plausible maximum {}'.format(
                speed, max_plausible_ms))

    return angle, speed, match.group('reference').strip().upper()


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
        # (arrival, angle, speed m/s, reference)
        self._samples = []          # type: List[Tuple[float, float, float, str]]
        self._last_update = None    # type: Optional[float]
        self._rejected = 0
        self._accepted = 0
        self._ignored = 0

    # -- ingest ---------------------------------------------------------------

    def update(self, now, text):
        # type: (float, str) -> bool
        """Take one datagram. Returns whether it yielded a wind reading.

        A datagram that is not a wind sentence at all is ignored quietly. One
        that IS a wind sentence but is unusable counts as a rejection and does
        NOT refresh the staleness clock -- an instrument emitting nonsense is a
        dead source, and must age into UNKNOWN exactly as silence does.
        """
        try:
            angle, speed, reference = parse_sentence(
                text, self._config.wind_max_plausible_ms)
        except NoWindSentence:
            self._ignored += 1
            return False
        except WindParseError:
            self._rejected += 1
            return False

        self._samples.append((now, angle, speed, reference))
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

        # The offset rotates the instrument's own zero mark to true north, so
        # it belongs only to a RELATIVE bearing. The OMC-140 reports R today;
        # if it is ever configured to emit T, adding the offset would introduce
        # exactly the error it exists to remove.
        if all(s[3] == REFERENCE_TRUE for s in window):
            return mean % 360.0

        # NEGATIVE OFFSETS ARE FINE, and the modulo is what makes them fine.
        # Python's % with a positive divisor always returns a non-negative
        # result -- -20.0 % 360.0 is 340.0 -- unlike C, Java or JavaScript,
        # where it would stay -20.0. Anyone porting this expression, or
        # "simplifying" the modulo away because the inputs look bounded, will
        # reintroduce a negative bearing that ASCOM does not allow.
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
        """Wind sentences that were unusable. A sick instrument."""
        return self._rejected

    @property
    def ignored_count(self):
        """Datagrams carrying no wind sentence -- the periodic $WIVER
        identification, mostly. Normal traffic, not a fault."""
        return self._ignored

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
