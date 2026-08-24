# -*- coding: utf-8 -*-
"""Wind statistics and verdicts.

Time is passed in rather than read from a clock, so these tests advance it
deterministically instead of sleeping.
"""

import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from greenhill.core import wind as W  # noqa: E402
from greenhill.core.config import WeatherConfig  # noqa: E402


from conftest import mwv_datagram, ver_datagram  # noqa: E402


def sentence(direction, speed):
    """A real anemometer datagram: wrapper, TAG block, MWV sentence."""
    return mwv_datagram(angle=direction, speed=speed)


def feed(monitor, seconds, speed, direction=90.0, start=0.0, rate=1.0):
    """Feed `seconds` of samples at `rate` Hz. Returns the final time."""
    now = start
    for _ in range(int(seconds * rate)):
        monitor.update(now, sentence(direction, speed))
        now += 1.0 / rate
    return now


class TestParsing:
    """Against the real datagram shape, captured at Greenhill."""

    def test_reads_a_real_datagram(self):
        angle, speed, reference = W.parse_sentence(
            'UdPbC\x00\\s:D383P1,s:WI4383*23\\$IIMWV,273,R,007.51,M,A*15', 100.0)
        assert angle == 273.0
        assert speed == pytest.approx(7.51)
        assert reference == 'R'

    def test_is_not_fooled_by_the_wrapper(self):
        # The legacy display splits the whole datagram on commas and reads
        # fields 2 and 4. That lands on the right values only because the TAG
        # block contributes exactly one comma; any change to the header shifts
        # both, silently, and the result still looks like a number.
        plain = W.parse_sentence(mwv_datagram(angle=273, speed=7.51,
                                              wrapper=False), 100.0)
        wrapped = W.parse_sentence(mwv_datagram(angle=273, speed=7.51), 100.0)
        assert plain == wrapped

    def test_the_identification_sentence_is_not_a_wind_reading(self):
        # $WIVER arrives every few minutes. Read by field index it yields
        # angle 1 and speed 'WI'; the legacy parser is saved only by float()
        # raising into a bare except.
        with pytest.raises(W.NoWindSentence):
            W.parse_sentence(ver_datagram(), 100.0)

    def test_a_datagram_with_no_sentence_is_not_an_error(self):
        # Distinct from a broken sentence: this is normal traffic.
        with pytest.raises(W.NoWindSentence):
            W.parse_sentence('nothing of interest here', 100.0)

    @pytest.mark.parametrize('units,value,expected_ms', [
        ('M', 7.51, 7.51),          # metres per second
        ('N', 10.0, 5.14444),       # knots
        ('K', 36.0, 10.0),          # kilometres per hour
        ('S', 10.0, 4.4704),        # miles per hour
    ])
    def test_speed_is_converted_from_the_units_on_the_wire(self, units, value,
                                                           expected_ms):
        # The sentence carries its units, so they are read rather than assumed.
        # If the instrument is ever reconfigured to knots, 7.51 knots does not
        # silently become 7.51 m/s.
        _, speed, _ = W.parse_sentence(
            mwv_datagram(speed=value, units=units), 100.0)
        assert speed == pytest.approx(expected_ms, rel=1e-4)

    def test_rejects_unknown_units(self):
        with pytest.raises(W.WindParseError, match='units'):
            W.parse_sentence(mwv_datagram(units='Z'), 100.0)

    def test_rejects_a_void_reading(self):
        # 'V' is the instrument telling us its own reading is no good.
        with pytest.raises(W.WindParseError, match='not valid'):
            W.parse_sentence(mwv_datagram(status='V'), 100.0)

    def test_rejects_a_bad_checksum(self):
        with pytest.raises(W.WindParseError, match='checksum'):
            W.parse_sentence(mwv_datagram(corrupt_checksum=True), 100.0)

    def test_accepts_any_talker(self):
        # II at Greenhill, but WI and others are legitimate.
        assert W.parse_sentence(mwv_datagram(talker='WI'), 100.0)[1] > 0

    @pytest.mark.parametrize('angle', [-1, 400])
    def test_rejects_impossible_angles(self, angle):
        with pytest.raises(W.WindParseError, match='0-360'):
            W.parse_sentence(mwv_datagram(angle=angle), 100.0)

    def test_rejects_an_implausible_speed(self):
        # 360 km/h is a misparse, not weather. It matters because it would trip
        # the gust threshold and close the dome on a formatting change.
        with pytest.raises(W.WindParseError, match='plausible'):
            W.parse_sentence(mwv_datagram(speed=150.0), 100.0)

    def test_a_non_numeric_speed_is_a_parse_error(self):
        with pytest.raises(W.WindParseError, match='speed'):
            W.parse_sentence('$IIMWV,273,R,abc,M,A', 100.0, verify_checksum=False)


class TestCircularMean:
    def test_averages_across_north(self):
        # The whole reason for circular statistics: the ordinary mean of 350
        # and 10 is 180, which points exactly the wrong way.
        mean, _ = W.circular_mean_and_scatter([350.0, 10.0])
        assert mean == pytest.approx(0.0, abs=0.01) or mean == pytest.approx(360.0, abs=0.01)

    def test_averages_an_ordinary_spread(self):
        mean, scatter = W.circular_mean_and_scatter([80.0, 90.0, 100.0])
        assert mean == pytest.approx(90.0, abs=0.1)
        assert scatter > 0

    def test_a_steady_wind_has_no_scatter(self):
        _, scatter = W.circular_mean_and_scatter([90.0] * 10)
        assert scatter == pytest.approx(0.0, abs=1e-6)

    def test_a_scattered_wind_has_more_than_a_steady_one(self):
        _, steady = W.circular_mean_and_scatter([90.0, 92.0, 88.0])
        _, wild = W.circular_mean_and_scatter([10.0, 180.0, 300.0])
        assert wild > steady


class TestStatistics:
    def test_mean_speed_over_the_window(self):
        monitor = W.WindMonitor(WeatherConfig())
        now = feed(monitor, 30, speed=4.0)
        assert monitor.mean_speed_ms(now) == pytest.approx(4.0)

    def test_old_samples_fall_out_of_the_mean(self):
        config = WeatherConfig(wind_mean_window_s=10.0, wind_gust_window_s=10.0)
        monitor = W.WindMonitor(config)
        now = feed(monitor, 10, speed=10.0)
        now = feed(monitor, 10, speed=2.0, start=now)
        # The 10 m/s samples are all older than the 10 s window.
        assert monitor.mean_speed_ms(now) == pytest.approx(2.0)

    def test_a_gust_is_a_short_average_not_one_sample(self):
        # An ultrasonic head produces the odd spiky reading. A gust threshold
        # that fired on a single sample would close the dome on sensor noise,
        # so the gust is the strongest three-second average.
        config = WeatherConfig(wind_gust_average_s=3.0)
        monitor = W.WindMonitor(config)
        now = feed(monitor, 30, speed=2.0)
        monitor.update(now, sentence(90, 40.0))     # one wild sample
        now += 1.0
        gust = monitor.gust_ms(now)
        assert gust < 20.0                          # diluted by its neighbours
        assert monitor.peak_ms(now) == pytest.approx(40.0)

    def test_a_real_gust_is_reported(self):
        config = WeatherConfig(wind_gust_average_s=3.0)
        monitor = W.WindMonitor(config)
        now = feed(monitor, 30, speed=2.0)
        now = feed(monitor, 6, speed=12.0, start=now)   # sustained, not a spike
        assert monitor.gust_ms(now) == pytest.approx(12.0, abs=0.5)

    def test_direction_has_the_north_offset_applied(self):
        config = WeatherConfig(wind_north_offset_deg=30.0)
        monitor = W.WindMonitor(config)
        now = feed(monitor, 30, speed=5.0, direction=100.0)
        assert monitor.direction_deg(now) == pytest.approx(130.0, abs=0.1)

    def test_slow_samples_are_left_out_of_the_direction(self):
        # Below about 1 m/s an anemometer's direction is essentially random,
        # and averaging those in swings the answer around.
        config = WeatherConfig(wind_direction_min_speed_ms=1.0,
                               wind_north_offset_deg=0.0)
        monitor = W.WindMonitor(config)
        now = feed(monitor, 20, speed=5.0, direction=90.0)
        now = feed(monitor, 20, speed=0.2, direction=270.0, start=now)
        assert monitor.direction_deg(now) == pytest.approx(90.0, abs=1.0)

    def test_direction_is_zero_in_a_calm(self):
        # Which is also what ASCOM asks for when the speed is zero.
        config = WeatherConfig(wind_direction_min_speed_ms=1.0)
        monitor = W.WindMonitor(config)
        now = feed(monitor, 20, speed=0.1, direction=270.0)
        assert monitor.direction_deg(now) == 0.0


class TestVerdict:
    def config(self, **kwargs):
        base = dict(wind_min_samples=5, wind_max_age_s=15.0)
        base.update(kwargs)
        return WeatherConfig(**base)

    def test_unknown_before_any_data(self):
        monitor = W.WindMonitor(self.config())
        verdict, reasons = monitor.verdict(0.0)
        assert verdict == W.UNKNOWN
        assert reasons

    def test_unknown_with_too_few_samples(self):
        monitor = W.WindMonitor(self.config())
        now = feed(monitor, 2, speed=1.0)
        assert monitor.verdict(now)[0] == W.UNKNOWN

    def test_calm_when_below_both_thresholds(self):
        monitor = W.WindMonitor(self.config())
        now = feed(monitor, 30, speed=2.0)
        assert monitor.verdict(now)[0] == W.CALM

    def test_windy_on_sustained_wind(self):
        monitor = W.WindMonitor(self.config())
        now = feed(monitor, 60, speed=6.0)          # over 5.56 m/s
        verdict, reasons = monitor.verdict(now)
        assert verdict == W.WINDY
        assert any('sustained' in r for r in reasons)

    def test_windy_on_gusts_alone(self):
        # A mean well under the sustained limit, with gusts over the gust
        # limit: the case the sustained threshold alone would miss.
        monitor = W.WindMonitor(self.config())
        now = feed(monitor, 60, speed=1.0)
        now = feed(monitor, 8, speed=10.0, start=now)
        verdict, reasons = monitor.verdict(now)
        assert verdict == W.WINDY
        assert any('gust' in r for r in reasons)
        assert monitor.mean_speed_ms(now) < 5.56

    def test_goes_unknown_when_the_stream_stops(self):
        # The failure the legacy display could not see: it simply froze on its
        # last good values and went on showing a calm night indefinitely.
        monitor = W.WindMonitor(self.config())
        now = feed(monitor, 30, speed=2.0)
        assert monitor.verdict(now)[0] == W.CALM
        assert monitor.verdict(now + 20.0)[0] == W.UNKNOWN

    def test_a_corrupt_instrument_does_not_count_as_liveness(self):
        # A sender emitting broken wind sentences is a dead source. If
        # rejections refreshed the staleness clock, a stuck instrument would
        # look healthy forever.
        monitor = W.WindMonitor(self.config())
        now = feed(monitor, 30, speed=2.0)
        for _ in range(30):
            now += 1.0
            monitor.update(now, mwv_datagram(corrupt_checksum=True))
        assert monitor.rejected_count == 30
        assert monitor.verdict(now)[0] == W.UNKNOWN

    def test_non_wind_traffic_does_not_count_as_liveness_either(self):
        # $WIVER proves the instrument is powered, but says nothing about the
        # wind. A stream of nothing but identification sentences must still age
        # into UNKNOWN.
        monitor = W.WindMonitor(self.config())
        now = feed(monitor, 30, speed=2.0)
        for _ in range(30):
            now += 1.0
            monitor.update(now, ver_datagram())
        assert monitor.ignored_count == 30
        assert monitor.rejected_count == 0      # not a fault, just not wind
        assert monitor.verdict(now)[0] == W.UNKNOWN


class TestNorthOffset:
    """The offset rotates the instrument's zero mark to true north, so it
    belongs only to a bearing reported as RELATIVE."""

    def test_applied_to_a_relative_bearing(self):
        config = WeatherConfig(wind_north_offset_deg=30.0,
                               wind_direction_min_speed_ms=1.0)
        monitor = W.WindMonitor(config)
        now = 0.0
        for _ in range(20):
            monitor.update(now, mwv_datagram(angle=100.0, speed=5.0,
                                             reference='R'))
            now += 1.0
        assert monitor.direction_deg(now) == pytest.approx(130.0, abs=0.1)

    def test_not_applied_to_a_true_bearing(self):
        # If the OMC-140 is ever configured to emit T, adding the offset would
        # introduce exactly the error it exists to remove.
        config = WeatherConfig(wind_north_offset_deg=30.0,
                               wind_direction_min_speed_ms=1.0)
        monitor = W.WindMonitor(config)
        now = 0.0
        for _ in range(20):
            monitor.update(now, mwv_datagram(angle=100.0, speed=5.0,
                                             reference='T'))
            now += 1.0
        assert monitor.direction_deg(now) == pytest.approx(100.0, abs=0.1)


class TestAgainstTheRealCapture:
    """Fifteen minutes of real Greenhill traffic, replayed.

    The parser this replaced would have rejected every one of these datagrams:
    it split on '*' first, which cuts at the TAG block's own checksum and
    leaves two fields. Wind would have been permanently UNKNOWN, and the site
    permanently unsafe.
    """

    def capture(self):
        import json
        import os
        path = os.path.join(os.path.dirname(os.path.dirname(
            os.path.abspath(__file__))), 'greenhill-capture.jsonl')
        if not os.path.exists(path):
            pytest.skip('capture file not present')
        with open(path, encoding='utf-8') as handle:
            for line in handle:
                record = json.loads(line)
                if record.get('stream') == 'wind' and 'text' in record:
                    yield record

    def test_every_wind_datagram_is_understood(self):
        accepted = ignored = rejected = 0
        for record in self.capture():
            try:
                W.parse_sentence(record['text'], 100.0)
                accepted += 1
            except W.NoWindSentence:
                ignored += 1
            except W.WindParseError:
                rejected += 1
        assert rejected == 0
        assert ignored == 1                 # one $WIVER identification
        assert accepted > 800

    def test_the_readings_are_plausible_weather(self):
        speeds = []
        for record in self.capture():
            try:
                speeds.append(W.parse_sentence(record['text'], 100.0)[1])
            except (W.NoWindSentence, W.WindParseError):
                pass
        # It was a windy quarter of an hour: about 6.4 m/s mean, gusting to 12.6.
        assert 1.0 < min(speeds) < 5.0
        assert 8.0 < max(speeds) < 20.0
