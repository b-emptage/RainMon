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


def sentence(direction, speed):
    """A sentence in the shape the legacy display expects: direction at field
    2, speed at field 4."""
    return '$WIMWV,x,{},R,{},M,A'.format(direction, speed)


def feed(monitor, seconds, speed, direction=90.0, start=0.0, rate=1.0):
    """Feed `seconds` of samples at `rate` Hz. Returns the final time."""
    now = start
    for _ in range(int(seconds * rate)):
        monitor.update(now, sentence(direction, speed))
        now += 1.0 / rate
    return now


class TestParsing:
    def test_reads_the_configured_fields(self):
        assert W.parse_sentence(sentence(123.4, 5.6), 2, 4, 100.0) == (123.4, 5.6)

    def test_ignores_an_nmea_checksum(self):
        assert W.parse_sentence('$WIMWV,x,90,R,3.0,M,A*4F', 2, 4, 100.0)[1] == 3.0

    def test_field_positions_are_configurable(self):
        # The anemometer's format is unconfirmed. If the recorder shows the
        # values live elsewhere, that must be a config change, not a release.
        assert W.parse_sentence('$X,45.0,7.5', 1, 2, 100.0) == (45.0, 7.5)

    @pytest.mark.parametrize('text', [
        '$WIMWV,x,notanumber,R,3.0,M',
        '$WIMWV,x,90,R,notanumber,M',
        'too,short',
        '',
    ])
    def test_rejects_unusable_sentences(self, text):
        with pytest.raises(W.WindParseError):
            W.parse_sentence(text, 2, 4, 100.0)

    @pytest.mark.parametrize('direction', [-1, 400])
    def test_rejects_impossible_directions(self, direction):
        with pytest.raises(W.WindParseError, match='0-360'):
            W.parse_sentence(sentence(direction, 3.0), 2, 4, 100.0)

    def test_rejects_a_negative_speed(self):
        with pytest.raises(W.WindParseError, match='negative'):
            W.parse_sentence(sentence(90, -1.0), 2, 4, 100.0)

    def test_rejects_an_implausible_speed(self):
        # 360 km/h is a misparsed field, not weather. It matters because such a
        # reading would trip the gust threshold and close the dome on nothing
        # more than a change in the sentence format.
        with pytest.raises(W.WindParseError, match='plausible'):
            W.parse_sentence(sentence(90, 150.0), 2, 4, 100.0)


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

    def test_garbage_does_not_count_as_liveness(self):
        # A sender emitting nonsense is a dead source. If rejected sentences
        # refreshed the staleness clock, a stuck instrument would look healthy
        # forever.
        monitor = W.WindMonitor(self.config())
        now = feed(monitor, 30, speed=2.0)
        for _ in range(30):
            now += 1.0
            monitor.update(now, 'garbage,,,')
        assert monitor.rejected_count == 30
        assert monitor.verdict(now)[0] == W.UNKNOWN
