# -*- coding: utf-8 -*-
"""Tunable settings for the weather service.

Every number the observatory might want to change lives here, so that changing
one is an edit to a config file on the Windows 11 box rather than a visit to
the Windows 7 machine.

Units are SI throughout -- metres per second, seconds, degrees, millimetres per
hour. The operators think in km/h and the conversions are noted, but nothing in
the code carries a km/h value: ASCOM ObservingConditions is defined in m/s and
Arcsecond stores m/s, so a second unit anywhere is a threshold waiting to be
wrong by a factor of 3.6.
"""

from typing import Optional

KMH_PER_MS = 3.6


class WeatherConfig(object):
    """Thresholds and timings. Defaults are the agreed Greenhill values."""

    def __init__(self,
                 # --- wind ---------------------------------------------------
                 wind_sustained_max_ms=20.0 / KMH_PER_MS,     # 5.56 m/s = 20 km/h
                 wind_gust_max_ms=30.0 / KMH_PER_MS,          # 8.33 m/s = 30 km/h
                 wind_mean_window_s=60.0,
                 wind_gust_window_s=120.0,
                 wind_gust_average_s=3.0,
                 wind_direction_min_speed_ms=1.0,
                 wind_north_offset_deg=30.0,
                 wind_direction_field=2,
                 wind_speed_field=4,
                 wind_max_plausible_ms=100.0,
                 wind_min_samples=10,
                 # --- rain ---------------------------------------------------
                 rain_wet_sections_trigger=2,
                 raindrop_window_s=10.0,
                 rain_min_observing_detectors=2,
                 rain_rate_when_wet_mm_h=1.0,
                 # --- staleness and latching ---------------------------------
                 rain_max_age_s=15.0,
                 wind_max_age_s=15.0,
                 settle_s=30.0,
                 rain_clear_s=600.0,
                 wind_clear_s=120.0,
                 # --- route 1: the direct dome close -------------------------
                 dome_close_enabled=False,
                 dome_address='',
                 dome_device_number=0,
                 dome_http_timeout_s=10.0,
                 dome_poll_interval_s=2.0,
                 dome_verify_timeout_s=45.0,
                 dome_retry_limit=3,
                 dome_escalated_retry_s=60.0):
        # Sustained wind is a 60 s mean; a gust is the strongest 3 s average in
        # the last two minutes, which is the ASCOM definition and also the
        # sensible one. An ultrasonic head is fast and noisy, and closing the
        # dome on a single spiky sample is how a safety system earns a reputation
        # for crying wolf.
        self.wind_sustained_max_ms = wind_sustained_max_ms
        self.wind_gust_max_ms = wind_gust_max_ms
        self.wind_mean_window_s = wind_mean_window_s
        self.wind_gust_window_s = wind_gust_window_s
        self.wind_gust_average_s = wind_gust_average_s

        # Below this speed an anemometer's direction reading is essentially
        # random, so those samples are left out of the direction average. The
        # legacy display used the same 1 m/s cut.
        self.wind_direction_min_speed_ms = wind_direction_min_speed_ms

        # PROVISIONAL. The legacy display adds 30 degrees to the raw direction,
        # with no note saying why -- presumably a mounting correction. It is
        # carried forward so behaviour does not change silently, but it must be
        # checked against a known reference before WindDirection is trusted.
        self.wind_north_offset_deg = wind_north_offset_deg

        # PROVISIONAL. The anemometer's sentence format is undocumented; these
        # are the positions the legacy display reads, and the recorder exists to
        # confirm them. Configurable rather than hard-coded so correcting them
        # is a config edit, not a release.
        self.wind_direction_field = wind_direction_field
        self.wind_speed_field = wind_speed_field
        self.wind_max_plausible_ms = wind_max_plausible_ms

        # Below this many samples the mean and gust are not meaningful, so wind
        # reports UNKNOWN -- which is unsafe -- rather than a confident average
        # of two readings.
        self.wind_min_samples = wind_min_samples

        # Sections, not detectors: each detector has two that trigger
        # independently, so one detector reporting W reaches this on its own.
        self.rain_wet_sections_trigger = rain_wet_sections_trigger

        # A single wet section that clears within this long is a real raindrop
        # evaporating off a heated sensor. One that persists longer is treated
        # as contamination -- droppings, an insect, a failed sensor.
        self.raindrop_window_s = raindrop_window_s

        # How many detectors must be reporting a usable wet/dry state for the
        # rain verdict to count. With three installed, requiring two means one
        # dead unit does not blind the observatory, while a single survivor is
        # not asked to speak for the whole site.
        self.rain_min_observing_detectors = rain_min_observing_detectors

        # There is no rain gauge here, only wet/dry sections. This is the
        # nominal rate published while the latch holds, so Arcsecond's
        # precipitation condition has something to compare against. It is an
        # encoding of a state, not a measurement, and says so in the device's
        # SensorDescription.
        self.rain_rate_when_wet_mm_h = rain_rate_when_wet_mm_h

        # Streams arrive about once a second. These allow a good many dropped
        # datagrams before declaring the source dead, because UDP loss is normal
        # and a nuisance NOGO every time a packet goes missing would be worse
        # than useless.
        self.rain_max_age_s = rain_max_age_s
        self.wind_max_age_s = wind_max_age_s

        # After a restart the service is unsafe until both streams have been
        # healthy this long. Short, because nothing is wet -- it only has to
        # gather enough samples to mean anything.
        self.settle_s = settle_s

        # After actual rain, everything must stay dry this long before the
        # latch releases. Comfortably above the 2-5 minute drying time of a
        # heated sensor, so the dome does not reopen onto sensors that are
        # merely still evaporating.
        self.rain_clear_s = rain_clear_s

        # Wind needs far less: nothing has to dry out, it only has to stop
        # gusting. Long enough not to flap around the threshold.
        self.wind_clear_s = wind_clear_s

        # OFF BY DEFAULT, and it must be turned on deliberately. This is the
        # only setting in the system that commands a roof, and a package that
        # started driving one the moment it was installed -- before anyone had
        # checked the address, watched a dry-run close, or confirmed the dome
        # was answering -- would be the wrong kind of helpful. The server logs
        # loudly at startup whichever state it is in, so a forgotten `false`
        # cannot pass unnoticed.
        self.dome_close_enabled = dome_close_enabled
        self.dome_address = dome_address                # 'host:11111'
        self.dome_device_number = dome_device_number
        self.dome_http_timeout_s = dome_http_timeout_s

        # How often to look at the dome while a close is in progress.
        self.dome_poll_interval_s = dome_poll_interval_s

        # How long to let a close run before issuing another. Must comfortably
        # exceed the dome's full travel time, or a slow close gets a second
        # command on top of it.
        self.dome_verify_timeout_s = dome_verify_timeout_s

        # Fast attempts before escalating. After this it keeps trying, but
        # slowly: an unreachable dome is not fixed by asking faster, and a log
        # filling at 1 Hz buries the one line that matters.
        self.dome_retry_limit = dome_retry_limit
        self.dome_escalated_retry_s = dome_escalated_retry_s

    @classmethod
    def from_mapping(cls, mapping):
        # type: (dict) -> WeatherConfig
        """Build from a config.toml section, ignoring keys we do not know.

        Unknown keys are ignored rather than rejected so that a config file
        written for a later version still starts an earlier one -- on a
        safety service, running with a stale setting beats not running.
        """
        known = cls().__dict__.keys()
        return cls(**{k: v for k, v in mapping.items() if k in known})

    def describe_wind_thresholds(self):
        # type: () -> str
        return 'sustained {:.2f} m/s ({:.0f} km/h), gust {:.2f} m/s ({:.0f} km/h)'.format(
            self.wind_sustained_max_ms, self.wind_sustained_max_ms * KMH_PER_MS,
            self.wind_gust_max_ms, self.wind_gust_max_ms * KMH_PER_MS)
