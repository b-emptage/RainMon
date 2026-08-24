# -*- coding: utf-8 -*-
# -----------------------------------------------------------------------------
# weatherdevice.py - the bridge between the safety core and ASCOM semantics.
#
# No HTTP here, and no sockets beyond the multicast reader it owns. Everything
# that decides anything lives in greenhill/core; this module only translates it
# into what ObservingConditions and SafetyMonitor are defined to return.
# -----------------------------------------------------------------------------
"""One weather service, published as two Alpaca devices.

`SafetyMonitor.IsSafe` and every `ObservingConditions` property are views onto
a single `SafetyEvaluator`. They cannot disagree, which is the point: the two
dome-close routes must never be able to form different opinions.

THE MONITORING THREAD RUNS FROM SERVER STARTUP AND IS NOT TIED TO `Connected`.

That is a deliberate departure from the dome server, where disconnecting
de-energises the motors. The difference is that this device owns no exclusive
hardware -- multicast can be received by any number of listeners -- and that
stopping safety monitoring because an ASCOM client went away would be absurd:
the direct dome-close route depends on this thread, and it has no client.

`Connected` therefore defaults to TRUE and is a courtesy flag over the ASCOM
surface only. It has to default that way in any case: Arcsecond polls
`Connected` but never sets it, so a device waiting to be connected would be
recorded as "Device disconnected" forever, with no values ever collected.
"""

import threading
import time
from logging import Logger
from typing import Dict, List, Optional

from greenhill import rain_protocol
from greenhill.core.dome_client import AlpacaDomeClient, DomeCloser
from greenhill.core.receiver import MulticastReader
from greenhill.core.safety import SafetyEvaluator

# AlpycaDevice's "Exception" classes are NOT Python exceptions -- they are
# plain response-payload objects, constructed and handed to PropertyResponse.
# Raising or catching one is a TypeError. So this module, which knows nothing
# about HTTP anyway, raises ordinary Python exceptions and the responders in
# observingconditions.py translate them into Alpaca error numbers.
class WeatherDeviceError(Exception):
    """Base for the faults this bridge reports upward."""


class SensorNotFitted(WeatherDeviceError):
    """A sensor the interface defines and this site does not have."""


class UnknownSensor(WeatherDeviceError):
    """A name that is not an ObservingConditions sensor at all."""


class ValueNotAvailable(WeatherDeviceError):
    """A sensor we do have, with no reading yet."""


class SettingNotAdjustable(WeatherDeviceError):
    """A setting fixed by the safety configuration."""


# Sensors this installation actually has, and what they are really measuring.
# The descriptions are published verbatim through SensorDescription, because a
# client that is about to make a decision on RainRate deserves to know it is
# reading an encoded state rather than a gauge.
SENSOR_DESCRIPTIONS = {
    'WindSpeed': 'Ultrasonic anemometer, 60 second mean, metres per second.',
    'WindGust': ('Ultrasonic anemometer, strongest 3 second average in the '
                 'last 2 minutes, metres per second.'),
    'WindDirection': ('Ultrasonic anemometer, 60 second circular mean of the '
                      'direction the wind is coming FROM, degrees. Samples '
                      'below 1 m/s are excluded because the reading is '
                      'meaningless there. A site-specific north offset is '
                      'applied and has NOT yet been verified against a known '
                      'reference.'),
    'Temperature': ('Mean of the rain detector thermistors, degrees C. These '
                    'are detector-body readings on actively heated sensors, '
                    'NOT a meteorological air temperature.'),
    'RainRate': ('DERIVED STATE, NOT A MEASUREMENT. There is no rain gauge at '
                 'this site, only wet/dry sections on heated detectors. Zero '
                 'means dry; a nominal non-zero value is published while the '
                 'rain safety latch holds, and also whenever the detectors '
                 'cannot be read at all, so that a sensor which has gone blind '
                 'is never mistaken for clear weather.'),
}

# Present in the interface, absent from this installation. Reading one throws
# PropertyNotImplemented rather than returning a plausible zero.
UNIMPLEMENTED_SENSORS = ('CloudCover', 'DewPoint', 'Humidity', 'Pressure',
                         'SkyBrightness', 'SkyQuality', 'SkyTemperature',
                         'StarFWHM')

# Which stream each sensor's freshness comes from.
_WIND_SENSORS = ('WindSpeed', 'WindGust', 'WindDirection')
_RAIN_SENSORS = ('RainRate', 'Temperature')

# How often the loop re-evaluates when no packets arrive. A verdict that only
# changed when data arrived could never notice data NOT arriving, which is the
# failure this whole system exists to catch.
EVALUATION_INTERVAL_S = 1.0


class GreenhillWeather(object):
    def __init__(self, weather_config, multicast_config, logger):
        # type: (object, Dict[str, object], Logger) -> None
        self._config = weather_config
        self._multicast = multicast_config
        self._logger = logger

        self._evaluator = SafetyEvaluator(weather_config)
        self._reader = None         # type: Optional[MulticastReader]
        self._thread = None         # type: Optional[threading.Thread]
        self._stop = threading.Event()
        self._lock = threading.Lock()
        self._state = None
        self._started_monotonic = None  # type: Optional[float]
        self._closer = None             # type: Optional[DomeCloser]

        # See the module docstring: true from the start, and never a gate on
        # the monitoring itself.
        self._connected = True

    # -- lifecycle ------------------------------------------------------------

    def start(self, simulate=False):
        # type: (bool) -> None
        """Join both streams and begin evaluating. Called once at startup.

        A stream that cannot be joined is logged loudly but is NOT fatal. It
        will age into UNKNOWN, which is unsafe -- and an unsafe service is far
        better than no service, which is what refusing to start would leave the
        observatory with.
        """
        if self._thread is not None:
            return

        self._started_monotonic = time.monotonic()
        self._stop.clear()

        if simulate:
            # Synthetic dry, calm weather fed straight into the evaluator, with
            # no sockets involved. This is what ASCOM Conform runs against: it
            # exercises the entire HTTP surface and the real safety core, while
            # being deterministic and needing neither hardware nor a network
            # that will forward multicast.
            self._thread = threading.Thread(target=self._run_simulated,
                                            name='weather-sim', daemon=True)
            self._thread.start()
            self._logger.warning(
                '==SIMULATED== feeding synthetic dry, calm weather. This '
                'server is NOT watching the sky.')
            # Route 1 is deliberately NOT armed here, whatever the config says.
            # A server running on invented weather must never be in a position
            # to command a real roof -- and simulate.py is what Conform and
            # bench work run against, on machines that may well be able to
            # reach the dome.
            self._logger.warning(
                '==DOME CLOSE NOT ARMED== simulated mode never commands the '
                'dome, regardless of dome_close_enabled.')
            return

        self._reader = MulticastReader()
        self._reader.add('wind', self._multicast['wind_group'],
                         self._multicast['wind_port'], self._multicast['interface'])
        self._reader.add('rain', self._multicast['rain_group'],
                         self._multicast['rain_port'], self._multicast['interface'])
        for name, error in self._reader.failures.items():
            self._logger.error(
                '==STREAM UNAVAILABLE== could not join the %s stream: %s. The '
                'service will report UNSAFE until it is receiving that data. '
                'Check the group address, the interface and the firewall.',
                name, error)

        self._thread = threading.Thread(target=self._run, name='weather',
                                        daemon=True)
        self._thread.start()
        self._logger.info('==WEATHER== monitoring started. %s',
                          self._config.describe_wind_thresholds())
        self._start_dome_closer()

    def _start_dome_closer(self):
        # type: () -> None
        """Arm route 1, if it has been turned on.

        Off unless the configuration says otherwise, and the state is logged
        either way at a level nobody can miss. This is the only thing in the
        package that commands a roof; a service that started driving one the
        moment it was installed -- before anyone had checked the address or
        watched a dry run -- would be the wrong kind of helpful. But a
        forgotten `false` would leave the observatory a route short without
        anything saying so, which is why the disarmed case shouts too.
        """
        if not self._config.dome_close_enabled:
            self._logger.warning(
                '==DOME CLOSE NOT ARMED== dome_close_enabled is false, so this '
                'service will NOT close the dome itself. Arcsecond is the only '
                'route, and it runs about 90 seconds behind.')
            return
        if not self._config.dome_address:
            self._logger.error(
                '==DOME CLOSE MISCONFIGURED== dome_close_enabled is true but '
                'dome_address is empty. Route 1 is NOT armed.')
            return

        client = AlpacaDomeClient(self._config.dome_address,
                                  self._config.dome_device_number,
                                  self._config.dome_http_timeout_s)
        # Passes a callable rather than a value: the closer runs on its own
        # thread and must read the verdict as it stands each time it looks, not
        # as it stood when it started.
        self._closer = DomeCloser(client, self._safety_verdict, self._logger,
                                  self._config, reasons=lambda: self.unsafe_reasons)
        self._closer.start()

    def _safety_verdict(self):
        # type: () -> Optional[bool]
        """True, False, or None when nothing has been evaluated yet.

        None matters: it is not the same as unsafe. The weather service already
        answers False the moment it has looked at anything at all, so None can
        only mean the closer is running ahead of the first evaluation -- and
        commanding a roof on no information is not a safe default, it is an
        arbitrary one.
        """
        state = self._snapshot()
        return None if state is None else bool(state.is_safe)

    def stop(self):
        # type: () -> None
        if self._closer is not None:
            self._closer.stop()
            self._closer = None
        self._stop.set()
        if self._thread is not None:
            self._thread.join(timeout=5.0)
            self._thread = None
        if self._reader is not None:
            self._reader.close()
            self._reader = None

    def _run(self):
        last_evaluated = 0.0
        while not self._stop.is_set():
            try:
                for stream, payload in self._reader.poll(0.5):
                    now = self._now()
                    if stream == 'rain':
                        try:
                            self._evaluator.rain.update(
                                now, rain_protocol.parse(payload))
                        except rain_protocol.ProtocolError as exc:
                            self._logger.debug('bad rain packet: %s', exc)
                    else:
                        try:
                            self._evaluator.wind.update(now, payload.decode('ascii'))
                        except UnicodeDecodeError:
                            pass

                now = self._now()
                if now - last_evaluated >= EVALUATION_INTERVAL_S:
                    last_evaluated = now
                    state = self._evaluator.update(now)
                    with self._lock:
                        self._state = state
                    if self._evaluator.changed(state):
                        self._log_transition(state)
            except Exception as exc:               # noqa: BLE001
                # The safety thread must not be able to die. Anything unhandled
                # is logged and the loop continues; the alternative is a server
                # that answers cheerfully with a frozen verdict.
                self._logger.exception('weather loop error: %s', exc)
                time.sleep(1.0)

    def _run_simulated(self):
        """Dry and calm, once a second, for Conform and for bench work."""
        sequence = 0
        while not self._stop.is_set():
            now = self._now()
            sequence += 1
            self._evaluator.rain.update(now, rain_protocol.RainPacket(
                sequence, 'simulated',
                [rain_protocol.Detector(name, 'D', 12.0)
                 for name in ('H127', 'H50', 'ACC')],
                True))
            self._evaluator.wind.update(now, '$WIMWV,x,90.0,R,2.0,M,A')
            state = self._evaluator.update(now)
            with self._lock:
                self._state = state
            if self._evaluator.changed(state):
                self._log_transition(state)
            self._stop.wait(0.25)

    def _log_transition(self, state):
        if state.is_safe:
            self._logger.info('==SAFE== conditions are within limits')
        else:
            self._logger.warning('==UNSAFE== %s', '; '.join(state.reasons))

    def _now(self):
        # type: () -> float
        return time.monotonic() - self._started_monotonic

    # -- ASCOM common ---------------------------------------------------------

    @property
    def connected(self):
        # type: () -> bool
        return self._connected

    def connect(self):
        # type: () -> None
        self._connected = True

    def disconnect(self):
        # type: () -> None
        """Drops the ASCOM surface only. Monitoring continues -- see the module
        docstring."""
        self._connected = False

    # -- state ----------------------------------------------------------------

    def _snapshot(self):
        with self._lock:
            return self._state

    @property
    def is_safe(self):
        # type: () -> bool
        """SafetyMonitor.IsSafe.

        False before the first evaluation. A safety monitor that has not yet
        looked must never answer "safe".
        """
        state = self._snapshot()
        return bool(state.is_safe) if state is not None else False

    @property
    def unsafe_reasons(self):
        # type: () -> List[str]
        state = self._snapshot()
        return list(state.reasons) if state is not None else ['starting up']

    def observing_conditions(self):
        # type: () -> Dict[str, Optional[float]]
        state = self._snapshot()
        if state is None:
            return {}
        return {
            'WindSpeed': state.wind_speed_ms,
            'WindGust': state.wind_gust_ms,
            'WindDirection': state.wind_direction_deg,
            'Temperature': state.temperature_c,
            'RainRate': state.rain_rate_mm_h,
        }

    def sensor_value(self, name):
        # type: (str) -> float
        """One ObservingConditions sensor.

        Raises ValueNotSet (through the caller) when the value is not yet
        known. Returning zero would be a reading, and a client cannot tell a
        fabricated zero from a real calm.
        """
        value = self.observing_conditions().get(name)
        if value is None:
            raise ValueNotAvailable('{} is not available yet'.format(name))
        return float(value)

    @property
    def average_period_hours(self):
        # type: () -> float
        return self._config.wind_mean_window_s / 3600.0

    def set_average_period(self, hours):
        # type: (float) -> None
        """The averaging window is fixed by the safety rules, so this accepts
        only the value it already reports.

        Rejected rather than silently ignored: a client that asked for a
        different averaging period and was told "fine" would go on to read
        numbers that are not what it asked for.
        """
        if abs(hours - self.average_period_hours) > 1e-6:
            raise SettingNotAdjustable(
                'AveragePeriod is fixed at {:.6f} h ({:.0f} s) by the safety '
                'configuration and cannot be changed through Alpaca.'.format(
                    self.average_period_hours, self._config.wind_mean_window_s))

    def time_since_last_update(self, sensor_name):
        # type: (str) -> float
        """Seconds since the named sensor last had a value.

        The real freshness channel for a polling client, and the reason a
        stalled sensor is visible from outside this process at all. An empty
        name gives the most recent update of any sensor.
        """
        state = self._snapshot()
        if state is None:
            raise ValueNotAvailable('no data yet')

        name = (sensor_name or '').strip()
        if not name:
            ages = [a for a in (state.rain_age_s, state.wind_age_s) if a is not None]
            if not ages:
                raise ValueNotAvailable('no data yet')
            return min(ages)

        if name in UNIMPLEMENTED_SENSORS:
            raise SensorNotFitted('{} is not fitted at this site'.format(name))
        if name in _WIND_SENSORS:
            age = state.wind_age_s
        elif name in _RAIN_SENSORS:
            age = state.rain_age_s
        else:
            raise UnknownSensor(
                '{} is not an ObservingConditions sensor name'.format(name))
        if age is None:
            raise ValueNotAvailable('{} has never been updated'.format(name))
        return age

    def sensor_description(self, sensor_name):
        # type: (str) -> str
        name = (sensor_name or '').strip()
        if name in SENSOR_DESCRIPTIONS:
            return SENSOR_DESCRIPTIONS[name]
        if name in UNIMPLEMENTED_SENSORS:
            raise SensorNotFitted('{} is not fitted at this site'.format(name))
        raise UnknownSensor(
            '{} is not an ObservingConditions sensor name'.format(name))

    def refresh(self):
        # type: () -> None
        """ASCOM Refresh().

        A no-op, honestly. The sensors PUSH to this service about once a
        second; there is no hardware here to interrogate on demand, and
        pretending otherwise would only add a way for a client to think it had
        forced an update it had not.
        """
        return

    # -- diagnostics ----------------------------------------------------------

    def diagnostics(self):
        # type: () -> Dict[str, object]
        """Everything the collapse into IsSafe throws away.

        ASCOM has one boolean and this site has four independent reasons to
        close, so a bare `false` cannot say which. Published through a vendor
        Action so an operator, or a UI, can see WHY without reading the log.
        """
        state = self._snapshot()
        if state is None:
            return {'ready': False, 'reasons': ['starting up']}
        return {
            'ready': True,
            'isSafe': state.is_safe,
            'reasons': list(state.reasons),
            'conditions': dict(state.conditions),
            'detectors': dict(state.detector_states),
            'rainAgeSeconds': state.rain_age_s,
            'windAgeSeconds': state.wind_age_s,
            'windSampleCount': self._evaluator.wind.sample_count,
            'windRejectedCount': self._evaluator.wind.rejected_count,
            'rainOutOfOrderCount': self._evaluator.rain.out_of_order_count,
            'rainBridgeRestarts': self._evaluator.rain.restart_count,
            'wetSections': self._evaluator.rain.wet_sections,
            'observingDetectors': self._evaluator.rain.observing_count,
            # Route 1's own state, so an operator can see from the ASCOM
            # surface whether the direct close is armed and what it last did.
            'domeClose': (self._closer.status() if self._closer is not None
                          else {'state': 'not armed'}),
        }
