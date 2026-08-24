"""Shared test rig for the Alpaca layer.

The whole server is drivable with no hardware and no network: the weather
device is fed synthetic packets directly, and Falcon's test client drives the
REAL routing, so what these tests exercise is the same code an ASCOM client
reaches.
"""
import logging
import os
import sys

import pytest

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEVICE_DIR = os.path.join(REPO_ROOT, 'device')
for path in (REPO_ROOT, DEVICE_DIR):
    if path not in sys.path:
        sys.path.insert(0, path)


@pytest.fixture
def weather_app():
    """(TestClient, weather device) with the same wiring main() performs.

    Every logger the server hands around must be set. shr.py logs each request
    on the way in, so a missing logger there turns every response into a 500 --
    which is exactly how the first run of this rig failed.
    """
    import falcon.testing as testing

    import alpaca_common
    import app
    import discovery
    import exceptions
    import log
    import management
    from shr import set_shr_logger

    logger = logging.getLogger('greenhill-weather-test')
    logger.addHandler(logging.NullHandler())
    log.logger = logger
    exceptions.logger = logger
    discovery.logger = logger
    alpaca_common.set_common_logger(logger)
    management.set_management_logger(logger)
    set_shr_logger(logger)

    device = app.start_weather_device(logger, simulate=True)
    try:
        yield testing.TestClient(app.build_app()), device
    finally:
        device.stop()
        app.weather_dev = None
        alpaca_common.set_weather_device(None)


@pytest.fixture
def client(weather_app):
    return weather_app[0]


@pytest.fixture
def device(weather_app):
    return weather_app[1]


def set_weather_state(device, is_safe=True, reasons=None, rain_rate=0.0,
                      wind_speed=3.0, wind_gust=4.0, wind_direction=120.0,
                      temperature=12.0, conditions=None, rain_age=1.0,
                      wind_age=1.0, detectors=None):
    """Pin the device's published state, so the ASCOM surface can be tested
    without waiting out real settle and latch periods.

    Reaches past the public API deliberately: these tests are about whether the
    HTTP layer faithfully reports what the safety core decided, and the core
    itself is covered by its own tests.
    """
    from greenhill.core.safety import SafetyState

    with device._lock:
        device._state = SafetyState(
            is_safe=is_safe,
            reasons=reasons if reasons is not None else [],
            rain_rate_mm_h=rain_rate,
            wind_speed_ms=wind_speed,
            wind_gust_ms=wind_gust,
            wind_direction_deg=wind_direction,
            temperature_c=temperature,
            conditions=conditions or {'rain': False, 'wind': False,
                                      'rain_data': False, 'wind_data': False},
            rain_age_s=rain_age,
            wind_age_s=wind_age,
            detector_states=detectors or {'H127': 'D', 'H50': 'D', 'ACC': 'D'})
    # The simulated feed would overwrite it on its next tick.
    device._stop.set()


# The real anemometer datagram, as captured at Greenhill: a proprietary UDP
# wrapper, an NMEA 0183 v4 TAG block, then the sentence. Tests build the real
# shape rather than a convenient one, because the shape IS most of what the
# parser has to cope with.
WRAPPER = 'UdPbC\x00\\s:D383P1,s:WI4383*23\\'


def nmea(body):
    """'$' + body + checksum, as the instrument sends it."""
    checksum = 0
    for char in body:
        checksum ^= ord(char)
    return '${}*{:02X}'.format(body, checksum)


def mwv_datagram(angle=90.0, speed=2.0, reference='R', units='M', status='A',
                 talker='II', wrapper=True, corrupt_checksum=False):
    """One MWV datagram. `speed` is in whatever `units` says, as on the wire."""
    sentence = nmea('{}MWV,{},{},{:06.2f},{},{}'.format(
        talker, angle, reference, speed, units, status))
    if corrupt_checksum:
        sentence = sentence[:-2] + ('00' if not sentence.endswith('00') else 'FF')
    return (WRAPPER if wrapper else '') + sentence


def ver_datagram():
    """The identification sentence the OMC-140 emits every few minutes. Carries
    no wind reading, and must not be mistaken for one."""
    return ('UdPbC\x00\\s:D383G0,s:WI4383*35\\'
            '$WIVER,1,1,WI,OBS,,OMC-14000002383,OMC-140,02.40B18,,0*46')
