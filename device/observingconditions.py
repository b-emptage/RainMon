# -*- coding: utf-8 -*-
# -----------------------------------------------------------------------------
# observingconditions.py - Alpaca API responders for the Greenhill weather.
#
# Telemetry only. The GO/NOGO lives on the SafetyMonitor device, because ASCOM
# ObservingConditions has no safety property and faking one into a measurement
# would make both devices lie.
# -----------------------------------------------------------------------------
from datetime import datetime, timezone

from falcon import Request, Response, before

import alpaca_common
from alpaca_common import (BaseAbsentSensor, BaseAction, BaseCommandBlind,
                           BaseCommandBool, BaseCommandString, BaseConnect,
                           BaseConnected, BaseConnecting, BaseDescription,
                           BaseDisconnect, BaseDriverInfo, BaseDriverVersion,
                           BaseInterfaceVersion, BaseName, BaseSensorValue,
                           BaseSupportedActions, maxdev, not_connected)
from exceptions import (DriverException, InvalidValueException,
                        NotImplementedException, ValueNotSetException)
from weatherdevice import (SensorNotFitted, SettingNotAdjustable,
                           UnknownSensor, ValueNotAvailable)
from shr import (MethodResponse, PreProcessRequest, PropertyResponse,
                 StateValue, get_request_field)


class ObservingConditionsMetadata:
    """Metadata describing the ObservingConditions device."""
    Name = 'Greenhill Weather'
    Version = '0.1.0'
    Description = 'Greenhill Observatory rain detectors and ultrasonic anemometer'
    DeviceType = 'ObservingConditions'
    # Stable across restarts: Arcsecond identifies devices by UniqueID and
    # silently drops any device that does not present one.
    DeviceID = '5c7e2b90-1f43-4a86-b0d9-64e83a1c7f52'
    Info = ('Greenhill Observatory weather telemetry\n'
            'Wind from the site ultrasonic anemometer; rain state from the\n'
            'heated detector array. Speeds are metres per second.\n'
            'RainRate is a DERIVED STATE, not a gauge reading -- see\n'
            'SensorDescription("RainRate") before using it.')
    MaxDeviceNumber = maxdev
    InterfaceVersion = 2        # IObservingConditionsV2 (ASCOM Platform 7)


# -- common members -----------------------------------------------------------

class action(BaseAction):
    metadata = ObservingConditionsMetadata


class commandblind(BaseCommandBlind):
    pass


class commandbool(BaseCommandBool):
    pass


class commandstring(BaseCommandString):
    pass


class connect(BaseConnect):
    pass


class disconnect(BaseDisconnect):
    pass


class connecting(BaseConnecting):
    pass


class connected(BaseConnected):
    pass


class description(BaseDescription):
    metadata = ObservingConditionsMetadata


class driverinfo(BaseDriverInfo):
    metadata = ObservingConditionsMetadata


class driverversion(BaseDriverVersion):
    metadata = ObservingConditionsMetadata


class interfaceversion(BaseInterfaceVersion):
    metadata = ObservingConditionsMetadata


class name(BaseName):
    metadata = ObservingConditionsMetadata


class supportedactions(BaseSupportedActions):
    pass


# -- sensors this site has ----------------------------------------------------

class windspeed(BaseSensorValue):
    sensor = 'WindSpeed'


class windgust(BaseSensorValue):
    sensor = 'WindGust'


class winddirection(BaseSensorValue):
    sensor = 'WindDirection'


class temperature(BaseSensorValue):
    sensor = 'Temperature'


class rainrate(BaseSensorValue):
    sensor = 'RainRate'


# -- sensors it does not -------------------------------------------------------
# Each throws PropertyNotImplemented. Arcsecond catches that, caches it and
# stops asking; anything else records the truth rather than a fabricated
# reading.

class cloudcover(BaseAbsentSensor):
    sensor = 'CloudCover'


class dewpoint(BaseAbsentSensor):
    sensor = 'DewPoint'


class humidity(BaseAbsentSensor):
    sensor = 'Humidity'


class pressure(BaseAbsentSensor):
    sensor = 'Pressure'


class skybrightness(BaseAbsentSensor):
    sensor = 'SkyBrightness'


class skyquality(BaseAbsentSensor):
    sensor = 'SkyQuality'


class skytemperature(BaseAbsentSensor):
    sensor = 'SkyTemperature'


class starfwhm(BaseAbsentSensor):
    sensor = 'StarFWHM'


# -- interface specifics -------------------------------------------------------

@before(PreProcessRequest(maxdev))
class averageperiod:
    """The averaging window, in hours.

    Reported honestly rather than as the conventional 0.0: these values ARE
    averaged, over the same 60 seconds the wind safety threshold uses, and a
    client told "no averaging" would draw the wrong conclusion from a smooth
    trace.
    """

    def on_get(self, req: Request, resp: Response, devnum: int):
        if not_connected(req, resp):
            return
        resp.text = PropertyResponse(
            alpaca_common.weather_dev.average_period_hours, req).json

    def on_put(self, req: Request, resp: Response, devnum: int):
        if not_connected(req, resp):
            return
        raw = get_request_field('AveragePeriod', req)       # 400 if missing
        try:
            hours = float(raw)
        except (TypeError, ValueError):
            resp.text = MethodResponse(
                req, InvalidValueException(
                    'AveragePeriod {!r} is not a number'.format(raw))).json
            return
        try:
            alpaca_common.weather_dev.set_average_period(hours)
            resp.text = MethodResponse(req).json
        except SettingNotAdjustable as ex:
            resp.text = MethodResponse(req, InvalidValueException(str(ex))).json
        except Exception as ex:
            resp.text = MethodResponse(
                req, DriverException(0x500, 'AveragePeriod failed', ex)).json


@before(PreProcessRequest(maxdev))
class sensordescription:
    def on_get(self, req: Request, resp: Response, devnum: int):
        if not_connected(req, resp):
            return
        sensor = get_request_field('SensorName', req)       # 400 if missing
        try:
            resp.text = PropertyResponse(
                alpaca_common.weather_dev.sensor_description(sensor), req).json
        except SensorNotFitted as ex:
            resp.text = PropertyResponse(
                None, req, NotImplementedException(str(ex))).json
        except UnknownSensor as ex:
            resp.text = PropertyResponse(
                None, req, InvalidValueException(str(ex))).json
        except Exception as ex:
            resp.text = PropertyResponse(
                None, req,
                DriverException(0x500, 'SensorDescription failed', ex)).json


@before(PreProcessRequest(maxdev))
class timesincelastupdate:
    """Seconds since a sensor last had a value.

    The freshness channel a polling client actually has. It is how a stalled
    sensor becomes visible from outside this process -- and the reason
    Arcsecond could, in principle, notice that a reading has stopped moving
    rather than treating the last one as current forever.
    """

    def on_get(self, req: Request, resp: Response, devnum: int):
        if not_connected(req, resp):
            return
        sensor = get_request_field('SensorName', req, False, '')
        try:
            resp.text = PropertyResponse(
                alpaca_common.weather_dev.time_since_last_update(sensor), req).json
        except SensorNotFitted as ex:
            resp.text = PropertyResponse(
                None, req, NotImplementedException(str(ex))).json
        except UnknownSensor as ex:
            resp.text = PropertyResponse(
                None, req, InvalidValueException(str(ex))).json
        except ValueNotAvailable as ex:
            resp.text = PropertyResponse(
                None, req, ValueNotSetException(str(ex))).json
        except Exception as ex:
            resp.text = PropertyResponse(
                None, req,
                DriverException(0x500, 'TimeSinceLastUpdate failed', ex)).json


@before(PreProcessRequest(maxdev))
class refresh:
    """A no-op, honestly.

    The sensors push to this service about once a second; there is no hardware
    here to interrogate on demand. Succeeding without doing anything is better
    than failing -- clients call Refresh() as a matter of course -- but it must
    not be mistaken for a forced read.
    """

    def on_put(self, req: Request, resp: Response, devnum: int):
        if not_connected(req, resp):
            return
        try:
            alpaca_common.weather_dev.refresh()
            resp.text = MethodResponse(req).json
        except Exception as ex:
            resp.text = MethodResponse(
                req, DriverException(0x500, 'Refresh failed', ex)).json


@before(PreProcessRequest(maxdev))
class devicestate:
    """Platform 7 bulk state read: everything a client polls, in one round trip.

    Only sensors that currently HAVE a value appear. The spec asks for the
    operational state, and a member with no reading has none to report.
    """

    def on_get(self, req: Request, resp: Response, devnum: int):
        if not_connected(req, resp):
            return
        try:
            readings = alpaca_common.weather_dev.observing_conditions()
            state = [StateValue(sensor, value)
                     for sensor, value in sorted(readings.items())
                     if value is not None]
            state.append(StateValue('TimeStamp', datetime.now(timezone.utc)
                                    .isoformat(timespec='milliseconds')
                                    .replace('+00:00', 'Z')))
            resp.text = PropertyResponse(state, req).json
        except Exception as ex:
            resp.text = PropertyResponse(
                None, req,
                DriverException(0x500, 'ObservingConditions.Devicestate failed',
                                ex)).json
