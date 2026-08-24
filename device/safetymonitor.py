# -*- coding: utf-8 -*-
# -----------------------------------------------------------------------------
# safetymonitor.py - Alpaca API responders for the Greenhill SafetyMonitor.
#
# The whole interface is one boolean. All the judgement behind it lives in
# greenhill/core/safety.py; this module is only the HTTP surface.
# -----------------------------------------------------------------------------
from datetime import datetime, timezone

from falcon import Request, Response, before

from alpaca_common import (BaseAction, BaseCommandBlind, BaseCommandBool,
                           BaseCommandString, BaseConnect, BaseConnected,
                           BaseConnecting, BaseDescription, BaseDisconnect,
                           BaseDriverInfo, BaseDriverVersion,
                           BaseInterfaceVersion, BaseName,
                           BaseSupportedActions, maxdev, not_connected,
                           weather_dev)
import alpaca_common
from exceptions import DriverException
from shr import PreProcessRequest, PropertyResponse, StateValue


class SafetyMonitorMetadata:
    """Metadata describing the SafetyMonitor device."""
    Name = 'Greenhill Weather Safety'
    Version = '0.1.0'
    Description = 'Greenhill Observatory rain and wind safety monitor'
    DeviceType = 'SafetyMonitor'
    # Stable across restarts: Arcsecond identifies devices by UniqueID and
    # silently drops any device that does not present one.
    DeviceID = 'f3a91c27-58d4-4b6e-9a02-7c31e8d5b104'
    Info = ('Greenhill Observatory weather safety monitor\n'
            'Fuses the rain detector array and the ultrasonic anemometer\n'
            'Fails closed: absent or stale data reports UNSAFE\n'
            'Use the Greenhill:GetWeatherStatus action for the reason')
    MaxDeviceNumber = maxdev
    InterfaceVersion = 3        # ISafetyMonitorV3 (ASCOM Platform 7)


# -- common members -----------------------------------------------------------
# One-line subclasses, because app.py routes a class only when it is DEFINED in
# this module. The behaviour lives once, in alpaca_common.

class action(BaseAction):
    metadata = SafetyMonitorMetadata


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
    metadata = SafetyMonitorMetadata


class driverinfo(BaseDriverInfo):
    metadata = SafetyMonitorMetadata


class driverversion(BaseDriverVersion):
    metadata = SafetyMonitorMetadata


class interfaceversion(BaseInterfaceVersion):
    metadata = SafetyMonitorMetadata


class name(BaseName):
    metadata = SafetyMonitorMetadata


class supportedactions(BaseSupportedActions):
    pass


# -- the device itself --------------------------------------------------------

@before(PreProcessRequest(maxdev))
class issafe:
    """The verdict.

    False whenever the service cannot see: no data yet, either stream stale,
    the rain bridge's serial port down, or too few detectors reporting. That is
    the entire point of the rewrite -- in the software this replaces, a dead
    sensor produced a wet count of zero, which was indistinguishable from a
    clear night.

    IsSafe never throws. A safety monitor that answered an error where a client
    expected a boolean would, in most client implementations, be treated as
    "no opinion" and skipped -- so a failure here would read as permission to
    keep observing. It answers False and puts the reason in the log and in the
    Greenhill:GetWeatherStatus action.
    """

    def on_get(self, req: Request, resp: Response, devnum: int):
        if not_connected(req, resp):
            return
        try:
            resp.text = PropertyResponse(alpaca_common.weather_dev.is_safe, req).json
        except Exception as ex:
            alpaca_common.logger.exception('IsSafe failed: %s', ex)
            resp.text = PropertyResponse(False, req).json


@before(PreProcessRequest(maxdev))
class devicestate:
    """Platform 7 bulk state read."""

    def on_get(self, req: Request, resp: Response, devnum: int):
        if not_connected(req, resp):
            return
        try:
            state = [
                StateValue('IsSafe', alpaca_common.weather_dev.is_safe),
                StateValue('TimeStamp', datetime.now(timezone.utc)
                           .isoformat(timespec='milliseconds')
                           .replace('+00:00', 'Z')),
            ]
            resp.text = PropertyResponse(state, req).json
        except Exception as ex:
            resp.text = PropertyResponse(
                None, req,
                DriverException(0x500, 'SafetyMonitor.Devicestate failed', ex)).json
