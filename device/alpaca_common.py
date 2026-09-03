# -*- coding: utf-8 -*-
# -----------------------------------------------------------------------------
# alpaca_common.py - responders shared by both device types.
#
# This server publishes TWO devices, ObservingConditions and SafetyMonitor,
# which are two views onto one weather service. Every ASCOM device carries the
# same dozen common members, and app.py routes a class to a URI only when the
# class is DEFINED in that device's module -- so each module declares its own
# one-line subclass of the bases here. That keeps the logic in one place
# without duplicating fifteen responder bodies twice over.
# -----------------------------------------------------------------------------
import json
from logging import Logger

from falcon import Request, Response, before

from exceptions import (ActionNotImplementedException, DriverException,
                        InvalidValueException, NotConnectedException,
                        NotImplementedException, ValueNotSetException)
from shr import MethodResponse, PreProcessRequest, PropertyResponse, get_request_field, to_bool
from weatherdevice import ValueNotAvailable

logger = None                   # type: Logger

# Single instance of each device type.
maxdev = 0

# The one weather service both device types read. Set by app.py at startup.
weather_dev = None

ACTION_GET_STATUS = 'Greenhill:GetWeatherStatus'
SUPPORTED_ACTIONS = [ACTION_GET_STATUS]


def set_common_logger(lgr):
    global logger
    logger = lgr


def set_weather_device(device):
    global weather_dev
    weather_dev = device


def not_connected(req, resp):
    # type: (Request, Response) -> bool
    """Emit the standard NotConnected response; report whether it was emitted.

    The response type must match the HTTP method -- a PUT comes back as a
    MethodResponse, not a PropertyResponse.
    """
    if weather_dev is None or not weather_dev.connected:
        err = NotConnectedException()
        if req.method == 'PUT':
            resp.text = MethodResponse(req, err).json
        else:
            resp.text = PropertyResponse(None, req, err).json
        return True
    return False


# --------------------------------------------------------------------------
# Common members. Each device module subclasses these and sets `metadata`.
# --------------------------------------------------------------------------

@before(PreProcessRequest(maxdev))
class BaseAction:
    metadata = None

    def on_put(self, req: Request, resp: Response, devnum: int):
        if not_connected(req, resp):
            return
        name = get_request_field('Action', req)         # 400 if missing
        if name != ACTION_GET_STATUS:
            resp.text = MethodResponse(
                req, ActionNotImplementedException(
                    'Action "{}" is not supported'.format(name))).json
            return
        try:
            # Everything the collapse into a single boolean throws away: which
            # of the four conditions is holding, how each detector reads, and
            # how old both streams are. ASCOM has one bit; the operator needs
            # the reason.
            resp.text = MethodResponse(
                req, value=json.dumps(weather_dev.diagnostics())).json
        except Exception as ex:
            resp.text = MethodResponse(
                req, DriverException(0x500, 'Action failed', ex)).json


@before(PreProcessRequest(maxdev))
class BaseCommandBlind:
    def on_put(self, req: Request, resp: Response, devnum: int):
        resp.text = MethodResponse(req, NotImplementedException()).json


@before(PreProcessRequest(maxdev))
class BaseCommandBool:
    def on_put(self, req: Request, resp: Response, devnum: int):
        resp.text = MethodResponse(req, NotImplementedException()).json


@before(PreProcessRequest(maxdev))
class BaseCommandString:
    def on_put(self, req: Request, resp: Response, devnum: int):
        resp.text = MethodResponse(req, NotImplementedException()).json


@before(PreProcessRequest(maxdev))
class BaseConnect:
    """Platform 7 asynchronous connect. Nothing here takes any time -- the
    monitoring thread is already running -- so Connecting is never True."""

    def on_put(self, req: Request, resp: Response, devnum: int):
        try:
            weather_dev.connect()
            resp.text = MethodResponse(req).json
        except Exception as ex:
            resp.text = MethodResponse(
                req, DriverException(0x500, 'Connect failed', ex)).json


@before(PreProcessRequest(maxdev))
class BaseDisconnect:
    def on_put(self, req: Request, resp: Response, devnum: int):
        try:
            weather_dev.disconnect()
            resp.text = MethodResponse(req).json
        except Exception as ex:
            resp.text = MethodResponse(
                req, DriverException(0x500, 'Disconnect failed', ex)).json


@before(PreProcessRequest(maxdev))
class BaseConnecting:
    def on_get(self, req: Request, resp: Response, devnum: int):
        resp.text = PropertyResponse(False, req).json


@before(PreProcessRequest(maxdev))
class BaseConnected:
    """Defaults to TRUE, and disconnecting does not stop the monitoring.

    Unlike the dome -- where disconnecting de-energises the motors -- this
    device owns no exclusive hardware, and Arcsecond polls Connected without
    ever setting it. A device that waited to be connected would be recorded as
    disconnected forever and never yield a single reading.
    """

    def on_get(self, req: Request, resp: Response, devnum: int):
        resp.text = PropertyResponse(
            weather_dev.connected if weather_dev else False, req).json

    def on_put(self, req: Request, resp: Response, devnum: int):
        conn = to_bool(get_request_field('Connected', req))     # 400 if not a bool
        try:
            if conn:
                weather_dev.connect()
            else:
                weather_dev.disconnect()
            resp.text = MethodResponse(req).json
        except Exception as ex:
            resp.text = MethodResponse(
                req, DriverException(
                    0x500, 'Connected={} failed'.format(conn), ex)).json


@before(PreProcessRequest(maxdev))
class BaseDescription:
    metadata = None

    def on_get(self, req: Request, resp: Response, devnum: int):
        resp.text = PropertyResponse(self.metadata.Description, req).json


@before(PreProcessRequest(maxdev))
class BaseDriverInfo:
    metadata = None

    def on_get(self, req: Request, resp: Response, devnum: int):
        resp.text = PropertyResponse(self.metadata.Info, req).json


@before(PreProcessRequest(maxdev))
class BaseDriverVersion:
    metadata = None

    def on_get(self, req: Request, resp: Response, devnum: int):
        resp.text = PropertyResponse(self.metadata.Version, req).json


@before(PreProcessRequest(maxdev))
class BaseInterfaceVersion:
    metadata = None

    def on_get(self, req: Request, resp: Response, devnum: int):
        resp.text = PropertyResponse(self.metadata.InterfaceVersion, req).json


@before(PreProcessRequest(maxdev))
class BaseName:
    metadata = None

    def on_get(self, req: Request, resp: Response, devnum: int):
        resp.text = PropertyResponse(self.metadata.Name, req).json


@before(PreProcessRequest(maxdev))
class BaseSupportedActions:
    def on_get(self, req: Request, resp: Response, devnum: int):
        # Must return a list, never PropertyNotImplemented. This is how a
        # generic client discovers the diagnostics action.
        resp.text = PropertyResponse(SUPPORTED_ACTIONS, req).json


# --------------------------------------------------------------------------
# ObservingConditions sensor bases. Defined here rather than in the device
# module because app.py routes every class DEFINED in a device module -- a
# helper base declared there would acquire a URI of its own.
# --------------------------------------------------------------------------

@before(PreProcessRequest(maxdev))
class BaseSensorValue:
    """A sensor this site actually has."""
    sensor = None

    def on_get(self, req: Request, resp: Response, devnum: int):
        if not_connected(req, resp):
            return
        try:
            resp.text = PropertyResponse(
                weather_dev.sensor_value(self.sensor), req).json
        except ValueNotAvailable as ex:
            # Known sensor, no reading yet. NOT zero: a fabricated zero is
            # indistinguishable from a real calm, and this one would be read by
            # something deciding whether to keep the dome open.
            resp.text = PropertyResponse(
                None, req, ValueNotSetException(str(ex))).json
        except Exception as ex:
            resp.text = PropertyResponse(
                None, req,
                DriverException(0x500, '{} failed'.format(self.sensor), ex)).json


@before(PreProcessRequest(maxdev))
class BaseAbsentSensor:
    """A sensor the interface defines and this site does not have.

    Throws PropertyNotImplemented rather than returning a plausible value.
    Arcsecond caches the result and stops asking, and any other client learns
    the truth instead of recording a fabricated humidity every minute.
    """
    sensor = None

    def on_get(self, req: Request, resp: Response, devnum: int):
        if not_connected(req, resp):
            return
        resp.text = PropertyResponse(
            None, req,
            NotImplementedException(
                '{} is not fitted at this site'.format(self.sensor))).json
