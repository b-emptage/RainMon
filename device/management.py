# -*- coding: utf-8 -*-
# -----------------------------------------------------------------------------
# management.py - Alpaca management API.
#
# From the AlpycaDevice template -- https://github.com/ASCOMInitiative/AlpycaDevice
# MIT, (c) 2022-2024 Bob Denny. See LICENSE-AlpycaDevice.txt.
# Changed to advertise BOTH device types served here.
# -----------------------------------------------------------------------------
from logging import Logger

from falcon import Request, Response

from config import Config
from observingconditions import ObservingConditionsMetadata
from safetymonitor import SafetyMonitorMetadata
from shr import DeviceMetadata, PropertyResponse

logger: Logger = None


def set_management_logger(lgr):
    global logger
    logger = lgr


class apiversions:
    def on_get(self, req: Request, resp: Response):
        resp.text = PropertyResponse([1], req).json


class description:
    def on_get(self, req: Request, resp: Response):
        resp.text = PropertyResponse({
            'ServerName': DeviceMetadata.Description,
            'Manufacturer': DeviceMetadata.Manufacturer,
            'Version': DeviceMetadata.Version,
            'Location': Config.location,
        }, req).json


class configureddevices:
    """Both devices this server publishes.

    They are two views onto one weather service, not two instruments: the
    SafetyMonitor carries the GO/NOGO and the ObservingConditions carries the
    telemetry. ASCOM has no safety property on ObservingConditions, so a single
    device could not express both honestly.

    The UniqueIDs must stay stable across restarts. Arcsecond identifies
    devices by UniqueID and silently drops any device that does not present
    one -- change these and the site's equipment registration breaks.
    """

    def on_get(self, req: Request, resp: Response):
        resp.text = PropertyResponse([
            {
                'DeviceName': ObservingConditionsMetadata.Name,
                'DeviceType': ObservingConditionsMetadata.DeviceType,
                'DeviceNumber': 0,
                'UniqueID': ObservingConditionsMetadata.DeviceID,
            },
            {
                'DeviceName': SafetyMonitorMetadata.Name,
                'DeviceType': SafetyMonitorMetadata.DeviceType,
                'DeviceNumber': 0,
                'UniqueID': SafetyMonitorMetadata.DeviceID,
            },
        ], req).json
