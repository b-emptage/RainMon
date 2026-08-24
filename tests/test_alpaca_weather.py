# -*- coding: utf-8 -*-
"""The Alpaca surface, driven through the real Falcon routing.

What an ASCOM client -- Arcsecond, NINA, Conform -- actually reaches. The
safety rules themselves are covered in test_safety.py; these tests are about
whether the HTTP layer reports them faithfully and fails in the ways the spec
requires.
"""

import json

import pytest

from conftest import set_weather_state

# Alpaca error numbers.
NOT_IMPLEMENTED = 0x400
INVALID_VALUE = 0x401
NOT_CONNECTED = 0x407
VALUE_NOT_SET = 0x402

PARAMS = {'ClientID': 1, 'ClientTransactionID': 7}

OC = '/api/v1/observingconditions/0/'
SM = '/api/v1/safetymonitor/0/'


def get(client, path, **extra):
    params = dict(PARAMS)
    params.update(extra)
    return client.simulate_get(path, params=params)


def put(client, path, body):
    form = dict(PARAMS)
    form.update(body)
    return client.simulate_put(
        path, body='&'.join('{}={}'.format(k, v) for k, v in form.items()),
        headers={'Content-Type': 'application/x-www-form-urlencoded'})


class TestManagement:
    def test_both_devices_are_advertised(self, client):
        devices = client.simulate_get('/management/v1/configureddevices').json['Value']
        assert {d['DeviceType'] for d in devices} == {'ObservingConditions',
                                                      'SafetyMonitor'}

    def test_every_device_presents_a_unique_id(self, client):
        # Arcsecond identifies devices by UniqueID and silently drops any
        # device without one. These must also stay STABLE across restarts, or
        # the site's equipment registration breaks.
        devices = client.simulate_get('/management/v1/configureddevices').json['Value']
        ids = [d['UniqueID'] for d in devices]
        assert all(ids) and len(set(ids)) == 2

    def test_api_versions(self, client):
        assert client.simulate_get('/management/apiversions').json['Value'] == [1]


class TestCommonMembers:
    @pytest.mark.parametrize('base', [OC, SM])
    def test_identity_members(self, client, base):
        assert get(client, base + 'name').json['Value']
        assert get(client, base + 'description').json['Value']
        assert get(client, base + 'driverinfo').json['Value']
        assert get(client, base + 'driverversion').json['Value']

    @pytest.mark.parametrize('base,version', [(OC, 2), (SM, 3)])
    def test_interface_version(self, client, base, version):
        assert get(client, base + 'interfaceversion').json['Value'] == version

    @pytest.mark.parametrize('base', [OC, SM])
    def test_supported_actions_is_a_list(self, client, base):
        # Must be a list, never PropertyNotImplemented: it is how a generic
        # client discovers the diagnostics action.
        value = get(client, base + 'supportedactions').json['Value']
        assert isinstance(value, list)
        assert 'Greenhill:GetWeatherStatus' in value

    @pytest.mark.parametrize('base', [OC, SM])
    def test_command_members_are_not_implemented(self, client, base):
        for member in ('commandblind', 'commandbool', 'commandstring'):
            body = put(client, base + member, {'Command': 'x', 'Raw': 'false'})
            assert body.json['ErrorNumber'] == NOT_IMPLEMENTED

    def test_the_transaction_id_is_echoed(self, client):
        # Alpaca clients match responses to requests by this.
        assert get(client, SM + 'issafe').json['ClientTransactionID'] == 7


class TestConnected:
    def test_defaults_to_connected(self, client):
        # Arcsecond polls Connected but never sets it. A device that waited to
        # be connected would be recorded as disconnected forever and never
        # yield a single reading.
        assert get(client, SM + 'connected').json['Value'] is True
        assert get(client, OC + 'connected').json['Value'] is True

    def test_both_devices_share_one_connection(self, client):
        put(client, SM + 'connected', {'Connected': 'false'})
        assert get(client, OC + 'connected').json['Value'] is False

    def test_disconnected_reads_report_not_connected(self, client):
        put(client, SM + 'connected', {'Connected': 'false'})
        assert get(client, SM + 'issafe').json['ErrorNumber'] == NOT_CONNECTED
        assert get(client, OC + 'windspeed').json['ErrorNumber'] == NOT_CONNECTED

    def test_disconnecting_does_not_stop_the_monitoring(self, client, device):
        # The dome de-energises its motors on disconnect. This device must NOT
        # do the equivalent: the direct dome-close route reads the safety state
        # and has no ASCOM client of its own, so stopping on disconnect would
        # quietly disarm it.
        put(client, SM + 'connected', {'Connected': 'false'})
        assert device._thread is not None
        assert device._thread.is_alive()

    def test_reconnecting_works(self, client):
        put(client, SM + 'connected', {'Connected': 'false'})
        put(client, SM + 'connected', {'Connected': 'true'})
        assert get(client, SM + 'issafe').json['ErrorNumber'] == 0

    def test_platform_7_connect_and_disconnect(self, client):
        assert put(client, SM + 'disconnect', {}).json['ErrorNumber'] == 0
        assert get(client, SM + 'connected').json['Value'] is False
        assert put(client, SM + 'connect', {}).json['ErrorNumber'] == 0
        assert get(client, SM + 'connected').json['Value'] is True
        assert get(client, SM + 'connecting').json['Value'] is False

    def test_a_non_boolean_connected_is_rejected(self, client):
        assert put(client, SM + 'connected', {'Connected': 'banana'}).status_code == 400


class TestSafetyMonitor:
    def test_reports_safe_when_the_core_says_so(self, client, device):
        set_weather_state(device, is_safe=True)
        assert get(client, SM + 'issafe').json['Value'] is True

    def test_reports_unsafe_when_the_core_says_so(self, client, device):
        set_weather_state(device, is_safe=False, reasons=['rain: 3 wet sections'])
        assert get(client, SM + 'issafe').json['Value'] is False

    def test_is_unsafe_before_the_first_evaluation(self, client, device):
        # A safety monitor that has not yet looked must never answer "safe".
        with device._lock:
            device._state = None
        assert get(client, SM + 'issafe').json['Value'] is False

    def test_issafe_answers_false_rather_than_throwing(self, client, device):
        # Most clients treat an error from a safety monitor as "no opinion" and
        # carry on observing, so an exception here would read as permission.
        class Exploding:
            @property
            def is_safe(self):
                raise RuntimeError('boom')
            connected = True

        import alpaca_common
        alpaca_common.set_weather_device(Exploding())
        try:
            body = get(client, SM + 'issafe').json
            assert body['Value'] is False
            assert body['ErrorNumber'] == 0
        finally:
            alpaca_common.set_weather_device(device)

    def test_devicestate(self, client, device):
        set_weather_state(device, is_safe=True)
        values = {v['Name']: v['Value']
                  for v in get(client, SM + 'devicestate').json['Value']}
        assert values['IsSafe'] is True
        assert values['TimeStamp'].endswith('Z')


class TestObservingConditions:
    def test_publishes_the_sensors_this_site_has(self, client, device):
        set_weather_state(device, wind_speed=5.5, wind_gust=7.25,
                          wind_direction=210.0, temperature=11.5, rain_rate=0.0)
        assert get(client, OC + 'windspeed').json['Value'] == 5.5
        assert get(client, OC + 'windgust').json['Value'] == 7.25
        assert get(client, OC + 'winddirection').json['Value'] == 210.0
        assert get(client, OC + 'temperature').json['Value'] == 11.5
        assert get(client, OC + 'rainrate').json['Value'] == 0.0

    @pytest.mark.parametrize('sensor', [
        'cloudcover', 'dewpoint', 'humidity', 'pressure', 'skybrightness',
        'skyquality', 'skytemperature', 'starfwhm'])
    def test_absent_sensors_are_not_implemented(self, client, sensor):
        # Not a plausible zero. Arcsecond caches PropertyNotImplemented and
        # stops asking; a fabricated humidity would be recorded every minute.
        assert get(client, OC + sensor).json['ErrorNumber'] == NOT_IMPLEMENTED

    def test_a_sensor_with_no_reading_yet_is_value_not_set(self, client, device):
        # Not zero: a fabricated calm is indistinguishable from a real one, and
        # something is about to decide whether to keep the dome open on it.
        set_weather_state(device, wind_speed=None)
        assert get(client, OC + 'windspeed').json['ErrorNumber'] == VALUE_NOT_SET

    def test_rain_rate_reports_the_latched_state(self, client, device):
        set_weather_state(device, is_safe=False, rain_rate=1.0,
                          detectors={'H127': 'D', 'H50': 'D', 'ACC': 'D'})
        # Sensors read dry, yet the rate is still up, because the latch holds.
        # This is what stops Arcsecond reopening the dome onto wet sensors.
        assert get(client, OC + 'rainrate').json['Value'] == 1.0

    def test_average_period_is_reported_honestly(self, client):
        # These values ARE averaged, over the same 60 s the wind threshold
        # uses. Reporting the conventional 0.0 would tell a client the trace is
        # instantaneous when it is not.
        assert get(client, OC + 'averageperiod').json['Value'] == pytest.approx(
            60.0 / 3600.0)

    def test_average_period_accepts_its_own_value(self, client):
        current = get(client, OC + 'averageperiod').json['Value']
        assert put(client, OC + 'averageperiod',
                   {'AveragePeriod': current}).json['ErrorNumber'] == 0

    def test_average_period_rejects_anything_else(self, client):
        # Rejected rather than ignored: a client told "fine" would go on to
        # read numbers that are not what it asked for.
        body = put(client, OC + 'averageperiod', {'AveragePeriod': '0.5'}).json
        assert body['ErrorNumber'] == INVALID_VALUE
        assert 'fixed' in body['ErrorMessage']

    def test_sensor_descriptions_are_published(self, client):
        for sensor in ('WindSpeed', 'WindGust', 'WindDirection', 'Temperature',
                       'RainRate'):
            value = get(client, OC + 'sensordescription',
                        SensorName=sensor).json['Value']
            assert value

    def test_the_rain_rate_description_says_it_is_derived(self, client):
        # A client about to make a decision on RainRate deserves to know there
        # is no gauge behind it.
        text = get(client, OC + 'sensordescription',
                   SensorName='RainRate').json['Value']
        assert 'NOT A MEASUREMENT' in text.upper()

    def test_descriptions_of_absent_sensors_are_not_implemented(self, client):
        assert get(client, OC + 'sensordescription',
                   SensorName='Humidity').json['ErrorNumber'] == NOT_IMPLEMENTED

    def test_an_unknown_sensor_name_is_an_invalid_value(self, client):
        assert get(client, OC + 'sensordescription',
                   SensorName='Bananas').json['ErrorNumber'] == INVALID_VALUE

    def test_time_since_last_update_per_sensor(self, client, device):
        set_weather_state(device, rain_age=3.0, wind_age=8.0)
        assert get(client, OC + 'timesincelastupdate',
                   SensorName='RainRate').json['Value'] == 3.0
        assert get(client, OC + 'timesincelastupdate',
                   SensorName='WindSpeed').json['Value'] == 8.0

    def test_time_since_last_update_with_no_name_is_the_most_recent(self, client,
                                                                    device):
        set_weather_state(device, rain_age=3.0, wind_age=8.0)
        assert get(client, OC + 'timesincelastupdate',
                   SensorName='').json['Value'] == 3.0

    def test_refresh_succeeds(self, client):
        # A no-op, but clients call it as a matter of course and failing would
        # be worse than doing nothing.
        assert put(client, OC + 'refresh', {}).json['ErrorNumber'] == 0

    def test_devicestate_carries_the_readings(self, client, device):
        set_weather_state(device, wind_speed=4.0, rain_rate=0.0)
        values = {v['Name']: v['Value']
                  for v in get(client, OC + 'devicestate').json['Value']}
        assert values['WindSpeed'] == 4.0
        assert values['RainRate'] == 0.0
        assert values['TimeStamp'].endswith('Z')

    def test_devicestate_omits_readings_it_does_not_have(self, client, device):
        set_weather_state(device, wind_speed=None)
        names = {v['Name'] for v in get(client, OC + 'devicestate').json['Value']}
        assert 'WindSpeed' not in names
        assert 'RainRate' in names


class TestDiagnosticsAction:
    def test_publishes_what_the_boolean_throws_away(self, client, device):
        # ASCOM has one bit and this site has four independent reasons to
        # close, so a bare `false` cannot say which.
        set_weather_state(device, is_safe=False, reasons=['rain: 3 wet sections'],
                          conditions={'rain': True, 'wind': False,
                                      'rain_data': False, 'wind_data': False},
                          detectors={'H127': 'W', 'H50': 'w', 'ACC': 'D'})
        body = put(client, SM + 'action',
                   {'Action': 'Greenhill:GetWeatherStatus', 'Parameters': ''}).json
        assert body['ErrorNumber'] == 0
        status = json.loads(body['Value'])
        assert status['isSafe'] is False
        assert status['conditions']['rain'] is True
        assert status['detectors'] == {'H127': 'W', 'H50': 'w', 'ACC': 'D'}

    def test_is_available_on_both_devices(self, client, device):
        set_weather_state(device)
        for base in (OC, SM):
            body = put(client, base + 'action',
                       {'Action': 'Greenhill:GetWeatherStatus',
                        'Parameters': ''}).json
            assert body['ErrorNumber'] == 0

    def test_an_unknown_action_is_rejected(self, client):
        body = put(client, SM + 'action',
                   {'Action': 'Nope', 'Parameters': ''}).json
        assert body['ErrorNumber'] == 0x40C     # ActionNotImplemented


class TestAlpacaProtocol:
    def test_a_bad_client_id_is_a_400(self, client):
        # Conform found this in the AlpycaDevice sample: a repeated ClientID
        # made Falcon collapse the value into a list, and int(list) raised,
        # producing a 500 where Alpaca requires 400.
        assert client.simulate_get(
            SM + 'issafe',
            params={'ClientID': ['1', ''], 'ClientTransactionID': 1}
        ).status_code == 400

    def test_a_missing_required_parameter_is_a_400(self, client):
        assert put(client, SM + 'connected', {}).status_code == 400

    def test_a_negative_device_number_does_not_route(self, client):
        assert get(client, '/api/v1/safetymonitor/-1/issafe').status_code == 404

    def test_an_unknown_device_number_is_rejected(self, client):
        assert get(client, '/api/v1/safetymonitor/3/issafe').status_code == 400


class TestServerIdentity:
    def test_the_server_describes_itself_as_the_weather_server(self, client):
        # device/ was cloned from the dome server, whose DeviceMetadata is
        # inside the vendored shr.py. Left alone, this server introduces itself
        # to every discovering client as the clamshell dome.
        value = client.simulate_get('/management/v1/description').json['Value']
        assert 'dome' not in value['ServerName'].lower()
        assert 'weather' in value['ServerName'].lower()



class TestRouteOneArming:
    """Whether the direct dome close is armed, and how loudly it says so.

    This is the only thing in the package that commands a roof, so both states
    are logged at a level nobody can miss: armed, because that is worth
    knowing; and NOT armed, because a forgotten `false` would leave the
    observatory a route short with nothing saying so.
    """

    def _service(self, **weather):
        import logging

        from greenhill.core.config import WeatherConfig
        from weatherdevice import GreenhillWeather

        logger = logging.getLogger('greenhill-arming-test')
        multicast = {'wind_group': '239.192.0.4', 'wind_port': 60004,
                     'rain_group': '239.192.0.5', 'rain_port': 60005,
                     'interface': '0.0.0.0'}
        return GreenhillWeather(WeatherConfig(**weather), multicast, logger)

    def test_disarmed_by_default_and_says_so(self, caplog):
        service = self._service()
        with caplog.at_level('INFO'):
            service._start_dome_closer()
        assert service._closer is None
        assert 'DOME CLOSE NOT ARMED' in caplog.text

    def test_arming_needs_an_address(self, caplog):
        # Enabled but unconfigured is the dangerous middle: it reads as armed
        # in the config file and is not.
        service = self._service(dome_close_enabled=True, dome_address='')
        with caplog.at_level('INFO'):
            service._start_dome_closer()
        assert service._closer is None
        assert 'DOME CLOSE MISCONFIGURED' in caplog.text

    def test_arms_when_configured(self, caplog):
        service = self._service(dome_close_enabled=True,
                                dome_address='127.0.0.1:11111')
        with caplog.at_level('INFO'):
            service._start_dome_closer()
        try:
            assert service._closer is not None
            assert 'DOME CLOSE ARMED' in caplog.text
        finally:
            service._closer.stop()

    def test_simulated_weather_never_arms_it(self, caplog):
        # simulate.py is what Conform and bench work run against, on machines
        # that may well be able to reach the real dome. A server running on
        # invented weather must not be able to command a roof, whatever the
        # config says.
        service = self._service(dome_close_enabled=True,
                                dome_address='127.0.0.1:11111')
        with caplog.at_level('INFO'):
            service.start(simulate=True)
        try:
            assert service._closer is None
            assert 'simulated mode never commands the dome' in caplog.text
        finally:
            service.stop()

    def test_the_verdict_passed_to_the_closer_is_none_before_evaluation(self):
        # None is not "unsafe". Commanding a roof before anything has been
        # evaluated would be acting on no information at all.
        service = self._service()
        assert service._safety_verdict() is None

    def test_the_diagnostics_action_reports_the_armed_state(self, client, device):
        # Readable from the ASCOM surface, so an operator can check without
        # logging in to the machine.
        set_weather_state(device)
        body = put(client, SM + 'action',
                   {'Action': 'Greenhill:GetWeatherStatus', 'Parameters': ''}).json
        assert json.loads(body['Value'])['domeClose']['state'] == 'not armed'
