# -*- coding: utf-8 -*-
"""Serial decode tests, against a fake port.

The reply strings here are the real ones from the hardware author's notes and
from the simulator in RainMonSimT.py, so these tests are checking the decode
against what the detectors actually say rather than against a restatement of
the code.
"""

import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from greenhill import rain_serial  # noqa: E402
from greenhill.rain_serial import (  # noqa: E402
    DecodeError, PortFailure, RainSensorPort, SerialTimeout,
    decode_mk3_temperatures, decode_status_reply, read_reply,
    temperature_from_adc)

MK1_DRY = '*BISDEE RAIN SENSOR 0 STATUS = D'
MK3_DRY = '*BISDEE RAIN SENSOR MK3  0 STATUS = D 012F'
MK3_WET = '*BISDEE RAIN SENSOR MK3  1 STATUS = W 012F'
MK3_TEMPS = '*BISDEE RAIN SENSOR MK3  0 Tamb=0256 Tnormal=08 Tdrying=0F'


class FakePort:
    """A serial port that replays scripted replies byte by byte.

    Deliberately byte-at-a-time, because that is how read_reply consumes it and
    the CR handling is the part worth exercising.
    """

    def __init__(self, replies=None, fail_on_write=False, fail_on_read=False):
        self.replies = list(replies or [])
        self.writes = []
        self.fail_on_write = fail_on_write
        self.fail_on_read = fail_on_read
        self.closed = False
        self.write_timeout = None
        self._buffer = b''

    def write(self, data):
        if self.fail_on_write:
            raise OSError('device disconnected')
        self.writes.append(data)
        if self.replies:
            reply = self.replies.pop(0)
            self._buffer = b'' if reply is None else (reply.encode('ascii') + b'\r')
        else:
            self._buffer = b''
        return len(data)

    def read(self, size=1):
        if self.fail_on_read:
            raise OSError('read error')
        if not self._buffer:
            return b''              # pyserial signals a timeout with empty bytes
        char, self._buffer = self._buffer[:1], self._buffer[1:]
        return char

    def reset_input_buffer(self):
        pass

    def close(self):
        self.closed = True


class FakeSerialModule:
    def __init__(self, port=None, raise_on_open=None):
        self.port = port
        self.raise_on_open = raise_on_open
        self.opened_with = None

    def Serial(self, *args, **kwargs):
        self.opened_with = (args, kwargs)
        if self.raise_on_open is not None:
            raise self.raise_on_open
        return self.port


class TestTemperature:
    def test_midscale_reading_is_near_room_temperature(self):
        assert temperature_from_adc(512) == pytest.approx(25.0, abs=1.0)

    def test_falls_as_the_adc_reading_rises(self):
        # The Beta equation is DECREASING in the ADC value: 512 is exactly
        # 25.0 C, 800 is near freezing. Worth pinning down, because the
        # direction is not obvious from the formula and an inverted
        # transcription would still look plausible on a mild night.
        readings = [temperature_from_adc(v) for v in (200, 400, 600, 800)]
        assert readings == sorted(readings, reverse=True)

    @pytest.mark.parametrize('adc', [0, -1, 1024, 2000])
    def test_rejects_physically_impossible_readings(self, adc):
        # 0 and 1024 are the divider's unreachable rails; letting them through
        # raises ZeroDivisionError or ValueError somewhere less obvious.
        with pytest.raises(DecodeError):
            temperature_from_adc(adc)


class TestStatusDecode:
    def test_decodes_mk1_without_a_temperature(self):
        assert decode_status_reply(MK1_DRY, 0) == ('D', None)

    def test_decodes_mk3_with_a_temperature(self):
        status, temperature = decode_status_reply(MK3_DRY, 0)
        assert status == 'D'
        assert temperature is not None

    def test_decodes_every_status_letter(self):
        for letter in ('P', 'M', 'I', 'D', 'W', 'w', 'E'):
            reply = '*BISDEE RAIN SENSOR 2 STATUS = ' + letter
            assert decode_status_reply(reply, 2)[0] == letter

    def test_rejects_a_reply_for_the_wrong_detector(self):
        # The detectors share one bus. A reply carrying someone else's number
        # means the request and the answer have drifted out of step, and taking
        # it would attribute one sensor's wetness to another.
        with pytest.raises(DecodeError, match='asked for'):
            decode_status_reply(MK3_WET, 0)

    def test_accepts_that_same_reply_for_its_own_detector(self):
        assert decode_status_reply(MK3_WET, 1)[0] == 'W'

    def test_rejects_a_foreign_prefix(self):
        with pytest.raises(DecodeError, match='prefix'):
            decode_status_reply('*SOMETHING ELSE     0 STATUS = D', 0)

    @pytest.mark.parametrize('reply', ['', 'short', MK1_DRY + 'trailing'])
    def test_rejects_wrong_lengths(self, reply):
        with pytest.raises(DecodeError):
            decode_status_reply(reply, 0)

    def test_rejects_an_unknown_status_letter(self):
        with pytest.raises(DecodeError, match='status letter'):
            decode_status_reply('*BISDEE RAIN SENSOR 0 STATUS = X', 0)

    def test_a_corrupt_temperature_does_not_cost_the_status(self):
        # The thermistor is diagnostic; the wet/dry letter is the safety input.
        # Losing the first must never cost the second.
        status, temperature = decode_status_reply(
            '*BISDEE RAIN SENSOR MK3  0 STATUS = W ZZZZ', 0)
        assert status == 'W'
        assert temperature is None


class TestMk3Temperatures:
    def test_decodes_all_three_fields(self):
        ambient, normal, drying = decode_mk3_temperatures(MK3_TEMPS)
        assert ambient is not None
        assert normal == 0x08
        assert drying == 0x0F

    def test_missing_fields_come_back_as_none(self):
        # None, not the original's -999 sentinel, so a failure cannot be
        # mistaken for a very cold night.
        assert decode_mk3_temperatures('nothing useful') == (None, None, None)


class TestReadReply:
    def test_reads_up_to_the_carriage_return(self):
        port = FakePort([MK1_DRY])
        port.write(b'*R0S\r')
        assert read_reply(port) == MK1_DRY

    def test_raises_on_timeout(self):
        port = FakePort([None])
        port.write(b'*R0S\r')
        with pytest.raises(SerialTimeout):
            read_reply(port)

    def test_survives_a_byte_above_ascii(self):
        # The original decoded each byte as UTF-8 as it arrived, so one noise
        # byte on the line raised UnicodeDecodeError and took the process with
        # it. Bytes are now collected and decoded once, with replacement.
        port = FakePort()
        port._buffer = b'\xffBISDEE\r'
        reply = read_reply(port)
        assert reply.endswith('BISDEE')
        assert len(reply) == 7          # the bad byte became one replacement char

    def test_bounds_a_reply_that_never_terminates(self):
        # A device streaming without a carriage return used to grow the buffer
        # without limit.
        port = FakePort()
        port._buffer = b'x' * (rain_serial.MAX_REPLY_BYTES + 50)
        with pytest.raises(DecodeError, match='carriage return'):
            read_reply(port)

    def test_a_read_error_is_a_port_failure_not_a_timeout(self):
        with pytest.raises(PortFailure):
            read_reply(FakePort(fail_on_read=True))


class TestRainSensorPort:
    def build(self, replies=None, **kwargs):
        port = FakePort(replies)
        module = FakeSerialModule(port)
        sensors = RainSensorPort(
            'COM7', [(0, 'H127'), (1, 'H50'), (2, 'ACC')],
            serial_module=module, **kwargs)
        return sensors, port

    def test_polls_every_configured_detector(self):
        sensors, port = self.build([
            '*BISDEE RAIN SENSOR 0 STATUS = D',
            '*BISDEE RAIN SENSOR 1 STATUS = w',
            '*BISDEE RAIN SENSOR 2 STATUS = W'])
        sensors.open()
        readings, port_ok = sensors.poll()
        assert port_ok is True
        assert [r.status for r in readings] == ['D', 'w', 'W']
        assert [r.identifier for r in readings] == ['H127', 'H50', 'ACC']

    def test_a_silent_detector_reports_e_and_the_others_still_read(self):
        sensors, port = self.build([
            '*BISDEE RAIN SENSOR 0 STATUS = D',
            None,                                   # detector 1 does not answer
            '*BISDEE RAIN SENSOR 2 STATUS = W'])
        sensors.open()
        readings, port_ok = sensors.poll()
        assert [r.status for r in readings] == ['D', 'e', 'W']
        assert port_ok is True                      # the LINK is fine
        assert readings[1].reason

    def test_a_dead_port_reports_e_for_everything_and_clears_port_ok(self):
        sensors, port = self.build()
        sensors.open()
        port.fail_on_write = True
        readings, port_ok = sensors.poll()
        assert port_ok is False
        assert [r.status for r in readings] == ['e', 'e', 'e']
        assert port.closed is True

    def test_polling_a_closed_port_does_not_raise(self):
        # The emit loop calls poll() unconditionally: a bridge that stopped
        # publishing because its port had gone would be indistinguishable from
        # a bridge that had died, which is the distinction the whole design
        # rests on.
        sensors, _ = self.build()
        readings, port_ok = sensors.poll()
        assert port_ok is False
        assert all(r.status == 'e' for r in readings)

    def test_open_failure_names_the_likely_cause(self):
        module = FakeSerialModule(raise_on_open=OSError('access denied'))
        sensors = RainSensorPort('COM7', [(0, 'H127')], serial_module=module)
        with pytest.raises(PortFailure, match='RainMonT.exe'):
            sensors.open()

    def test_a_repeatedly_failing_detector_is_backed_off(self):
        # Otherwise every dead detector costs a full read timeout on every
        # cycle, and three dead detectors stretch a 1 s cadence to nearly 3 s.
        sensors, port = self.build(failure_backoff_threshold=2,
                                   failure_backoff_cycles=5)
        sensors.open()
        for _ in range(2):
            port.replies = [None, None, None]
            sensors.poll()

        writes_before = len(port.writes)
        port.replies = [None, None, None]
        readings, _ = sensors.poll()
        assert len(port.writes) == writes_before        # nothing was asked
        assert all(r.status == 'e' for r in readings)
        assert all('backed off' in r.reason for r in readings)

    def test_a_detector_that_comes_back_is_picked_up(self):
        # The original probed once at startup and froze the active list, so a
        # detector unplugged at boot stayed dead until someone restarted it.
        sensors, port = self.build(failure_backoff_threshold=2,
                                   failure_backoff_cycles=1)
        sensors.open()
        port.replies = [None, None, None]
        sensors.poll()

        port.replies = ['*BISDEE RAIN SENSOR 0 STATUS = D',
                        '*BISDEE RAIN SENSOR 1 STATUS = D',
                        '*BISDEE RAIN SENSOR 2 STATUS = D']
        readings, port_ok = sensors.poll()
        assert [r.status for r in readings] == ['D', 'D', 'D']
        assert port_ok is True
