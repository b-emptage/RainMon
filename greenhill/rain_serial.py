# -*- coding: utf-8 -*-
"""Serial protocol for the Bisdee Tier rain detectors.

Lifted from the decode logic in RainMonT.py and separated from the Tk UI it was
embedded in, so it can be tested without hardware and run without a display.
Python 3.8 compatible: this runs on the Windows 7 bridge.

Wire protocol, as documented by the hardware author:

    *RnS<CR>   status of detector n   ->  *BISDEE RAIN SENSOR n STATUS = c
                                     or  *BISDEE RAIN SENSOR MK3  n STATUS = c aaaa
    *RnA<CR>   MK3 temperatures      ->  *BISDEE RAIN SENSOR MK3  n Tamb=aaaa
                                             Tnormal=hh Tdrying=hh
    *RnI<CR>   re-initialise detector n
    *P<CR>     park all detectors

`aaaa` is the detector thermistor as a 4-digit ASCII hex ADC reading.

Two behaviours are preserved deliberately from the original, because they are
what the real hardware does rather than what a clean design would choose:

* MK3 units insert "MK3  " at offset 20 of the status reply. It is stripped so
  both hardware generations decode through one path.
* The trailing ADC field is present only on MK2/MK3, giving replies of 37
  characters where MK1 gives 32.

What is NOT preserved is the original's habit of collapsing every failure into
the single status letter "e". The letter still goes out -- the receiver's
safety logic only needs to know the detector is not observing -- but the reason
is logged, because "no reply in 0.9 s" and "replied with the wrong detector
number" call for very different trips up the hill.
"""

import logging
import math
import time
from typing import List, Optional, Tuple

STATUS_PREFIX = '*BISDEE RAIN SENSOR '
STATUS_REPLY_LENGTH = 32            # prefix(20) + index(1) + " STATUS = "(10) + letter(1)
STATUS_REPLY_WITH_ADC_LENGTH = 37   # ... + space + 4 hex digits
MK3_MARKER_SLICE = slice(20, 25)
MK3_MARKER = 'MK3  '

CARRIAGE_RETURN = 13

# Thermistor constant for the detector temperature sensors.
BETA = 4100.0

# A reply that never terminates would otherwise grow a buffer forever. The
# longest legitimate reply is the MK3 temperature line at well under 80 bytes.
MAX_REPLY_BYTES = 256

logger = logging.getLogger(__name__)


class SerialTimeout(Exception):
    """The detector did not answer within the port's read timeout."""


class DecodeError(Exception):
    """The detector answered, but not with something we recognise."""


class PortFailure(Exception):
    """The serial port itself failed -- not one detector, the whole link."""


def temperature_from_adc(adc_value):
    # type: (int) -> float
    """Thermistor ADC reading -> degrees C.

    Beta equation, as in the original. The reading is a ratiometric divider, so
    0 and 1024 are both physically unreachable and mathematically fatal (a
    division by zero and a log of zero respectively); they are rejected rather
    than allowed to raise something less obvious deeper in the call.
    """
    if adc_value <= 0 or adc_value >= 1024:
        raise DecodeError('ADC reading {} out of range'.format(adc_value))
    return -273.15 + BETA * 298.15 / (BETA - 298.15 * math.log((1024.0 - adc_value) / adc_value))


def decode_status_reply(reply, index):
    # type: (str, int) -> Tuple[str, Optional[float]]
    """Decode one *RnS reply into (status letter, temperature or None).

    Raises DecodeError on anything that does not match the documented shape.
    Validation is deliberately no stricter than the original's -- prefix,
    length, detector index and status letter -- because there is MK1 hardware
    in the field that this code cannot be tested against, and rejecting a
    reading the old software accepted would be a regression in the direction
    that matters least.
    """
    if not isinstance(reply, str):
        raise DecodeError('reply is not text')

    # MK3 units announce themselves mid-string. Strip the marker so both
    # generations reach the length checks below in the same shape.
    if reply[MK3_MARKER_SLICE] == MK3_MARKER:
        reply = reply[:20] + reply[25:]

    temperature = None
    if len(reply) == STATUS_REPLY_WITH_ADC_LENGTH:
        adc_field = reply[-4:]
        reply = reply[:STATUS_REPLY_LENGTH]
        try:
            temperature = round(temperature_from_adc(int(adc_field, 16)), 1)
        except (ValueError, DecodeError):
            # A bad temperature must not cost us the status letter: the
            # thermistor is diagnostic, the wet/dry state is the safety input.
            logger.debug('detector %d: unusable ADC field %r', index, adc_field)
            temperature = None

    if len(reply) != STATUS_REPLY_LENGTH:
        raise DecodeError('reply is {} characters, expected {} or {}'.format(
            len(reply), STATUS_REPLY_LENGTH, STATUS_REPLY_WITH_ADC_LENGTH))
    if reply[0:20] != STATUS_PREFIX:
        raise DecodeError('reply does not start with the status prefix')
    if reply[20:21] != '{:1d}'.format(index):
        raise DecodeError('reply is for detector {!r}, asked for {}'.format(
            reply[20:21], index))

    status = reply[-1]
    # Imported here rather than at module scope to keep this module usable on
    # its own; the two modules are otherwise independent.
    from greenhill.rain_protocol import VALID_STATUS
    if status not in VALID_STATUS:
        raise DecodeError('unknown status letter {!r}'.format(status))

    return status, temperature


def decode_mk3_temperatures(reply):
    # type: (str) -> Tuple[Optional[float], Optional[int], Optional[int]]
    """Decode one *RnA reply into (ambient C, normal offset, drying offset).

    Note that ambient uses a plain linear scaling of the ADC, not the Beta
    equation used for the detector thermistor above. That asymmetry is in the
    original and in the hardware; it is not a transcription slip.

    Diagnostic only -- these are heater setpoints, not safety inputs. Fields
    that cannot be read come back as None rather than the original's -999
    sentinels, so a caller cannot mistake a failure for a very cold night.
    """
    ambient = None      # type: Optional[float]
    normal = None       # type: Optional[int]
    drying = None       # type: Optional[int]

    marker = reply.find('Tamb=')
    if marker >= 0:
        try:
            ambient = round(
                (500.0 * float(int(reply[marker + 5:marker + 9], 16)) / 1024.0) - 273.15, 2)
        except ValueError:
            ambient = None

    marker = reply.find('Tnormal=')
    if marker >= 0:
        try:
            normal = int(reply[marker + 8:marker + 10], 16)
        except ValueError:
            normal = None

    marker = reply.find('Tdrying=')
    if marker >= 0:
        try:
            drying = int(reply[marker + 8:marker + 10], 16)
        except ValueError:
            drying = None

    return ambient, normal, drying


def read_reply(port):
    # type: (object) -> str
    """Read one CR-terminated reply from an open serial port.

    Byte-at-a-time, as in the original, because the detectors answer with no
    length prefix and no fixed size. Differences from the original, both of
    which were latent crashes rather than design choices:

    * bytes are collected and decoded once at the end, so a byte above 0x7F on
      a noisy line cannot raise UnicodeDecodeError mid-read;
    * the buffer is bounded, so a device that streams without ever sending a
      carriage return cannot exhaust memory.
    """
    buffer = bytearray()
    while True:
        try:
            char = port.read()
        except Exception as exc:
            raise PortFailure('read failed: {}'.format(exc))

        if not char:
            raise SerialTimeout('no reply within the port timeout')
        if char[0] == CARRIAGE_RETURN:
            try:
                port.reset_input_buffer()
            except Exception:
                pass        # a flush we cannot do is not worth losing the reply over
            return buffer.decode('ascii', errors='replace')

        buffer.extend(char)
        if len(buffer) > MAX_REPLY_BYTES:
            raise DecodeError(
                'no carriage return within {} bytes'.format(MAX_REPLY_BYTES))


class Reading(object):
    """One detector's reading from one poll cycle."""

    __slots__ = ('index', 'identifier', 'status', 'temperature_c', 'reason')

    def __init__(self, index, identifier, status, temperature_c=None, reason=None):
        # type: (int, str, str, Optional[float], Optional[str]) -> None
        self.index = index
        self.identifier = identifier
        self.status = status
        self.temperature_c = temperature_c
        self.reason = reason        # why it is "e", when it is

    def __repr__(self):
        return 'Reading({}, {!r}, {!r}, {!r})'.format(
            self.index, self.identifier, self.status, self.temperature_c)


class RainSensorPort(object):
    """Owns the serial link to the detectors.

    Polls every CONFIGURED detector on every cycle rather than probing once at
    startup and freezing the active list, which is what the original did. A
    detector that was unplugged when the software started stayed dead there
    until someone restarted it -- on a bridge meant to run for months without
    attention, that is the wrong default. Here a detector that comes back is
    picked up on its next poll, and one that is genuinely gone simply reports
    "e" until it is fixed.

    The cost of that is a full read timeout per dead detector per cycle, which
    would stretch the emit cadence just when the observatory most wants it
    steady. So a detector that has failed `failure_backoff_threshold` times in
    a row is only retried every `failure_backoff_cycles` cycles, and reports
    "e" without being asked in between.
    """

    def __init__(self, port_name, detectors, baud=9600, read_timeout=0.9,
                 write_timeout=0.4, failure_backoff_threshold=5,
                 failure_backoff_cycles=30, serial_module=None):
        # type: (str, List[Tuple[int, str]], int, float, float, int, int, object) -> None
        """`detectors` is an ordered list of (index, identifier) pairs, e.g.
        [(0, 'H127'), (1, 'H50'), (2, 'ACC')]."""
        self.port_name = port_name
        self.detectors = list(detectors)
        self.baud = baud
        self.read_timeout = read_timeout
        self.write_timeout = write_timeout
        self.failure_backoff_threshold = failure_backoff_threshold
        self.failure_backoff_cycles = failure_backoff_cycles

        # Injectable so the tests can drive a fake port, and so this module
        # imports cleanly on a machine with no pyserial installed.
        if serial_module is None:
            import serial as serial_module     # noqa: F401  (late import by design)
        self._serial = serial_module

        self._port = None
        self._consecutive_failures = {index: 0 for index, _ in self.detectors}
        self._cycles_since_retry = {index: 0 for index, _ in self.detectors}

    @property
    def is_open(self):
        # type: () -> bool
        return self._port is not None

    def open(self):
        # type: () -> None
        """Open the port. Raises PortFailure with a diagnosis worth reading."""
        if self._port is not None:
            return
        try:
            self._port = self._serial.Serial(
                self.port_name, self.baud, timeout=self.read_timeout,
                rtscts=False, dsrdtr=False)
            self._port.write_timeout = self.write_timeout
            self._port.reset_input_buffer()
        except Exception as exc:
            self._port = None
            raise PortFailure(
                "could not open {}: {}. On Windows a serial port can only be "
                "held by one process -- check that the legacy RainMonT.exe is "
                "not still running.".format(self.port_name, exc))

    def close(self):
        # type: () -> None
        if self._port is not None:
            try:
                self._port.close()
            except Exception as exc:
                logger.warning('error closing %s: %s', self.port_name, exc)
            finally:
                self._port = None

    def _should_skip(self, index):
        # type: (int) -> bool
        if self._consecutive_failures[index] < self.failure_backoff_threshold:
            return False
        if self._cycles_since_retry[index] >= self.failure_backoff_cycles:
            self._cycles_since_retry[index] = 0
            return False
        self._cycles_since_retry[index] += 1
        return True

    def _read_one(self, index):
        # type: (int) -> Tuple[str, Optional[float]]
        try:
            self._port.write('*R{:1d}S\r'.format(index).encode('ascii'))
        except Exception as exc:
            raise PortFailure('write failed: {}'.format(exc))
        return decode_status_reply(read_reply(self._port), index)

    def poll(self):
        # type: () -> Tuple[List[Reading], bool]
        """Poll every configured detector once.

        Returns (readings, port_ok). `port_ok` is False when the link itself
        failed, as distinct from individual detectors not answering; the port
        is closed in that case and reopened by the next call to `open`.
        """
        readings = []       # type: List[Reading]
        port_ok = self._port is not None

        for index, identifier in self.detectors:
            if not port_ok:
                readings.append(Reading(index, identifier, 'e',
                                        reason='serial port not open'))
                continue

            if self._should_skip(index):
                readings.append(Reading(index, identifier, 'e',
                                        reason='backed off after repeated failures'))
                continue

            try:
                status, temperature = self._read_one(index)
            except PortFailure as exc:
                # The link is gone, not just this detector. Give up on the rest
                # of the cycle rather than waiting out a timeout per detector.
                logger.error('serial port %s failed: %s', self.port_name, exc)
                self.close()
                port_ok = False
                readings.append(Reading(index, identifier, 'e', reason=str(exc)))
                continue
            except (SerialTimeout, DecodeError) as exc:
                self._consecutive_failures[index] += 1
                if self._consecutive_failures[index] == self.failure_backoff_threshold:
                    logger.warning(
                        'detector %d (%s) has failed %d times in a row; backing '
                        'off to one retry every %d cycles',
                        index, identifier, self.failure_backoff_threshold,
                        self.failure_backoff_cycles)
                readings.append(Reading(index, identifier, 'e', reason=str(exc)))
                continue

            if self._consecutive_failures[index]:
                logger.info('detector %d (%s) is answering again', index, identifier)
            self._consecutive_failures[index] = 0
            self._cycles_since_retry[index] = 0
            readings.append(Reading(index, identifier, status, temperature))

        return readings, port_ok

    def read_mk3_temperatures(self, index):
        # type: (int) -> Tuple[Optional[float], Optional[int], Optional[int]]
        """Diagnostic *RnA read. Not on the safety path; never called by the
        emit loop, because it doubles the serial traffic per detector."""
        if self._port is None:
            raise PortFailure('serial port not open')
        try:
            self._port.write('*R{:1d}A\r'.format(index).encode('ascii'))
        except Exception as exc:
            raise PortFailure('write failed: {}'.format(exc))
        return decode_mk3_temperatures(read_reply(self._port))
