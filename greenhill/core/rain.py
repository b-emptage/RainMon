# -*- coding: utf-8 -*-
"""Rain verdict from the sensor bridge's multicast stream.

The rule the observatory operates on, in its own terms:

    close when 2 of the 6 sections are wet, OR when one section has become wet
    and evaporated again within about ten seconds

The second clause is the interesting one and it is not a rounding of the first.
The sensors are heated, so a real raindrop lands, registers, and dries within
seconds. Something that goes wet and STAYS wet is not weather -- it is bird
droppings, an insect, or a failed sensor -- and closing the dome for it every
night is how a safety system ends up switched off.

ONE DELIBERATE CHANGE FROM THE LEGACY BEHAVIOUR. In the old software a single
wet section that persisted for an hour and then dried still closed the dome,
because a rescheduling timer meant the "within ten seconds" window was never
actually enforced. Here it is: a single wet section that clears inside
`raindrop_window_s` is rain, and one that persists longer and then clears is
contamination that finally dried. That is what was described as the intent, and
it is what the code now does. `raindrop_window_s` is configurable if the
observatory wants the old behaviour back.

Times come from the caller, and they are ARRIVAL times. Nothing here trusts the
timestamp inside a packet: the Windows 7 box has no guaranteed clock sync, and
a machine an hour out would otherwise look permanently fresh.
"""

from typing import Dict, List, Optional, Tuple

UNAVAILABLE = 'unavailable'
NO_RAIN = 'no_rain'
RAIN = 'rain'

# A sequence number that jumps backwards by more than this is the bridge having
# restarted, not a datagram arriving out of order.
SEQUENCE_RESTART_GAP = 100


class RainMonitor(object):
    def __init__(self, config):
        self._config = config
        self._last_arrival = None       # type: Optional[float]
        self._last_packet = None
        self._last_sequence = None      # type: Optional[int]

        # When the current run of exactly-one-wet-section began, or None.
        self._single_wet_since = None   # type: Optional[float]
        # Rain stays asserted this long after the last evidence for it, so a
        # raindrop lasting one packet cannot be missed by a caller that happens
        # to evaluate between packets. The long latch lives in safety.py; this
        # is only about not dropping the observation on the floor.
        self._rain_asserted_until = None  # type: Optional[float]
        self._reasons = []              # type: List[str]

        self._out_of_order = 0
        self._restarts = 0

    # -- ingest ---------------------------------------------------------------

    def update(self, now, packet):
        # type: (float, object) -> bool
        """Take one validated packet, timestamped by ARRIVAL. Returns whether
        it was used."""
        if self._last_sequence is not None:
            if packet.sequence <= self._last_sequence:
                if packet.sequence < self._last_sequence - SEQUENCE_RESTART_GAP:
                    # The bridge restarted and its counter went back to zero.
                    self._restarts += 1
                    self._reset_wetness_history()
                else:
                    # A duplicate or a reordered datagram. Dropping it matters:
                    # an old "dry" arriving after a newer "wet" would otherwise
                    # overwrite the newer reading.
                    self._out_of_order += 1
                    return False

        self._last_sequence = packet.sequence
        self._last_arrival = now
        self._last_packet = packet
        self._evaluate(now, packet)
        return True

    def _reset_wetness_history(self):
        self._single_wet_since = None

    def _evaluate(self, now, packet):
        observing = len(packet.observing_detectors)
        if not packet.port_ok:
            self._reasons = ['rain bridge reports its serial port is down']
            self._single_wet_since = None
            return
        if observing < self._config.rain_min_observing_detectors:
            self._reasons = [
                'only {} rain detector(s) reporting, need {}'.format(
                    observing, self._config.rain_min_observing_detectors)]
            self._single_wet_since = None
            return

        self._reasons = []
        sections = packet.total_wet_sections

        if sections >= self._config.rain_wet_sections_trigger:
            self._single_wet_since = None
            self._assert_rain(now, '{} wet sections (threshold {})'.format(
                sections, self._config.rain_wet_sections_trigger))
            return

        if sections > 0:
            # Exactly one wet section. Not rain yet -- what happens next
            # decides whether it was a drop or a smear.
            if self._single_wet_since is None:
                self._single_wet_since = now
            return

        # Nothing wet now. Was something wet a moment ago?
        if self._single_wet_since is not None:
            duration = now - self._single_wet_since
            self._single_wet_since = None
            if duration <= self._config.raindrop_window_s:
                self._assert_rain(now, 'single section wet for {:.1f}s then dry '
                                       '(raindrop)'.format(duration))

    def _assert_rain(self, now, reason):
        self._rain_asserted_until = now + self._config.raindrop_window_s
        self._reasons = [reason]

    # -- readings -------------------------------------------------------------

    @property
    def wet_sections(self):
        # type: () -> Optional[int]
        return self._last_packet.total_wet_sections if self._last_packet else None

    @property
    def observing_count(self):
        # type: () -> int
        return len(self._last_packet.observing_detectors) if self._last_packet else 0

    @property
    def detector_states(self):
        # type: () -> Dict[str, str]
        if self._last_packet is None:
            return {}
        return {d.id: d.status for d in self._last_packet.detectors}

    def temperature_c(self):
        # type: () -> Optional[float]
        """Mean of the detectors reporting a temperature.

        The detectors are spread across the site and this is the only ambient
        figure the installation has, so ObservingConditions publishes it as
        Temperature. It is a detector-body reading on a heated sensor, not a
        meteorological air temperature, and the device says so.
        """
        if self._last_packet is None:
            return None
        values = [d.temperature_c for d in self._last_packet.detectors
                  if d.temperature_c is not None]
        return sum(values) / len(values) if values else None

    @property
    def out_of_order_count(self):
        return self._out_of_order

    @property
    def restart_count(self):
        return self._restarts

    # -- health ---------------------------------------------------------------

    def age_s(self, now):
        # type: (float) -> Optional[float]
        if self._last_arrival is None:
            return None
        return now - self._last_arrival

    def is_stale(self, now):
        # type: (float) -> bool
        age = self.age_s(now)
        return age is None or age > self._config.rain_max_age_s

    def verdict(self, now):
        # type: (float) -> Tuple[str, List[str]]
        """(UNAVAILABLE | NO_RAIN | RAIN, reasons).

        UNAVAILABLE covers no data, stale data, a failed serial port, and too
        few detectors still speaking. The caller treats all of it as unsafe:
        the one thing this system must never do is report a clear night because
        it has stopped being able to see.
        """
        if self._last_arrival is None:
            return UNAVAILABLE, ['no rain data received']
        if self.is_stale(now):
            return UNAVAILABLE, ['rain data is {:.0f}s old'.format(self.age_s(now))]
        if self._reasons and self._rain_asserted_until is None:
            return UNAVAILABLE, list(self._reasons)

        if self._rain_asserted_until is not None and now < self._rain_asserted_until:
            return RAIN, list(self._reasons)

        # Rain has expired, but the bridge may still be reporting a fault.
        if self._last_packet is not None:
            if not self._last_packet.port_ok:
                return UNAVAILABLE, ['rain bridge reports its serial port is down']
            observing = len(self._last_packet.observing_detectors)
            if observing < self._config.rain_min_observing_detectors:
                return UNAVAILABLE, [
                    'only {} rain detector(s) reporting, need {}'.format(
                        observing, self._config.rain_min_observing_detectors)]

        return NO_RAIN, []
