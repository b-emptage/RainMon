# -*- coding: utf-8 -*-
"""What the observatory says out loud.

Separated from both the window and the speech engine so the decisions can be
tested without a display or a sound card -- which is most of what matters here,
because the failure mode of a spoken alert is not silence, it is nagging. An
alert that repeats every second gets muted, and a muted alert protects nobody.

The rules are the legacy ones, with two additions. The original announced rain
and announced dry; it had no way to say that the sensors had stopped answering,
because it could not tell. This can, and that is the announcement most worth
having: an astronomer who hears nothing assumes it is dry.
"""

from typing import List, Optional

# Categories, so an observatory that finds one of them tiresome can silence it
# without losing the others.
RAIN = 'rain'
SAFETY = 'safety'
FAULTS = 'faults'


class Snapshot(object):
    """Everything the alert policy is allowed to look at."""

    __slots__ = ('wet_sections', 'active_sections', 'rain_ok', 'wind_ok',
                 'is_safe', 'reasons')

    def __init__(self, wet_sections=None, active_sections=0, rain_ok=False,
                 wind_ok=False, is_safe=False, reasons=None):
        # type: (Optional[int], int, bool, bool, bool, Optional[List[str]]) -> None
        self.wet_sections = wet_sections        # None when we cannot tell
        self.active_sections = active_sections
        self.rain_ok = rain_ok
        self.wind_ok = wind_ok
        self.is_safe = is_safe
        self.reasons = reasons or []


class AlertPolicy(object):
    """Turns a stream of snapshots into the few sentences worth saying.

    Everything is edge-triggered except the rain reminder, which repeats on a
    timer for as long as the sensors are wet -- the one case where an astronomer
    who missed the first announcement still needs to know.
    """

    def __init__(self, repeat_seconds=30.0, categories=(RAIN, SAFETY, FAULTS),
                 startup_grace_seconds=20.0):
        self._repeat_seconds = repeat_seconds
        self._categories = set(categories)
        # Nothing is announced as broken until the streams have had a chance to
        # arrive. Every source looks dead for the first fraction of a second,
        # and a window that greets whoever opens it with "rain sensors not
        # responding" teaches them to ignore the sentence that matters most.
        # Longer than the staleness bound, so a genuinely dead sensor IS
        # announced shortly after the window opens -- that is current news,
        # unlike rain, which would be stale.
        self._startup_grace_seconds = startup_grace_seconds
        self._previous = None           # type: Optional[Snapshot]
        self._last_rain_phrase_at = None    # type: Optional[float]
        self._first_seen_at = None      # type: Optional[float]
        # Only a fault we announced can recover. Otherwise the transition from
        # "no data yet" to "working" is announced as a recovery from a failure
        # nobody was ever told about.
        self._announced_faults = set()

    def _enabled(self, category):
        return category in self._categories

    def update(self, now, snapshot):
        # type: (float, Snapshot) -> List[str]
        """The phrases to speak for this snapshot. Usually none."""
        phrases = []        # type: List[str]
        previous = self._previous

        if previous is None:
            # Say nothing on the first snapshot. Announcing the weather at
            # start-up would mean the window greets whoever opens it with
            # "rain detected" for rain that stopped an hour ago.
            self._previous = snapshot
            self._first_seen_at = now
            if self._wet(snapshot):
                self._last_rain_phrase_at = now
            return phrases

        phrases.extend(self._rain_phrases(now, previous, snapshot))
        phrases.extend(self._fault_phrases(now, snapshot))
        phrases.extend(self._safety_phrases(previous, snapshot))

        self._previous = snapshot
        return phrases

    # -- rain -----------------------------------------------------------------

    @staticmethod
    def _wet(snapshot):
        return bool(snapshot.wet_sections)

    def _rain_phrases(self, now, previous, snapshot):
        if not self._enabled(RAIN):
            return []

        if self._wet(snapshot):
            first = not self._wet(previous)
            due = (self._last_rain_phrase_at is None
                   or now - self._last_rain_phrase_at >= self._repeat_seconds)
            if first or due:
                self._last_rain_phrase_at = now
                return ['Rain detected. {} of {} sections wet.'.format(
                    snapshot.wet_sections, snapshot.active_sections)]
            return []

        if self._wet(previous) and snapshot.wet_sections == 0:
            # Only from a known-wet state to a known-dry one. Going wet ->
            # unreadable is a fault, not good news, and saying "sensors dry"
            # there would be the most misleading sentence this could utter.
            self._last_rain_phrase_at = None
            return ['Rain sensors dry.']
        return []

    # -- faults ---------------------------------------------------------------

    def _fault_phrases(self, now, snapshot):
        if not self._enabled(FAULTS):
            return []

        settled = (self._first_seen_at is not None
                   and now - self._first_seen_at >= self._startup_grace_seconds)

        phrases = []
        for ok, name in ((snapshot.rain_ok, 'Rain sensors'),
                         (snapshot.wind_ok, 'Wind sensor')):
            if not ok and settled and name not in self._announced_faults:
                self._announced_faults.add(name)
                phrases.append('{} not responding.'.format(name))
            elif ok and name in self._announced_faults:
                self._announced_faults.discard(name)
                phrases.append('{} responding again.'.format(name))
        return phrases

    # -- safety ---------------------------------------------------------------

    def _safety_phrases(self, previous, snapshot):
        if not self._enabled(SAFETY):
            return []
        if previous.is_safe and not snapshot.is_safe:
            return ['Conditions unsafe.']
        if snapshot.is_safe and not previous.is_safe:
            return ['Conditions safe.']
        return []
