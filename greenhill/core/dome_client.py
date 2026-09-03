# -*- coding: utf-8 -*-
"""Route 1: close the dome directly, over Alpaca.

The fast path. When the safety verdict goes false this issues CloseShutter to
the dome server on the other machine and then WATCHES until the dome says it is
shut. Arcsecond will reach the same conclusion and run its own close procedure,
but on a 60 second weather poll plus a 30 second safety heartbeat -- up to about
ninety seconds behind this.

Both routes are meant to fire. Closing a closed dome is a no-op, so the overlap
costs nothing; what it buys is that neither route is a single point of failure.

Three rules this module is built around:

**It only ever closes.** Reopening is an observing decision and belongs to
Arcsecond's recovery procedure. One opener, no argument.

**It never disconnects.** The dome de-energises its motors on Disconnect, on
purpose -- so a client dropping its connection cannot leave a shell running
unsupervised. Which means disconnecting while a close is in progress would stop
the close. This client sets Connected once and leaves it alone, including on
shutdown.

**It waits for a positive "closed".** The dome reports a shell stopped half way
as `shutterOpen`, deliberately, because ASCOM has no partial state and calling a
dome shut when it is not is the one error that leaves it open in the rain. So
"not yet closed" is the honest reading of anything except `shutterClosed`, and
that is what this waits for.

Standard library only: `urllib`, not `requests`. One less dependency on a
machine that must keep running.
"""

import json
import threading
import time
from typing import Callable, Dict, List, Optional

try:                                    # Python 3
    from urllib.error import HTTPError, URLError
    from urllib.parse import urlencode
    from urllib.request import Request, urlopen
except ImportError:                     # pragma: no cover
    raise

# ASCOM ShutterState. Note that `shutterClosed` is 1 -- 4 is the ERROR state,
# and confusing the two would have this reporting success on a faulted dome.
SHUTTER_OPEN = 0
SHUTTER_CLOSED = 1
SHUTTER_OPENING = 2
SHUTTER_CLOSING = 3
SHUTTER_ERROR = 4

SHUTTER_NAMES = {
    SHUTTER_OPEN: 'open',
    SHUTTER_CLOSED: 'closed',
    SHUTTER_OPENING: 'opening',
    SHUTTER_CLOSING: 'closing',
    SHUTTER_ERROR: 'error',
}

# Alpaca asks each client application to use a stable ClientID so a device can
# tell its callers apart in its own log.
CLIENT_ID = 1782

IDLE = 'idle'
CLOSING = 'closing'
CLOSED = 'closed'
FAILING = 'failing'


class DomeError(Exception):
    """The dome could not be reached, or refused a request."""


class AlpacaDomeClient(object):
    """The bare Alpaca calls this needs. Deliberately small."""

    def __init__(self, address, device_number=0, timeout=10.0):
        # type: (str, int, float) -> None
        self.base = 'http://{}/api/v1/dome/{}'.format(
            address.rstrip('/'), device_number)
        self.timeout = timeout
        self._transaction = 0
        self._lock = threading.Lock()

    def _next_transaction(self):
        with self._lock:
            self._transaction += 1
            return self._transaction

    def _unwrap(self, raw, what):
        try:
            payload = json.loads(raw.decode('utf-8'))
        except ValueError as exc:
            raise DomeError('{}: reply was not JSON ({})'.format(what, exc))
        if not isinstance(payload, dict):
            raise DomeError('{}: reply was not an Alpaca envelope'.format(what))
        number = payload.get('ErrorNumber', 0)
        if number:
            raise DomeError('{}: dome returned error 0x{:X} {}'.format(
                what, number, payload.get('ErrorMessage', '')))
        return payload.get('Value')

    def get(self, member):
        # type: (str) -> object
        query = urlencode({'ClientID': CLIENT_ID,
                           'ClientTransactionID': self._next_transaction()})
        url = '{}/{}?{}'.format(self.base, member, query)
        try:
            with urlopen(url, timeout=self.timeout) as response:
                return self._unwrap(response.read(), 'GET ' + member)
        except (HTTPError, URLError, OSError) as exc:
            raise DomeError('GET {}: {}'.format(member, exc))

    def put(self, member, **fields):
        # type: (str, object) -> object
        body = dict(fields)
        body['ClientID'] = CLIENT_ID
        body['ClientTransactionID'] = self._next_transaction()
        request = Request(
            '{}/{}'.format(self.base, member),
            data=urlencode(body).encode('ascii'),
            headers={'Content-Type': 'application/x-www-form-urlencoded'},
            method='PUT')
        try:
            with urlopen(request, timeout=self.timeout) as response:
                return self._unwrap(response.read(), 'PUT ' + member)
        except (HTTPError, URLError, OSError) as exc:
            raise DomeError('PUT {}: {}'.format(member, exc))

    # -- the three things route 1 does ---------------------------------------

    def is_connected(self):
        # type: () -> bool
        return bool(self.get('connected'))

    def ensure_connected(self):
        # type: () -> None
        """Connect if we are not. Never the reverse -- see the module note."""
        if not self.is_connected():
            self.put('connected', Connected=True)

    def shutter_status(self):
        # type: () -> int
        value = self.get('shutterstatus')
        try:
            return int(value)
        except (TypeError, ValueError):
            raise DomeError('shutterstatus was {!r}, not a state'.format(value))

    def close_shutter(self):
        # type: () -> None
        self.put('closeshutter')


class DomeCloser(object):
    """Watches the safety verdict and closes the dome when it goes false.

    Runs on its own thread. It must not share the weather loop's, because every
    call here is blocking HTTP to another machine and a dome that stopped
    answering would otherwise stall the safety evaluation itself -- the one
    thing that must keep running.
    """

    def __init__(self, client, is_safe, logger, config, reasons=None,
                 clock=time.monotonic):
        # type: (AlpacaDomeClient, Callable[[], Optional[bool]], object, object, Optional[Callable[[], List[str]]], Callable[[], float]) -> None
        """`is_safe` returns True, False, or None when the verdict is not yet
        known. None is NOT treated as unsafe here: the weather service already
        answers False until it has looked, so a None means this closer started
        before the service did, and commanding a roof on that basis would be
        acting on no information at all."""
        self._client = client
        self._is_safe = is_safe
        # Why it is unsafe, for the log. A close is the most consequential
        # thing this package does, and an operator finding a shut dome in the
        # morning should not have to correlate timestamps across two logs to
        # learn what shut it.
        self._reasons = reasons or (lambda: [])
        self._logger = logger
        self._config = config
        self._clock = clock

        self._thread = None         # type: Optional[threading.Thread]
        self._stop = threading.Event()
        self._lock = threading.Lock()

        self._state = IDLE
        self._attempts = 0
        # TWO clocks, and they are not the same thing. `_next_poll_at` is when
        # to look at the dome again; `_reissue_after` is the deadline past
        # which another CloseShutter is warranted. Sharing one timer made the
        # closer go blind for the whole verify timeout right after commanding a
        # close -- exactly the window in which it most needs to be watching.
        self._next_poll_at = 0.0
        self._reissue_after = 0.0
        self._last_status = None    # type: Optional[int]
        self._last_error = None     # type: Optional[str]
        self._closed_at = None      # type: Optional[float]
        self._escalated = False
        self._commands_issued = 0

    # -- lifecycle ------------------------------------------------------------

    def start(self):
        if self._thread is not None:
            return
        self._stop.clear()
        self._thread = threading.Thread(target=self._run, name='dome-closer',
                                        daemon=True)
        self._thread.start()
        self._logger.info(
            '==DOME CLOSE ARMED== route 1 will close %s when conditions go '
            'unsafe', self._client.base)

    def stop(self):
        self._stop.set()
        if self._thread is not None:
            self._thread.join(timeout=5.0)
            self._thread = None
        # Deliberately NOT disconnecting from the dome. Disconnect de-energises
        # its motors, so doing it on our way out could stop a close that is
        # still in progress.

    def _run(self):
        while not self._stop.is_set():
            try:
                self._tick()
            except Exception as exc:            # noqa: BLE001
                # This thread must not be able to die: it is the fast close
                # route, and a dead one fails silently.
                self._logger.exception('dome closer error: %s', exc)
            self._stop.wait(self._config.dome_poll_interval_s)

    # -- the state machine ----------------------------------------------------

    def _tick(self):
        safe = self._is_safe()

        if safe is None:
            return

        if safe:
            if self._state != IDLE:
                self._logger.info(
                    '==DOME CLOSE STAND DOWN== conditions are safe again. This '
                    'service does not reopen the dome; that is Arcsecond\'s '
                    'recovery procedure.')
                self._reset()
            return

        # Unsafe from here down.
        if self._state == CLOSED:
            return                              # already verified shut

        now = self._clock()
        if now < self._next_poll_at:
            return
        self._next_poll_at = now + self._config.dome_poll_interval_s

        try:
            self._client.ensure_connected()
            status = self._client.shutter_status()
        except DomeError as exc:
            self._record_failure('cannot read the dome: {}'.format(exc))
            return

        with self._lock:
            self._last_status = status

        if status == SHUTTER_CLOSED:
            self._mark_closed()
            return

        if status == SHUTTER_CLOSING:
            # Already on its way. Keep watching, but do not command again --
            # the dome's own supervision is on the limit switches, and a second
            # command would only add noise.
            with self._lock:
                self._state = CLOSING
            return

        if status == SHUTTER_ERROR:
            # A latched fault. The dome documents that closing is NEVER blocked
            # by one -- it must not be possible to lock the roof open -- so the
            # close is still worth issuing. Say so loudly: somebody has to go
            # and look at it either way.
            self._logger.error(
                '==DOME FAULT== the dome reports shutterError while conditions '
                'are unsafe. Issuing the close anyway; a latched fault does not '
                'block closing. THIS NEEDS AN ENGINEER.')

        # Open, or faulted, and not on its way. Command it -- unless a previous
        # command is still within its grace period, in which case give it time
        # rather than stacking another on top.
        if now >= self._reissue_after:
            self._issue_close(now, status)

    def _issue_close(self, now, status):
        self._attempts += 1
        try:
            self._client.close_shutter()
        except DomeError as exc:
            self._record_failure('CloseShutter failed: {}'.format(exc))
            return

        with self._lock:
            self._state = CLOSING
            self._commands_issued += 1
            self._last_error = None
        try:
            why = '; '.join(self._reasons()) or 'no reason recorded'
        except Exception:                       # noqa: BLE001
            why = 'no reason recorded'
        self._logger.warning(
            '==DOME CLOSING== %s. The dome reads %s; CloseShutter issued '
            '(attempt %d).', why, SHUTTER_NAMES.get(status, status),
            self._attempts)
        # Keep polling at the normal rate; only the RE-ISSUE waits. The grace
        # period must comfortably exceed the dome's full travel, or a slow
        # close collects a second command halfway through.
        self._reissue_after = now + self._config.dome_verify_timeout_s

    def _mark_closed(self):
        with self._lock:
            first = self._state != CLOSED
            self._state = CLOSED
            self._closed_at = self._clock()
            self._last_error = None
        if first:
            self._logger.warning('==DOME CLOSED== confirmed shut by the dome.')
        self._attempts = 0
        self._escalated = False

    def _record_failure(self, message):
        with self._lock:
            self._state = FAILING
            self._last_error = message
        self._attempts += 1

        if self._attempts <= self._config.dome_retry_limit:
            self._logger.error('==DOME CLOSE FAILED== %s (attempt %d of %d)',
                               message, self._attempts,
                               self._config.dome_retry_limit)
            self._next_poll_at = self._clock() + self._config.dome_poll_interval_s
            return

        # Past the retry limit. Keep trying anyway, but slowly and quietly --
        # an unreachable dome is not going to be fixed by asking faster, and a
        # log filling at 1 Hz buries the message that matters. Never give up
        # entirely: the roof still has to shut.
        if not self._escalated:
            self._escalated = True
            self._logger.critical(
                '==DOME UNREACHABLE== %d attempts to close have failed: %s. '
                'THE DOME MAY BE OPEN IN BAD WEATHER. Retrying every %.0fs. '
                'Arcsecond will attempt its own close procedure, but if that '
                'is also failing the dome needs closing by hand.',
                self._attempts, message, self._config.dome_escalated_retry_s)
        self._next_poll_at = self._clock() + self._config.dome_escalated_retry_s
        # An unreachable dome tells us nothing about whether the last command
        # landed, so let the next reachable tick command it again immediately.
        self._reissue_after = 0.0

    def _reset(self):
        with self._lock:
            self._state = IDLE
            self._last_error = None
            self._closed_at = None
        self._attempts = 0
        self._next_poll_at = 0.0
        self._reissue_after = 0.0
        self._escalated = False

    # -- diagnostics ----------------------------------------------------------

    def status(self):
        # type: () -> Dict[str, object]
        with self._lock:
            return {
                'state': self._state,
                'shutterStatus': (SHUTTER_NAMES.get(self._last_status)
                                  if self._last_status is not None else None),
                'attempts': self._attempts,
                'commandsIssued': self._commands_issued,
                'escalated': self._escalated,
                'lastError': self._last_error,
                'target': self._client.base,
            }
