# -*- coding: utf-8 -*-
"""Spoken alerts, on a queue.

Lifted from the Speaker class in RainMonT.py, which the observatory already
knows works over a remote session -- that took some getting right and is not
worth rediscovering. Two changes: the audio-file path is gone (nothing used
it), and the absence of Windows speech is now a supported state rather than an
exception, so the window runs on any machine.

Speech happens on its own thread. The window must never wait for a sentence to
finish: SAPI blocks for as long as it takes to say the words, and a display
that freezes for two seconds every time it rains is a display nobody trusts.
"""

import logging
import threading
import time

try:                            # Python 3
    from queue import Queue
except ImportError:             # pragma: no cover
    from Queue import Queue     # type: ignore

logger = logging.getLogger('greenhill.ui.speaker')

_STOP = object()


class Speaker(object):
    """Queues text and says it, if this machine can speak.

    On Windows, SAPI through win32com -- the same call the legacy monitor
    used. Anywhere else, and on a Windows box without pywin32, the text is
    logged instead. The window does not care which; an observatory running the
    display on a laptop should still see everything, just not hear it.
    """

    def __init__(self, enabled=True, rate=-3):
        self._enabled = enabled
        self._rate = rate
        self._queue = Queue()
        self._voice = None
        self._available = False
        self._thread = None
        self._spoken = []           # for tests and diagnostics
        self._lock = threading.Lock()

    # -- lifecycle ------------------------------------------------------------

    def start(self):
        # type: () -> None
        if self._thread is not None:
            return
        self._thread = threading.Thread(target=self._run, name='speaker',
                                        daemon=True)
        self._thread.start()

    def stop(self):
        # type: () -> None
        if self._thread is None:
            return
        self._queue.put(_STOP)
        self._thread.join(timeout=3.0)
        self._thread = None

    @property
    def available(self):
        # type: () -> bool
        """Whether anything is actually audible. False until the thread has
        tried to build a voice."""
        return self._available

    @property
    def spoken(self):
        with self._lock:
            return list(self._spoken)

    # -- api ------------------------------------------------------------------

    def say(self, text):
        # type: (str) -> None
        """Queue one sentence. Never blocks."""
        if not self._enabled or not text:
            return
        self._queue.put(text)

    # -- the thread -----------------------------------------------------------

    def _build_voice(self):
        """A SAPI voice, or None.

        COM must be initialised on the thread that uses it, and the voice must
        be created there too -- handing one across threads is what made the
        earlier version fail under a remote session.
        """
        try:
            import pythoncom
            import win32com.client
        except ImportError:
            logger.info('No Windows speech available; alerts will be logged '
                        'rather than spoken.')
            return None, None
        try:
            pythoncom.CoInitialize()
            voice = win32com.client.Dispatch('SAPI.SpVoice')
            voice.Rate = self._rate
            return voice, pythoncom
        except Exception as exc:            # noqa: BLE001
            logger.warning('Could not start Windows speech (%s); alerts will '
                           'be logged rather than spoken.', exc)
            return None, None

    def _run(self):
        voice, com = self._build_voice()
        self._available = voice is not None
        try:
            while True:
                item = self._queue.get()
                if item is _STOP:
                    break
                with self._lock:
                    self._spoken.append(item)
                    del self._spoken[:-50]
                try:
                    if voice is not None:
                        voice.Speak(item)
                    else:
                        logger.info('[alert] %s', item)
                except Exception as exc:        # noqa: BLE001
                    # One failed sentence must not silence every later one, and
                    # must not take the thread down with it.
                    logger.warning('Speech failed for %r: %s', item, exc)
                    time.sleep(0.5)
        finally:
            if com is not None:
                try:
                    com.CoUninitialize()
                except Exception:               # noqa: BLE001
                    pass
