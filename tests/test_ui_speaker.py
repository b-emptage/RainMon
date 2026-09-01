# -*- coding: utf-8 -*-
"""The speaker, on machines that cannot speak.

The window has to run anywhere -- a laptop, a Mac, a Windows box without
pywin32 -- and show everything even when it can say nothing.
"""

import logging
import os
import sys
import time

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from greenhill.ui.speaker import Speaker  # noqa: E402


def drain(speaker, expected, timeout=2.0):
    """Wait for the speech thread to work through the queue."""
    deadline = time.time() + timeout
    while time.time() < deadline and len(speaker.spoken) < expected:
        time.sleep(0.01)
    return speaker.spoken


class TestWithoutAWindowsVoice:
    def test_starts_and_stops_cleanly(self):
        speaker = Speaker()
        speaker.start()
        speaker.stop()

    def test_speech_is_logged_when_it_cannot_be_heard(self, caplog):
        speaker = Speaker()
        with caplog.at_level(logging.INFO):
            speaker.start()
            speaker.say('Rain detected.')
            drain(speaker, 1)
            speaker.stop()
        assert 'Rain detected.' in caplog.text

    def test_reports_that_it_is_not_audible(self):
        # So the window can say so rather than implying alerts are covered.
        speaker = Speaker()
        speaker.start()
        drain(speaker, 0)
        time.sleep(0.1)
        speaker.stop()
        if sys.platform != 'win32':
            assert speaker.available is False


class TestQueueing:
    def test_say_never_blocks_the_caller(self):
        # The window must not wait for a sentence: SAPI blocks for as long as
        # it takes to say the words, and a display that freezes every time it
        # rains is a display nobody trusts.
        speaker = Speaker()
        speaker.start()
        started = time.time()
        for index in range(50):
            speaker.say('sentence {}'.format(index))
        assert time.time() - started < 0.5
        speaker.stop()

    def test_everything_queued_is_spoken_in_order(self):
        speaker = Speaker()
        speaker.start()
        for index in range(5):
            speaker.say('sentence {}'.format(index))
        spoken = drain(speaker, 5)
        speaker.stop()
        assert spoken[:5] == ['sentence {}'.format(i) for i in range(5)]

    def test_muting_drops_everything(self):
        speaker = Speaker(enabled=False)
        speaker.start()
        speaker.say('Rain detected.')
        time.sleep(0.1)
        speaker.stop()
        assert speaker.spoken == []

    def test_empty_text_is_ignored(self):
        speaker = Speaker()
        speaker.start()
        speaker.say('')
        speaker.say(None)
        time.sleep(0.1)
        speaker.stop()
        assert speaker.spoken == []

    def test_stopping_without_starting_is_harmless(self):
        Speaker().stop()

    def test_starting_twice_is_harmless(self):
        speaker = Speaker()
        speaker.start()
        speaker.start()
        speaker.stop()


class TestFailureIsolation:
    def test_one_failed_sentence_does_not_silence_the_rest(self, caplog):
        # A voice that throws once must not take the thread -- and every later
        # alert -- down with it.
        speaker = Speaker()
        calls = {'n': 0}

        class Flaky:
            def Speak(self, text):
                calls['n'] += 1
                if calls['n'] == 1:
                    raise RuntimeError('voice busy')

        speaker._build_voice = lambda: (Flaky(), None)
        with caplog.at_level(logging.WARNING):
            speaker.start()
            speaker.say('first')
            speaker.say('second')
            drain(speaker, 2)
            speaker.stop()
        assert calls['n'] == 2
        assert 'Speech failed' in caplog.text
