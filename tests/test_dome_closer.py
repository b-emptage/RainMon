# -*- coding: utf-8 -*-
"""Route 1: the direct dome close.

This is the module that commands a roof, so the tests are mostly about
restraint -- what it must NOT do -- and about what happens when the dome does
not answer.

The Alpaca client is replaced by a fake dome that models the two behaviours
that matter: a close takes time, and the dome reports a shell stopped part way
as `open`, never as `closed`.
"""

import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from greenhill.core import dome_client as D  # noqa: E402
from greenhill.core.config import WeatherConfig  # noqa: E402


class FakeDome:
    """A dome that answers over the same surface AlpacaDomeClient exposes."""

    def __init__(self, status=D.SHUTTER_OPEN, connected=False):
        self.base = 'http://dome.test/api/v1/dome/0'
        self.status = status
        self.connected = connected
        self.close_commands = 0
        self.connect_commands = 0
        self.disconnect_commands = 0
        self.fail_with = None           # set to a message to make calls fail
        self.close_takes = 0            # ticks of "closing" before it shuts
        self._closing_ticks = 0

    def _maybe_fail(self):
        if self.fail_with:
            raise D.DomeError(self.fail_with)

    def is_connected(self):
        self._maybe_fail()
        return self.connected

    def ensure_connected(self):
        self._maybe_fail()
        if not self.connected:
            self.connected = True
            self.connect_commands += 1

    def disconnect(self):                       # never called, and asserted so
        self.disconnect_commands += 1

    def shutter_status(self):
        self._maybe_fail()
        if self.status == D.SHUTTER_CLOSING:
            self._closing_ticks += 1
            if self._closing_ticks >= self.close_takes:
                self.status = D.SHUTTER_CLOSED
        return self.status

    def close_shutter(self):
        self._maybe_fail()
        self.close_commands += 1
        if self.close_takes:
            self.status = D.SHUTTER_CLOSING
            self._closing_ticks = 0
        else:
            self.status = D.SHUTTER_CLOSED


class Clock:
    def __init__(self):
        self.now = 0.0

    def __call__(self):
        return self.now

    def advance(self, seconds):
        self.now += seconds


class Log:
    """Captures log lines so the tests can assert what an operator would see."""

    def __init__(self):
        self.lines = []

    def _record(self, level, message, *args):
        self.lines.append((level, message % args if args else message))

    def info(self, m, *a): self._record('info', m, *a)
    def warning(self, m, *a): self._record('warning', m, *a)
    def error(self, m, *a): self._record('error', m, *a)
    def critical(self, m, *a): self._record('critical', m, *a)
    def exception(self, m, *a): self._record('error', m, *a)

    def text(self, level=None):
        return '\n'.join(line for lvl, line in self.lines
                         if level is None or lvl == level)


class Rig:
    """A closer driven a tick at a time, with no threads."""

    def __init__(self, dome=None, safe=True, **overrides):
        settings = dict(dome_poll_interval_s=2.0, dome_verify_timeout_s=45.0,
                        dome_retry_limit=3, dome_escalated_retry_s=60.0)
        settings.update(overrides)
        self.dome = dome or FakeDome()
        self.safe = safe
        self.clock = Clock()
        self.log = Log()
        self.closer = D.DomeCloser(self.dome, lambda: self.safe, self.log,
                                   WeatherConfig(**settings), clock=self.clock)

    def tick(self, times=1, advance=2.0):
        for _ in range(times):
            self.closer._tick()
            self.clock.advance(advance)


class TestRestraint:
    """What it must not do."""

    def test_does_nothing_while_conditions_are_safe(self):
        rig = Rig(safe=True)
        rig.tick(10)
        assert rig.dome.close_commands == 0
        assert rig.dome.connect_commands == 0

    def test_does_nothing_before_the_verdict_is_known(self):
        # None means this closer started before the weather service did.
        # Commanding a roof on no information at all is worse than waiting.
        rig = Rig()
        rig.safe = None
        rig.tick(10)
        assert rig.dome.close_commands == 0

    def test_never_reopens(self):
        # Reopening is an observing decision and belongs to Arcsecond's
        # recovery procedure. One opener, no argument.
        rig = Rig(safe=False)
        rig.tick(3)
        assert rig.dome.status == D.SHUTTER_CLOSED
        rig.safe = True
        rig.tick(10)
        assert rig.dome.status == D.SHUTTER_CLOSED

    def test_never_disconnects(self):
        # Disconnect de-energises the dome's motors, so doing it while a close
        # is in progress would stop the close.
        rig = Rig(safe=False)
        rig.tick(5)
        rig.safe = True
        rig.tick(5)
        rig.closer.stop()
        assert rig.dome.disconnect_commands == 0

    def test_does_not_command_an_already_closed_dome(self):
        # Route 2 fires too, and may well get there first. Closing a closed
        # dome must be a no-op, not a second command.
        rig = Rig(dome=FakeDome(status=D.SHUTTER_CLOSED), safe=False)
        rig.tick(10)
        assert rig.dome.close_commands == 0

    def test_does_not_re_issue_while_the_dome_is_already_closing(self):
        dome = FakeDome(status=D.SHUTTER_OPEN)
        dome.close_takes = 5
        rig = Rig(dome=dome, safe=False)
        rig.tick(4)
        assert dome.close_commands == 1

    def test_stays_quiet_once_closed(self):
        rig = Rig(safe=False)
        rig.tick(20)
        assert rig.dome.close_commands == 1


class TestClosing:
    def test_closes_when_conditions_go_unsafe(self):
        rig = Rig(safe=True)
        rig.tick(2)
        rig.safe = False
        rig.tick(2)
        assert rig.dome.close_commands == 1
        assert rig.dome.status == D.SHUTTER_CLOSED

    def test_connects_first_if_it_is_not_connected(self):
        rig = Rig(safe=False)
        rig.tick(2)
        assert rig.dome.connect_commands == 1

    def test_waits_for_a_positive_closed(self):
        # The dome reports a shell stopped part way as `open`, deliberately --
        # ASCOM has no partial state, and calling a dome shut when it is not is
        # the one error that leaves it open in the rain.
        dome = FakeDome(status=D.SHUTTER_OPEN)
        dome.close_takes = 3
        rig = Rig(dome=dome, safe=False)
        rig.tick(2)
        assert rig.closer.status()['state'] == D.CLOSING
        rig.tick(4)
        assert rig.closer.status()['state'] == D.CLOSED

    def test_reports_the_close_at_warning_level(self):
        # An operator should see this without turning the log level up.
        rig = Rig(safe=False)
        rig.tick(3)
        assert 'DOME CLOSING' in rig.log.text('warning')
        assert 'DOME CLOSED' in rig.log.text('warning')

    def test_closes_again_on_a_second_unsafe_episode(self):
        rig = Rig(safe=False)
        rig.tick(3)
        rig.safe = True
        rig.tick(3)
        rig.dome.status = D.SHUTTER_OPEN        # Arcsecond reopened it
        rig.safe = False
        rig.tick(3)
        assert rig.dome.close_commands == 2


class TestFaultedDome:
    def test_closes_anyway_when_the_dome_reports_a_fault(self):
        # The dome documents that a latched fault NEVER blocks closing -- it
        # must not be possible to lock the roof open.
        rig = Rig(dome=FakeDome(status=D.SHUTTER_ERROR), safe=False)
        rig.tick(2)
        assert rig.dome.close_commands == 1

    def test_says_a_fault_needs_an_engineer(self):
        rig = Rig(dome=FakeDome(status=D.SHUTTER_ERROR), safe=False)
        rig.tick(2)
        assert 'DOME FAULT' in rig.log.text('error')


class TestUnreachableDome:
    def test_retries_a_failing_close(self):
        dome = FakeDome()
        dome.fail_with = 'connection refused'
        rig = Rig(dome=dome, safe=False)
        rig.tick(3)
        assert rig.closer.status()['state'] == D.FAILING
        assert rig.closer.status()['attempts'] >= 3

    def test_escalates_once_past_the_retry_limit(self):
        dome = FakeDome()
        dome.fail_with = 'connection refused'
        rig = Rig(dome=dome, safe=False, dome_retry_limit=2,
                  dome_escalated_retry_s=60.0)
        rig.tick(6, advance=60.0)
        critical = rig.log.text('critical')
        assert 'DOME UNREACHABLE' in critical
        assert 'MAY BE OPEN IN BAD WEATHER' in critical
        # Once, not once per attempt: the message that matters must not be
        # buried under copies of itself.
        assert critical.count('DOME UNREACHABLE') == 1

    def test_never_gives_up(self):
        # The roof still has to shut. It backs off; it does not stop.
        dome = FakeDome()
        dome.fail_with = 'connection refused'
        rig = Rig(dome=dome, safe=False, dome_retry_limit=2,
                  dome_escalated_retry_s=60.0)
        rig.tick(10, advance=60.0)
        before = rig.closer.status()['attempts']
        rig.tick(5, advance=60.0)
        assert rig.closer.status()['attempts'] > before

    def test_recovers_when_the_dome_comes_back(self):
        dome = FakeDome()
        dome.fail_with = 'connection refused'
        rig = Rig(dome=dome, safe=False, dome_retry_limit=2)
        rig.tick(4, advance=60.0)
        dome.fail_with = None
        rig.tick(3, advance=60.0)
        assert dome.close_commands == 1
        assert rig.closer.status()['state'] == D.CLOSED

    def test_a_failure_to_read_does_not_stop_the_close_being_tried(self):
        # If we cannot confirm the dome is shut, we act as though it is not.
        dome = FakeDome(status=D.SHUTTER_CLOSED)
        dome.fail_with = 'timeout'
        rig = Rig(dome=dome, safe=False)
        rig.tick(2)
        assert rig.closer.status()['state'] == D.FAILING


class TestAlpacaEnvelope:
    """The client's handling of what a real dome sends back."""

    def make(self, payload):
        client = D.AlpacaDomeClient('dome.test:11111')
        return client, payload

    def test_an_alpaca_error_becomes_a_dome_error(self):
        client = D.AlpacaDomeClient('dome.test:11111')
        with pytest.raises(D.DomeError, match='0x40B'):
            client._unwrap(
                b'{"Value":null,"ErrorNumber":1035,"ErrorMessage":"bad state"}',
                'PUT closeshutter')

    def test_a_successful_envelope_yields_its_value(self):
        client = D.AlpacaDomeClient('dome.test:11111')
        assert client._unwrap(
            b'{"Value":1,"ErrorNumber":0,"ErrorMessage":""}', 'GET x') == 1

    def test_a_non_json_reply_is_a_dome_error(self):
        client = D.AlpacaDomeClient('dome.test:11111')
        with pytest.raises(D.DomeError, match='not JSON'):
            client._unwrap(b'<html>proxy error</html>', 'GET x')

    def test_transaction_ids_increment(self):
        client = D.AlpacaDomeClient('dome.test:11111')
        assert client._next_transaction() < client._next_transaction()

    def test_the_url_is_the_alpaca_dome_endpoint(self):
        assert D.AlpacaDomeClient('10.0.0.9:11111', 0).base == \
            'http://10.0.0.9:11111/api/v1/dome/0'


class TestShutterStateValues:
    def test_closed_is_one_and_error_is_four(self):
        # Worth pinning down. The plan for this work said "closed (4)", which
        # is the ERROR state -- a closer built on that would have reported
        # success on a faulted dome and stopped watching.
        assert D.SHUTTER_CLOSED == 1
        assert D.SHUTTER_ERROR == 4


class TestTheTwoClocks:
    """Polling and re-issuing are separate, and conflating them is a real bug.

    The first version of this closer used one timer for both, so after
    commanding a close it stopped looking at the dome for the whole verify
    timeout -- 45 seconds of not watching, beginning at the exact moment it had
    most reason to watch.
    """

    def test_keeps_watching_after_commanding_a_close(self):
        dome = FakeDome(status=D.SHUTTER_OPEN)
        dome.close_takes = 2
        rig = Rig(dome=dome, safe=False, dome_poll_interval_s=2.0,
                  dome_verify_timeout_s=45.0)
        rig.tick(1)
        assert dome.close_commands == 1
        # Well inside the 45 s re-issue window, it has still noticed the close
        # complete.
        rig.tick(4, advance=2.0)
        assert rig.closer.status()['state'] == D.CLOSED

    def test_does_not_re_issue_inside_the_grace_period(self):
        # A dome that stays `open` for a while is not necessarily ignoring us;
        # it may simply be slow. Stacking commands on a moving roof helps
        # nobody.
        dome = FakeDome(status=D.SHUTTER_OPEN)

        def refuse_to_move():
            dome.close_commands += 1        # accepted, but nothing happens
        dome.close_shutter = refuse_to_move

        rig = Rig(dome=dome, safe=False, dome_poll_interval_s=2.0,
                  dome_verify_timeout_s=45.0)
        rig.tick(10, advance=2.0)           # 20 s of polling
        assert dome.close_commands == 1

    def test_re_issues_once_the_grace_period_expires(self):
        dome = FakeDome(status=D.SHUTTER_OPEN)

        def refuse_to_move():
            dome.close_commands += 1
        dome.close_shutter = refuse_to_move

        rig = Rig(dome=dome, safe=False, dome_poll_interval_s=2.0,
                  dome_verify_timeout_s=10.0)
        rig.tick(12, advance=2.0)           # 24 s: two grace periods
        assert dome.close_commands >= 2


class TestTheLogExplainsItself:
    def test_the_close_line_says_why(self):
        # A shut dome found in the morning should not need two logs and a
        # timestamp comparison to explain itself.
        rig = Rig(safe=False)
        rig.closer._reasons = lambda: ['rain: 3 wet sections (threshold 2)']
        rig.tick(2)
        assert 'rain: 3 wet sections' in rig.log.text('warning')

    def test_a_broken_reason_source_does_not_stop_the_close(self):
        # The reason is for the log. It must never be able to prevent the
        # thing it is describing.
        def explode():
            raise RuntimeError('reasons unavailable')
        rig = Rig(safe=False)
        rig.closer._reasons = explode
        rig.tick(2)
        assert rig.dome.close_commands == 1
