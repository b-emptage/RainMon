# Route 1 — the direct dome close

When the safety verdict goes false, the weather service commands the dome
server on the other machine to close, then watches until the dome confirms it
is shut. A second or two, end to end.

Arcsecond reaches the same conclusion independently and runs its own close
procedure — but on a 60 s weather poll plus a 30 s safety heartbeat, so up to
about ninety seconds behind. **Both routes are meant to fire.** Closing a closed
dome is a no-op, so the overlap costs nothing, and what it buys is that neither
route is a single point of failure.

## What it does

1. Ensure `Connected` on the dome (it refuses to report `ShutterStatus`
   otherwise).
2. Read `ShutterStatus`. If already `shutterClosed`, do nothing.
3. Otherwise `CloseShutter`, then keep polling until the dome reports
   `shutterClosed`.
4. If it never gets there, retry — then escalate and keep retrying slowly.

`shutterClosed` is **1**. `4` is `shutterError`. The plan for this work
originally said "closed (4)", which would have had the closer report success on
a faulted dome and stop watching.

## Three rules it is built around

**It only ever closes.** Reopening is an observing decision and belongs to
Arcsecond's recovery procedure. One opener, no argument. If conditions clear,
this service logs a stand-down and does nothing else.

**It never disconnects.** The dome de-energises its motors on `Disconnect` —
deliberately, so a client dropping its connection cannot leave a shell running
unsupervised. Which means disconnecting *while a close is running would stop
the close*. This client sets `Connected` once and leaves it alone, including on
shutdown.

**It waits for a positive "closed".** The dome reports a shell stopped half way
as `shutterOpen`, on purpose: ASCOM has no partial state, and calling a dome
shut when it is not is the one error that leaves it open in the rain. So
anything other than `shutterClosed` reads as "not yet", which is what this
waits for.

## Two clocks, not one

Polling and re-issuing are separate, and conflating them was a real bug during
development: sharing one timer meant that after commanding a close the closer
stopped looking at the dome for the whole verify timeout — 45 seconds of not
watching, beginning at the moment it had most reason to watch.

| Setting | Default | What it governs |
|---|---|---|
| `dome_poll_interval_s` | 2 s | how often to look at the dome |
| `dome_verify_timeout_s` | 45 s | how long a close may run before commanding another |

`dome_verify_timeout_s` **must comfortably exceed the dome's full travel
time**, or a slow close collects a second command halfway through. Measure the
real travel before trusting the default.

## When things go wrong

| Situation | Behaviour |
|---|---|
| Dome already closing | Keep watching; do not command again. Its own supervision is on the limit switches. |
| Dome reports `shutterError` | **Issue the close anyway.** The dome documents that a latched fault never blocks closing — it must not be possible to lock the roof open. Logged at ERROR: somebody has to go and look either way. |
| Cannot read the dome | Treated as "not closed". If we cannot confirm it is shut, we act as though it is not. |
| Close keeps failing | Retries `dome_retry_limit` times fast, then **escalates once** at CRITICAL and keeps retrying every `dome_escalated_retry_s`. It never gives up — the roof still has to shut — but it stops filling the log, because a message repeated at 1 Hz buries itself. |

## Arming it

**Off by default.** `dome_close_enabled = false` in `device/config.toml`. This
is the only setting in the package that commands a roof, and a service that
started driving one the moment it was installed — before anyone had checked the
address, watched a dry-run close, or confirmed the dome was answering — would be
the wrong kind of helpful.

The startup log shouts either way, so a forgotten `false` cannot pass unnoticed:

```
==DOME CLOSE ARMED== route 1 will close http://10.0.0.9:11111/api/v1/dome/0 ...
==DOME CLOSE NOT ARMED== dome_close_enabled is false, so this service will NOT
   close the dome itself. Arcsecond is the only route, and it runs about 90
   seconds behind.
```

`Greenhill:GetWeatherStatus` reports the same thing under `domeClose`, so it can
be checked from the ASCOM surface without logging in to the machine.

**Simulated mode never arms it**, whatever the config says. `simulate.py` is
what Conform and bench work run against, on machines that may well be able to
reach the real dome, and a server running on invented weather must not be in a
position to command a roof.

## Operational consequence worth knowing

**The service starts unsafe, so restarting it will close the dome.** That is
the fail-closed design working as intended — it has not yet seen the sky, and
it does not assume. But it means restarting the weather service mid-night shuts
the roof, and it stays shut until conditions have been good for the settle
period and Arcsecond's recovery procedure reopens it.

The close log line says why, so this is at least self-explanatory in the
morning:

```
==DOME CLOSING== rain_data: rain data is 15s old. The dome reads open;
                 CloseShutter issued (attempt 1).
```

If that proves disruptive in practice, the alternative is to arm route 1 only
after the service has been safe once — protection would then rely on Arcsecond
for the first thirty seconds after a restart. That is a trade, not an
improvement, and it is the observatory's call.

## Connection is per client — fixed

ASCOM has one `Connected` property, and the dome now has two clients:
Arcsecond, and this service. Sharing a single flag meant either one setting it
false would **de-energise the dome's motors under the other** — possibly
mid-close.

`Greenhill-DomeShutter` now tracks connection state per ClientID. The board
opens for the first client and is released only when the last one lets go, so
Arcsecond disconnecting can no longer interrupt a close in progress. The rule
that was being protected — a shell must never be left running with nobody
watching — still holds; it just distinguishes "a client went away" from
"everyone went away".

This client uses ClientID **1782** and never disconnects, including on
shutdown. It does not need to: the dome expires clients that have gone silent
for five minutes, so a crash here releases the board on its own.

Verified against the running dome server: both clients connect, Arcsecond
disconnects, and this service can still read `ShutterStatus` and command
`CloseShutter`.

## Testing

```bash
python -m pytest tests/test_dome_closer.py -q
```

No network. The dome is replaced by a fake that models the two behaviours that
matter: a close takes time, and a partially open shell reports `open`.

Verified end to end against the **real dome server** running on its simulated
board: rain on the multicast stream, through the fused verdict, to
`CloseShutter`, to a confirmed `shutterClosed`.

**Not yet done:** a dry-run close against the actual dome. Do that before
arming it, and check the real travel time against `dome_verify_timeout_s` while
you are there.
