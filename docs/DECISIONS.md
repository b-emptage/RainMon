# Decisions log

Running record of what was decided and why, kept for the report to Bryn.
Results and choices, not implementation detail.

## Architecture

**The Windows 7 box stops deciding anything.** It reads the detectors and
multicasts what they said. All thresholds, latching and dome commands moved to
the Windows 11 box. Result: the safety logic runs on a supported OS, the old
machine needs no Alpaca stack and no installer, and it should never need
touching again.

**The rain sensors now look like the anemometer.** The anemometer was already
multicasting to the LAN and `wind_sensor.py` was only ever a display client, so
the wind side needed nothing built — the Windows 11 box subscribes to the
instrument directly.

**Two Alpaca devices, not one.** `SafetyMonitor` carries the GO/NOGO;
`ObservingConditions` carries wind and rain telemetry. ASCOM has no safety flag
on ObservingConditions, so a single device could not express both honestly.

**Both close routes derive from one fused state.** Direct Alpaca close to the
dome (~1–2 s) is primary; Arcsecond's safety procedure (~90 s) is the backup.
Only Arcsecond reopens — the weather service never does, so the two routes
cannot fight.

## Safety behaviour

**Everything fails closed.** In the system as it stands, a dead serial link, a
disconnected sensor or an errored detector all produce a wet count of zero,
which is indistinguishable from a clear night. Under the new design, absent or
stale data is unsafe, not safe.

**"Alive but broken" and "gone" are different faults.** The bridge keeps
broadcasting when its serial port fails, flagging it. Silence means the machine
itself is gone. Both are unsafe; they send an engineer to different places.

**Only D, w and W count as observations.** A detector that is parked, moving,
initialising or in error reports nothing about the sky and is never counted as
dry.

**Rain rate published to Arcsecond reflects the latched state, not the
instantaneous reading.** Otherwise Arcsecond's recovery procedure reopens the
dome the moment the sensors dry, while the sensors are still wet.

## Constraints settled

* Windows 7 box: Python 3.8, 32-bit, existing PyInstaller chain — pinned, not
  upgraded. The box's own Python 2.7 is irrelevant; a frozen build carries its
  own interpreter.
* Weather Alpaca server on port **11112** — 11111 is taken by the Windows 11
  equipment server.
* Wind thresholds are entered in Arcsecond in **m/s**: 5.56 and 8.33, not 20
  and 30.
* Arcsecond condition set must be mode **ANY**, not the default ALL, or rain
  alone never fires.
* RESET stays a manual operation at the machine.

## Raised for Arcsecond (separate track) — all three fixed

Three issues in `arcsecond-back` undermined the backup route. **All are fixed
and merged into `staging`.**

1. An unreachable sensor wrote an empty weather point, whose missing values
   were skipped by the evaluator, which then decided GO — **sensor failure read
   as good weather.** Live outdoor readings now fail closed.
2. No staleness bound — the newest reading counted as live regardless of age.
   Readings now expire after five minutes.
3. No `IsSafe` quantity, so a SafetyMonitor could not reach a safety decision.
   `site__is_safe` now exists.

Item 1 mattered most: the Windows 11 box hosts arcsecond-local, the equipment
server and the weather service, so it is a single point of failure for both
close routes.

**Accepted, and out of scope here:** a dome-side deadman — the dome closes
itself if it has not heard from the weather service in N minutes. It is the
only protection that survives the Windows 11 box failing, and the observatory
has taken it into its upgrade plan.

## Phase 1 — rain bridge (done)

**The bridge polls every detector every cycle** instead of probing once at
startup. The old software froze its detector list at boot, so a unit unplugged
at that moment stayed dead until someone restarted the software. A detector
that fails repeatedly is backed off so it cannot stretch the one-second
cadence.

**Two latent crashes fixed** in the serial read: one noise byte on the line
could take the process down, and a device streaming without a terminator could
exhaust memory.

**Delivered:** the bridge, a stream recorder, a seven-scenario simulator,
76 hardware-free tests, and operator documentation.

**Open question for Bryn** *(settled — see Phase 5)*: the anemometer's sentence
format is undocumented. The current display reads direction and speed from fixed
positions 2 and 4 behind a silent catch-all. Running the recorder for a day
settles it, and every wind threshold depends on the answer.

## Phase 2 — the safety verdict (done)

**The rain rule now enforces the ten-second window that was always intended.**
In the current software a single wet section that persisted for an hour and
then dried still closed the dome, because a rescheduling timer meant the window
was never really applied. A drop that evaporates quickly is rain; a section
that stays wet is contamination and stays contamination. The window is
configurable if the observatory prefers the old behaviour.

**A gust is a three-second average, not a single sample.** An ultrasonic head
produces occasional spiky readings, and a threshold that fired on one of those
would close the dome on sensor noise. This is also the ASCOM definition, so the
figure published to Arcsecond and the figure used locally are the same number.

**Each cause latches separately, with a delay that fits it** — ten minutes for
rain because the sensors take two to five to dry, two minutes for wind because
nothing has to dry, thirty seconds for a stream that dropped out because a
network blip should not cost the observatory ten minutes of sky. A single delay
would have been wrong for two of the three.

**A latch is never released by the absence of evidence.** Found while testing:
if it rained and then the bridge died, the ten-minute countdown ran through the
blackout and released — declaring the observatory safe on the strength of ten
minutes in which nobody could see the sensors. The countdown now freezes
whenever the sensors cannot be read, so the clear period is ten minutes of
*observed* dryness.

**A blind rain sensor publishes rain.** Arcsecond's evaluator skips a condition
whose value is missing and, with nothing else triggered, decides GO. So a
device that can be reached but cannot see would read to Arcsecond as good
weather. Publishing the wet value instead turns that into a NOGO. This is a
workaround for the Arcsecond issue raised above, and it does not cover the
weather service itself being unreachable.

**Two dead detectors are enough to stop trusting the reading.** With three
installed, one failure does not blind the observatory, but a single survivor is
not asked to speak for the whole site.

**Replay is how the rules get checked.** Any recording from the observatory can
be run through the complete safety logic at speed, with no hardware and no
network — so a rule change can be tested against a night that actually
happened rather than against anyone's memory of it.

**Delivered:** rain and wind decoding, the fusion and latching, a live monitor
and a replay tool, and 158 hardware-free tests. Verified end to end over real
multicast.

**Still provisional, pending the recorder run** *(both since settled)*: the
anemometer's field positions and its 30-degree north offset. The capture
replaced the field indices with real NMEA parsing, and the observatory has
confirmed the offset is correct.

## Phase 3 — the Alpaca devices (done)

**Two devices on one server, port 11112.** `SafetyMonitor` carries the GO/NOGO,
`ObservingConditions` carries the telemetry, and both are views onto a single
safety evaluator so they cannot disagree. Port 11112 because the equipment
Alpaca server on that machine already holds 11111; for the same reason UDP
discovery will not bind there, which is expected and harmless.

**The monitoring runs whether or not any client is connected**, and `Connected`
defaults to true. Two reasons: Arcsecond polls `Connected` but never sets it, so
a device waiting to be connected would be recorded as disconnected forever and
never yield a reading; and the direct dome-close route has no ASCOM client of
its own, so tying the thread to a client connection would quietly disarm it.
This is a deliberate difference from the dome server, where disconnecting
de-energises the motors on purpose.

**`IsSafe` never returns an error.** Most clients treat an error from a safety
monitor as "no opinion" and carry on observing, so a failure there would read as
permission. It answers false and puts the reason elsewhere.

**Sensors this site does not have report "not implemented" rather than zero.**
Arcsecond caches that and stops asking. A fabricated humidity would otherwise be
recorded every minute as though it were measured.

**A vendor action publishes the reason.** ASCOM gives a safety monitor one bit,
and this site has four independent reasons to close. `Greenhill:GetWeatherStatus`
returns which condition is holding, how each detector reads and how old both
streams are — so an operator can see why without reading the log.

**Two defects found in the ASCOM sample code we inherited.** Its "Exception"
classes are not Python exceptions at all but response payloads, so raising one
fails; and its uncaught-exception handler used a calling convention Falcon
dropped three major versions ago, meaning it crashed instead of reporting
whatever had gone wrong. Both fixed here, and the second one fixed in
Greenhill-DomeShutter as well, which inherited it from the same sample.

**Conform runs in CI against both device types**, plus a Python 3.8 job that
byte-compiles and tests everything shipped to the Windows 7 box — the shared
protocol module is imported by both machines, and that job is what stops a
newer idiom reaching the one that cannot run it.

**Delivered:** both Alpaca devices, the monitoring service, a simulator, CI for
conformance and for the Windows 7 interpreter, and 213 hardware-free tests.
Verified end to end over real multicast: sensor stream to fused verdict to
latch to HTTP.

**Not yet verified:** Conform has only been run against the simulator. It should
be run against the real installation before the observatory relies on it.

## Phase 4 — the direct dome close (done)

**Both routes fire, and that is the point.** This service closes the dome
directly in a second or two; Arcsecond reaches the same conclusion about ninety
seconds later and closes it again. Closing a closed dome is a no-op, so the
overlap costs nothing and neither route is a single point of failure.

**It only ever closes.** Reopening stays with Arcsecond's recovery procedure, so
the two routes can never argue about whether the dome should be open.

**It never disconnects from the dome.** Disconnecting de-energises the dome's
motors — by design, so a client dropping out cannot leave a shell running
unwatched — which means disconnecting mid-close would stop the close.

**It waits for the dome to confirm it is shut**, rather than assuming the
command worked. The dome reports a half-open shell as "open", deliberately, so
that is what "not yet closed" looks like and the closer keeps waiting.

**A latched fault does not stop it trying.** The dome guarantees that closing is
never blocked by a fault, because it must not be possible to lock the roof open.
The close is issued anyway and logged as needing an engineer.

**An unreachable dome escalates once and then keeps trying quietly.** It never
gives up — the roof still has to shut — but a critical message repeated every
second buries itself.

**Off by default.** This is the only thing in the package that commands a roof.
It must be turned on deliberately, after a dry-run close on the real dome. The
startup log shouts whichever state it is in, so a forgotten setting cannot pass
unnoticed, and simulated mode refuses to arm it at all — bench and conformance
runs must never be able to drive a real dome on invented weather.

**Two things to be aware of operationally:**

*Restarting the weather service will close the dome.* It starts unsafe because
it has not yet seen the sky. That is the fail-closed design working, but it
means a restart mid-night shuts the roof until conditions have been good for the
settle period. The close log now says why, so it is self-explanatory in the
morning. Changing this would mean relying on Arcsecond alone for the first
thirty seconds after a restart — a trade, not an improvement, and the
observatory's call.

*The dome now has two Alpaca clients.* Originally they shared one `Connected`
flag, so either could de-energise the dome's motors under the other, possibly
mid-close. **Fixed:** the dome tracks connection state per client and releases
the hardware only when the last one disconnects.

**Delivered:** the Alpaca dome client and the close state machine, wired into
the weather service and reportable from the ASCOM surface, with 249 tests.
Verified end to end against the real dome server on its simulated board — rain
on the wire, through the fused verdict, to a confirmed closed dome.

**Corrected along the way:** the plan for this work said to wait for
`shutterClosed (4)`. Closed is 1; 4 is the error state. A closer built on that
would have reported success on a faulted dome and stopped watching.

**Still to do before arming:** a dry-run close on the actual dome, and a check
of its real travel time against the 45 s re-issue window.


## Follow-ups on the dome server

Two changes were made to `Greenhill-DomeShutter` while building the weather
service, both prompted by it being the dome's second Alpaca client.

**Connection is now per client.** ASCOM has one `Connected` property and the
dome now has two clients — Arcsecond, and the weather service that closes it.
Sharing one flag meant either could de-energise the motors under the other,
possibly mid-close. The board now opens for the first client and is released
only when the last lets go, which preserves the rule that mattered (a shell
must never be left running with nobody watching) while distinguishing "a client
went away" from "everyone went away". Clients that go silent for five minutes
are expired, because Alpaca is stateless HTTP and a crashed client never says
goodbye — and because Arcsecond's client library picks a random identity at
startup, so it presents a new one after every restart.

**The error handler no longer loses the error.** It used a calling convention
Falcon dropped three major versions ago, so the handler whose job was reporting
a fault crashed instead — discarding the real exception. Found by building a
second server on the same code and watching a responder fail; conformance
testing could not have found it, because Conform never provokes a driver fault.

## On fail-closed

Confirmed and kept as it stands. The system reports unsafe unless it can
positively see that conditions are good, and this is now the first thing both
the README and the safety document say. The consequence the observatory should
expect: **restarting the weather service closes the dome**, because it comes up
having not yet seen the sky. That is the design working, not a fault, but it is
disruptive if unexpected.
## Phase 5 — Arcsecond (done)

**The three holes in Arcsecond's own safety evaluator are fixed**, and merged
into `staging` in `arcsecond-back`. The first of them mattered
most: an unreachable sensor wrote a weather reading with no values, the
evaluator skipped conditions it could not evaluate, and with nothing triggered
the answer was GO — so a sensor failure was indistinguishable from good weather.
Live outdoor readings now fail closed. Readings also expire after five minutes,
where before the newest row counted as current however old it was. And
`site__is_safe` now exists, so a SafetyMonitor can actually stop the
observatory; until now one was polled, stored, and never consulted.

**One condition, not four.** The recommended Greenhill set is a single
`site__is_safe Smaller 1`. The SafetyMonitor has already fused rain, wind,
sensor health and stream staleness, so restating those thresholds in Arcsecond
would put the same numbers in two places that can drift — and the local copy is
the one that also drives the direct close, so a drift would have the two routes
disagreeing about when to close.

**The anemometer was captured, and it changed the wind code.** Fifteen minutes
of real traffic showed the instrument is an Observator OMC-140 sending NMEA MWV
inside a proprietary wrapper. **The parser built on assumption would have
rejected every single datagram** — it would have left wind permanently unknown
and the site permanently unsafe. Nothing short of real traffic would have found
that. The sentence is now parsed as NMEA, so units and validity are read from
the packet rather than assumed, and the north offset is applied only to the
relative bearing it was meant for.

The rain stream over the same window was flawless: 893 packets, no gaps, all
three detectors reporting dry.

**The wind thresholds were queried and confirmed.** Against 20 km/h sustained
and 30 km/h gust, the site would have been NOGO for that entire quarter of an
hour — over the sustained limit in 95% of evaluations and the gust limit in
88%. That was raised in case the limits were too tight for an exposed ridge;
the observatory confirms they are the intended figures, and it was simply that
windy.

**The two devices are registered in arcsecond-local and reporting.** The
checklist that got them there is in `docs/ARCSECOND_SETUP.md`.

## Phase 6 — the weather window (done)

The astronomer keeps the display and the spoken alerts, and they no longer
depend on anything else being alive.

**It listens and nothing else.** The window joins the same two multicast
streams everything else reads and draws what they say. It holds no serial port,
sends no dome commands, and never talks to Arcsecond. Three consequences worth
having:

* it runs on **any machine on the LAN, and on several at once** — the old
  monitor could only run on the box holding the serial port;
* it keeps working when the weather service or Arcsecond are down, which is
  exactly when someone wants to look at it;
* **closing it stops nothing.** In the software this replaces the window *was*
  the safety system, and closing it stopped the dome protection with it.

**The display is the one the operators already read** — the same site map, the
same status colours, the same compass and the same green-to-red speed ramp.
Readouts stay in km/h because that is what the observatory thinks in;
everything beneath, and everything published over Alpaca, is m/s, and the
conversion happens in one place.

**The window's verdict is its own, and it says so.** It computes from the same
core as the weather service, so the numbers agree, but it does not control the
dome. Pointed at the weather service it also shows the observatory's real
verdict and flags a disagreement — the two latches are not in step, having
started at different times, and an astronomer looking at a green panel beside a
closed roof deserves an explanation rather than a puzzle.

**Three deliberate silences**, because the failure mode of a spoken alert is
not silence but nagging — one that repeats gets muted, and a muted alert
protects nobody. Nothing is said on the first update, or the window would greet
whoever opens it with rain that stopped an hour ago. Nothing is called broken
for the first twenty seconds, because every source looks dead before its first
packet. And "rain sensors dry" is only ever said from a known-wet state to a
known-dry one: wet then unreadable is a fault, and announcing dryness there
would be the most misleading sentence the system could utter.

**The new announcement is the fault one** — "rain sensors not responding". The
old monitor could not say it, because it could not tell a dry night from a dead
sensor. It is the alert to keep if only one is kept.

**It runs without a voice.** On a Mac, a Linux box, or Windows without pywin32,
alerts are logged and everything else works. The window is kept 3.8-compatible
so it can also run on the Windows 7 box, where the display it replaces has
always lived.

**Delivered:** the window, the alert policy, the speaker, a frozen-build spec,
27 new tests (381 in total), and the safety core added to the Python 3.8 CI job
so nothing that ships to the old machine can drift past it.

---

## Where things stand

The phase sections above are a record of what was decided as the work went
along, and some of their open questions have since been answered in place. This
section is the one to trust for what is actually outstanding.

**All six phases are complete.** The rain bridge runs on the Windows 7 box, the
weather service and its two Alpaca devices run on the Windows 11 box, both are
registered in arcsecond-local and reporting, and the weather window is
available to anyone who wants it. 381 tests.

**Settled since they were first raised:**

* The anemometer's format — a capture identified an Observator OMC-140 sending
  NMEA MWV inside a proprietary wrapper, and the parser was rewritten around it.
* The 30-degree north offset — confirmed correct by the observatory.
* The wind thresholds — queried after a capture showed the site would have been
  NOGO throughout; confirmed as the intended figures.
* The three fail-open holes in Arcsecond's safety evaluator — fixed and merged.
* Per-client Alpaca connections on the dome, and its error handler.

**Still outstanding:**

1. **The direct dome close is not armed.** `dome_close_enabled = false` and
   `dome_address` is empty. Arm it after a dry-run close on the real dome, and
   check the dome's true travel time against the 45 s re-issue window while you
   are there.
2. **ASCOM Conform has only been run against the simulator**, never against the
   real installation.
3. **A dome-side deadman**, which the observatory has taken into its upgrade
   plan. It is the only protection that survives the Windows 11 box failing,
   and it is not part of this work.
