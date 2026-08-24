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

## Raised for Arcsecond (separate track)

Three issues in `arcsecond-back` that undermine the backup route:

1. An unreachable sensor writes an empty weather point, whose missing values
   are skipped by the evaluator, which then decides GO. **Sensor failure
   currently reads as good weather.**
2. No staleness bound — the newest reading counts as live regardless of age.
3. No `IsSafe` quantity, so a SafetyMonitor cannot reach a safety decision.

Item 1 matters most: the Windows 11 box now hosts arcsecond-local, the
equipment server and the weather service, so it is a single point of failure
for both close routes.

**Recommended, out of scope here:** a dome-side deadman — the dome closes
itself if it has not heard from the weather service in N minutes. It is the
only protection that survives the Windows 11 box failing.

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

**Open question for Bryn:** the anemometer's sentence format is undocumented.
The current display reads direction and speed from fixed positions 2 and 4
behind a silent catch-all. Running the recorder for a day settles it, and
every wind threshold depends on the answer.

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

**Still provisional, pending the recorder run:** the anemometer's field
positions and its 30-degree north offset. Both are configuration now, so
correcting them is a config edit rather than a release — but wind direction
should not be trusted until the offset has been checked against a known
reference.

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
whatever had gone wrong. Both fixed here. **The second one is still present in
Greenhill-DomeShutter and should be fixed there too.**

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
