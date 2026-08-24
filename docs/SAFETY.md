# The safety verdict

What decides whether the Greenhill dome closes. Everything else — the two
Alpaca devices, both close routes, the display — is a view onto this.

Runs on the Windows 11 box, alongside arcsecond-local and the equipment Alpaca
server. It reads two multicast streams and produces one boolean.

---

## Fail closed — read this first

**The system reports UNSAFE unless it can positively see that conditions are
good.** Not "safe unless it detects a problem". The difference is the whole
point of the rewrite, and it changes what several ordinary events do.

Every one of these reports **unsafe**:

| | |
|---|---|
| The service has just started | It has not looked at the sky yet. |
| Either sensor stream is silent for 15 s | The Windows 7 bridge or the anemometer has stopped. |
| The rain bridge reports its serial port down | It is alive; its sensors are not. |
| Fewer than 2 rain detectors report a usable state | `P`, `M`, `I`, `E` and `e` say nothing about the sky. |
| Too few wind samples to average | Not enough to mean anything. |
| Anything unexpected | The verdict starts false and is only ever cleared by evidence. |

In the software this replaces, **every one of those cases produced a wet count
of zero**, which was indistinguishable from a clear night. A dead serial link,
a disconnected sensor and an errored detector all read as "no rain", and
nothing closed the dome.

### What this means in practice

**Restarting the weather service will close the dome.** It comes up unsafe,
because it has not yet seen anything, and route 1 acts on that. The dome stays
shut until conditions have been good for the settle period and Arcsecond's
recovery procedure reopens it. This is the design working, not a fault — but
it is disruptive if you do not expect it, so **do not restart this service
mid-observation without expecting the roof to close.**

**A latch is never released by the absence of evidence.** If it rains and then
the rain bridge dies, the ten-minute countdown does not run through the
blackout — it freezes, and restarts only when the sensors can be seen again.
The clear period is ten minutes of *observed* dryness, never ten minutes of not
looking.

**A blind rain sensor publishes rain.** When the detectors cannot be read,
`RainRate` is raised rather than dropped, so a device that Arcsecond can reach
but that cannot see still produces a NOGO. See *What gets published* below.

**Both close routes fire.** Closing a closed dome is a no-op, so nothing is
lost, and neither route is a single point of failure.

### What fail-closed does NOT cover

If the **Windows 11 box itself** dies, this service is not running and neither
is arcsecond-local. Nothing tells the dome to close. The only protection that
survives that is a dome-side deadman — the dome closing itself if it has not
heard from the weather service in N minutes. That is recommended and not yet
built; see [DECISIONS.md](DECISIONS.md).

---

## The rules

### Rain

Counted in **sections**, not detectors. Each detector has two that trigger
independently: `w` is one wet, `W` is both. Three detectors are installed, so
six sections — and **one detector reporting `W` reaches the threshold on its
own.**

| Observation | Verdict |
|---|---|
| 2 or more wet sections | **rain** |
| exactly 1 wet section, clears within 10 s | **rain** — a drop evaporating off a heated sensor |
| exactly 1 wet section, persists | not rain — droppings, an insect, a failed sensor |
| exactly 1 wet section, persists then clears | not rain |

The third and fourth rows are the false-positive filter, and the reason the
rule is not simply "any wetness closes the dome".

**One deliberate change from the current software.** Today a single wet section
that persisted for an hour and then dried still closed the dome: a rescheduling
timer meant the ten-second window was never actually enforced. It is enforced
now, which is what was described as the intent. `raindrop_window_s` is
configurable if the observatory wants the old behaviour back.

### Wind

| Measure | Limit |
|---|---|
| sustained — 60 s mean | 5.56 m/s (20 km/h) |
| gust — strongest 3 s average in the last 2 minutes | 8.33 m/s (30 km/h) |

A gust is a short average, not a single sample. An ultrasonic head produces the
odd spiky reading, and a threshold that fired on one of those would close the
dome on sensor noise. This also matches the ASCOM definition, so the figure
published to Arcsecond and the figure used locally are the same number.

Everything is metres per second, everywhere. ASCOM and Arcsecond both require
m/s, and a km/h value anywhere would eventually become a threshold wrong by a
factor of 3.6.

### Not being able to tell

Every way of not knowing reports **unsafe**, never "no rain" — see
*Fail closed* above, which is the single biggest behavioural change from the
software this replaces.

## Latching

Each cause latches separately, with a clear delay that fits it.

| Cause | Clear delay | Why |
|---|---|---|
| rain | 600 s | a heated sensor takes 2–5 minutes to dry; reopening onto sensors that are still evaporating defeats the point |
| wind | 120 s | nothing has to dry, it only has to stop gusting |
| a stream that dropped out | 30 s | a network blip must not cost ten minutes of sky |

Two things about the timing are worth knowing before anyone times it with a
stopwatch.

**The averaging window adds its own lag.** Wind must first fall out of the 60 s
mean (or the 120 s gust window) before the 120 s latch begins counting, so a
real gust holds for up to about four minutes in total.

**A latch is never released by the absence of evidence.** If it rains and then
the bridge dies, the countdown does not run through the blackout — it freezes,
and restarts when the sensors can be seen again. The clear period is ten
minutes of *observed* dryness, not ten minutes of not looking.

**On startup everything is unsafe**, and stays unsafe until both streams have
been healthy for the settle period. The startup trip clears in 30 s rather than
the full rain delay, because it represents the absence of good news rather than
the presence of bad news — a restart on a clear night should not cost ten
minutes.

## What gets published

`SafetyMonitor.IsSafe` is the verdict. `ObservingConditions` carries
`WindSpeed`, `WindGust`, `WindDirection`, `Temperature` and `RainRate`.

**`RainRate` reflects the latched state, not the instantaneous reading**, for
two separate reasons:

Arcsecond fires its recovery procedure on the transition back to GO. A rate
that dropped to zero the moment the sensors dried would reopen the dome while
they were still evaporating.

And it is **raised when the rain sensors cannot be read at all**. That is
deliberate and defensive: Arcsecond's evaluator currently *skips* a condition
whose value is missing and, with nothing else triggered, decides GO — so a
blind sensor would read to it as good weather. Publishing the wet value means a
device that can be reached but cannot see still produces a NOGO.

This does **not** cover the weather service itself being unreachable. Only the
Arcsecond-side fix, or a dome-side deadman, covers that.

There is no rain gauge at Greenhill, only wet/dry sections, so `RainRate` is an
encoding of a state rather than a measurement. The device says so in its
`SensorDescription`.

## Watching it

```
python3 tools/monitor.py                          # live, from the network
python3 tools/monitor.py --replay capture.jsonl   # from a recording
```

Replay is the important one. Point it at a file from `tools/record_streams.py`
and the whole safety core runs against real traffic, at speed, with no hardware
and no network. A night of weather can be re-examined as often as needed, and
any change to the rules can be checked against what actually happened rather
than against what someone remembers happening.

```
      0s  UNSAFE  rain[DDD] rate=1.0  wind     - m/s gust     - @   0deg
           rain_data: clearing, 30s to go
           wind_data: no wind data received
     39s  SAFE    rain[DDD] rate=0.0  wind   2.0 m/s gust   2.0 @ 120deg
     60s  UNSAFE  rain[DWw] rate=1.0  wind   2.0 m/s gust   2.0 @ 120deg
           rain: 3 wet sections (threshold 2)
    699s  SAFE    rain[DDD] rate=0.0  wind   2.0 m/s gust   2.0 @ 120deg
```

Neither mode commands anything. They watch.

## Settings

All of it is in `greenhill/core/config.py`, wired to `config.toml` in Phase 3.
Two values there are **provisional** and both concern the anemometer, whose
sentence format has never been written down:

* `wind_direction_field` / `wind_speed_field` — the positions the legacy
  display reads. Configurable so that correcting them is a config edit rather
  than a release. `tools/record_streams.py` prints a real sentence field by
  field to settle it.
* `wind_north_offset_deg` — the legacy display adds 30°, with no note saying
  why. Carried forward so behaviour does not change silently, but
  `WindDirection` should not be trusted until it has been checked against a
  known reference.

## Testing

```
python -m pytest tests/ -q
```

No hardware, no network, any OS. The safety rules are driven by synthetic
packet streams with time passed in rather than slept through, so a ten-minute
latch is tested in microseconds.
