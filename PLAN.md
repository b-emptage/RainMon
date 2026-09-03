# Greenhill weather safety — implementation plan

Merges `Greenhill-RainMon` and `Greenhill-WindSensor` into one package that
exposes native Alpaca `ObservingConditions` and `SafetyMonitor` devices, and
closes the dome on a NOGO by two independent routes.

Status: **plan agreed, not yet implemented.**

---

## 1. Architecture

The Windows 7 box stops deciding anything. It becomes a sensor bridge that
reads the rain detectors over serial and multicasts raw observations, in the
same shape the anemometer already uses. All fusion, thresholds, latching and
Alpaca serving move to the Windows 11 box.

```
[Anemometer]  ──multicast 239.192.0.4:60004──┐   (existing, unchanged)
                                              │
[Win7 32-bit]  rain_broadcaster.exe           │
  COM7 ──> multicast 239.192.0.5:60005 ──────>│
  emit-only; decides nothing                  ▼
                          [Win11]  greenhill-weather  (new)
                            • joins both multicast streams
                            • core/safety.py — fusion, latch, watchdogs
                            • Alpaca :11112 → ObservingConditions #0
                                              SafetyMonitor #0
                            • route 1: Alpaca client → dome CloseShutter
                                              │
                  polled 60 s ────────────────┤
                          [arcsecond-local]  (same Win11 box, loopback)
                            → WeatherStatePoint
                            → 30 s safety heartbeat → NOGO → close procedure
                                              │  route 2
                                              ▼
                          [Dome box]  Alpaca Dome :11111
```

### Why the Win7 box only broadcasts

* No Alpaca stack, no falcon, no 32-bit wheel problems, no TOML.
* The safety logic runs on a supported OS.
* Fail-closed is a property of the transport: no packet in N seconds → unsafe.
  In the current system a dead serial link is indistinguishable from "dry".
* Anything that may ever need tuning lives on Win11, so the old box is never
  touched again.

### Ports

| What | Where | Port |
|---|---|---|
| Anemometer multicast | LAN | 239.192.0.4:60004 (UDP) |
| Rain multicast | LAN | 239.192.0.5:60005 (UDP) |
| Weather Alpaca server | Win11 | **11112** |
| Main equipment Alpaca server | Win11 | 11111 (existing — do not collide) |
| Dome Alpaca server | dome box | 11111 |
| Alpaca discovery | Win11 | 32227, expected to be held by the incumbent server. Bind failure is a warning; Arcsecond uses explicit addresses. |

---

## 2. Win7 — rain broadcaster

One script. `pyserial` + stdlib. Frozen with the **existing, proven Python 3.8
32-bit + PyInstaller chain — pin the exact versions, do not upgrade them.**
That chain is already known to produce an executable that runs on this machine;
a toolchain bump is the kind of change that fails only on the target OS.

Responsibilities:

* Own COM7 exclusively. Probe detectors, poll status (`*RnS`) and MK3
  temperatures (`*RnA`), decode as today.
* Multicast one datagram per second.
* Emit **raw observations only**. No thresholds, no latching, no decisions.
* Emit-only: it never listens. RESET (`*RnI`) stays a manual operation at the
  machine — deliberate, since a detector lockup is worth inspecting in person.

### Packet format

JSON, one datagram, ~200 bytes, well inside one MTU:

```json
{"v":1,"seq":4213,"t":"2026-08-17T10:23:45Z",
 "dets":[{"id":"H127","st":"D","tC":12.4},
         {"id":"H50","st":"w","tC":11.9},
         {"id":"ACC","st":"W","tC":13.1}],
 "port_ok":true}
```

* `st` is the detector's status letter: `P M I D W w E e` (unchanged semantics).
* `seq` gives packet-loss detection; `t` gives staleness detection.
* `port_ok` false means the serial port itself failed — distinct from
  "detectors report error", and both distinct from silence.
* Detector list is variable length: 3 installed today, expandable to 4.

Chosen over positional CSV because it parses with real error handling. The
anemometer's format cannot be changed, but its parser should not be copied: it
reads `parts[2]` and `parts[4]` by fixed index behind a bare `except:` that
silently yields `(None, None)`.

---

## 3. Win11 — greenhill-weather

```
greenhill-weather/
  device/          # cloned from Greenhill-DomeShutter/device/
    app.py  config.py  config.toml  discovery.py  log.py
    management.py  shr.py  exceptions.py  setup.py
    observingconditions.py   safetymonitor.py    # Alpaca responders
    weatherdevice.py                             # bridge: ASCOM semantics
  core/
    rain.py        # decode rain multicast, wet-state machine
    wind.py        # decode anemometer multicast, circular stats
    safety.py      # fusion + latch + watchdogs — the only thing that decides
    dome_client.py # route 1
  ui/app.py        # optional Tk viewer — a listener, no hardware access
  sim/             # RainMonSimT.py + rain/wind packet injectors
  tests/           # hardware-free, driven by replayed captures
```

Clone `device/` from `Greenhill-DomeShutter` rather than upstream AlpycaDevice:
that repo already fixes two defects Conform found in the sample — the
duplicated `ClientID` returning HTTP 500, and the HTTP/1.0 keep-alive bug that
resets pooled .NET client connections.

### Safety fusion

`core/safety.py` holds the single source of truth. Both Alpaca devices are
views onto it.

* **Fail closed.** `IsSafe` starts false and goes true only once both streams
  have delivered fresh, valid data. Silence, stale packets, `port_ok:false`, or
  all-detectors-error → unsafe. Never "dry and calm".
* **Rain rule** (unchanged in intent from `checkWetAndClose`, reimplemented as
  an explicit state machine):
  * wet sections ≥ 2 → unsafe.
    Sections, not detectors: `w` = 1, `W` = 2. One detector fully wet trips it.
  * exactly 1 wet section that then dries within 10 s → unsafe (a real
    raindrop evaporating off a heated sensor).
  * exactly 1 wet section that persists → not rain. Bird droppings, insects, or
    a failed sensor.
* **Wind rule**: 60 s mean > **5.56 m/s** (20 km/h), or gust > **8.33 m/s**
  (30 km/h). Anemometer confirmed to report m/s. Minimal hysteresis.
* **Latching**, `dry_clear_seconds`, default **600 s**. Cleared only when all
  sections are dry *continuously* for that long AND wind is below threshold.
  Comfortably above the 2–5 minute drying time. This replaces the arbitrary
  300 s TCP-spam guard in the current code, which was solving a real problem
  with the wrong mechanism.
* **We never reopen.** Route 1 is close-only. Reopening is an observing
  decision and belongs to Arcsecond's `recovery_procedure`. One opener.

### Alpaca surface

`ObservingConditions #0`

| Member | Value |
|---|---|
| `WindSpeed` | 60 s mean, m/s |
| `WindGust` | peak, m/s |
| `WindDirection` | degrees the wind comes *from*, north-offset applied, 0 when calm |
| `Temperature` | MK3 `Tamb`, °C |
| `RainRate` | mm/hr — **publishes the latched state**, see below |
| `TimeSinceLastUpdate` | real staleness, per sensor |
| everything else | `PropertyNotImplementedException` |

`SafetyMonitor #0` — `IsSafe`, the fused state.

Both must report `Connected = true` and stay connected: Arcsecond only ever
*reads* `Connected`, never sets it, and a device reporting false is recorded as
"Device disconnected" with no values.

**`RainRate` must publish the latch, not the instantaneous reading.** Arcsecond
fires its `recovery_procedure` on the NOGO→GO transition, so a `RainRate` that
drops to 0 the moment the sensors dry would reopen the dome while the local
state still considers it unsafe. Holding it non-zero for the life of the latch
is what keeps the two routes from fighting.

`RainRate` is an *encoding*, not a measurement — there is no rain gauge, only
wet/dry sections. Say so in `SensorDescription("RainRate")`.

### Route 1 — direct close

Alpaca client on Win11 against the dome box:

1. `PUT /api/v1/dome/0/closeshutter`
2. poll `GET .../shutterstatus` until it reads `shutterClosed` (**1** --
   4 is `shutterError`)
3. backoff and retry; escalate to log + speaker if it never gets there

Connect once and **stay connected** — the dome's `Disconnect` de-energises the
motors, so connect/command/disconnect around a moving shutter is unsafe.

The dome deliberately reports `shutterOpen` for a shell stopped half way, so
waiting for positive `shutterClosed` correctly treats a partial close as
"not yet shut".

Latency ≈ 1–2 s, versus up to ~90 s for route 2. Route 1 is primary; route 2 is
the backup. Both are idempotent, so both firing is harmless.

---

## 4. Route 2 — Arcsecond

arcsecond-local runs on the same Win11 box, so the weather beat polls our
server over loopback.

Cadence, from `arcsecond/installation/periodic_tasks.py`:
weather collection **60 s**, safety heartbeat **30 s** → worst case ~90 s from
rain to procedure dispatch, plus execution.

Condition set configuration:

* **mode `ANY`**, not the default `ALL`. With `ALL`, `is_nogo` requires *every*
  evaluated condition to trigger, so rain alone would never fire.
* `SITE_WIND_SPEED` Larger **5.56**, `SITE_WIND_GUST` Larger **8.33**.
  **Units are m/s.** Entering 20 and 30 gives 72 and 108 km/h thresholds.
* `SITE_PRECIPITATION_INTENSITY` Larger 0, fed from our `RainRate`.
* Bind the close procedure, and a `recovery_procedure` for reopening.

---

## 5. Defects that shaped this design

### In the current RainMon

| | |
|---|---|
| `RainMonT.py:191` | The entire close block sits inside `if ... and self.TCP_connected:`. If the socket has dropped, rain is detected, logged and announced — and nothing closes. |
| `RainMonT.py:172` | Detector errors return `e`, so `wetSensorCount == 0` and `checkWetAndClose` is never called. A dead serial link looks exactly like "perfectly dry". |
| `RainMonT.py:200` | With one wet section, the 1 Hz repeater schedules a *new* 10 s callback every second, each of which reschedules. The "10 second window" is a growing cascade. |
| `RainMonT.py:196` | After a close, `close_issued` blocks all evaluation for 300 s with no confirmation the dome actually closed. |
| `RainMonT.py:617` | `if "open" in response.lower()` releases the safety latch. `"shutter open failed"` matches. |
| `RainMonT.py:919` | `base_dir` is assigned only in the frozen branch but used unconditionally — `NameError` when run as `.py` with `LOG=True`. Masked because the `.exe` ships. |
| `RainMonT.py:379` | Blocking serial reads on the Tk thread: 4 detectors × 0.9 s timeout inside a 1000 ms loop, which also distorts every `after()`-based timing. |
| `wind_sensor.py:36` | Bare `except:` returning `(None, None)`; no staleness detection, so a stopped multicast freezes the display on the last good values indefinitely. |

### In arcsecond-back — separate track, needed before this runs unattended

1. **A dead sensor reads as GO.** `weather/tasks.py:105` writes a
   `WeatherStatePoint` with all-null columns when the device is unreachable;
   `safety/services/evaluator.py:264` then *skips* any condition observing
   `None`, and with nothing triggered `is_nogo = False`. Sensor failure is
   indistinguishable from good weather.
2. **No staleness bound.** `evaluator.py:49` takes the newest row regardless of
   age. A three-hour-old reading counts as live.
3. **No `IsSafe` quantity.** `EVALUATED_QUANTITIES` (`evaluator.py:29`) covers
   weather columns, site closure, sun altitude, telescope pointing, camera
   temperature and target magnitude. A `SafetyMonitor` is pollable as equipment
   but never reaches a `SafetyDecision`. Adding `SITE_IS_SAFE` would let the
   boolean travel as a boolean and return `RainRate` to honest telemetry.

Item 1 matters most under the new architecture: Win11 now hosts
arcsecond-local, the equipment Alpaca server and the weather service, so it is
a single point of failure for both close routes.

---

## 6. Phases

| | Work | Where | Gate |
|---|---|---|---|
| 1 | `rain_broadcaster` + a recorder capturing both live streams to file | Win7 / any | real packets captured |
| 2 | `core/`: rain decode, wind decode, fusion, latch, watchdogs — hardware-free, driven by replayed captures | any | tests green |
| 3 | Alpaca server, both devices, cloned from `Greenhill-DomeShutter/device/` | Win11 | both Conform suites clean |
| 4 | Route 1: dome client with verify-and-retry | Win11 | dry-run close |
| 5 | Register both devices in arcsecond-local; build the condition set | Win11 | e2e close |
| 6 | GUI + speaker redeployed as multicast listeners; decommission `RainMonT.exe` | — | — |
| B | The three arcsecond-back fixes | backend | — |

Phase 2 is the bulk of the work, and it is all on a modern machine with no
hardware attached.

---

## 7. Risks and open items

**Cutover is a hard swap.** COM7 is exclusive, so the old `RainMonT.exe` and
the new broadcaster cannot run together — there is no shadow-mode night. Do the
swap in a dry-forecast maintenance window and prove the chain first with
`RainMonSimT.py` over a virtual COM pair (com0com), which can drive wet states
through broadcaster → Win11 → the real dome without waiting for weather.

**Win11 is a single point of failure for both routes.** It hosts
arcsecond-local, the equipment Alpaca server and the weather service. If it
dies, nothing tells the dome to close. The only structural fix is a **dome-side
deadman**: the dome server closes if it has not heard a periodic "conditions
are safe" heartbeat within N minutes. This is a change to
`Greenhill-DomeShutter`, out of scope here, and needs a generous timeout so a
network blip does not close the roof mid-exposure. Recommended, to be scheduled
separately.

**The LAN becomes safety-critical.** Both boxes must be on the same L2 segment,
or IGMP routing configured for 239.192.0.0/14. Windows Firewall on Win11 must
allow the inbound UDP. The streams are unauthenticated — anyone on the
observatory LAN could inject a "dry" packet; acceptable here, but as a decision
rather than an oversight. Packet *loss* is safe by construction: loss becomes
silence becomes unsafe.

**Frozen-build config path.** `device/config.py` resolves `config.toml`
relative to the module file, which under PyInstaller lands inside `_MEIPASS` —
a read-only temp directory that vanishes on exit, leaving operators unable to
edit thresholds. The current `RainMon.ini` is read from
`os.path.dirname(sys.executable)` for exactly this reason. Only relevant if the
Win11 service is ever frozen; it is not for the Win7 broadcaster, which has no
TOML.

**Where do the GUI and speaker live?** Now that the GUI is a pure multicast
listener it can run anywhere, or in several places at once. If the Win11 box
runs unattended, spoken alerts there need a logged-in session — so the operator
display may be better left on the machine where someone actually sits.
Undecided; costs nothing to defer.

---

## 8. Commissioning checklist

1. Confirm both boxes receive both multicast streams (L2/IGMP, firewall).
2. Drive `RainMonSimT.py` over a virtual COM pair and verify the full chain:
   broadcaster → Win11 fusion → `IsSafe` false → dome reaches `shutterClosed`.
3. Cut the rain multicast and confirm `IsSafe` goes false on staleness.
4. Kill the broadcaster and confirm the same.
5. Verify the Arcsecond condition set fires — check units are m/s and mode is
   `ANY` by inspecting a recorded `SafetyDecision`.
6. Verify the latch: after simulated rain stops, confirm `RainRate` stays
   non-zero for `dry_clear_seconds` and that the dome does not reopen early.
7. Run both ASCOM Conform suites against the weather server.
8. Confirm the anemometer's north offset (currently a hard-coded `+30` at
   `wind_sensor.py:195`) against a known reference before trusting
   `WindDirection`.
