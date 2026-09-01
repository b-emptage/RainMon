# Greenhill weather — Alpaca device server

Publishes the safety verdict and the weather telemetry as native ASCOM Alpaca
devices, on the Windows 11 box alongside arcsecond-local and the equipment
server.

Built on the [AlpycaDevice](https://github.com/ASCOMInitiative/AlpycaDevice)
template (MIT, © Bob Denny — see `device/LICENSE-AlpycaDevice.txt`), taken by
way of `Greenhill-DomeShutter` rather than from upstream, because that server
already carries two fixes Conform found in the sample.

## Two devices, one service

| Device | Number | Carries |
|---|---|---|
| `SafetyMonitor` | 0 | `IsSafe` — the GO/NOGO |
| `ObservingConditions` | 0 | `WindSpeed`, `WindGust`, `WindDirection`, `Temperature`, `RainRate` |

Both are views onto a single `SafetyEvaluator`, so they cannot disagree. That
matters: the two dome-close routes read from them, and must never form
different opinions.

ASCOM `ObservingConditions` has no safety property, which is why the boolean
lives on its own device rather than being encoded into a measurement.

`http://<host>:11112/api/v1/{safetymonitor|observingconditions}/0/...`, with
`/management/v1/configureddevices` for discovery by address.

**Port 11112, not the conventional 11111** — the equipment Alpaca server on
this machine already holds that. For the same reason, UDP discovery on 32227
is *expected to fail to bind* here; it is logged as a warning and nothing else,
and Arcsecond is configured with an explicit address regardless.

## Layout

| File | Role |
|---|---|
| `greenhill/core/` | The safety rules. No HTTP, no ASCOM. See [SAFETY.md](SAFETY.md). |
| `device/weatherdevice.py` | Bridge: ASCOM semantics over the safety core, and the monitoring thread. |
| `device/observingconditions.py`, `device/safetymonitor.py` | Alpaca HTTP responders. Thin. |
| `device/alpaca_common.py` | The dozen members every ASCOM device carries, shared by both. |
| `device/app.py` | Startup, routing, discovery, single-instance guard. |
| `device/config.toml` | Network, the multicast groups, **and every safety threshold**. |
| `device/simulate.py` | Runs the server against synthetic weather. |

Vendored unchanged from AlpycaDevice: `shr.py`, `exceptions.py`, `discovery.py`,
`setup.py`. Modified: `config.py` (multicast and weather sections), `log.py`
(log filename), `management.py` (advertises both device types).

Each device module declares one-line subclasses of the bases in
`alpaca_common`, because `app.init_routes` only routes a class that is
*defined* in the module it is scanning. That is why `class connected(BaseConnected): pass`
appears twice and the behaviour only once.

## Running

```bash
pip install falcon
cd device && python app.py          # against the real sensor streams
cd device && python simulate.py     # against synthetic dry, calm weather
```

## Decisions that differ from the dome server

**The monitoring thread runs from startup and is not tied to `Connected`.** The
dome de-energises its motors on disconnect, deliberately. This device must not
do the equivalent: it owns no exclusive hardware — multicast can be received by
any number of listeners — and the direct dome-close route reads the safety state
with no ASCOM client of its own. Tying the thread to `Connected` would quietly
disarm it.

**`Connected` defaults to true.** It has to. Arcsecond polls `Connected` but
never sets it, so a device waiting to be connected would be recorded as
"Device disconnected" forever and never yield a single reading.

**`IsSafe` never throws.** Most clients treat an error from a safety monitor as
"no opinion" and carry on observing, so an exception here would read as
permission. It answers `false` and puts the reason in the log and in the
diagnostics action.

**`AveragePeriod` is reported honestly** as 60 s rather than the conventional
0.0. These values *are* averaged, over the same window the wind threshold uses,
and a client told "no averaging" would misread a smooth trace. Setting it to
anything else is rejected rather than ignored — a client told "fine" would go
on to read numbers that are not what it asked for.

**Absent sensors throw `PropertyNotImplemented`**, never a plausible zero.
Arcsecond caches that and stops asking; anything else would record a fabricated
humidity every minute. A sensor we *do* have but have no reading for yet throws
`ValueNotSet` for the same reason — a fabricated zero is indistinguishable from
a real calm, and something is about to decide whether to keep the dome open.

## The diagnostics action

ASCOM gives a safety monitor one bit, and this site has four independent reasons
to close. `Greenhill:GetWeatherStatus` publishes what the collapse throws away:

```
Greenhill:GetWeatherStatus   ""  -> {"isSafe": false,
                                     "reasons": ["rain: clearing, 412s to go"],
                                     "conditions": {"rain": true, "wind": false,
                                                    "rain_data": false,
                                                    "wind_data": false},
                                     "detectors": {"H127": "D", ...},
                                     "rainAgeSeconds": 0.4,
                                     "windSampleCount": 223, ...}
```

Available on both devices. It is how an operator, or a UI, sees *why* without
reading the log.

## Two traps worth knowing about

**AlpycaDevice's "Exception" classes are not Python exceptions.**
`InvalidValueException` and friends in `exceptions.py` are plain
response-payload objects — you construct one and hand it to `PropertyResponse`.
`raise` or `except` on one is a `TypeError`. So `weatherdevice.py` raises
ordinary Python exceptions (`SensorNotFitted`, `UnknownSensor`,
`ValueNotAvailable`, `SettingNotAdjustable`) and the responders translate them.

**The sample's uncaught-exception handler was itself broken.**
`HTTPInternalServerError` has been keyword-only since Falcon 3, so the
positional call in the AlpycaDevice sample raises a `TypeError` from inside the
handler whose job is to report errors — replacing the real fault with a
confusing one at the worst possible moment. Fixed here.
**`Greenhill-DomeShutter` still carries this bug** (`device/app.py`), and it
should be fixed there too.

## Testing

```bash
pip install -r requirements-dev.txt
python -m pytest tests/ -q
```

No hardware, no network, any OS. `tests/test_alpaca_weather.py` drives the real
Falcon routing through a test client.

### ASCOM Conform Universal

`.github/workflows/conformance.yml` runs both Conform suites against
`simulate.py`, **for both device types**, on every push — they share their
common responders, so a fault would otherwise appear on whichever happened to be
tested. It runs in CI rather than locally because ConformU ships as a Linux x64
binary.

The same workflow byte-compiles and tests the bridge modules on **Python 3.8**.
`greenhill/rain_protocol.py` is imported by both machines, and the Windows 7 box
cannot run anything newer; that job is what stops a 3.9 idiom reaching it.

**Conform has only been run against the simulator.** Run it against the real
installation before the observatory trusts it — the simulator feeds perfect
weather and cannot reproduce a stalled sensor or a dropped stream.

## Not yet done

* No Windows service packaging. Note that a service stop arrives as neither a
  signal nor a console event, so packaging one means wiring its stop request
  into the same shutdown path.
* Conform has only been run against `simulate.py`, never against the real
  installation.

Done since this document was first written: the direct dome close (Phase 4, see
[DOME_CLOSE.md](DOME_CLOSE.md)), registration in arcsecond-local, and the
anemometer's sentence format, which a capture settled — it is parsed as NMEA
now, not by field index. The north offset is confirmed correct by the
observatory.
