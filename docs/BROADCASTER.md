# The rain-sensor bridge

The Windows 7 half of the Greenhill weather safety system. It reads the rain
detectors over serial and multicasts what they said, once a second, to whoever
is listening. That is all it does.

It does not decide anything. No thresholds, no latching, no dome commands, no
network listener. Every rule that might one day want changing lives on the
Windows 11 weather service, so this machine should not need touching again
after it is installed.

The design mirrors the anemometer, which has always worked this way: the
instrument multicasts to the LAN and anyone who cares subscribes. `wind_sensor.py`
in the WindSensor repo is a *consumer* of that stream, not a producer.

## Layout

| File | Role |
|---|---|
| `rain_broadcaster.py` | Entry point. Thin, and at the repo root, because PyInstaller freezes a script. |
| `greenhill/broadcaster.py` | Config, logging, the emit loop. |
| `greenhill/rain_serial.py` | The detector serial protocol. No networking. |
| `greenhill/rain_protocol.py` | The wire format. **Imported by both ends** — keep it Python 3.8 compatible. |
| `rain_broadcaster.ini` | Settings. Must sit beside the executable. |
| `tools/record_streams.py` | Captures both multicast streams to a file. |
| `tools/rain_packet_sim.py` | Emits synthetic packets, so the receiver can be built without hardware. |

## Running

```
pip install -r requirements-broadcaster.txt
python rain_broadcaster.py
```

`--config PATH` to point at a different .ini, `--verbose` for debug logging.

### Building the executable

**Pin the toolchain and do not upgrade it.** The Greenhill box runs Windows 7
32-bit, whose last supported CPython is **3.8**; the executable is built from a
3.8 32-bit environment with the PyInstaller version already in use for
`RainMonT.exe`. That combination is proven in the field. A toolchain bump is
the kind of change that fails only on the target OS, which is the worst place
to discover it.

The box's own Python is 2.7 and is irrelevant — a frozen build carries its own
interpreter.

```
pyinstaller --onefile rain_broadcaster.py
```

`rain_broadcaster.ini` and `rain_broadcaster.log` live **beside** the resulting
`.exe`, never inside it. A path resolved relative to the module would land in
`_MEIPASS`: a read-only temporary directory that is deleted on exit, leaving
the operator unable to edit the settings or find the log. The legacy
`RainMon.ini` is read from the executable's directory for exactly this reason.

## The wire format

One JSON datagram per poll cycle to **239.192.0.5:60005**, adjacent to the
anemometer's 239.192.0.4:60004.

```json
{"v": 1, "seq": 4213, "t": "2026-08-17T10:23:45.123Z",
 "dets": [{"id": "H127", "st": "D", "tC": 12.4},
          {"id": "H50",  "st": "w", "tC": 11.9},
          {"id": "ACC",  "st": "W", "tC": 13.1}],
 "port_ok": true, "poll_ms": 63}
```

Status letters are the detectors' own: `P` parked, `M` moving, `I`
initialising, `D` dry, `w` half wet, `W` fully wet, `E` device error. `e` is
added by the bridge and means "no usable reply".

Three properties carry weight:

**`dets` always lists every configured detector**, including ones that did not
answer. The receiver therefore knows the expected population, and a short list
is a malformed packet rather than a quiet decision to evaluate the weather on
fewer sensors than the observatory has.

**`port_ok` separates two faults that look identical from a distance.**
`port_ok: false` means the bridge is alive and its serial port is not. Silence
means the bridge itself is gone. Both are unsafe, and the receiver must treat
them as such — but they send an engineer to different places.

**`t` is not a staleness clock.** It is this machine's wall clock, and nothing
guarantees it is synchronised. A box an hour out would look permanently stale
or, far worse, permanently fresh. **Staleness is measured by the receiver
against arrival time.** `t` is for logs, and for spotting a sender whose clock
has stopped.

### Wetness counting

Each detector has two sections that trigger independently, so `w` contributes
one wet section and `W` contributes two. The operational rule is "2 of 6
**sections**" — which means **a single detector reporting `W` reaches the
threshold on its own**. That is intended.

`D`, `w` and `W` are the only letters that say anything about wetness. A
detector that is parked, moving, initialising or in error is reporting nothing
about the sky, and **must never be counted as dry**. `rain_protocol.is_observation()`
is that distinction, and it is the fail-closed hinge of the whole system: in
the legacy software an errored detector produced a wet count of zero, which was
indistinguishable from a clear night.

## What the bridge does when things go wrong

| Fault | Behaviour |
|---|---|
| One detector does not answer | Reports `e` for that detector. The others still read. `port_ok` stays true — the link is fine. |
| A detector fails repeatedly | After 5 consecutive failures it is retried only every 30 cycles, still reporting `e` in between. Otherwise each dead detector costs a full 0.9 s read timeout every second, and three of them stretch a 1 s cadence to nearly 3 s. |
| A failed detector recovers | Picked up automatically on its next poll. The legacy software probed once at startup and froze the active list, so a detector unplugged at boot stayed dead until someone restarted it. |
| The serial port fails | Port closed, `port_ok: false`, every detector `e`. **Emission continues.** Reopened on a backoff of 1, 2, 5, 10, 30 s. |
| The port cannot be opened at all | Same, and the error names the likely cause: on Windows only one process may hold a COM port, so the usual answer is that `RainMonT.exe` is still running. |
| Multicast send fails | Logged (first failure, then every 60th) and the loop carries on. |
| The log cannot be written | Falls back to the console with a `==LOGGING==` line on stderr. Nothing about a log file may stop the bridge from running. |
| A second copy is started | Refuses, via a loopback-port mutex on 50816 (the dome server uses 50815). A start that finds the port held first waits up to 10 s, saying so on stderr, so a restart typed right after a Ctrl-C comes up on its own instead of racing the old bridge's shutdown. A refusal after the grace is real: the old process is still alive. |
| Ctrl-C does not stop it | The first Ctrl-C asks the loop to finish its cycle and close the serial port; it also restores the default handler, so a **second** Ctrl-C kills the process outright — even one wedged inside a serial driver call, where the polite handler can never run — and the kill releases the mutex port for the next start. |

## During changeover

**`RainMonT.exe` and the bridge cannot run at the same time.** A Windows COM
port belongs to one process, so whichever starts second fails — and until the
Windows 11 service is commissioned, stopping the old executable means the
observatory has no rain protection at all.

Do the swap in a dry-forecast maintenance window, and prove the chain first
with `RainMonSimT.py` over a virtual COM pair (com0com), which can drive wet
states end to end without waiting for weather.

## Testing

```
pip install -r requirements-dev.txt
python -m pytest tests/ -q
```

No hardware, any OS. The serial port is faked and the reply strings are the
real ones from the hardware author's notes.

### Capturing the live streams

Run this before building anything that consumes the streams:

```
python3 tools/record_streams.py --out capture.jsonl
```

It joins both groups, writes one JSON object per datagram, validates rain
packets as they arrive, and prints a summary on Ctrl-C — packet rates, the
largest gap, sequence gaps, and protocol errors.

It also **dissects a sample anemometer sentence field by field**. There is an
open question it exists to settle: `wind_sensor.py` reads direction from
`parts[2]` and speed from `parts[4]` by fixed index, behind a bare `except:`
that silently yields no reading, and nobody has written down what the
instrument actually emits. Everything downstream depends on the answer —
ASCOM `WindSpeed` is metres per second, Arcsecond stores metres per second, and
the 20/30 km/h thresholds become 5.56/8.33 m/s.

### Without hardware

```
python3 tools/rain_packet_sim.py --scenario raindrop
```

Scenarios exist to exercise the cases the safety logic must tell apart:

| Scenario | What it is | Expected verdict |
|---|---|---|
| `dry` | all detectors `D` | safe |
| `raindrop` | one section wet, dry again within 3 s | **rain** — a drop evaporating off a heated sensor |
| `fouling` | one section wet indefinitely | **not rain** — droppings, an insect, or a failed sensor |
| `rain` | two sections wet | rain |
| `downpour` | everything wet, then a slow dry | rain, then the latch holds through the 2–5 minute drying time |
| `port-down` | every detector `e`, `port_ok: false` | unsafe |
| `silence` | emits briefly, then stops | unsafe **on staleness** |

The last two are the ones the legacy software got wrong, and they are the
reason `port_ok` and the receiver's arrival-time clock exist.

On a machine with no usable network interface, pass `--interface 127.0.0.1` to
both the simulator and the recorder to run the whole thing over loopback.
