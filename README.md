# Greenhill weather safety

Rain and wind monitoring for the Greenhill Observatory (Bisdee Tier), exposed
to the observatory as native ASCOM Alpaca devices.

Two machines:

* **Windows 7 box** — the rain-sensor bridge. Owns the detectors' serial port
  and multicasts their readings to the LAN. Decides nothing.
  See [docs/BROADCASTER.md](docs/BROADCASTER.md).
* **Windows 11 box** — the weather service. Subscribes to the rain stream and
  to the anemometer's own multicast, fuses them into a safety state, and serves
  `ObservingConditions` and `SafetyMonitor` over Alpaca.
  See [docs/SAFETY.md](docs/SAFETY.md), [docs/ALPACA.md](docs/ALPACA.md) and
  [docs/DOME_CLOSE.md](docs/DOME_CLOSE.md).

On a NOGO the dome is closed by two independent routes: directly, by an Alpaca
client against the dome server (~1–2 s), and by Arcsecond's own safety
procedure (~90 s). Both derive from the same fused state.

**The system fails closed.** It reports UNSAFE unless it can positively see
that conditions are good — so a silent sensor, a dead serial port or a fresh
restart all close the dome rather than reading as a clear night. That last one
has a consequence worth knowing before it surprises you: **restarting the
weather service closes the dome.** See
[docs/SAFETY.md](docs/SAFETY.md#fail-closed--read-this-first).

[PLAN.md](PLAN.md) has the full design, the findings that shaped it, and the
phase plan.

## Status

| Phase | | |
|---|---|---|
| 1 | Rain bridge and stream recorder | **done** |
| 2 | Fusion, latching and watchdogs (Win11) | **done** |
| 3 | Alpaca server: ObservingConditions + SafetyMonitor | **done** |
| 4 | Route 1: direct dome close | **done** |
| 5 | Arcsecond registration and condition set | next |
| 6 | UI and speaker relocated; `RainMonT.exe` retired | |

## Legacy

`RainMonT.py`, `RainMonSimT.py`, `RainMon.ini` and `RainMonT.spec` are the
system in service today, and are untouched. `RainMonSimT.py` stays useful past
the changeover: over a virtual COM pair it can drive wet states through the
whole new chain without waiting for weather.
