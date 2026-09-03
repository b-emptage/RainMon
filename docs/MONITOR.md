# The weather window

The display the astronomer actually looks at. Rain sensors on the site map,
wind on the compass, and the important things said out loud.

```bash
python greenhill_monitor.py
python greenhill_monitor.py --mute
```

## What it is, and is not

**It listens and nothing else.** No serial port, no dome commands, no
Arcsecond. It joins the same two multicast streams everything else reads, and
draws what they say.

That has consequences worth stating plainly:

* It runs on **any machine on the observatory LAN**, and on as many at once as
  anyone wants — multicast is not exclusive. The old monitor could only run on
  the box holding COM7.
* It keeps working when **the weather service, Arcsecond, or both are down**,
  which is exactly when somebody wants to look at it.
* **Closing it stops nothing.** In the software this replaces, the window *was*
  the safety system: closing it stopped the monitoring and the dome protection
  with it. Now it is a viewer, and the two are separate processes on separate
  machines.
* It is 3.8-compatible, so it can run on the Windows 7 box where the display it
  replaces has always lived, or on anything newer.

## What is on screen

**The site map** with each detector's status letter at its position, in the
legacy colours — `D` green, `w` royal blue, `W` blue, `E`/`e` red. `NYI` shows
`-` because it is not installed.

**The table**: identifier, status, temperature per detector, and how old the
rain data is. That last figure turns red when the stream goes stale.

**The compass**: bearing arrow, a wedge showing how much the direction is
wandering, and three readouts in km/h — instantaneous, 60-second mean, and
gust. Coloured green through red across 5–30 km/h, the same ramp the wind
display always used.

> The readouts are km/h because that is what the observatory thinks in.
> Everything under the surface, and everything published over Alpaca, is m/s.
> The conversion happens here and nowhere else.

**The banner**: SAFE or UNSAFE, with the reasons.

## Whose verdict is it?

The window computes its own, from the same `greenhill.core` the weather service
uses — so the numbers agree. But **this window does not control the dome.**

Set `safety_address` to the weather service (`host:11112`) and it will also
poll the real `SafetyMonitor` and show the observatory's verdict beside its
own, saying so when they differ:

```
NOTE: the observatory reports UNSAFE -- this window is a local reading
      and does not control the dome.
```

They can legitimately differ: the two processes start at different times, so
their latches are not in step. An astronomer looking at a green panel beside a
closed roof deserves an explanation rather than a puzzle. Leave the address
blank to keep the window entirely self-contained.

## Spoken alerts

The same SAPI voice the observatory already uses, on its own thread — the
window never waits for a sentence to finish, because SAPI blocks for as long as
it takes to say the words.

| Category | Says |
|---|---|
| `rain` | "Rain detected. 3 of 6 sections wet." on the transition, then every 30 s while wet. "Rain sensors dry." when it clears. |
| `faults` | "Rain sensors not responding." / "Wind sensor not responding.", and the recovery. |
| `safety` | "Conditions unsafe." / "Conditions safe." |

Drop any of them from `alert_categories` if they become tiresome. **`faults` is
the one to keep if you keep only one** — it is the announcement the old monitor
could not make, because it could not tell the difference between a dry night
and a dead sensor. An astronomer who hears nothing assumes it is dry.

Three deliberate silences:

* **Nothing is said on the first update.** Otherwise opening the window greets
  you with "rain detected" for rain that stopped an hour ago.
* **Nothing is called broken for the first 20 seconds.** Every source looks dead
  before its first packet arrives. After that a genuinely dead sensor *is*
  announced — unlike stale rain, a sensor that is dead right now is current news.
* **"Rain sensors dry" is only said from a known-wet state to a known-dry one.**
  Wet then unreadable is a fault, and announcing dryness there would be the most
  misleading sentence this could utter.

On a machine with no Windows speech — a Mac, a Linux box, Windows without
pywin32 — alerts are written to the log instead and everything else works
normally.

## Settings

`greenhill_monitor.ini`, beside the executable. The one that matters:

**`interface`** — set it if the machine has more than one network adapter.
Left blank the OS chooses, and if it chooses wrong the join succeeds and no
packets ever arrive, which looks exactly like a dead sensor. This fault has
already caught the observatory once. `python tools/mcast_diag.py` reports which
interface actually receives.

## Building it

```bash
pyinstaller greenhill_monitor.spec
```

Windowed, with the site map bundled. The `.ini` is deliberately **not**
bundled: it must sit beside the executable where it can be edited.

Use the pinned 3.8 32-bit toolchain if the window is to run on the Windows 7
box; any current Python otherwise.

## Related

`tools/monitor.py` is the same information as a console line, and can replay a
capture. Useful for commissioning and for looking at a night after the fact;
this window is for looking at the night as it happens.
