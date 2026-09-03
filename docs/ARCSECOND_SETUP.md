# Registering Greenhill's weather with Arcsecond

Route 2: Arcsecond polls the weather devices, decides GO/NOGO, and runs a close
procedure. It backs up the direct close, which is roughly ninety seconds
quicker.

Everything here is configuration on the arcsecond-local install running on the
Windows 11 box — the same machine as the weather server, so the polling is over
loopback.

## 1. Register the Alpaca server and its devices

The weather server is a **separate Alpaca server** from the equipment one, on
its own port, because 11111 is already taken on that machine.

| | |
|---|---|
| Address | `127.0.0.1:11112` |
| Devices | `ObservingConditions` #0, `SafetyMonitor` #0 |

Both must be attached to the **Greenhill observing site**, not to a telescope:
`site__*` conditions read site-scoped equipment.

Check it took, from the machine itself:

```bash
curl "http://127.0.0.1:11112/management/v1/configureddevices"
curl "http://127.0.0.1:11112/api/v1/safetymonitor/0/issafe?ClientID=1&ClientTransactionID=1"
```

The weather beat polls ObservingConditions every 60 s into `WeatherStatePoint`;
the equipment beat polls SafetyMonitor every 30 s into `EquipmentStatePoint`.
Both are self-hosted-only tasks, which is why this needs arcsecond-local rather
than the cloud.

## 2. The condition set

Site-scoped, attached to Greenhill, **mode `ANY`**.

> `ALL` is the model default and it is wrong here. In `ALL` mode a NOGO
> requires *every* evaluated condition to trigger, so rain alone would never
> fire while the wind was calm.

### Recommended: one condition

```
quantity   site__is_safe
condition  Smaller
value      1
```

That is it. `SafetyMonitor.IsSafe` has already fused rain, wind, sensor health
and stream staleness on the Windows 11 box, using the rules in
[SAFETY.md](SAFETY.md). Restating those thresholds here would duplicate them in
two places that can drift apart — and the local copy is the one that also
drives the direct close, so a drift would leave the two routes disagreeing.

`site__is_safe` did not exist until recently: a SafetyMonitor was polled and
stored and never consulted, because no quantity read it. Fixed in
`arcsecond-back` on the `safety-fail-closed` branch, along with two related
holes — see below.

### Optional backstop

If you want Arcsecond to form its own opinion rather than trusting the monitor:

| quantity | condition | value | |
|---|---|---|---|
| `site__wind_speed` | Larger | `5.5556` | 20 km/h |
| `site__wind_gust` | Larger | `8.3333` | 30 km/h |
| `site__precipitation_intensity` | Larger | `0` | any rain |

**These are metres per second.** Entering 20 and 30 gives thresholds of 72 and
108 km/h, which nothing at this site will ever reach. Alpaca
ObservingConditions is defined in m/s and `WeatherStatePoint.wind_speed` stores
it unconverted.

Note the duplication this creates: change a threshold in `device/config.toml`
and you must change it here too, or the two routes will disagree about when to
close.

## 3. Procedures

Bind a **close procedure** to the set — park the mount and close the shutter.

Bind a **recovery procedure** too, if the dome should reopen by itself.
Arcsecond fires it on the NOGO→GO transition, and it is the *only* thing that
reopens: the weather service closes and never opens, deliberately, so the two
routes cannot argue about whether the dome should be open.

The weather service holds `RainRate` non-zero for as long as its rain latch
holds — ten minutes of observed dryness — so a recovery procedure cannot reopen
onto sensors that are still evaporating.

## 4. Timing

| | |
|---|---|
| Weather beat | 60 s |
| Equipment beat | 30 s |
| Safety heartbeat | 30 s |
| **Worst case, conditions to close** | **~90 s** |

Against roughly 1–2 s for the direct close. Route 1 is the fast one; this is
the one that still works when route 1 does not.

A procedure runs **once per NOGO episode**, not once per heartbeat — otherwise
a single overcast night would dispatch it about 120 times an hour.

## 5. Backend changes this depends on

Three fixes on the `safety-fail-closed` branch of `arcsecond-back`. **The first
matters most: without it, this whole route is unsound.**

1. **An unreachable sensor was GO.** A device that could not be reached wrote a
   weather point with every column null; the evaluator *skipped* conditions
   with a null value; with nothing triggered the decision was GO. A sensor
   failure was indistinguishable from good weather. Live outdoor quantities now
   fail closed.
2. **No staleness bound.** The newest reading counted as live however old it
   was, so sensors that died at dusk answered for the rest of the night.
   Five minutes now, settable with `SAFETY_STATE_MAX_AGE_SECONDS`.
3. **`site__is_safe` did not exist**, so a SafetyMonitor could not stop
   anything.

A NOGO that fired because nothing could be read records `observed_value: null`
and `unavailable: true`, so it is possible to tell "the wind was too strong"
from "nobody could tell how strong the wind was".

## 6. What this does not cover

If the **Windows 11 box** fails, this route fails with it — arcsecond-local,
the weather service and the equipment server all live there. Nothing tells the
dome to close.

The only protection that survives that is a **dome-side deadman**: the dome
closing itself if it has not heard from the weather service in N minutes. It is
recommended, not built, and it is the last structural gap in the design.

## 7. Before trusting any of it

- [ ] Confirm both devices appear and are polled — check `WeatherStatePoint`
      and `EquipmentStatePoint` rows are arriving with values, not errors.
- [ ] Force a NOGO (stop the rain bridge; the stream goes stale and `IsSafe`
      goes false within ~45 s) and confirm a `SafetyDecision` records NOGO and
      dispatches the procedure.
- [ ] Confirm the recovery procedure reopens, and only after the rain latch has
      released.
- [ ] **Settle the wind thresholds.** In a 15 minute capture the site exceeded
      the sustained limit in 95% of evaluations and the gust limit in 88% — it
      would have been NOGO throughout. Either it was genuinely that windy, or
      20/30 km/h is too tight for the ridge. Worth knowing before the dome
      starts closing on it.
