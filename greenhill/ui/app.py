# -*- coding: utf-8 -*-
"""The Greenhill weather window.

The two legacy displays -- the rain monitor's site map and the wind monitor's
compass -- in one window, fed from the multicast streams instead of from
hardware.

IT LISTENS AND NOTHING ELSE. No serial port, no dome commands, no Arcsecond. It
can run on any machine on the observatory LAN, several at once, and it keeps
working when the weather service, Arcsecond, or both are down -- which is
exactly when somebody wants to look at it. Closing it stops nothing.

Everything on screen is computed here, from the same `greenhill.core` the
weather service uses, so the numbers agree. The VERDICT is a different matter:
this window's is its own reading, and the one that actually closes the dome
belongs to the weather service. When that service is reachable the window shows
both and says so if they differ, because an astronomer looking at a green panel
beside a closed dome deserves an explanation rather than a puzzle.
"""

import json
import logging
import math
import os
import sys
import time
from typing import Dict, List, Optional

try:
    import tkinter as tk
    from tkinter import font as tkfont
except ImportError:                             # pragma: no cover
    tk = None

from greenhill import rain_protocol
from greenhill.core import rain as rain_module
from greenhill.core import wind as wind_module
from greenhill.core.receiver import MulticastReader
from greenhill.core.safety import SafetyEvaluator
from greenhill.ui.alerts import AlertPolicy, Snapshot

logger = logging.getLogger('greenhill.ui')

# The legacy palette, kept because the operators read these colours at a glance.
STATUS_COLOURS = {
    'P': ('white', 'black'),
    'M': ('black', 'gray'),
    'I': ('black', 'cyan'),
    'D': ('black', 'limegreen'),
    'W': ('white', 'blue'),
    'w': ('black', 'royalblue'),
    'E': ('white', 'red'),
    'e': ('white', 'red'),
    '-': ('black', 'gray'),
}

# Where each detector sits on the site map, from the legacy monitor. Keyed by
# the identifier that travels on the wire, so a detector moving position is a
# line here rather than an index to count out.
DETECTOR_POSITIONS = {
    'H127': (98, 293),
    'H50': (205, 258),
    'ACC': (123, 95),
    'NYI': (42, 398),
}

MAP_IMAGE = 'BT_SiteVectorMap.png'
MAP_SIZE = (268, 432)
COMPASS_SIZE = (300, 360)

BACKGROUND = '#333333'
DIM = '#999999'
EDGE = '#aaaaaa'

# Wind speed colour ramp, in km/h, as the legacy wind display used it.
RAMP_MIN_KMH = 5.0
RAMP_MAX_KMH = 30.0


def asset_path(name):
    # type: (str) -> str
    """Beside the executable when frozen, beside the repo otherwise."""
    if getattr(sys, 'frozen', False):
        base = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
    else:
        base = os.path.dirname(os.path.dirname(os.path.dirname(
            os.path.abspath(__file__))))
    return os.path.join(base, name)


def speed_colour(kmh):
    # type: (float) -> str
    """Green through to red across the ramp. The legacy curve, unchanged --
    the operators know what 'orange' means here."""
    span = max(RAMP_MAX_KMH - RAMP_MIN_KMH, 1e-6)
    fraction = min(max(kmh - RAMP_MIN_KMH, 0.0), span) / span
    percent = 100.0 * fraction
    red = int(25.5 * math.sqrt(percent))
    green = int(255 - 2.55 * percent)
    return '#{:02x}{:02x}{:02x}'.format(min(red, 255), max(green, 0), 50)


class AuthoritativeVerdict(object):
    """Polls the weather service's SafetyMonitor, if one is configured.

    Optional by design. The window's whole point is to work when things are
    down, so an unreachable service is a state to display, not an error.
    """

    def __init__(self, address, interval=5.0, timeout=2.0):
        self.address = (address or '').strip()
        self._interval = interval
        self._timeout = timeout
        self._next_poll = 0.0
        self.is_safe = None         # type: Optional[bool]
        self.error = None           # type: Optional[str]

    @property
    def enabled(self):
        return bool(self.address)

    def poll(self, now):
        # type: (float) -> None
        if not self.enabled or now < self._next_poll:
            return
        self._next_poll = now + self._interval
        url = ('http://{}/api/v1/safetymonitor/0/issafe'
               '?ClientID=1783&ClientTransactionID=1'.format(
                   self.address.rstrip('/')))
        try:
            from urllib.request import urlopen
            with urlopen(url, timeout=self._timeout) as response:
                payload = json.loads(response.read().decode('utf-8'))
            if payload.get('ErrorNumber'):
                raise ValueError(payload.get('ErrorMessage') or 'device error')
            self.is_safe = bool(payload.get('Value'))
            self.error = None
        except Exception as exc:                # noqa: BLE001
            self.is_safe = None
            self.error = str(exc)


class WeatherWindow(object):
    """The window. Rendering only -- every decision is made in greenhill.core."""

    def __init__(self, root, config, settings):
        self.root = root
        self.config = config
        self.settings = settings

        self.evaluator = SafetyEvaluator(config)
        self.reader = MulticastReader()
        self.alerts = AlertPolicy(
            repeat_seconds=float(settings.get('alert_repeat_seconds', 30.0)),
            categories=settings.get('alert_categories',
                                    ('rain', 'safety', 'faults')))
        self.speaker = settings['speaker']
        self.authority = AuthoritativeVerdict(settings.get('safety_address', ''))

        self.started = time.monotonic()
        self.detector_labels = {}       # type: Dict[str, tk.Label]
        self.status_cells = {}          # type: Dict[str, tk.Label]
        self.temp_cells = {}            # type: Dict[str, tk.Label]
        self._arrow = None
        self._wedges = []               # type: List[int]
        self._smoothed_angle = 0.0

        self._join_streams()
        self._build()
        self._tick()

    # -- streams --------------------------------------------------------------

    def _join_streams(self):
        interface = self.settings.get('interface', '') or '0.0.0.0'
        self.reader.add('wind', self.settings['wind_group'],
                        int(self.settings['wind_port']), interface)
        self.reader.add('rain', self.settings['rain_group'],
                        int(self.settings['rain_port']), interface)
        for name, error in self.reader.failures.items():
            logger.error('could not join the %s stream: %s', name, error)

    def now(self):
        return time.monotonic() - self.started

    # -- layout ---------------------------------------------------------------

    def _build(self):
        self.root.title('Greenhill Weather')
        self.root.configure(bg=BACKGROUND)

        frame = tk.Frame(self.root, bg=BACKGROUND)
        frame.pack(fill='both', expand=True, padx=8, pady=8)

        self._build_map(frame)
        self._build_table(frame)
        self._build_compass(frame)
        self._build_banner()

    def _build_map(self, parent):
        self.map_canvas = tk.Canvas(parent, width=MAP_SIZE[0], height=MAP_SIZE[1],
                                    bg=BACKGROUND, highlightthickness=0)
        self.map_canvas.grid(row=0, column=0, sticky='nw')
        try:
            self.map_photo = tk.PhotoImage(file=asset_path(MAP_IMAGE))
            self.map_canvas.create_image(0, 0, anchor='nw', image=self.map_photo)
        except Exception as exc:                # noqa: BLE001
            # A missing site map costs a picture, not the display.
            logger.warning('site map unavailable: %s', exc)
            self.map_canvas.create_text(MAP_SIZE[0] // 2, MAP_SIZE[1] // 2,
                                        text='site map unavailable', fill=DIM)

        for name, (x, y) in DETECTOR_POSITIONS.items():
            label = tk.Label(self.map_canvas, width=1, padx=5, text='-',
                             font=('Courier', 16, 'bold'), bg='white',
                             relief='sunken')
            self.map_canvas.create_window(x, y, window=label)
            self.detector_labels[name] = label

    def _build_table(self, parent):
        panel = tk.Frame(parent, bg=BACKGROUND)
        panel.grid(row=0, column=1, sticky='nw', padx=12)

        for column, heading in enumerate(('ID', 'Status', 'Temp')):
            tk.Label(panel, text=heading, font=('Courier', 11, 'bold'),
                     bg=BACKGROUND, fg='white').grid(row=0, column=column,
                                                     padx=6, pady=4)

        for row, name in enumerate(DETECTOR_POSITIONS, start=1):
            tk.Label(panel, text=name, font=('Courier', 9, 'bold'),
                     bg=BACKGROUND, fg='white').grid(row=row, column=0, padx=6,
                                                     pady=3, sticky='w')
            status = tk.Label(panel, width=3, text='-', relief='raised',
                              font=('Courier', 9, 'bold'))
            status.grid(row=row, column=1, padx=6, pady=3)
            self.status_cells[name] = status
            temp = tk.Label(panel, width=7, text='-', bg='white',
                            relief='sunken', font=('Courier', 9))
            temp.grid(row=row, column=2, padx=6, pady=3)
            self.temp_cells[name] = temp

        self.rain_health = tk.Label(panel, text='', bg=BACKGROUND, fg=DIM,
                                    font=('Helvetica', 9), justify='left',
                                    wraplength=200)
        self.rain_health.grid(row=len(DETECTOR_POSITIONS) + 1, column=0,
                              columnspan=3, sticky='w', pady=(10, 0))

    def _build_compass(self, parent):
        self.compass = tk.Canvas(parent, width=COMPASS_SIZE[0],
                                 height=COMPASS_SIZE[1], bg=BACKGROUND,
                                 highlightthickness=0)
        self.compass.grid(row=0, column=2, sticky='nw')

        for text, (x, y) in (('N', (150, 30)), ('E', (270, 150)),
                             ('S', (150, 270)), ('W', (30, 150))):
            self.compass.create_text(x, y, text=text, fill='white',
                                     font=('Helvetica', 16))
        self.compass.create_line(150, 40, 150, 260, fill=EDGE, tags='overlay')
        self.compass.create_line(40, 150, 260, 150, fill=EDGE, tags='overlay')
        self.compass.create_oval(50, 50, 250, 250, outline=EDGE, tags='overlay')

        self.compass.create_rectangle(4, 300, 298, 348, outline=EDGE)
        self.readout_now = self.compass.create_text(
            50, 325, text='-', font=('Helvetica', 16), fill='white')
        self.readout_mean = self.compass.create_text(
            150, 325, text='-', font=('Helvetica', 16), fill='white')
        self.readout_gust = self.compass.create_text(
            250, 325, text='-', font=('Helvetica', 16), fill='white')
        for x, caption in ((50, 'inst'), (150, '60s'), (250, 'gust')):
            self.compass.create_rectangle(x - 20, 295, x + 20, 305,
                                          outline=BACKGROUND, fill=BACKGROUND)
            self.compass.create_text(x, 300, text=caption,
                                     font=('Helvetica', 10), fill=DIM)
        self.compass.create_rectangle(130, 345, 170, 355, outline=BACKGROUND,
                                      fill=BACKGROUND)
        self.compass.create_text(150, 350, text='km/h', font=('Helvetica', 9),
                                 fill=DIM)

    def _build_banner(self):
        banner = tk.Frame(self.root, bg=BACKGROUND)
        banner.pack(fill='x', padx=8, pady=(0, 8))
        self.verdict = tk.Label(banner, text='starting up',
                                font=('Helvetica', 15, 'bold'),
                                bg='gray', fg='white', padx=12, pady=6)
        self.verdict.pack(fill='x')
        self.detail = tk.Label(banner, text='', bg=BACKGROUND, fg=DIM,
                               font=('Helvetica', 9), justify='left',
                               anchor='w', wraplength=860)
        self.detail.pack(fill='x', pady=(4, 0))

    # -- the loop -------------------------------------------------------------

    def _tick(self):
        try:
            self._pump()
        except Exception as exc:                # noqa: BLE001
            # A display that dies leaves an astronomer with a frozen window and
            # no hint that it stopped being true.
            logger.exception('display update failed: %s', exc)
        self.root.after(250, self._tick)

    def _pump(self):
        now = self.now()
        for stream, payload in self.reader.poll(0.0):
            if stream == 'rain':
                try:
                    self.evaluator.rain.update(now, rain_protocol.parse(payload))
                except rain_protocol.ProtocolError as exc:
                    logger.debug('bad rain packet: %s', exc)
            else:
                try:
                    self.evaluator.wind.update(now, payload.decode('ascii'))
                except UnicodeDecodeError:
                    pass

        state = self.evaluator.update(now)
        self.authority.poll(now)
        self._render(now, state)
        self._announce(now, state)

    def _announce(self, now, state):
        rain_verdict, _ = self.evaluator.rain.verdict(now)
        wind_verdict, _ = self.evaluator.wind.verdict(now)
        snapshot = Snapshot(
            wet_sections=(self.evaluator.rain.wet_sections
                          if rain_verdict != rain_module.UNAVAILABLE else None),
            active_sections=self.evaluator.rain.observing_count * 2,
            rain_ok=rain_verdict != rain_module.UNAVAILABLE,
            wind_ok=wind_verdict != wind_module.UNKNOWN,
            is_safe=state.is_safe,
            reasons=state.reasons)
        for phrase in self.alerts.update(now, snapshot):
            logger.info('%s', phrase)
            self.speaker.say(phrase)

    # -- rendering ------------------------------------------------------------

    def _render(self, now, state):
        self._render_detectors(state)
        self._render_wind(now)
        self._render_banner(now, state)

    def _render_detectors(self, state):
        states = state.detector_states or {}
        for name in DETECTOR_POSITIONS:
            status = states.get(name, '-')
            foreground, background = STATUS_COLOURS.get(status,
                                                        STATUS_COLOURS['-'])
            self.detector_labels[name].configure(text=status, fg=foreground,
                                                 bg=background)
            self.status_cells[name].configure(text=status, fg=foreground,
                                              bg=background)

        temperature = state.temperature_c
        for name in DETECTOR_POSITIONS:
            present = name in states
            self.temp_cells[name].configure(
                text='{:.1f}C'.format(temperature)
                if present and temperature is not None else '-')

        age = state.rain_age_s
        if age is None:
            self.rain_health.configure(text='no rain data received',
                                       fg='#ff8888')
        elif age > self.config.rain_max_age_s:
            self.rain_health.configure(
                text='rain data {:.0f}s old'.format(age), fg='#ff8888')
        else:
            self.rain_health.configure(
                text='rain data {:.0f}s old'.format(age), fg=DIM)

    def _render_wind(self, now):
        mean = self.evaluator.wind.mean_speed_ms(now)
        gust = self.evaluator.wind.gust_ms(now)
        latest = self.evaluator.wind.latest_speed_ms(now)
        direction = self.evaluator.wind.direction_deg(now)
        stale = self.evaluator.wind.is_stale(now)

        def show(item, value):
            if value is None or stale:
                self.compass.itemconfig(item, text='-', fill=DIM)
            else:
                kmh = value * 3.6
                self.compass.itemconfig(item, text='{:.1f}'.format(kmh),
                                        fill=speed_colour(kmh))

        show(self.readout_now, latest)
        show(self.readout_mean, mean)
        show(self.readout_gust, gust)

        self._draw_spread(now, stale)
        self._draw_arrow(direction, mean, stale)
        self.compass.tag_raise('overlay')

    def _draw_spread(self, now, stale):
        for wedge in self._wedges:
            self.compass.delete(wedge)
        self._wedges = []
        if stale:
            return
        scatter = self.evaluator.wind.direction_scatter_deg(now)
        if scatter <= 0.0:
            return
        centre = self.evaluator.wind.direction_deg(now)
        radius = 100
        # Tk arcs measure anticlockwise from east; bearings run clockwise from
        # north.
        start = 90.0 - (centre + scatter)
        self._wedges.append(self.compass.create_arc(
            150 - radius, 150 - radius, 150 + radius, 150 + radius,
            start=start, extent=2 * scatter, fill='#444444', outline=''))

    def _draw_arrow(self, direction, mean, stale):
        if self._arrow is not None:
            self.compass.delete(self._arrow)
            self._arrow = None
        if stale or mean is None:
            return

        target = math.radians(direction)
        delta = (target - self._smoothed_angle + math.pi) % (2 * math.pi) - math.pi
        self._smoothed_angle += delta * 0.25

        radius = 70
        sin, cos = math.sin(self._smoothed_angle), math.cos(self._smoothed_angle)
        colour = speed_colour(mean * 3.6)
        self._arrow = self.compass.create_line(
            150 - radius * sin, 150 + radius * cos,
            150 + radius * sin, 150 - radius * cos,
            fill=colour, width=3, arrow='last')

    def _render_banner(self, now, state):
        if state.is_safe:
            self.verdict.configure(text='SAFE', bg='#00701f')
        else:
            self.verdict.configure(text='UNSAFE', bg='#8b0000')

        lines = list(state.reasons)
        if self.authority.enabled:
            if self.authority.is_safe is None:
                lines.append('observatory verdict unavailable ({})'.format(
                    self.authority.error or 'not polled yet'))
            elif self.authority.is_safe != state.is_safe:
                # Worth saying loudly. The observatory's verdict is the one that
                # closes the dome; a green panel beside a shut roof otherwise
                # looks like a fault in the roof.
                lines.append(
                    'NOTE: the observatory reports {} -- this window is a '
                    'local reading and does not control the dome.'.format(
                        'SAFE' if self.authority.is_safe else 'UNSAFE'))
        self.detail.configure(text='   '.join(lines) if lines else
                              'rain and wind within limits')

    def close(self):
        self.reader.close()
