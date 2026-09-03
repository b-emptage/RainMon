#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""The Greenhill weather window -- rain, wind, and spoken alerts.

    python greenhill_monitor.py
    python greenhill_monitor.py --mute
    pyinstaller --onefile --windowed --add-data "BT_SiteVectorMap.png;." greenhill_monitor.py

Listens to the observatory's two multicast streams and shows what they say. It
commands nothing, needs neither Arcsecond nor the weather service to be
running, and can be opened on as many machines as anyone likes. Closing it
stops nothing.
"""

import argparse
import configparser
import logging
import os
import sys

from greenhill.core.config import WeatherConfig
from greenhill.ui.speaker import Speaker

CONFIG_FILENAME = 'greenhill_monitor.ini'
CONFIG_SECTION = 'monitor'

DEFAULTS = {
    'wind_group': '239.192.0.4',
    'wind_port': '60004',
    'rain_group': '239.192.0.5',
    'rain_port': '60005',
    'interface': '',
    'speaker': 'true',
    'alert_repeat_seconds': '30',
    'alert_categories': 'rain, safety, faults',
    'safety_address': '',
    'log_level': 'INFO',
}


def config_directory():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def load_settings(path=None):
    parser = configparser.ConfigParser()
    parser[CONFIG_SECTION] = dict(DEFAULTS)
    if path is None:
        path = os.path.join(config_directory(), CONFIG_FILENAME)
    if os.path.exists(path):
        try:
            parser.read(path)
        except configparser.Error as exc:
            print('==CONFIG== {} is malformed ({}); using defaults.'.format(
                path, exc), file=sys.stderr)

    settings = dict(DEFAULTS)
    if parser.has_section(CONFIG_SECTION):
        settings.update({key: parser.get(CONFIG_SECTION, key)
                         for key in parser.options(CONFIG_SECTION)})
    settings['alert_categories'] = tuple(
        part.strip().lower()
        for part in settings['alert_categories'].split(',') if part.strip())
    settings['alert_repeat_seconds'] = float(settings['alert_repeat_seconds'])
    return settings


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.split('\n')[0])
    parser.add_argument('--config', help='path to greenhill_monitor.ini')
    parser.add_argument('--mute', action='store_true',
                        help='show everything, say nothing')
    parser.add_argument('--verbose', action='store_true')
    args = parser.parse_args(argv)

    settings = load_settings(args.config)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else
        getattr(logging, settings['log_level'].upper(), logging.INFO),
        format='%(asctime)s %(levelname)s %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S')

    try:
        import tkinter as tk
    except ImportError:
        print('This display needs Tk, which is not available in this Python.',
              file=sys.stderr)
        return 2

    # Imported here so --help works on a machine with no display.
    from greenhill.ui.app import WeatherWindow

    speak = settings['speaker'].strip().lower() in ('1', 'true', 'yes', 'on')
    speaker = Speaker(enabled=speak and not args.mute)
    speaker.start()
    settings['speaker'] = speaker

    root = tk.Tk()
    window = WeatherWindow(root, WeatherConfig(), settings)
    root.protocol('WM_DELETE_WINDOW', root.destroy)
    try:
        root.mainloop()
    finally:
        window.close()
        speaker.stop()
    return 0


if __name__ == '__main__':
    sys.exit(main())
