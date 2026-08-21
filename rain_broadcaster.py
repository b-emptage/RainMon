#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Entry point for the Greenhill rain-sensor bridge (Windows 7 box).

    python rain_broadcaster.py
    pyinstaller --onefile rain_broadcaster.py

Kept at the repo root, and thin, because PyInstaller freezes a script rather
than a package and the resulting executable must sit beside its .ini and log.
See greenhill/broadcaster.py for what it actually does.
"""

import sys

from greenhill.broadcaster import main

if __name__ == '__main__':
    sys.exit(main())
