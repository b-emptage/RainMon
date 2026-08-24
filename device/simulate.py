# -*- coding: utf-8 -*-
"""Run the weather server against synthetic dry, calm weather.

    python simulate.py

No hardware, no network, no multicast -- the sensor data is fed straight into
the safety core. This is what ASCOM Conform runs against, and what to use on a
bench to check the HTTP surface.

It is NOT watching the sky, and the log says so on every start.
"""

import app

if __name__ == '__main__':
    app.main(simulate=True)
