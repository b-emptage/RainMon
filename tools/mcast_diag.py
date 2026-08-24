#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Which interface actually receives which multicast stream?

Run this ON THE BOX that is only seeing one sensor. It joins the wind and rain
groups on EACH local interface in turn -- and, separately, on INADDR_ANY the way
the service does -- and prints how many datagrams each combination received in a
few seconds. If the box is multi-homed and the OS is joining on the wrong NIC,
this shows it in one screen: one interface will get both streams, INADDR_ANY may
get one or none.

    python3 tools/mcast_diag.py
    python3 tools/mcast_diag.py --seconds 6

No dependencies beyond the standard library, so it runs on the Win11 box as-is.
"""

import argparse
import select
import socket
import sys
import time

try:
    from greenhill import rain_protocol  # noqa: E402
    import os
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    WIND = (rain_protocol.WIND_MULTICAST_GROUP, rain_protocol.WIND_MULTICAST_PORT)
    RAIN = (rain_protocol.MULTICAST_GROUP, rain_protocol.MULTICAST_PORT)
except Exception:
    WIND = ('239.192.0.4', 60004)
    RAIN = ('239.192.0.5', 60005)


def local_ipv4_addresses():
    """Best-effort list of this host's IPv4 interface addresses.

    Uses the hostname table plus a UDP-connect probe toward each sensor, which
    reveals the address of the interface the OS would actually route to it.
    """
    found = set()
    try:
        _, _, addrs = socket.gethostbyname_ex(socket.gethostname())
        found.update(a for a in addrs if not a.startswith('127.'))
    except OSError:
        pass
    for group, _port in (WIND, RAIN):
        probe = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        try:
            probe.connect((group, 1))          # no packet leaves; just routing
            found.add(probe.getsockname()[0])
        except OSError:
            pass
        finally:
            probe.close()
    found.discard('0.0.0.0')
    return sorted(found)


def join(group, port, interface):
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM, socket.IPPROTO_UDP)
    s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    if hasattr(socket, 'SO_REUSEPORT'):
        try:
            s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEPORT, 1)
        except OSError:
            pass
    s.bind(('', port))
    s.setsockopt(socket.IPPROTO_IP, socket.IP_ADD_MEMBERSHIP,
                 socket.inet_aton(group) + socket.inet_aton(interface))
    s.setblocking(False)
    return s


def listen(interface, seconds):
    """Return (wind_count, rain_count) received via `interface` in `seconds`."""
    socks = {}
    try:
        socks[join(WIND[0], WIND[1], interface)] = 'wind'
        socks[join(RAIN[0], RAIN[1], interface)] = 'rain'
    except OSError as exc:
        for s in socks:
            s.close()
        return None, None, str(exc)

    counts = {'wind': 0, 'rain': 0}
    deadline = time.time() + seconds
    while time.time() < deadline:
        readable, _, _ = select.select(list(socks), [], [], 0.25)
        for s in readable:
            try:
                s.recv(65535)
                counts[socks[s]] += 1
            except OSError:
                pass
    for s in socks:
        s.close()
    return counts['wind'], counts['rain'], None


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.split('\n')[0])
    parser.add_argument('--seconds', type=float, default=5.0,
                        help='how long to listen on each interface')
    args = parser.parse_args(argv)

    print('wind group {}:{}   rain group {}:{}'.format(WIND[0], WIND[1],
                                                       RAIN[0], RAIN[1]))
    print('listening {:.0f}s per interface. ~1 packet/s per healthy stream.'
          .format(args.seconds))
    print('')

    candidates = ['0.0.0.0 (INADDR_ANY -- what the service uses)']
    interfaces = ['0.0.0.0'] + local_ipv4_addresses()
    seen = set()
    print('{:<40} {:>8} {:>8}'.format('interface', 'wind', 'rain'))
    print('-' * 58)
    for iface in interfaces:
        if iface in seen:
            continue
        seen.add(iface)
        label = iface + (' (INADDR_ANY)' if iface == '0.0.0.0' else '')
        wind, rain, err = listen(iface, args.seconds)
        if err is not None:
            print('{:<40} {:>8} {:>8}   ERROR: {}'.format(label, '-', '-', err))
            continue
        flag = ''
        if wind and not rain:
            flag = '  <- WIND only'
        elif rain and not wind:
            flag = '  <- RAIN only'
        elif wind and rain:
            flag = '  <- both (use this one)'
        elif not wind and not rain:
            flag = '  <- nothing'
        print('{:<40} {:>8} {:>8}{}'.format(label, wind, rain, flag))

    print('')
    print('If one specific interface shows BOTH but INADDR_ANY (0.0.0.0) does '
          'not,\nset  interface = "<that address>"  in device/config.toml.')


if __name__ == '__main__':
    sys.exit(main())
