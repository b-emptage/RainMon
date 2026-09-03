# -*- coding: utf-8 -*-
"""Multicast plumbing, shared by the tools and by the service.

Deliberately thin. Everything above this is driven by (arrival time, bytes) and
knows nothing about sockets, which is what lets the whole safety core be tested
by replaying a capture file instead of a network.
"""

import select
import socket
from typing import Dict, Iterator, List, Optional, Tuple


def join_multicast_group(group, port, interface='0.0.0.0'):
    # type: (str, int, str) -> socket.socket
    """Join one group and return a non-blocking socket.

    SO_REUSEADDR before bind is what lets several processes receive the same
    stream at once -- the legacy wind display, a recorder and the weather
    service can all listen together. That property is what makes the whole
    multicast design workable, and it is why nothing here ever needs exclusive
    access to a sensor.
    """
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM, socket.IPPROTO_UDP)
    sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    if hasattr(socket, 'SO_REUSEPORT'):
        # BSD and macOS need this as well as SO_REUSEADDR to share a bound port.
        try:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEPORT, 1)
        except OSError:
            pass
    sock.bind(('', port))
    sock.setsockopt(socket.IPPROTO_IP, socket.IP_ADD_MEMBERSHIP,
                    socket.inet_aton(group) + socket.inet_aton(interface))
    sock.setblocking(False)
    return sock


class MulticastReader(object):
    """Reads several named multicast streams through one select loop."""

    def __init__(self):
        self._sockets = {}      # type: Dict[socket.socket, str]
        self.failures = {}      # type: Dict[str, str]

    def add(self, name, group, port, interface='0.0.0.0'):
        # type: (str, str, int, str) -> bool
        """Join a stream. Returns False and records why on failure.

        A stream that cannot be joined is NOT fatal here: the caller decides.
        For the weather service the answer is to carry on and let the missing
        stream age into UNKNOWN, which is unsafe -- refusing to start would
        leave the observatory with no safety service at all, which is worse.
        """
        try:
            self._sockets[join_multicast_group(group, port, interface)] = name
            return True
        except OSError as exc:
            self.failures[name] = str(exc)
            return False

    @property
    def stream_count(self):
        return len(self._sockets)

    @property
    def stream_names(self):
        # type: () -> List[str]
        """Names of the streams whose join succeeded -- the ones actually being
        listened to, as opposed to those recorded in `failures`."""
        return list(self._sockets.values())

    def poll(self, timeout=0.5):
        # type: (float) -> List[Tuple[str, bytes]]
        """Return whatever has arrived, as (stream name, payload)."""
        if not self._sockets:
            return []
        readable, _, _ = select.select(list(self._sockets), [], [], timeout)
        received = []
        for sock in readable:
            try:
                payload, _ = sock.recvfrom(65535)
            except OSError:
                continue
            received.append((self._sockets[sock], payload))
        return received

    def close(self):
        for sock in list(self._sockets):
            sock.close()
        self._sockets.clear()
