# -*- coding: utf-8 -*-
# -----------------------------------------------------------------------------
# app.py - Alpaca device server for the Greenhill weather safety system.
#
# Adapted from the AlpycaDevice sample app.py --
# https://github.com/ASCOMInitiative/AlpycaDevice -- MIT, (c) 2022-2024 Bob Denny.
# See LICENSE-AlpycaDevice.txt.
#
# Taken by way of Greenhill-DomeShutter rather than from upstream, because that
# server already carries two fixes Conform found in the sample: a duplicated
# ClientID returning HTTP 500 instead of 400, and an HTTP/1.0 status line that
# reset every pooled .NET client connection.
#
# Changes here: serves TWO device types instead of one, and owns a weather
# monitoring thread that runs whether or not any client is connected.
# -----------------------------------------------------------------------------
import inspect
import os
import signal
import socket
import sys
import traceback
from enum import IntEnum
from socketserver import ThreadingMixIn
from wsgiref.simple_server import (ServerHandler, WSGIRequestHandler,
                                   WSGIServer, make_server)

# The safety core lives in the greenhill package at the repo root, one level up
# from here. This server is run as `python app.py` from inside device/, which
# is the AlpycaDevice convention, so the root is not otherwise on the path.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import alpaca_common                                        # noqa: E402
import discovery                                            # noqa: E402
import exceptions                                           # noqa: E402
import log                                                  # noqa: E402
import management                                           # noqa: E402
import observingconditions                                  # noqa: E402
import safetymonitor                                        # noqa: E402
import setup                                                # noqa: E402
from config import Config                                   # noqa: E402
from discovery import DiscoveryResponder                    # noqa: E402
from falcon import App, HTTPInternalServerError, Request, Response  # noqa: E402
from shr import set_shr_logger                              # noqa: E402
from weatherdevice import GreenhillWeather                  # noqa: E402

from greenhill.core.config import WeatherConfig             # noqa: E402

API_VERSION = 1

# Loopback port held as a process mutex. The dome server uses 50815 and the
# rain bridge 50816; this is the next along. Two weather servers would both
# join the multicast groups and both answer on the Alpaca port, and which one a
# client reached would depend on which won the bind.
_SINGLE_INSTANCE_PORT = 50817
_instance_guard = None
_shutdown_signal = None

weather_dev = None


def acquire_single_instance_lock() -> bool:
    global _instance_guard
    guard = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        guard.bind(('127.0.0.1', _SINGLE_INSTANCE_PORT))
        guard.listen(1)
    except OSError:
        guard.close()
        return False
    _instance_guard = guard
    return True


def _signal_shutdown(signum, frame):
    """Unwind main() so its finally block stops the monitoring thread."""
    global _shutdown_signal
    _shutdown_signal = signum
    signal.signal(signum, signal.SIG_DFL)   # a second signal kills us outright
    # Raised on the main thread inside serve_forever()'s select, so it
    # propagates exactly as KeyboardInterrupt already does. Do NOT call
    # httpd.shutdown() here: it waits on the loop this thread is running.
    sys.exit(128 + signum)


def install_signal_handlers() -> list:
    installed = []
    for signame in ('SIGTERM', 'SIGBREAK'):
        sig = getattr(signal, signame, None)
        if sig is None:
            continue
        try:
            signal.signal(sig, _signal_shutdown)
        except (ValueError, OSError):
            continue
        installed.append(signame)
    return installed


class ThreadingWSGIServer(ThreadingMixIn, WSGIServer):
    """One thread per connection. ASCOM clients poll several properties at a
    time, and wsgiref's plain server would serialise them all behind whichever
    request happened to be slowest."""
    daemon_threads = True


class Http11ServerHandler(ServerHandler):
    """wsgiref.handlers.BaseHandler hardcodes http_version = '1.0', and it --
    not WSGIRequestHandler.protocol_version -- writes the status line."""
    http_version = '1.1'


class LoggingWSGIRequestHandler(WSGIRequestHandler):
    # Every ASCOM .NET client pools connections through HttpClient. Served as
    # HTTP/1.0 the connection closes after each response and the client's next
    # request on that pooled socket is reset, surfacing as an error on whatever
    # member happened to be next.
    protocol_version = 'HTTP/1.1'
    timeout = 60                # idle keep-alive connections hold a thread each

    def handle(self):
        """Serve every request on this connection, not just the first.

        wsgiref's own handle() serves exactly one and returns, so there is no
        keep-alive to be had from it whatever protocol_version says.
        """
        self.close_connection = True
        self.handle_one_request()
        while not self.close_connection:
            self.handle_one_request()

    def handle_one_request(self):
        try:
            self.raw_requestline = self.rfile.readline(65537)
        except (TimeoutError, socket.timeout, ConnectionError, OSError):
            self.close_connection = True
            return

        if not self.raw_requestline:                # client hung up
            self.close_connection = True
            return
        if len(self.raw_requestline) > 65536:
            self.requestline = ''
            self.request_version = ''
            self.command = ''
            self.send_error(414)
            self.close_connection = True
            return
        if not self.parse_request():                # an error was already sent
            self.close_connection = True
            return

        handler = Http11ServerHandler(
            self.rfile, self.wfile, self.get_stderr(), self.get_environ(),
            multithread=True)
        handler.request_handler = self
        try:
            handler.run(self.server.get_app())
        except (ConnectionError, OSError):
            self.close_connection = True
            return

    def log_message(self, format: str, *args):
        # Requests are logged on the way in by shr.log_request, which keeps the
        # log in causal order. Suppress the wsgiref duplicate.
        pass


def init_routes(app: App, devname: str, module):
    """Route each responder class in `module` to its Alpaca URI by class name.

    Only classes DEFINED in the module are routed -- which is why each device
    module declares one-line subclasses of the shared bases in alpaca_common
    rather than importing them directly.
    """
    for cname, ctype in inspect.getmembers(module, inspect.isclass):
        if ctype.__module__ == module.__name__ and not issubclass(ctype, IntEnum):
            app.add_route(
                f'/api/v{API_VERSION}/{devname}/{{devnum:int(min=0)}}/{cname.lower()}',
                ctype())


def custom_excepthook(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    log.logger.error(f'An uncaught {exc_type.__name__} exception occurred:')
    log.logger.error(exc_value)
    if Config.verbose_driver_exceptions and exc_traceback:
        for line in traceback.format_tb(exc_traceback):
            log.logger.error(repr(line))


def falcon_uncaught_exception_handler(req: Request, resp: Response,
                                      ex: BaseException, params):
    exc = sys.exc_info()
    custom_excepthook(exc[0], exc[1], exc[2])
    # Keyword arguments, not positional: falcon made these keyword-only in 3.0,
    # so the positional form in the AlpycaDevice sample raises a TypeError of
    # its own -- from inside the handler whose job is to report errors, which
    # replaces the real fault with a confusing one at the worst moment.
    raise HTTPInternalServerError(
        title='Internal Server Error',
        description='Alpaca endpoint responder failed. See logfile.')


def build_app() -> App:
    """The Falcon app, with both device types routed. Separated from main() so
    the tests can drive the real routing through a test client."""
    falc_app = App()
    init_routes(falc_app, 'observingconditions', observingconditions)
    init_routes(falc_app, 'safetymonitor', safetymonitor)
    falc_app.add_route('/management/apiversions', management.apiversions())
    falc_app.add_route(f'/management/v{API_VERSION}/description',
                       management.description())
    falc_app.add_route(f'/management/v{API_VERSION}/configureddevices',
                       management.configureddevices())
    falc_app.add_route('/setup', setup.svrsetup())
    falc_app.add_route(
        f'/setup/v{API_VERSION}/observingconditions/{{devnum}}/setup',
        setup.devsetup())
    falc_app.add_route(f'/setup/v{API_VERSION}/safetymonitor/{{devnum}}/setup',
                       setup.devsetup())
    falc_app.add_error_handler(Exception, falcon_uncaught_exception_handler)
    return falc_app


def start_weather_device(logger, simulate=False):
    """Create the weather service and start it monitoring.

    Started here, at server startup, and NOT when a client connects. The
    monitoring is the point of this process; the ASCOM surface is how other
    software reads it. In particular the direct dome-close route has no ASCOM
    client at all, so tying the thread to Connected would disarm it.
    """
    global weather_dev
    if weather_dev is not None:
        weather_dev.stop()
    weather_dev = GreenhillWeather(
        WeatherConfig.from_mapping(Config.weather or {}),
        {'wind_group': Config.wind_group, 'wind_port': Config.wind_port,
         'rain_group': Config.rain_group, 'rain_port': Config.rain_port,
         'interface': Config.interface or '0.0.0.0'},
        logger)
    alpaca_common.set_weather_device(weather_dev)
    weather_dev.start(simulate=simulate)
    return weather_dev


def main(simulate=False):
    # The guard comes FIRST, before the logger. init_logging rotates the log,
    # and rotating means renaming a file the incumbent still has open: on
    # Windows that raises, so startup would die with an unexplained log-file
    # error instead of the message below.
    if not acquire_single_instance_lock():
        print('==STARTUP FAILED== Another Greenhill weather server is already '
              'running. Two of them would both join the sensor streams and '
              'both answer on the Alpaca port; refusing to start.',
              file=sys.stderr)
        sys.exit(1)

    logger = log.init_logging()
    log.logger = logger
    exceptions.logger = logger
    discovery.logger = logger
    alpaca_common.set_common_logger(logger)
    management.set_management_logger(logger)
    set_shr_logger(logger)

    sys.excepthook = custom_excepthook
    start_weather_device(logger, simulate=simulate)

    installed = install_signal_handlers()
    if 'SIGTERM' not in installed:
        logger.warning('==SIGNALS== No SIGTERM handler could be installed; '
                       'stop the server with Ctrl-C.')
    elif sys.platform == 'win32':
        logger.warning('==SIGNALS== Windows delivers no SIGTERM. Stop the '
                       'server with Ctrl-C or Ctrl-Break.')

    # Discovery is a convenience. On this box the equipment Alpaca server is
    # expected to already hold UDP 32227, so losing the bind is NORMAL here and
    # must not stop the weather from being served: Arcsecond is configured with
    # an explicit address regardless.
    try:
        DiscoveryResponder(Config.ip_address, Config.port)
    except Exception as ex:
        logger.warning(
            f'==DISCOVERY DISABLED== Could not bind the Alpaca discovery port: '
            f'{ex}. Expected on this machine, where the equipment server holds '
            f'it. The weather devices are still fully readable at '
            f'{Config.ip_address or "0.0.0.0"}:{Config.port}; clients must be '
            f'given that address explicitly.')

    try:
        with make_server(Config.ip_address, Config.port, build_app(),
                         server_class=ThreadingWSGIServer,
                         handler_class=LoggingWSGIRequestHandler) as httpd:
            logger.info(f'==STARTUP== Greenhill weather server on '
                        f'{Config.ip_address}:{Config.port} serving '
                        f'ObservingConditions and SafetyMonitor. '
                        f'Time stamps are UTC.')
            httpd.serve_forever()
    finally:
        if _shutdown_signal is not None:
            logger.info(f'==SIGNAL== {signal.Signals(_shutdown_signal).name} '
                        f'received; stopping.')
        if weather_dev is not None:
            weather_dev.stop()
        logger.info('==SHUTDOWN== Greenhill weather server stopped.')


if __name__ == '__main__':
    main()
