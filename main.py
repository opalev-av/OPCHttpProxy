import json
import os
import sys

import OpenOPC, pywintypes
from flask import Flask, request
import win32serviceutil
import servicemanager
import win32service
import logging
import time
import argparse

pywintypes.datetime = pywintypes.TimeType

CONFIG_FILE = 'config.json'
LOG_FILE = 'log.txt'

LISTEN = '127.0.0.1'
PORT = 3090

OPCServer = ""
white_list_tag = []

log = logging.getLogger('log')
log.setLevel(logging.DEBUG)


def create_config_file(file_name):

    if os.path.exists(file_name):
        return

    template_config = {'OPCServer':'', 'white_list':[], 'listen':'127.0.0.1', 'port':3090}

    with open(file_name,'w') as f:
        json.dump(template_config, f)


def get_tags_values(tags):
    result = {'tags': []}

    query_tags = []

    for tag in tags:
        if tag in white_list_tag:
            query_tags.append(tag)

    if len(query_tags) == 0:
        return result
    try:
        #opc = OpenOPC.client()
        opc = OpenOPC.open_client()
    except OpenOPC.OPCError:
        return {'Error': "Error initialise OPC Automation"}
    try:
        opc.connect(OPCServer)
    except OpenOPC.OPCError:
        return {'Error': f'Could not connect {OPCServer}'}
    values_tags = opc.read(query_tags)
    for val in values_tags:
        result['tags'].append({'tag': val[0], 'value': val[1], 'status': val[2]})
    opc.close()
    return result


app = Flask(__name__)


@app.route('/get-tags', methods=['POST'])
def request_get_tags():
    try:
        log.debug(f'request {request.remote_addr}, {request.data.decode("utf-8")}')
        data = json.loads(request.data)
        result = get_tags_values(data['tags'])
    except json.JSONDecodeError:
        return json.dumps({'Error': "non json"})
    except KeyError:
        return json.dumps({'Error': "key 'tags' not found"})

    return json.dumps(result,ensure_ascii=False)


def load_config(config_file_name):
    global OPCServer
    global white_list_tag
    global LISTEN
    global PORT

    create_config_file(config_file_name)

    with open(config_file_name, 'r', encoding='utf-8') as f:
        config = json.load(f)

    OPCServer = config['OPCServer']
    white_list_tag = config['white_list']
    LISTEN = config['listen']
    PORT = config['port']


class Service(win32serviceutil.ServiceFramework):
    _svc_name_ = "OPCHttpService"
    _svc_description_ = "OPC HTTP Service"
    _svc_display_name_ = 'OPC HTTP Serivce'

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        log.debug('Stoping')
        self.ReportServiceStatus(win32service.SERVICE_STOPPED)

    def SvcDoRun(self):
        self.ReportServiceStatus(win32service.SERVICE_START_PENDING)
        log.debug(f'Starting, {LISTEN}:{PORT}')
        self.service = app
        self.ReportServiceStatus(win32service.SERVICE_RUNNING)
        self.service.run(host=LISTEN, port=PORT)

def init_logging(log_file):
    handler_log = logging.FileHandler(log_file)
    formatter_log = logging.Formatter("%(asctime)s; %(message)s")
    handler_log.setFormatter(formatter_log)
    log.addHandler(handler_log)

def init():

    parser = argparse.ArgumentParser()
    parser.add_argument('install', nargs='?')
    parser.add_argument('remove', nargs='?')
    parser.add_argument('start', nargs='?')
    parser.add_argument('stop', nargs='?')
    parser.add_argument('restart', nargs='?')
    parser.add_argument('-conf', nargs='?')
    parser.add_argument('-log', nargs='?')

    params = parser.parse_args()

    if params.log != None:
        init_logging(params.log)

    if params.conf != None:
        load_config(params.conf)

    if len(sys.argv) == 1 or params.conf != None:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(Service)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(Service)

if __name__ == '__main__':
    init()
