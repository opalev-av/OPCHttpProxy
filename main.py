import json
import os
import OpenOPC, pywintypes
from flask import Flask, request
import time

pywintypes.datetime = pywintypes.TimeType

CONFIG_FILE = 'config.json'

OPCServer = ""
white_list_tag = []

def get_tags_values(tags):

    result = {'tags':[]}

    query_tags = []

    for tag in tags:
        if tag in white_list_tag:
            query_tags.append(tag)

    if len(query_tags) == 0:
        return result

    opc = OpenOPC.client()
    opc.connect(OPCServer)
    values_tags = opc.read(query_tags)
    for val in values_tags:
        result['tags'].append({'tag':val[0], 'value':val[1], 'status':val[2], 'time':[3]})
    opc.close()
    return result

app = Flask(__name__)


@app.route('/get-tags', methods=['POST'])
def request_get_tags():
    data = json.loads(request.data)
    result = get_tags_values(data['tags'])
    return json.dumps(result)


def load_config():
    global OPCServer
    global white_list_tag
    config_file_name = f'{os.path.dirname(os.path.realpath(__name__))}/{CONFIG_FILE}'

    with open(config_file_name,'r') as f:
        config = json.load(f)
    OPCServer = config['OPCServer']
    white_list_tag = config['white_list']

load_config()



