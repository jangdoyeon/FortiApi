"""
Create Fortigate-Object from excel
"""

__author__ = "dyjang"

import pyfortiapi
from pprint import pprint
import copy
import json
import time
import openpyxl as opx
import sys
from os import path


class Pcolors:
    """
    for colorful cli output
    """

    RED = '\033[95m'
    BLUE = '\033[94m'
    GREEN = '\033[92m'
    YELLO = '\033[93m'
    ORANGE = '\033[91m'
    GRAY = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


SETTING_FILE = 'setting.json'
json_load = json.load(open(SETTING_FILE))

try:

    if sys.argv[1] == 'initaccount':
        with open(SETTING_FILE, 'w') as f:
            username = input('F/W USERNAME : ')
            password = input('F/W PASSWORLD : ')
            json_data = {
                'username': username,
                'passworld': password,
                'xls': 'data.xlsx',
                'sheet': 'Sheet1'
            }
            f.write(json.dumps(json_data))
            print(f'{Pcolors.GREEN}[+] Setting Complete')
            for key, value in json_data.items():
                print(f'{Pcolors.GRAY}[-] {key} : {value}')
            sys.exit()
    if not path.exists(SETTING_FILE):
        print(
            f'{Pcolors.YELLO}[E] Please initialize setting\n{Pcolors.GRAY}[-] fortiapi(.py) initaccount')
        sys.exit()
    _FW = sys.argv[1]
    _VDOM = 'root'
    if len(sys.argv) == 3:
        _VDOM = sys.argv[2]
except IndexError:
    print("""
 _____          _   _      _    ____ ___
|  ___|__  _ __| |_(_)    / \  |  _ \_ _|
| |_ / _ \| '__| __| |   / _ \ | |_) | |
|  _| (_) | |  | |_| |  / ___ \|  __/| |
|_|  \___/|_|   \__|_| /_/   \_\_|  |___|

Usage:
  fortiapi(.py) <fw host name> [vdom name]

Example:
  fortiapi(.py) office root

F/W Host Name List (from setting.json):
  {0}

VDOM Name List:
  root (default)

Initialize:
  fortiapi(.py) initaccount
    """.format('\n  '.join(list(json_load["fw"].keys()))))
    sys.exit()


_USERNAME = json_load["username"]
_PASSWORD = json_load["passworld"]
_IP_ADDR = ''
_SHEET_PATH = json_load["xls"]
_SHEET_NAME = json_load["sheet"]

fw_keys_arr = json_load["fw"].keys()

if _FW in fw_keys_arr:
    _IP_ADDR = json_load["fw"][_FW]
else:
    print(f'{Pcolors.YELLO}[E] Incorrect name of FW')


def get_addr_obj(device, obj_name):
    """get addr obj"""

    object_name = obj_name
    addresses = device.get_firewall_address(object_name)
    pprint(addresses)


def set_addr_obj(device, name, type, detail, comment):  # name, type, subnet, comment
    """set addr obj"""

    dic = dict()
    if type == 'subnet':
        dic['name'] = name
        dic['type'] = type
        dic['subnet'] = detail
        dic['comment'] = comment
    elif type == 'fqdn':
        dic['name'] = name
        dic['type'] = type
        dic['fqdn'] = detail
        dic['comment'] = comment
    jsonstring = json.dumps(dic)

    result = device.create_firewall_address(name, jsonstring)
    if result == 424:  # alread exist
        update_payload = dict()
        update_payload['comment'] = comment
        update_jsonstring = json.dumps(update_payload)
        update_result = device.update_firewall_address(name, update_jsonstring)
        print(
            f'{Pcolors.GREEN}[+] Object update result code : {update_result}')
        print(f'{Pcolors.GREEN}[-] Update comment : {update_jsonstring}')
    elif result != 200:
        print(result, jsonstring)


def create_addr_obj(device, ws):
    cnt = 0  # for object count
    for row in ws.rows:
        line = list()
        if row[0].value == None:
            return 0  # first culumn in row == None ?
        for val in row:
            if val.value == None:
                break
            line.append(val.value)

        name, type, detail, coment, input_detail = line
        print(
            f'{Pcolors.GREEN}[-] {cnt+1}. Object name : {name}, Object detail : {detail}')
        set_addr_obj(device, name, type, detail, coment)
        cnt = cnt + 1


def get_policy(device):
    policy = device.get_firewall_policy(filters='')
    pprint(policy)


def set_policy(device):
    payload = {'policyid': 900,
               'name': 'Test Policy',
               'srcintf': [{'name': 'port1'}],
               'dstintf': [{'name': 'port16'}],
               'srcaddr': [{'name': 'all'}],
               'dstaddr': [{'name': 'all'}],
               'action': 'accept',
               'status': 'enable',
               'schedule': 'always',
               'service': [{'name': 'ALL'}],
               'nat': 'enable',
               'fsso': 'enable',
               'wsso': 'disable',
               'rsso': 'enable'}
    payload = repr(payload)
    device.create_firewall_policy(500, payload)


def excel_to_dic(ws):
    """Convert excel to dictionary"""

    all_arr = []
    for row in ws.rows:
        row_arr = []
        for cell in row:
            row_arr.append(cell.value)
        all_arr.append(row_arr)

    key_arr = all_arr[0]
    dic = {}
    dic1 = {}
    data = []

    for i in range(1, len(all_arr)):
        idx = 0
        for c in all_arr[i]:
            dic[key_arr[idx]] = c
            idx += 1
        dic1 = dic.copy()
        data.append(dic1)
    print(data)


def init():
    device = pyfortiapi.FortiGate(
        ipaddr=_IP_ADDR, username=_USERNAME, password=_PASSWORD, vdom=_VDOM)
    wb = opx.load_workbook(_SHEET_PATH, data_only=True)
    ws = wb[_SHEET_NAME]
    return device, ws


if __name__ == '__main__':
    device, ws = init()
    if ws:
        print(f'{Pcolors.BLUE}[-] Successfully got SHEET')
    create_addr_obj(device, ws)
