# coding=utf-8

import requests
import time
import os
import sys
import websocket
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__),".."))
sys.path.append(BASE_DIR)
from utils.log import log, ERROR


def save_html(content):
    with open('2.html', 'w', encoding='utf-8') as f:
        f.write(content)


def recive_from(url):
    websocket.enableTrace(True)
    ws = websocket.WebSocketApp(url,
                                on_message=on_message,
                                on_error=on_error,
                                on_close=on_close)
    ws.on_open = on_open
    ws.run_forever(ping_timeout=30)


def on_message(ws, message):  # 服务器有数据更新时，主动推送过来的数据
    print(message)


def on_error(ws, error):  # 程序报错时，就会触发on_error事件
    print(error)


def on_close(ws):
    print("Connection closed ……")


def on_open(ws):  # 连接到服务器之后就会触发on_open事件，这里用于send数据
    req = '{"event":"subscribe", "channel":"btc_usdt.deep"}'
    print(req)
    ws.send(req)


   



if __name__ == '__main__':
    # https://www.binance.com/fapi/v1/time
    recive_from("wss://fstream.binance.com/stream?<symbol>@depth<levels>@100ms")

