#!/usr/bin/python 
# -*- coding: utf-8 -*-

# 青岛西海岸新区致远中学附属大学 强智教务系统API
# By Github@WinrunnerMax https://github.com/WindrunnerMax/SHST/
# API完整版见上述Github页面，这里仅保留与课表相关的API，并针对相关适用性进行适当修改

import time
import requests
import re
import os
import json
import datetime

class SW(object):
    """docstring for SW"""
    def __init__(self, account, password, url):
        super(SW, self).__init__()
        self.url = url
        self.account = account
        self.password = password
        self.session = self.login()
    
    HEADERS = {
        "User-Agent":"Mozilla/5.0 (Linux; U; Mobile; Android 6.0.1;C107-9 Build/FRF91 )",
        "Referer": "http://www.baidu.com",
        "Accept-encoding": "gzip, deflate, br",
        "Accept-language": "zh-CN,zh-TW;q=0.8,zh;q=0.6,en;q=0.4,ja;q=0.2",
        "Cache-control": "max-age=0"
    }

    def login(self):
        params = {
            "method" : "authUser",
            "xh" : self.account,
            "pwd" : self.password
        }
        session = requests.Session()
        req = session.get(self.url, params = params, timeout = 5, headers = self.HEADERS)
        s = json.loads(req.text)
        # print(s)
        if s["flag"] != "1" : exit(0)
        self.HEADERS["token"] = s["token"]
        return session

    def get_handle(self,params):
        req = self.session.get(self.url, params = params ,timeout = 5 ,headers = self.HEADERS)
        return req

    def get_student_info(self):
        params = {
            "method" : "getUserInfo",
            "xh" : self.account
        }
        req = self.get_handle(params)
        # print(req.text)
    
    def get_current_time(self):
        params = {
            "method" : "getCurrentTime",
            "currDate" : datetime.datetime.now().strftime("%Y-%m-%d")
        }
        req = self.get_handle(params)
        # print(req.text)
        return req.text

    def get_class_info(self,zc = -1):
        s = json.loads(self.get_current_time())
        params = {
            "method" : "getKbcxAzc",
            "xnxqid" : s["xnxqh"],
            "zc" : s["zc"] if zc == -1 else zc,
            "xh" : self.account
        }
        req = self.get_handle(params)
        # print(req.text)
        return req.text
