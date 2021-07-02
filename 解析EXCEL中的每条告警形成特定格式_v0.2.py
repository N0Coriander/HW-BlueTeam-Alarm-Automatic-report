#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Author : N0Coriander
# Version : v0.2
# address : https://github.com/N0Coriander
# Date : 2021/5/31 16:58
# Desc : 手动复制excel中的某行数据，自动获取剪贴板中的内容更换成指定格式，然后再塞回剪贴板中，同时创建txt文本并塞进去

import pyperclip
import re
import os

while 1:
    input('[+] 先复制excel中的某一行，然后按回车键开始。。')
    info = pyperclip.paste()  # 读取剪贴板内容
    validinfo = re.sub(r'	0x.+', '', info)
    # print(validinfo)
    newtime = validinfo[0:19]  # 最近告警时间
    end = validinfo.split('	')
    name = end[1]    # 客户名称
    # expression = re.compile(r'((([01]?\d?\d|2[0-4]\d|25[0-5])\.){3}([01]?\d?\d|2[0-4]\d|25[0-5]))')
    expression = re.compile(r'((\d+\.){3}\d+)')
    result = expression.findall(validinfo)
    ip = []
    for i in result:
        ip.append(i[0])  # 提取IP
    if len(ip) > 1:
        eventname = end[-1] # 告警名称
        pyperclip.copy(f"""客户：{name}
报告序号：
事件名称：{eventname}
最近告警时间：{newtime}
攻击者IP：{ip[1]}
受害者：{ip[0]}
研判结论：攻击成功/登录成功/脆弱性/...
建议：防火墙封禁该攻击IP/确认是否为业务触发，若不是立即下线并上机排查/强化口令xxx""")
        path = os.getcwd()
        f = open(path + f'/{name}-通告.txt', 'a')
        f.write(pyperclip.paste() + '\n\n')
        f.close()
    else:
        eventname = end[-1]  # 告警名称
        pyperclip.copy(f"""客户：{name}
报告序号：
事件名称：{eventname}
受害者：{ip[0]}
建议：确认是否为业务触发，若不是立即下线并上机排查""")
        path = os.getcwd()
        f = open(path + f'/{name}-通告.txt', 'a')
        f.write(pyperclip.paste() + '\n\n')
        f.close()
    print('[!] 内容已提取，请自主粘贴')
    continue

