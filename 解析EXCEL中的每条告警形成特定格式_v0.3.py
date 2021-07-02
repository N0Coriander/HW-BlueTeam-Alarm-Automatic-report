#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Author : N0Coriander
# address : https://github.com/N0Coriander
# Date : 2021/5/31 18:04
# Desc : 手动复制excel中的某行数据，自动获取剪贴板中的内容更换成指定格式，然后再塞回剪贴板中，同时创建txt文本并塞进去

import pyperclip
import os
while 1:
    input('[+] 先复制excel中的某一行，然后按回车键开始。。')
    info = pyperclip.paste()  # 读取剪贴板内容
    end = info.split('	')[0:6]
    name = end[1]
    if end[3] == '':
        pyperclip.copy(f"""客户：{end[1]}
报告序号：
事件名称：{end[5]}
受害者：{end[2]}
建议：确认是否为业务触发，若不是立即下线并上机排查""")
        path = os.getcwd()
        f = open(path + f'/{name}-通告.txt', 'a')
        f.write(pyperclip.paste() + '\n\n')
        f.close()
    else:
        pyperclip.copy(f"""客户：{end[1]}
报告序号：
事件名称：{end[5]}
最近告警时间：{end[0]}
攻击者IP：{end[3]}
受害者：{end[2]}
研判结论：攻击成功/登录成功/脆弱性/...
建议：防火墙封禁该攻击IP/确认是否为业务触发，若不是立即下线并上机排查/强化口令xxx""")
        path = os.getcwd()
        f = open(path + f'/{name}-通告.txt', 'a')
        f.write(pyperclip.paste() + '\n\n')
        f.close()
    print('[!] 内容已提取，请自主粘贴')
    continue

