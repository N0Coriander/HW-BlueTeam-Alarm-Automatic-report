#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Author : N0Coriander
# address : https://github.com/N0Coriander
# Date : 2021/6/2 09:10
# Desc : 运行脚本后，根据提示先手动复制excel中的某行数据，然后键盘回车，会自动获取剪贴板中的内容并更换成指定格式，然后再塞回剪贴板中，同时创建txt文本并塞进去
# Compare : 优化部分代码

import pyperclip
import os
while 1:
    input('[+] 先复制excel中的某一行，然后按回车键开始')
    info = pyperclip.paste()  # 读取剪贴板内容
    end = info.split('	')[0:6]
    name = end[1]
    if end[3] == '':
        # 放入剪贴板
        pyperclip.copy(f"""客户：{end[1]}
事件名称：{end[5]}
报告序号：
受害者：{end[2]}
建议：确认是否为业务触发，若不是立即下线并上机排查""")
        # 根据客户名称创建txt文本，并将生成的内容追加进去
        with open(os.getcwd() + f'/{name}-通告.txt', 'a') as f:
            f.write(pyperclip.paste() + '\n\n')
    else:
        pyperclip.copy(f"""客户：{end[1]}
事件名称：{end[5]}
报告序号：
最近告警时间：{end[0]}
攻击者IP：{end[3]}
受害者：{end[2]}
研判结论：攻击成功/登录成功/脆弱性/...
建议：防火墙封禁该攻击IP/确认是否为业务触发，若不是立即下线并上机排查/强化口令xxx""")
        with open(os.getcwd() + f'/{name}-通告.txt', 'a') as f:
            f.write(pyperclip.paste() + '\n\n')
    print('[!] 内容已提取，请自主粘贴')
    continue
