# HW-BlueTeam-Alarm-Automatic-report
> HW中蓝队成员每天都需要将告警转换成指定格式上报，同时在下班后还需要提交日报。该工具针对以上两个场景实现了自动化输出
Desc : 用户自己先在导出来的EXCEL告警表中进行筛选，复制要生成通告的数据(支持多行)，然后输入你想生成的格式对应的数字，脚本会自动将内容填充进你的剪贴板，你就可以打开微信群直接粘贴发送了！同时还会在脚本目录下创建word文档，将生成的通告追加进去，便于当天护网结束后可以直接提交日报！<br>

* 成功类：按事件名称归类，罗列最近告警时间、攻击者IP、受害IP<br>
* 爆破类：按事件名称归类，罗列汇总的攻击、受害IP，增加次数<br>
* 远程工具：按事件名称归类，只统计受害IP<br>
* 外打内：事件名称统一输出为"存在扫描行为"，只统计攻击者IP<br>
* 内打内：事件名称统一输出为"内网IP攻击行为"，罗列攻击者IP、受害IP<br>
