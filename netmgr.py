#!/usr/bin/env python3
# -*- coding:utf-8 -*-
##############################################
#         网络管理模块（DHCP处理）
#       陈逸少（jmchxy@gmail.com）
##############################################
import sys
import os
import re
import json
import pymysql
import pandas as pd
from jlib.ipmac import sort_macinfo, format_mac

#==========================================================
# 辅助函数
#==========================================================

#==========================================================
# 功能函数
#==========================================================

#----------------------------------------------
#   数据表字段
#   macinfo["ip"];		//ip地址
#   macinfo["mac"];		//mac地址
#   macinfo["diskid"];	//硬盘序列号
#   macinfo["pcname"];	//计算机名
#   macinfo["owner"];	//使用者
#	macinfo["username"];//用户名
#   macinfo["room"];	//房间号
#   macinfo["comment"];	//其它信息，交换机端口号
#----------------------------------------------

#----------------------------------------------
# 输出数据信息
#----------------------------------------------
def print_macinfo(df):
	print(df.to_csv(index=False))

#----------------------------------------------
# 输出上网行为管理用户列表
#----------------------------------------------
def export_userlist(df):
	result = "Username,Showname,Binding IP Address,Binding MAC Address\n"
	for i in range(0, len(df)):
		if str(df['username'][i]) == "nan" or str(df['username'][i]) == "":
			pass
		else:
			result += "{},{},{},\r\n".format(df['username'][i], df['owner'][i], df['ip'][i])
	#返回字符串
	return result

#----------------------------------------------
# 输出 dhcp 绑定命令(Linux dhcpd.conf 格式)
#----------------------------------------------
def dhcpd_bind(df):
	result = "#=========================\n# IP-MAC 绑定 \n#=========================\n"
	for i in range(0, len(df)):
		#如果mac地址为空，不生成绑定信息
		if str(df['mac'][i]) == "" or str(df['mac'][i]) == "nan":
			continue
		#描述信息
		pcdesc = "{}|{}|{}".format(df['owner'][i], df['room'][i], df['comment'][i])
		#MAC地址格式化
		pcmac = format_mac(df['mac'][i], ":", False)
		#计算机名称
		if str(df['pcname'][i]) == "":
			pcname = "PC" + pcmac.replace(':', '')
		else:
			pcname1 = df['pcname'][i].split(".")
			pcname = pcname1[0].replace('-', '')
		#生成绑定配置数据
		result += "\n#{}\nhost {} {}\n".format(pcdesc, pcname, "{")
		result += "\thardware ethernet {};\n\tfixed-address {};\n{}\n".format(pcmac, df['ip'][i], "}")
	#返回字符串
	return result

#----------------------------------------------
# 输出 dhcp 绑定命令(Windows Server 格式)
#   netsh exec bind-file
#----------------------------------------------
def windows_bind(df, dnsServer):
	result = ""
	for i in range(0, len(df)):
		#如果mac地址为空，不生成绑定信息
		if str(df['mac'][i]) == "" or str(df['mac'][i]) == "nan":
			continue
		#描述信息
		pcdesc = "{}|{}|{}".format(df['owner'][i], df['room'][i], df['comment'][i])
		#MAC地址格式化
		pcmac = format_mac(df['mac'][i], "", False)
		#计算机名称
		if str(df['pcname'][i]) == "":
			pcname = "PC" + pcmac
		else:
			pcname = df['pcname'][i]
		#计算子网
		ips = df['ip'][i].split('.')
		ips[3] = '0'
		scope = ".".join(ips)
		#生成绑定命令行
		result += "Dhcp Server {} Scope {} Add reservedip {} {} \"{}\" \"{}\" \"BOTH\"\r\n".format(
			dnsServer, scope, df['ip'][i], pcmac, pcname, pcdesc)
	#返回字符串
	return result

#----------------------------------------------
# 输出 Windows Server dhcp 允许列表
#----------------------------------------------
def windows_macfitler(df, dnsServer):
	result = ""
	for i in range(0, len(df)):
		#如果mac地址为空，不生成MAC列表
		if str(df['mac'][i]) == "" or str(df['mac'][i]) == "nan":
			continue
		#生成允许列表
		pcdesc = "{}_{}_{}".format(df['owner'][i], df['room'][i], df['comment'][i])
		pcmac = df['mac'][i].replace('-', '').lower()
		if df["owner"][i][0:2] != '保留':
			result += "Dhcp Server {} v4 Add Filter Allow {} \"{}\"\r\n".format(dnsServer, pcmac, pcdesc)
		else:
			pass
	return result

#----------------------------------------------
# 输出 Windows Server dhcp maclist 列表(Windows2008)
#----------------------------------------------
def export_maclist(df):
	result = ""
	for i in range(0, len(df)):
		#如果mac地址为空，不生成MAC列表
		if str(df['mac'][i]) == "" or str(df['mac'][i]) == "nan":
			continue
		#生成允许列表
		pcdesc = "{}_{}_{}".format(df['owner'][i], df['room'][i], df['comment'][i])
		pcmac = df['mac'][i].replace('-', '').lower()
		if df["owner"][i][0:2] != '保留':
			result += "{}    # {}\r\n".format(pcmac, pcdesc)
		else:
			pass
	return result

##########################################
## 主模块
## export PYTHONIOENCODING=utf8
##########################################
if __name__ == '__main__':
	import argparse
	parser = argparse.ArgumentParser(description="网络管理维护脚本，源文件为Excel文件", formatter_class=argparse.RawTextHelpFormatter)
	#输出：输出方法和文件名
	parser.add_argument('-a', '--action', default="print", help='''指定要执行的动作：
    print（输出读取的信息，默认）, 
    bind（生成Linux dhcpd 的 ip-mac 绑定配置器文件）,
    winbind（生成 Windows Server dhcp 的绑定命令文件）,
    macfilter(生成 Windows Server dhcp 的筛选命令文件),
    maclist（生成 Windows Server dhcp 的MAC过滤文件）,
    user（上网行为管理列表）''')
	parser.add_argument('-d', '--dnsserver', default="10.99.2.103", help="指定 Windows 绑定的 DNS Server 的IP地址")
	parser.add_argument('-o', '--out', default="dhcpbind.conf", help="输出绑定命令的文件名")
	#输入，文件名和表名
	parser.add_argument('filename', type=str, help="输入的Excel文件名")
	parser.add_argument('sheet', type=str, help="输入的Excel表名")
	#解析参数
	args = parser.parse_args()
	#输入：数据来源类型，数据库或文件
	#输入
	try:
		df = pd.read_excel(args.filename, args.sheet, dtype={'username':str, 'diskid':str})
		#执行动作并输出
		if args.action == 'print':
			print_macinfo(df)
		elif args.action == 'bind':	#生成Linux绑定格式
			bindtxt = dhcpd_bind(df)
			outfile = args.sheet + ".conf"
			with open(outfile, "wb") as f:
				f.write(bindtxt.encode('utf-8'))
			print("Successful! copy file {} to Linux dhcp server /etc/dhcp/bind/".format(outfile))
		elif args.action == 'winbind':	#生成Windows绑定命令
			bindtxt = windows_bind(df, args.dnsserver)
			outfile = args.sheet + "_bind.txt"
			with open(outfile, "wb") as f:
				f.write(bindtxt.encode('gbk'))
			print("Successful! Execute “netsh exec {}” on Windows DNS Server to bind.".format(outfile))
		elif args.action == 'macfilter':	#生成Windows筛选命令
			bindtxt = windows_macfitler(df, args.dnsserver)
			outfile = args.sheet + "_filter.txt"
			with open(outfile, "wb") as f:
				f.write(bindtxt.encode('gbk'))
			print("Successful! Execute “netsh exec {}” on Windows DNS Server to bind.".format(outfile))
		elif args.action == 'maclist':	#输出DHCP MacList
			listtext = "MAC_ACTION = {ALLOW}\r\n#允许的Mac地址列表\r\n" + export_maclist(df)
			with open("MACList.txt", "wb") as f:
				f.write(listtext.encode('gbk'))
		elif args.action == 'user':	#输出上网行为管理列表
			listtext = export_userlist(df)
			with open("organizedFrame.csv", "wb") as f:
				f.write(listtext.encode('gbk'))
		else:
			print("\033[31m错误：未定义的动作！\033[0m")
			exit(-1)
	except Exception as err:
		print("\033[31m{}\033[0m".format(err))
		exit(-1)
