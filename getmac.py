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
from jlib.ipmac import sort_macinfo, format_mac

#----------------------------------------------
#   macinfo 关联数组
#   macinfo["ip"];		//ip地址
#   macinfo["mac"];		//mac地址
#   macinfo["diskid"];	//硬盘序列号
#   macinfo["pcname"];	//计算机名
#   macinfo["owner"];	//使用者
#   macinfo["username"];//使用者用户名
#   macinfo["room"];	//房间号
#   macinfo["xcode"];		//资产编号
#   macinfo["comment"];	//其它信息
#----------------------------------------------

#----------------------------------------------
# 输出计算机信息
# print("ip,mac,diskid,pcname,owner,username,room,xcode,comment")
#----------------------------------------------
def macinfo_to_text(macinfo):
	#硬盘系列号
	if "diskid" in macinfo:
		diskid = macinfo["diskid"]
	else:
		diskid = "DISKID"
	#用户名
	if "username" in macinfo:
		username = macinfo["username"]
	else:
		username = ""
	#资产编号
	if "xcode" in macinfo:
		xcode = macinfo["xcode"]
	else:
		xcode = ""
	#输出
	text = '"{}","{}","{}","{}","{}","{}","{}","{}","{}"\n'.format(macinfo['ip'], 
			format_mac(macinfo['mac']), diskid, macinfo['pcname'],
			macinfo['owner'], username, macinfo['room'], xcode, macinfo['comment'])
	return text


#输出计算机信息列表
def format_macinfo_list(macinfo_list):
	text = "ip,mac,diskid,pcname,owner,username,room,xcode,comment\n"
	for macinfo in macinfo_list:
		text += macinfo_to_text(macinfo)
	return text


#----------------------------------------------
# 从数据库中获取指定人的用户名
#----------------------------------------------
def get_username(name, db):
	dbconn = pymysql.connect(db["host"], db["username"], db['password'], db['database'], charset=db['charset'], use_unicode=True)
	cursor = dbconn.cursor(pymysql.cursors.DictCursor)
	sql = "SELECT * from cms_users WHERE nickname=%s LIMIT 1"
	count = cursor.execute(sql, (name, ))
	if count == 1:
		user = cursor.fetchone()
		return user["username"]
	return ""


#----------------------------------------------
# 从数据库中获取MAC地址对应的计算机信息
#----------------------------------------------
def get_pc_info(mac, db):
	dbconn = pymysql.connect(db["host"], db["username"], db['password'], db['database'], charset=db['charset'], use_unicode=True)
	cursor = dbconn.cursor(pymysql.cursors.DictCursor)
	sql = "SELECT * from cms_macinfo WHERE mac=%s LIMIT 1"
	count = cursor.execute(sql, (mac, ))
	if count == 1:
		macinfo = cursor.fetchone()
		return macinfo
	return None

#----------------------------------------------
# 从dhcp导出文件中获取计算机信息
#   Windows Server dhcp服务导出格式
#----------------------------------------------
def get_from_dhcp(filename, db):
	maclist = []
	with open(filename, "r", encoding="gbk") as sf:
		sf.readline()  #读取第一行并丢弃
		#处理每一行文本
		while True:
			line = sf.readline().strip('\r\n')	#去掉行尾结束符
			if len(line) == 0:			#空行
				break;
			#分割字符串
			lines = line.split("\t")	#分割字段
			if len(lines) < 5:
				continue
			#保存数据
			row = {}
			row['ip']  = lines[0]
			row['mac'] = lines[4]
			row['diskid'] = ''
			row['pcname'] = lines[1]
			#描述信息
			descs = lines[5].split("|")
			row['owner'] = descs[0]
			if len(descs) >= 2:
				row['room'] = descs[1]
			else:
				row['room'] = ""
			#连接的交换机端口
			if len(descs) >= 3:
				row['comment'] = descs[2]
				row['port'] = descs[2]
			else:
				row['comment'] = ""
				row['port'] = ""
			#获取拥有者的用户名
			row["username"] = get_username(row['owner'], db)
			#获取硬盘ID和资产编号
			pc_info = get_pc_info(format_mac(row['mac'], "-", True), db)
			if pc_info != None:
				row["diskid"] = pc_info["diskid"]
				row["xcode"] = pc_info["xcode"]
			maclist.append(row)
	return maclist


##########################################
## 主模块
## export PYTHONIOENCODING=utf8
##########################################
if __name__ == '__main__':
	from jlib.vardump import var_dump
	import argparse
	parser = argparse.ArgumentParser(description="从Windows DHCP 及数据库获取计算机信息")
	#配置文件（包含数据库信息）
	parser.add_argument('-c', '--config', help="配置文件, json 格式")
	#输出方式
	parser.add_argument('-a', '--action', default="dump", help="指定要执行的动作：dump（默认）, print（打印输出）, export(导出为csv文件)")
	#输出文件编码
	parser.add_argument('-e', '--encoding', default="utf-8", help="导出文件的字符编码，默认 utf-8")
	#输出文件文件名
	parser.add_argument('-o', '--out', default="export.csv", help="导出文件的文件名")
	#输入文件名
	parser.add_argument('filename', type=str, help="输入的dhcp数据文件名，txt格式")
	#解析参数
	args = parser.parse_args()
	#数据库配置
	db = {}
	#------------------------------------------------
	#从配置文件获取数据库配置
	#------------------------------------------------
	#获取配置文件路径
	if not args.config:
		args.config = sys.path[0] + "/config/config.json"
	#读取配置信息
	try:
		with open(args.config, 'rb') as f:
			config_data = json.load(f)
			db = config_data.get("db", {})
	except Exception as err:
		log_error("从文件“{}”读取数据库配置发生错误 : {}.".format(args.config, err))
		exit(1)
	#读取数据
	macinfo = get_from_dhcp(args.filename, db)
	
	#执行动作并输出
	if args.action == 'print':
		macinfo_text = format_macinfo_list(macinfo)
		print(macinfo_text)
	elif args.action == 'dump':
		print(json.dumps(macinfo, indent=4, sort_keys=True, ensure_ascii=False, default=lambda obj: obj.__dict__))
	elif args.action == 'export':
		macinfo_text = format_macinfo_list(macinfo)
		with open(args.out, "wb") as f:
			f.write(macinfo_text.encode(args.encoding))
		print("Successful! export data to file {} .".format(args.out))
	else:
		print("\033[31m错误：未定义的动作！\033[0m错误：")
		exit(-1)
