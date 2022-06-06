#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2022/6/6 20:10
# @Author  : lianghongwei
# @File    : xunjian.py
# @Software: PyCharm
# @Description :

import openpyxl, pandas

class BackupConfig(object):
	def __init__(self):
		"""初始参数"""
		self.device_file = '设备信息表.xlsx'

	def load_excel(self):
		"""
		加载excel文件
		:return:
		"""
		try:
			wb = openpyxl.load_workbook(self.device_file)
			return wb
		except FileNotFoundError:
			print('{}文件不存在'.format(self.device_file))

	def get_device_info(self):
		"""
		获取设备基本信息：
		:return:
		"""
		try:
			# 方法1：openpyxl
			wb = self.load_excel()
			ws1 = wb[wb.sheetnames[0]]
			# 通过参数min_row、max_col限制区域
			for row in ws1.iter_rows(min_row=2, max_col=9):
				if str(row[1].value).strip() == '#':
					# 跳过注释行
					continue
				info_dict = {'ip': row[2].value,
				             'protocol': row[3].value,
				             'port': row[4].value,
				             'username': row[5].value,
				             'password': row[6].value,
				             'secret': row[7].value,
				             'device_type': row[8].value,
				             # 'cmd_list': self.get_cmd_info(wb[row[8].value.strip().lower()]),
				             }
				yield info_dict
				# yield info_dict
				# 方法2:pandas
				# names = ['comment', 'ip', 'protocol', 'port', 'username', 'password', 'secret', 'device_type']
				# df = pandas.read_excel(self.device_file, usecols='B:I', names=names, keep_default_na=False )
				# data = df.to_dict(orient='records')
				# for row in data:
				# 	row['cmd_list'] = self.get_cmd_info(row['device_type'])
				# 	yield  row
		except Exception as e:
			print('Error:', e)
		finally:
			wb.close()
	def get_cmd_info(self):
		pass
	def run_cmd(self):
		pass
	def connect_test(self):
		pass
	def connect(self):
		pass
if __name__=='__main__':
	# 执行主程序
	go = BackupConfig().get_device_info()
	print(go.__next__())
	print(go.__next__())
	print(go.__next__())
	print(go.__next__())
	# BackupConfig().connect()
