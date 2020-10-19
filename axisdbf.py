#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
@project: axislib
@file: AxisDataReader.py
@author: modeldevtools
@date: 10/18/2020
"""

import os
from time import sleep

import pandas as pd
import win32com.client
from colorama import Fore
from tqdm import tqdm

# Options
pd.set_option('display.expand_frame_repr', False)

# Constants
adStateOpen = 1

color_bars = [
	Fore.BLACK,
	Fore.RED,
	Fore.GREEN,
	Fore.YELLOW,
	Fore.BLUE,
	Fore.MAGENTA,
	Fore.CYAN,
	Fore.WHITE
	]


class Options:
	default_options = {
		'filename': None,
		'table': None,
		'sql': None,
		'connected': 0,
		'debug': False,
		'verbose': False
		}

	def __init__(self, **kwargs):
		self.options = dict(Options.default_options)
		self.options.update(kwargs)

	def __getitem__(self, key):
		return self.options[key]


class AxisDbf(object):

	def __init__(self, filename=None, table=None, sql=None, verbose=False):
		for i in tqdm(range(100), desc='Initializing AxisDbf class:'):
			sleep(0.01)
		self.filename = filename
		self.table = table
		self.sql = sql
		self.verbose = verbose
		self.conn = None
		self.cmd = None
		self.connected = 0

		if self.filename is not None:
			self.table = os.path.basename(filename).split('.')[0]

	def print_verbose(self, message, params):
		"""print the message only if verbose is enabled"""
		if self.verbose:
			print(message.format(params))

	def connect(self, f=None):
		if f is not None:
			self.filename = f
		self.table = table = os.path.basename(self.filename).split('.')[0]

		if self.verbose is True:
			pbar = tqdm(len(range(100)), desc='Connecting to Visual FoxPro database: {}'.format(self.filename))
			for i in range(100):
				sleep(.02)
				pbar.update(len(range(100)))
			pbar.close()
		try:
			self.conn = conn = win32com.client.Dispatch('ADODB.Connection')
			self.dsn = dsn = 'Provider=VFPOLEDB.1;Data Source=%s' % \
			                 self.filename
			self.conn.Open(dsn)
			pbar.update(100)
			pbar.close()
			if self.conn.State == adStateOpen:
				self.connected = 1
				self.print_verbose(message='\nConnection Successful!{0}',
				                   params='')
			else:
				self.print_verbose(message='Connection failed. Verify '
				                           'connection string parameters and '
				                           'try again.{0}',
				                   params='')

		except:
			self.print_verbose(message='Unable to open database connection. '
			                           'Please verify that {0} is not in '
			                           'use.', params=self.filename)

	def disconnect(self):
		if self.verbose is True:
			pbar = tqdm(len(range(100)), desc='Disconnecting to Visual FoxPro '
			                                  'database: {}'.format(self.filename))
			for i in range(100):
				sleep(.02)
				pbar.update(len(range(100)))
			pbar.close()

		if self.connected == 1:
			self.conn.Close()

	def execute(self, sql):
		self.recordset = None
		self.records = []
		self.rowid = 0
		self.fieldnames = []

		if self.table is None:
			self.table = os.path.basename(self.filename).split('.')[0]

		if sql is not None:
			self.sql = sql
		else:
			self.sql = "SELECT * FROM '" + self.table + "';"

		if self.verbose is True:
			pbar = tqdm(len(range(100)), desc='Executing Visual FoxPro '
			                                  'command: {}'.format(self.sql))
			for i in range(100):
				sleep(.02)
				pbar.update(len(range(100)))
			pbar.close()

		self.cmd = win32com.client.Dispatch('ADODB.Command')
		self.cmd.ActiveConnection = self.conn
		self.cmd.CommandText = self.sql

		self.recordset = win32com.client.Dispatch('ADODB.Recordset')
		self.recordset = self.cmd.Execute()[0]
		self.rowcount = self.cmd.Execute()[1]

		for x in range(self.recordset.Fields.Count):
			self.fieldnames.append(self.recordset.Fields.Item(x).Name)

		values_list = []
		try:
			data = self.recordset.GetRows()
			self.rowcount = len(data[0])
			pbar = tqdm(total=len(range(0, self.rowcount)),
			            desc='Enumerating Recordset')
			for y in range(0, self.rowcount):
				for x in data:
					v = x[y].strip() if isinstance(x[y], str) else x[y]
					values_list.append(v)
				self.records.append(tuple(values_list))
				values_list = []
				pbar.update(1)
			pbar.close()
			self.records = tuple(self.records)
		except UnboundLocalError:
			pass
		except:
			pass

	def fetch_all(self):
		lst = []
		try:
			pbar = tqdm(len(range(0,self.rowcount)), desc='Fetching Records')
			for x in self.records:
				lst.append(x)
				pbar.update(1)
			pbar.close()
		except IndexError:
			pass
		return lst

	def fetch_one(self):
		i = self.rowid
		j = i + 1
		self.rowid = j
		try:
			return tuple(self.records[i])
		except IndexError:
			pass

	def fetch_fieldnames(self):
		field_list = []
		pbar = tqdm(len(range(self.recordset.Fields.Count)),
		            desc='Collecting Field Names')
		for x in range(self.recordset.Fields.Count):
			field_list.append(self.recordset.Fields.Item(x).Name)
			pbar.update(1)
		pbar.close()
		return field_list

	def datatable(self, display=False):
		pbar = tqdm(len(range(0, self.rowcount)), desc='Creating Datatable')
		header = self.fetch_fieldnames()
		data = self.fetch_all()
		self.datatable = dt = pd.DataFrame(data, columns=header)
		pbar.update(len(0,self.rowcount))
		pbar.close()
		return dt


def axisdbf_test():
	filename = "C:\EL\WA\RPA\DATASETS\sample\AxisLib_ Tables.DBF"
	ax = AxisDbf(verbose=True)
	ax.connect(f=filename)
	ax.execute(sql="Select * from 'AxisLib_ Tables'")
	ax.fetch_fieldnames()
	ax.fetch_all()
	ax.fetch_one()
	ax.datatable()


def axisdbf_example():
	conn = win32com.client.Dispatch('ADODB.Connection')
	db = "C:\EL\WA\RPA\DATASETS\sample\AxisLib_ Tables.DBF"
	dsn = 'Provider=VFPOLEDB.1;Data Source=%s' % db
	conn.Open(dsn)
	cmd = win32com.client.Dispatch('ADODB.Command')
	cmd.ActiveConnection = conn
	cmd.CommandText = "Select * from 'AxisLib_ Tables'"

	rs, rowcount = cmd.Execute()  # This returns a tuple: (<RecordSet>,
	# number_of_records)

	for x in range(rs.Fields.Count):
		print(rs.Fields.item(x).Name)

	while rowcount:
		for x in range(rs.Fields.Count):
			print('%s --> %s' % (
				rs.Fields.item(x).Name, rs.Fields.item(x).Value))
		rs.MoveNext()  # <- Extra indent
		rowcount = rowcount - 1


if __name__ == '__main__':
	axisdbf_test()
# axisdbf_example()
