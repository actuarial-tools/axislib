#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
@project: axislib
@file: MoSes.py
@date: 8/18/2019
"""

"""
author: modeldevtools
purpose: Convert MoSes Model To SQL
"""
from __future__ import print_function

import csv
import glob
import os
import sqlite3
import sys

import pandas as pd
from dbfread import DBF

pd.set_option('display.expand_frame_repr', False)
csv.field_size_limit(2147483647)
maxInt = sys.maxsize
decrement = True

while decrement:
	# decrease the maxInt value by factor 10
	# as long as the OverflowError occurs.

	decrement = False
	try:
		csv.field_size_limit(maxInt)
	except OverflowError:
		maxInt = int(maxInt / 10)
		decrement = True


def encode_decode(x, enc='cp850', dec='utf-8'):
	"""
	DBF returns a unicode string encoded as args.input_encoding.
	We convert that back into bytes and then decode as args.output_encoding.
	"""
	if not isinstance(x, str):
		# DBF converts columns into non-str like int, float
		x = str(x)

	x = x.encode(enc).decode(dec)
	return x


def convertDBFtoCSV(input_file_path, outdir):
	print("Converting %s to csv" % input_file_path)
	dbf_basename = os.path.basename(input_file_path)
	csv_basename = dbf_basename[:-4] + ".csv"
	output_file_path = os.path.join(outdir, csv_basename)

	with open(output_file_path, 'w') as csvfile:
		input_reader = DBF(input_file_path, ignore_missing_memofile=True)
		output_writer = csv.DictWriter(csvfile, delimiter=',',
		                               lineterminator='\n',
		                               fieldnames=[x for x in
		                                           input_reader.field_names])
		output_writer.writeheader()

		for record in input_reader:
			row = {k: v for k, v in record.items()}
			output_writer.writerow(row)


def BatchCSVtoSQLite(db, csvDir):
	cnxn = sqlite3.connect(os.path.join(csvDir, "model.db"))
	cnxn.text_factory = str
	cursor = cnxn.cursor()

	for csvfile in glob.glob(os.path.join(csvDir, "*.csv")):
		try:
			tablename = os.path.splitext(os.path.basename(csvfile))[0]

			with open(csvfile, "r") as f:
				reader = csv.reader(f)

				header = True
				for row in reader:
					if header:
						header = False
						sqlCmd = "DROP TABLE IF EXISTS %s" % tablename
						cursor.execute(sqlCmd)
						sqlCmd = "CREATE TABLE %s (%s)" % (tablename,
						                                   ", ".join([
							                                   "%s text" %
							                                   column
							                                   for
							                                   column
							                                   in
							                                   row]))
						cursor.execute(sqlCmd)

						for column in row:
							if column.lower().endswith("_id"):
								index = "%s__%s" % (tablename, column)
								sqlCmd = "CREATE INDEX %s on %s (%s)" % (
									index, tablename, column)
								cursor.execute(sqlCmd)

						sqlInsert = "INSERT INTO %s VALUES (%s)" % (
							tablename, ", ".join(["?" for column in row]))
						rowLength = len(row)
					else:
						# skip lines that don'tursor have the right number of
						# columns
						if len(row) == rowLength:
							cursor.execute(sqlInsert, row)
				cnxn.commit()
		except:
			print("OverflowError:", csvfile)
	cursor.close()
	cnxn.close()


if __name__ == "__main__":
	path_moses_app = r"C:\Users\usrname\Desktop\project\client\MoSes\MoSes " \
	                 r"Model\ESGModel"
	path_project = r"/data"

	filepaths = [os.path.join(path_moses_app, f) for f
	             in os.listdir(path_moses_app)
	             if f.lower().endswith(".dbf")]

	for filepath in filepaths:
		try:
			convertDBFtoCSV(input_file_path=filepath,
			                outdir=os.path.join(path_project, "database"))
		except:
			print("Unicode Decode Error:", filepath)
	dbName = os.path.join(path_project, "database\model.db")
	csvDirectory = os.path.join(path_project, "database")
	BatchCSVtoSQLite(db=dbName, csvDir=csvDirectory)
