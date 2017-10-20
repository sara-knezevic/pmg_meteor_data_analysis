# ovo kao final treba da bude
# sve ovde da se smesti

import os, shutil, sys
from openpyxl import *
import numpy as np
import time, datetime

def create_excel(directory, filename):
	wb = Workbook()
	ws = wb.active
	wb.save(os.path.join(directory, filename))

def convert_to_seconds(date):
	return date.tm_sec + date.tm_min * 60 + date.tm_hour * 3600 + date.tm_mday * 24 * 3600

def store_observers(sheet):
	global observers
	observers = set()

	R = 1

	for row in sheet.iter_rows(min_row = 1):
		observers.add(sheet.cell(row = R, column = 7).value)
		R += 1

	observers = list(observers)
	return observers

def cell(worksheet, r, c):
	return worksheet.cell(row = r, column = c).value

def timings(posmatraci, vremena):
	time_format = "%Y/%m/%d %H:%M:%S"
	day, night = date.split("-")

	time_of_observers = {}

	iter_row = 1
	for row in vremena.iter_rows(min_row = 1):
		o = vremena.cell(row = iter_row, column = 1).value

		tx = ""
		ty = ""

		# starts
		if (int(cell(vremena, iter_row, 2)[:2]) > 17):
			tx = "2017/08/" + day + " " + cell(vremena, iter_row, 2)
			# ends
			if (int(cell(vremena, iter_row, 3)[:2]) > 17):
				ty = "2017/08/" + day + " " + cell(vremena, iter_row, 3)
			else:
				ty = "2017/08/" + night + " " + cell(vremena, iter_row, 3)

		else:
			tx = "2017/08/" + night + " " + cell(vremena, iter_row, 2)
			# ends
			if (int(cell(vremena, iter_row, 3)[:2]) > 17):
				ty = "2017/08/" + day + " " + cell(vremena, iter_row, 3)
			else:
				ty = "2017/08/" + night + " " + cell(vremena, iter_row, 3)

		t_start = time.strptime(tx[:-1], time_format)
		t_end = time.strptime(ty[:-1], time_format)

		time_of_observers.setdefault(o, []).append(t_start)
		time_of_observers.setdefault(o, []).append(t_end)

		iter_row += 1

	return time_of_observers

def compare_meteors(sheet):
	meteors = {}
	time_format = "%Y-%m-%d %H:%M:%S"
	r = 1

	meteors.setdefault("Istok", [])
	meteors.setdefault("Zapad", [])
	meteors.setdefault("Sever", [])

	for row in sheet.iter_rows(min_row = 1):
		group = cell(sheet, r, 5)
		t = str(cell(sheet, r, 1))[:-8] + str(cell(sheet, r, 3))

		meteors.setdefault(group, []).append(t)

		r += 1

	compared_times = {}

	# to be compared : 

	return 0

posmatranja = load_workbook("/home/sarasdfg/Petnica/Kodovi/posmatranja.xlsx")
vremena = load_workbook("/home/sarasdfg/Petnica/Kodovi/timings.xlsx")
dates = ["8-9", "9-10", "10-11", "11-12", "14-15", "15-16", "16-17", "17-18"]

date = dates[0]
p_ws = posmatranja[date]
t_ws = vremena[date]

observers = store_observers(p_ws)
time_of_observers = timings(p_ws, t_ws)

compare_meteors(p_ws)
