from argparse import ArgumentParser
from collections import namedtuple
from datetime import timedelta, date
import xlrd

Person = namedtuple("Person", ["Name", "Room"])
Date = namedtuple("Date", ["Day", "Month", "Year"])
Interval = namedtuple("Interval", ["Start", "End"])
Time = namedtuple("Time", ["Hour", "Minute"])

def createHourSpendingPlot(times):
	try:
		import matplotlib.pyplot as plt
	except:
		print ("matplotlib not supported.")
		return
	minHour = 23
	maxHour = 0
	for key in times.keys():
		if key.Hour == -1:
			continue
		minHour = min(minHour, key.Hour)
		maxHour = max(maxHour, key.Hour)
	
	hourToMoney = dict()
	for hour in range(minHour, maxHour+1):
		hourToMoney.setdefault(hour, 0.0)
	
	for key in times.keys():
		if key.Hour == -1:
			continue
		hourToMoney[key.Hour] = hourToMoney[key.Hour] + times[key]
	
	xs = sorted(hourToMoney.keys())
	ys = []
	for x in xs:
		ys.append(hourToMoney[x])
	width = 1.0
	plt.bar(xs, ys, width, align='center')
	plt.xticks(xs, xs)
	plt.xlabel('hour')
	plt.ylabel('dkk')
	plt.show()

def createDaySpendingPlot(dates):
	try:
		import matplotlib.pyplot as plt
	except:
		print ("Matplotlib not supported.")
		return
	week = {"Monday": 0, "Tuesday": 1, "Wednesday": 2, "Thursday": 3, "Friday": 4, "Saturday": 5, "Sunday": 6}
			
	dayToMoney = dict()
	for key in dates.keys():
		dayToMoney[key.strftime('%A')] = dayToMoney.setdefault(key.strftime('%A'), 0.0) + dates[key]
	
	real = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
	labels = sorted(dayToMoney, key=real.index)
	ys = []
	xs = []
	for label in labels:
		xs.append(week[label])
		ys.append(dayToMoney[label])

	width = 1.0
	plt.bar(xs, ys, width, align='center')
	plt.xticks(xs, labels, rotation=15)
	plt.ylabel('dkk')
	
	plt.show()

def createMonthSpendingPlot(dates):
	try:
		import matplotlib.pyplot as plt
	except:
		print ("Matplotlib not supported.")
		return

	ys = []
	monthSpend = dict()
	for key in dates.keys():
		monthSpend[key.month] = monthSpend.setdefault(key.month, 0.0) + dates[key]
	xs = sorted(monthSpend.keys())
	labels = []
	for x in xs:
		ys.append(monthSpend[x])
		labels.append(date(2016, x, 1).strftime('%B'))

	width = 1.0
	plt.bar(xs, ys, width, align='center')
	plt.xticks(xs, labels)
	plt.ylabel('dkk')

	plt.show()
	

def daterange(startDate, endDate):
	for n in range(int((endDate-startDate).days)):
		yield startDate + timedelta(n)

def checkInt(value):
	try:
		return int(value)
	except ValueError:
		print ("%s is not integer like expected." % (value))
	return -1

def checkFloat(value):
	try:
		return float(value)
	except ValueError:
		print ("%s is not float like expected." % (value))
	return 0.0

def loadWorksheet(workbook, sheet):
	return workbook.sheet_by_name(sheet)

def loadWorkbook(location):
	return xlrd.open_workbook(location)

def findColumn(sheet, columnName, rowNum=0):
	for column in range(sheet.ncols):
		if (sheet.cell(rowNum, column).value.lower() == columnName.lower()):
			return column
	return -1

def findRow(sheet, rowName, columnNum=0):
	for row in range(sheet.nrows):
		if (sheet.cell(row, columnNum).value.lower() == rowName.lower()):
			return row
	return -1

def readReceipts(receiptSheet):
	columns = {'name': None, 'room': None, 'amount': None, 'day': None, 'month': None, 'year': None, 'hours': None, 'minutes': None}
	for key in columns.keys():
		column = findColumn(receiptSheet, key)
		if column == -1:
			print ("Could not find column '%s' in sheet '%s'." % (key, receiptSheet.name))
			return
		else:
			columns[key] = column

	persons = dict()
	dates = dict()
	times = dict()

	for row in range(1, receiptSheet.nrows):
		name = receiptSheet.cell(row, columns['name']).value.strip()
		room = checkInt(receiptSheet.cell(row, columns['room']).value)
		amount = checkFloat(receiptSheet.cell(row, columns['amount']).value)
		day = checkInt(receiptSheet.cell(row, columns['day']).value)
		month = checkInt(receiptSheet.cell(row, columns['month']).value)
		year = checkInt(receiptSheet.cell(row, columns['year']).value)
		hour = checkInt(receiptSheet.cell(row, columns['hours']).value)
		minute = checkInt(receiptSheet.cell(row, columns['minutes']).value)
		person = Person(Name=name, Room=room)
		if month != -1 and day != -1 and year != -1:
			date_ = date(year, month, day)
		else:
			print ("Invalid date: name=%s, room=%s in sheet=%s" % (name, room, receiptSheet.name))
		time = Time(Hour=hour, Minute=minute)

		persons[person] = persons.setdefault(person, 0.0) + amount
		dates[date_] = dates.setdefault(date_, 0.0) + amount
		times[time] = times.setdefault(time, 0.0) + amount

	return persons, dates, times

def readPeople(peopleSheet):
	columns = {"name": None, "room": None, "start day": None, "start month": None, "start year": None, "end day": None, "end month": None, "end year": None}
	for key in columns.keys():
		column = findColumn(peopleSheet, key, 0)
		if column == -1:
			print ("Could not find column '%s' in sheet '%s'." % (key, peopleSheet.name))
		else:
			columns[key] = column	
	periodRow = findRow(peopleSheet, "period start", 0)
	periodStartDay = checkInt(peopleSheet.cell(periodRow, columns['start day']).value)
	periodStartMonth = checkInt(peopleSheet.cell(periodRow, columns['start month']).value)
	periodStartYear = checkInt(peopleSheet.cell(periodRow, columns['start year']).value)
	periodEndDay = checkInt(peopleSheet.cell(periodRow, columns['end day']).value)
	periodEndMonth = checkInt(peopleSheet.cell(periodRow, columns['end month']).value)
	periodEndYear = checkInt(peopleSheet.cell(periodRow, columns['end year']).value)

	periodStartDate = date(periodStartYear, periodStartMonth, periodStartDay)
	periodEndDate = date(periodEndYear, periodEndMonth, periodEndDay)

	fullInterval = Interval(Start=periodStartDate, End=periodEndDate)

	periods = dict()

	for row in range(1, peopleSheet.nrows):
		if row == periodRow:
			continue
		name = peopleSheet.cell(row, columns['name']).value.strip()
		room = checkInt(peopleSheet.cell(row, columns['room']).value)
		startDay = checkInt(peopleSheet.cell(row, columns['start day']).value)
		startMonth = checkInt(peopleSheet.cell(row, columns['start month']).value)
		startYear = checkInt(peopleSheet.cell(row, columns['start year']).value)
		endDay = checkInt(peopleSheet.cell(row, columns['end day']).value)
		endMonth = checkInt(peopleSheet.cell(row, columns['end month']).value)
		endYear = checkInt(peopleSheet.cell(row, columns['end year']).value)
		

		try:
			interval = Interval(Start=date(startYear, startMonth, startDay), End=date(endYear, endMonth, endDay))
		except:
			print ("Invalid date: name=%s, room=%s in sheet=%s" % (name, room, peopleSheet.name))
		person = Person(Name=name, Room=room)

		periods.setdefault(person, interval)
		
	return periods, fullInterval 

def readRemainder(remainderSheet):
	columns = {"name": None, "room": None, "remainder": None}
	for key in columns.keys():
		column = findColumn(remainderSheet, key, 0)
		if column == -1:
			print ("Could not find column '%s' in sheet '%s'." % (key, remainderSheet.name))
			continue
		columns[key] = column
	
	remainders = dict()
	for row in range(1, remainderSheet.nrows):
		name = remainderSheet.cell(row, columns['name']).value.strip()
		room = checkInt(remainderSheet.cell(row, columns['room']).value)
		remainder = checkFloat(remainderSheet.cell(row, columns['remainder']).value)

		person = Person(Name=name, Room=room)
		remainders[person] = remainders.setdefault(person, 0.0) + remainder
	return remainders

def isBetween(lower, between, upper):
	return lower <= between <= upper

def calculateAmounts(persons, periods, fullInterval):
	totalAmount = sum([persons[key] for key in persons.keys()])
	dailyAmount = totalAmount / float((fullInterval.End - fullInterval.Start).days)

	payments = dict()

	for d in daterange(fullInterval.Start, fullInterval.End):
		amount = 0
		for person in periods.keys():
			if isBetween(periods[person].Start, d, periods[person].End):
				amount += 1  
		dayAmount = dailyAmount / float(amount)
		for person in periods.keys():
			if isBetween(periods[person].Start, d, periods[person].End):
				payments[person] = payments.setdefault(person, 0.0) + dayAmount

	return payments

def main(document, target, stats):
	workbook = loadWorkbook(document)
	if workbook.nsheets < 3:
		print ("Detected %s sheets in document. Minimum of 3 required." % (workbook.nsheets))
		return
	receiptSheet = loadWorksheet(workbook, "Receipts")
	peopleSheet = loadWorksheet(workbook, "People")
	remainderSheet = loadWorksheet(workbook, "From Last")

	persons, dates, times = readReceipts(receiptSheet)	
	periods, fullInterval = readPeople(peopleSheet)
	remainders = readRemainder(remainderSheet)

	payments = calculateAmounts(persons, periods, fullInterval)

	printPayments(payments, persons, periods, remainders, target)

	if stats:
		createMonthSpendingPlot(dates)
		createDaySpendingPlot(dates)
		createHourSpendingPlot(times)

	return 0

def printPayments(payments, persons, periods, remainders, target):
	sortedKeys = sorted(periods.keys(), key=lambda person: person.Room)
	nljust, rljust, flljust, bfljust, ppljust, tpljust = 6, 6, 11, 12, 12, 8
	for key in sortedKeys:
		remainderValue = "{0:.2f}".format(remainders.get(key, 0.0))
		personsValue = "{0:.2f}".format(persons.get(key, 0.0))
		paymentsValue = "{0:.2f}".format(payments.get(key, 0.0))
		toPay = "{0:.2f}".format(payments.get(key, 0.0) - persons.get(key, 0.0) + remainders.get(key, 0.0))

		nljust = max(nljust, len(key.Name)+2)
		rljust = max(rljust, len(str(key.Room))+2)
		flljust = max(flljust, len(remainderValue)+2)
		bfljust = max(bfljust, len(personsValue)+2)
		ppljust = max(ppljust, len(paymentsValue)+2)
		tpljust = max(tpljust, len(toPay)+2)

	with open(target, 'w+') as out:
		out.write("'To pay' is the amount to pay.\nNegative means you get money from the kitchen account.\nPositive means you need to pay the kitchen account.\n")
		out.write("In total spent = %.2f\n" % (sum(payments[key] for key in payments.keys())))
		out.write("Name".ljust(nljust) + "Room".ljust(rljust) + "From last".ljust(flljust) + "Bought for".ljust(bfljust) + "Per person".ljust(ppljust) + "To pay".ljust(tpljust) + "\n")
		for key in sortedKeys:
			remainderValue = "{0:.2f}".format(remainders.get(key, 0.0))
			personsValue = "{0:.2f}".format(persons.get(key, 0.0))
			paymentsValue = "{0:.2f}".format(payments.get(key, 0.0))
			toPay = "{0:.2f}".format(payments.get(key, 0.0) - persons.get(key, 0.0) + remainders.get(key, 0.0))
			out.write("%s%s%s%s%s%s\n" % (key.Name.ljust(nljust), str(key.Room).ljust(rljust), remainderValue.ljust(flljust), personsValue.ljust(bfljust), paymentsValue.ljust(ppljust), toPay.ljust(tpljust)))
	
def printDictionary(dictionary):
	for key in dictionary:
		print (key, "=>", dictionary[key])

"""
For this to work a excel document (tested with .xls) is needed
with 3 sheets. 'Receipts', 'People' and 'From Last'.
'Receipts' should contain the data from the receipts. Like name, room and amount at minimum. day month year hours and minutes will also work (As columns).
'People' should contain the data on who is part of the account for the period, and for what period they were. Will need name, room, start day, start month, start year, end day, end month and end year as columns. This sheet should also contain one row with name Period start, which should indicate the full extent of the period.
'From Last' should contain the data on what hasn't been paid back from last time. name, room and remainder columns are needed.

For all sheets the name of the columns should be the top most column.
"""
if __name__ == "__main__":
	argparser = ArgumentParser(description='Do the kitchen account.')
	argparser.add_argument('-f', type=str, help='Path to the spreadsheet containing the data.')
	argparser.add_argument('-t', type=str, help='Path to target file.')
	argparser.add_argument('-stats', action='store_true', help='Statistics switch.')
	args = argparser.parse_args()

	if args.f is None or args.t is None:
		print ("Missing one or more arguments.")
	else:
		main(args.f, args.t, args.stats)
