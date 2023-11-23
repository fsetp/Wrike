################################################################################
#
#
import pandas as pd
import datetime
import calendar
import helper

########################################
#
XLS_NAME	= 'monthly.xlsx'
SHEET_NAME	= 'Monthly'

########################################
#
class MonthlyLib():

	####################################
	# Constoructor
	#
	def __init__(self):
		self.schTaskTitle		= None
		self.schTaskStartDate	= None
		self.schTaskEndDate		= None

		self.mstTaskTitle		= None
		self.mstTaskPeriod		= None
		self.mstUnitPrice		= None

		self.jorTaskTitle		= None
		self.jorWorkDate		= None
		self.jorWorkerName		= None
		self.jorWorkTime		= None
		self.jorDescription		= None

	####################################
	# Initialize
	#
	def Initialize(self):
		print('Initialize Monthly Object.')
		print('')

	####################################
	#
	def SetScheduleTaskTitle(self, series):
		self.schTaskTitle = series

	####################################
	#
	def SetScheduleTaskStartDate(self, series):
		self.schTaskStartDate = series

	####################################
	#
	def SetScheduleTaskEndDate(self, series):
		self.schTaskEndDate = series

	####################################
	#
	def SetMasterTaskTitle(self, series):
		self.mstTaskTitle = series

	####################################
	#
	def SetMasterTaskPeriod(self, series):
		self.mstTaskPeriod = series

	####################################
	#
	def SetMasterUnitPrice(self, series):
		self.mstUnitPrice = series

	####################################
	#
	def SetJournalTaskTitle(self, series):
		self.jorTaskTitle = series

	####################################
	#
	def SetJournalWorkDate(self, series):
		self.jorWorkDate = series

	####################################
	#
	def SetJournalWorkerName(self, series):
		self.jorWorkerName = series

	####################################
	#
	def SetJournalWorkTime(self, series):
		self.jorWorkTime = series

	####################################
	#
	def SetJournalDescription(self, series):
		self.jorjorDescription = series

	####################################
	#
	def GetPrice(self, title):
		rows = self.mstTaskTitle.shape
		for row in range(rows[0]):
			if (title == self.mstTaskTitle[row]):
				return self.mstUnitPrice[row]
		return 0

	####################################
	#
	def GetMonthlyCount(self, year, month):
		count = 0

		first_date = datetime.datetime(year, month, 1).date()
		days = calendar.monthrange(year, month)
		last_date = datetime.datetime(year, month, days[1]).date()

		rows = self.jorTaskTitle.shape

		total = 0
		for row in range(rows[0]):
			if (self.jorWorkDate[row] >= first_date and self.jorWorkDate[row] <= last_date):
				count += 1

		return count

	####################################
	#
	def MakeMonthlyReport(self, year, month):
		first_date = datetime.datetime(year, month, 1).date()
#		print(first_date)
		days = calendar.monthrange(year, month)
		last_date = datetime.datetime(year, month, days[1]).date()
#		print(last_date)

		rows = self.jorTaskTitle.shape
#		print(type(rows[0]))

		print('date, tilte, time, price, suntotal')
		total = 0
		for row in range(rows[0]):
			if (self.jorWorkDate[row] >= first_date and self.jorWorkDate[row] <= last_date):
				time = self.jorWorkTime[row]
#				print(type(time))
				hours = time.days * 24 + time.seconds / 3600
				price = self.GetPrice(self.jorTaskTitle[row])
				subtotal = hours * price
				total += subtotal
				print(f'{self.jorWorkDate[row]}, {self.jorTaskTitle[row]}, {time}, {price}, {subtotal}')

		print(f'total : {total}')

	####################################
	#
	def OutputMonthlyReport2Excel(self):

		print('Create Monthly Report Excel File.')
		print('Making Information Object(s) ...', end = ' ')

		WorkDate = pd.Series(dtype = object)
		TaskTitle  = pd.Series(dtype = object)
		TaskTime = pd.Series(dtype = object)
		UnitPrice = pd.Series(dtype = object)
		SubTotal = pd.Series(dtype = object)

		years = (2023, 2024)
		for year in years:
			for month in range(1, 13):
				if (self.GetMonthlyCount(year, month) > 0):

					first_date = datetime.datetime(year, month, 1).date()
					days = calendar.monthrange(year, month)
					last_date = datetime.datetime(year, month, days[1]).date()
					rows = self.jorTaskTitle.shape
					total = 0
					for row in range(rows[0]):
						if (self.jorWorkDate[row] >= first_date and self.jorWorkDate[row] <= last_date):
							time = self.jorWorkTime[row]
							hours = time.days * 24 + time.seconds / 3600
							price = self.GetPrice(self.jorTaskTitle[row])
							subtotal = hours * price
							total += subtotal

							WorkDate[str(row)]  = self.jorWorkDate[row]
							TaskTitle[str(row)] = self.jorTaskTitle[row]
							TaskTime[str(row)]  = time
							UnitPrice[str(row)] = price
							SubTotal[str(row)]  = subtotal

		idx_df = []
		idx_df.append('日付')
		idx_df.append('タスク名')
		idx_df.append('タスク時間')
		idx_df.append('タスク単価')
		idx_df.append('小計')

		df = pd.concat([	WorkDate,
							TaskTitle,
							TaskTime,
							UnitPrice,
							SubTotal	],
							axis = 1)
		df.columns = idx_df
		print('Done.')

		print(f'Writing Excel File [{XLS_NAME}] ...', end = ' ')
		df.to_excel(XLS_NAME, sheet_name = SHEET_NAME, index = False)
		print('Done.')
		print('')
