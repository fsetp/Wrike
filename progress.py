################################################################################
#
#
import pandas as pd
import helper

########################################
#
XLS_NAME	= 'output.xlsx'
SHEET_NAME	= 'Task Progress'

########################################
#
class ProgressLib():

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

		self.prgTaskTitle		= None
		self.prgTaskStartDate	= None
		self.prgTaskEndDate		= None
		self.prgTaskPeriod		= None
		self.prgmstTaskPeriod	= None
		self.prgPercentPeriod	= None

	####################################
	# Initialize
	#
	def Initialize(self):
		print('Initialize Prgress Object.')
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
#		print(self.jorWorkTime)

	####################################
	#
	def SetJournalDescription(self, series):
		self.jorjorDescription = series

	####################################
	#
	def SetTaskTitleSeries(self, series):
		self.TaskTitle = series
		rows = series.shape
		if (self.rows == 0):
			self.rows = rows

		elif (self.rows != rows):
			print('Warning! Progress Rows Different.')

	####################################
	#
	def SetTaskPeriodSeries(self, series):
		self.TaskPeriod = series
		rows = series.shape
		if (self.rows == 0):
			self.rows = rows

		elif (self.rows != rows):
			print('Warning! Progress Rows Different.')

	###################################
	#
	def MakeReport(self):
		print('Making Report Object.')

		print('Progress Search.')
		sch_rows_t = self.schTaskTitle.shape
		sch_rows = sch_rows_t[0]
		print(f'Schedule {sch_rows} Row(s).')
		jor_rows_t = self.jorTaskTitle.shape
		jor_rows = jor_rows_t[0]
		print(f'Journal {jor_rows} Row(s).')

		#
		print('Scaning Journal Object...', end = ' ')

		index = 0
		self.prgTaskTitle		= pd.Series(dtype = object)
		self.prgTaskStartDate	= pd.Series(dtype = object)
		self.prgTaskEndDate		= pd.Series(dtype = object)
		self.prgLatestWorkDate	= pd.Series(dtype = object)
		self.prgWorkDateJudge	= pd.Series(dtype = object)
		self.prgTaskPeriod		= pd.Series(dtype = object)
		self.prgmstTaskPeriod	= pd.Series(dtype = object)
		self.prgPercentPeriod	= pd.Series(dtype = object)

		idx_df = []
		idx_df.append('タスク名')
		idx_df.append('タスク開始日')
		idx_df.append('タスク終了日')
		idx_df.append('タスク最終日')
		idx_df.append('タスク判定')
		idx_df.append('タスク時間残')
		idx_df.append('タスク時間')
		idx_df.append('タスク残割合')

		#
		for sch_row in range(sch_rows):
			if (self.mstTaskPeriod[sch_row] != helper.TimeDeltaZero()):

				#
				schTaskTitle		= self.schTaskTitle[sch_row]
				schTimeDelta		= self.mstTaskPeriod[sch_row]
				schTaskStartDate	= self.schTaskStartDate[sch_row]
				schTaskEndDate		= self.schTaskEndDate[sch_row]
				jorTimeDelta		= schTimeDelta
				available			= False

				# Subract Journal Time Delta from Schedule Time Period
				jorLatestWorkDate = schTaskStartDate
				if (jorLatestWorkDate is not None):
					for jor_row in range(jor_rows):
						# 
#						print(type(self.schTaskTitle[sch_row]))
#						print(type(self.jorTaskTitle[jor_row]))
						if (self.schTaskTitle[sch_row] == self.jorTaskTitle[jor_row]):
							jorTimeDelta = jorTimeDelta - self.jorWorkTime[jor_row]
							available = True

						#
						jorWorkDate = self.jorWorkDate[jor_row]
#						print(type(jorLatestWorkDate))
#						print(type(jorWorkDate))
						if (helper.IsDateType(jorLatestWorkDate) and helper.IsDateType(jorWorkDate)):
							if (jorLatestWorkDate < jorWorkDate):
								jorLatestWorkDate = jorWorkDate

				# 
				if (available == True):

					#

					# Task Title
					self.prgTaskTitle[str(index)] = schTaskTitle

					# Task Start Date
					self.prgTaskStartDate[str(index)] = schTaskStartDate
#					self.prgTaskStartDate[str(index)].astype(str)

					# Task End Date
					self.prgTaskEndDate[str(index)] = schTaskEndDate
#					self.prgTaskEndDate[str(index)].astype(str)

					# Latest Work Date
					self.prgLatestWorkDate[str(index)] = jorLatestWorkDate

					# Work Date Expired Judgement
					strJudge = ''
#					print(type(jorLatestWorkDate))
#					print(type(schTaskStartDate))
					if (jorLatestWorkDate < schTaskStartDate):
						strJudge = 'Early'
					elif (jorLatestWorkDate > schTaskStartDate):
						strJudge = 'Expired'
					else:
						strJudge = 'In Schedule'
					self.prgWorkDateJudge[str(index)] = strJudge

					# Task Period Remain
#					value = jorTimeDelta.days * 24 + jorTimeDelta.seconds / 3600
					value = helper.Delta2Hours(jorTimeDelta)
					self.prgTaskPeriod[str(index)] = value
					self.prgTaskPeriod[str(index)].astype(float)
					print(jorTimeDelta)

					# Task Period
#					value = schTimeDelta.days * 24 + schTimeDelta.seconds / 3600
					value = helper.Delta2Hours(schTimeDelta)
					self.prgmstTaskPeriod[str(index)] = value
					self.prgmstTaskPeriod[str(index)].astype(float)

					# Period Remain Rate
					self.prgPercentPeriod[str(index)] = jorTimeDelta / schTimeDelta
					self.prgPercentPeriod[str(index)].astype(float)

					index += 1

		#
		df = pd.concat([	self.prgTaskTitle,
							self.prgTaskStartDate,
							self.prgTaskEndDate,
							self.prgLatestWorkDate,
							self.prgWorkDateJudge,
							self.prgTaskPeriod,
							self.prgmstTaskPeriod,
							self.prgPercentPeriod	],
							axis = 1)
		df.columns = idx_df
		print('Done.')

#		print(df)

		print(f'Writing Excel File [{XLS_NAME}] ...', end = ' ')
		df.to_excel(XLS_NAME, sheet_name = SHEET_NAME, index = False)
		print('Done.')

# https://note.nkmk.me/python-pandas-assign-append/
# https://note.nkmk.me/python-datetime-usage/
