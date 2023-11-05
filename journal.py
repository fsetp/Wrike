################################################################################
#
#
import pandas as pd
import helper

########################################
#
CSV_NAME	= 'db1750.csv'

WORK_DATE_COLUMN	= '日付'
WORKER_NAME_COLUMN	= '作業者'
TASK_TITLE_COLUMN	= 'タスク'
WORK_TIME_COLUMN	= '作業時間'
DESCRIPTION_COLUMN	= '備考'

########################################
#
class JournalLib():

	####################################
	# Constoructor
	#
	def __init__(self):
		self.DataFrame		= None
		self.WorkDate		= None
		self.WorkerName		= None
		self.TaskTitle		= None
		self.WorkTime		= None
		self.Description	= None
		self.rows			= 0
		self.columns		= 0

	####################################
	# Initialize
	#
	def Initialize(self):
		print('Initialize Journal Object.')

		# Read Csvl File
		print('Loading Csv File [', CSV_NAME, '] ...', end = ' ')
		self.DataFrame = pd.read_csv(CSV_NAME)
		print('Done.')

		# Get Rows and Columns
		self.rows, self.columns = self.DataFrame.shape
		print(f'{self.rows} Raw(s).')

		print('Set Date Cell to Date ...', end = ' ')
		self._ForceSetDate(self.DataFrame, WORK_DATE_COLUMN)
		print('Done.')
#		print(self.DataFrame)

		# Set time delta in Task Period Cell
		print('Set Time Delta Cell to Work Time ...', end = ' ')
		total = self._ForceSetTimeDelta(self.DataFrame, WORK_TIME_COLUMN)
		print('Done.')
		print('Total Period ', total)

		# Parent Task
		self.WorkDate = self.DataFrame[WORK_DATE_COLUMN]

		#
		self.WorkerName = self.DataFrame[WORKER_NAME_COLUMN]

		#
		self.TaskTitle = self.DataFrame[TASK_TITLE_COLUMN]

		#
		self.WorkTime = self.DataFrame[WORK_TIME_COLUMN]
#		print(self.WorkTime)

		#
		self.Description = self.DataFrame[DESCRIPTION_COLUMN]

		print('')
		return self.DataFrame

	####################################
	#
	def GetWorkDate(self):
		return self.WorkDate

	####################################
	#
	def GetWorkerName(self):
		return self.WorkerName

	####################################
	#
	def GetTaskTitle(self):
		return self.TaskTitle

	####################################
	#
	def GetWorkTime(self):
		return self.WorkTime

	####################################
	#
	def GetDescription(self):
		return self.Description

	####################################
	#
	def GetRows(self):
		return self.rows

	####################################
	#
	def GetWorkDateAtRow(self, row):
		if (self.rows > row):
			date = helper.DateText2Date(self.WorkDate[row])
			return date

		return None

	####################################
	#
	def GetWorkerNameAtRow(self, row):
		if (self.rows > row):
			return self.WorkerName[row]

		return ''

	####################################
	#
	def GetTaskTitleAtRow(self, row):
		if (self.rows > row):
			return self.TaskTitle[row]

		return ''

	####################################
	#
	def GetWorkTimeAtRow(self, row):
		if (self.rows > row):
			delta = helper.TimeText2TimeDelta(self.WorkTime[row])
			return delta

		return helper.TimeDeltaZero()

	####################################
	#
	def GetDescriptionAtRow(self, row):
		if (self.rows > row):
			return self.Description[row]

		return ''

	########################################
	#
	def ShowTaskTitle(self):
		for row in range(self.rows):
			print(self.TaskTitle[row])

	####################################
	#
	def ShowTaskTitleSeries(self):
		print(self.TaskTitle)

	####################################
	# Conversion str to timedelta in Task Period Cell
	#
	def _ForceSetTimeDelta(self, df, column):

		total_delta = helper.TimeDeltaZero()
		for row in range(self.rows):
			text = df.iat[row, df.columns.get_loc(column)]
#			print(text)

			delta = helper.TimeText2TimeDelta(text)
			df.iat[row, df.columns.get_loc(column)] = delta
			total_delta += delta

		return total_delta

	####################################
	# Conversion str to timedelta in Task Period Cell
	#
	def _ForceSetDate(self, df, column):

		for row in range(self.rows):
			text = df.iat[row, df.columns.get_loc(column)]
			if (type(text) is str):
				date = helper.DateText2Date(text).date()
				df.iat[row, df.columns.get_loc(column)] = date

			else:
				df.iat[row, df.columns.get_loc(column)] = helper.DateZero()


