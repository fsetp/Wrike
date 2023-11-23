################################################################################
#
#
import pandas as pd
import helper

########################################
#
#XLS_NAME	= 'UCJV60 (export).xls'
XLS_NAME	= 'UCJV60 ES1追加 (export).xls'
SHEET_NAME	= 'Tasks'

INFO_XLS_NAME	= 'TaskInfo.xlsx'
INFO_SHEET_NAME	=  'Task Info'

PARENT_TASK_COLUMN		= 'Parent task'
TASK_TITLE_COLUMN		= 'Title'
TASK_PERIOD_COLUMN		= 'Description'
START_DATE_COLUMN		= 'Start Date'
DURATION_COLUMN			= 'Duration'
DURATION_HOURS_COLUMN	= 'Duration (Hours)'
TIME_SPENT_COLUMN		= 'Time Spent (Hours)'
END_DATE_COLUMN			= 'End Date'
DEPENDS_ON_COLUMN		= 'Depends On'

########################################
#
class ScheduleLib():

	####################################
	# Constoructor
	#
	def __init__(self):
		self.DataFrame		= None
		self.ParentTask		= None
		self.TaskTitle		= None
		self.StartDate		= None
		self.Duration		= None
		self.Duration_Hours	= None
		self.TimeSpent		= None
		self.EndDate		= None
		self.DependsOn		= None
		self.TaskPeriod		= None
		self.rows			= 0
		self.columns		= 0

	####################################
	# Initialize
	#
	def Initialize(self):
		print('Initialize Schedule Object.')

		# Read Excel File
		print('Loading Excel File [', XLS_NAME, '] ...', end = ' ')
		self.DataFrame = pd.read_excel(XLS_NAME, SHEET_NAME)
#		print(self.DataFrame)

		# Get Rows and Columns
		self.rows, self.columns = self.DataFrame.shape
		print(f'{self.rows} Raw(s).')

		# Set str in Parent Task Cell
		# Set str in Task Title Cell
		print('Set String Cell to Parent Task and Task Title  ...', end = ' ')
		self._ForceSetString(self.DataFrame, PARENT_TASK_COLUMN)
		self._ForceSetString(self.DataFrame, TASK_TITLE_COLUMN)
		print('Done.')

		# Set date in Start Data Cell
		# Set date in End Date Cell
		print('Set Date Cell to Start and End Date ...', end = ' ')
		self._ForceSetDate(self.DataFrame, START_DATE_COLUMN)
		self._ForceSetDate(self.DataFrame, END_DATE_COLUMN)
		print('Done.')

		# Set time delta in Task Period Cell
		print('Set Time Delta Cell to Period ...', end = ' ')
		total = self._ForceSetTimeDelta(self.DataFrame, TASK_PERIOD_COLUMN)
		print('Done.')
		print('Total Period ', total)

		# Parent Task
		self.ParentTask = self.DataFrame[PARENT_TASK_COLUMN]

		# Task Title
		self.TaskTitle = self.DataFrame[TASK_TITLE_COLUMN]

		# Start Date
		self.StartDate = self.DataFrame[START_DATE_COLUMN]

		# Duration
		self.Duratione = self.DataFrame[DURATION_COLUMN]

		# Duration Hours
		self.Duration_Hourse = self.DataFrame[DURATION_HOURS_COLUMN]

		# 
		self.TimeSpent = self.DataFrame[TIME_SPENT_COLUMN]

		# End Date
		self.EndDate = self.DataFrame[END_DATE_COLUMN]
#		print(self.EndDate)

		# Depebds on Task
		self.DependsOn = self.DataFrame[DEPENDS_ON_COLUMN]

		# Task Period
		self.TaskPeriod = self.DataFrame[TASK_PERIOD_COLUMN]

		print('')
		return self.DataFrame

	####################################
	#
	def GetTaskTitle(self):
		return self.TaskTitle

	####################################
	#
	def GetTaskPeriod(self):
		return self.TaskPeriod

	####################################
	#
	def GetTaskStartDate(self):
		return self.StartDate

	####################################
	#
	def GetTaskEndDate(self):
		return self.EndDate

	####################################
	#
	def OutputTaskPeriod2Excel(self):

		print('Create Task Information Excel File.')
		print('Making Information Object(s) ...', end = ' ')

		TaskTitle  = pd.Series(dtype = object)
		TaskPeriod = pd.Series(dtype = object)
		rows = self.TaskPeriod.shape
		for row in range(rows[0]):
			TaskTitle[str(row)] = self.TaskTitle[row]

			TimeDelta = self.TaskPeriod[row]

#			value = TimeDelta.days * 24 + TimeDelta.seconds / 3600
			value = helper.Delta2Hours(TimeDelta)
			TaskPeriod[str(row)] = value
			TaskPeriod[str(row)].astype(float)

		idx_df = []
		idx_df.append('タスク名')
		idx_df.append('タスク時間')

		df = pd.concat([TaskTitle,
						TaskPeriod	],
						axis = 1)
		df.columns = idx_df
		print('Done.')

		print(f'Writing Excel File [{INFO_XLS_NAME}] ...', end = ' ')
		df.to_excel(INFO_XLS_NAME, sheet_name = INFO_SHEET_NAME, index = False)
		print('Done.')
		print('')

	####################################
	#
	def _ForceSetString(self, df, column):
		rows, columns = df.shape
		for row in range(rows):
			if (type(df.iat[row, df.columns.get_loc(column)]) is not str):
				df.iat[row, df.columns.get_loc(column)] = ''

	####################################
	# Conversion str to timedelta in Task Period Cell
	#
	def _ForceSetTimeDelta(self, df, column):

		total_delta = helper.TimeDeltaZero()
		rows, columns = df.shape
		for row in range(rows):
			text = df.iat[row, df.columns.get_loc(column)]
			delta = helper.HoursText2TimeDelta(text)
			df.iat[row, df.columns.get_loc(column)] = delta
			total_delta += delta

		return total_delta

	####################################
	# Conversion str to timedelta in Task Period Cell
	#
	def _ForceSetDate(self, df, column):

		rows, columns = df.shape
		for row in range(rows):
			stamp = df.iat[row, df.columns.get_loc(column)]
#			print(type(stamp))
#			text = f'{stamp.year}-{stamp.month}-{stamp.day}'
			text = helper.Timestamp2Date(stamp)
#			print(type(text))
			if (type(text) is str):
				date = helper.DateText2DateIso(text)
				df.iat[row, df.columns.get_loc(column)] = date
#				print(type(date))

			else:
				df.iat[row, df.columns.get_loc(column)] = helper.DateZero()

