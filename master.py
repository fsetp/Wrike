################################################################################
#
#
import pandas as pd
import helper

########################################
#
XLS_NAME	= 'TaskMaster(1223).xlsx'
SHEET_NAME	= 'Task Info'

TASK_TITLE_COLUMN		= 'タスク名'
TASK_PERIOD_COLUMN		= 'タスク時間'
UNIT_PRICE_COLUMN		= 'タスク単価'

########################################
#
class TaskMasterLib():

	####################################
	# Constoructor
	#
	def __init__(self):
		self.TaskTitle		= None
		self.TaskPeriod		= None
		self.UnitPrice		= None
		self.rows			= 0
		self.columns		= 0

	####################################
	# Initialize
	#
	def Initialize(self):
		print('Initialize Task Master Object.')

		# Read Excel File
		print('Loading Excel File [', XLS_NAME, '] ...', end = ' ')
		self.DataFrame = pd.read_excel(XLS_NAME, SHEET_NAME)
#		print(self.DataFrame)

		# Get Rows and Columns
		self.rows, self.columns = self.DataFrame.shape
		print(f'{self.rows} Raw(s).')

#		print('Set NaN Cell to Float ...', end = ' ')
#		self._ForceSetFloat(self.DataFrame, TASK_PERIOD_COLUMN)
#		self._ForceSetFloat(self.DataFrame, TASK_PRICE_COLUMN)
#		print('Done.')

		# Set time delta in Task Period Cell
		print('Set Time Delta Cell to Period ...', end = ' ')
		total = self._ForceSetTimeDelta(self.DataFrame, TASK_PERIOD_COLUMN)
		print('Done.')
		print('Total Period ', total)
#		print(self.DataFrame)

		# Task Title
		self.TaskTitle = self.DataFrame[TASK_TITLE_COLUMN]

		# Task Period
		self.TaskPeriod = self.DataFrame[TASK_PERIOD_COLUMN]

		# Task Price
		self.UnitPrice = self.DataFrame[UNIT_PRICE_COLUMN]

		print('')

#		self.ShowTitlePrice()

#		print(self.DataFrame)
#		return self.DataFrame

	####################################
	#
	def ShowTitlePrice(self):
		#
		rows = self.TaskTitle.shape
		total = 0
		for row in range(rows[0]):
			print(f'{self.TaskTitle[row]} : {self.UnitPrice[row]}')

	####################################
	#
	def GetPrice(self, title):
		rows = self.TaskTitle.shape
		for row in range(rows[0]):
			if (title == self.TaskTitle[row]):
#				print(f'{title} {self.TaskTitle[row]} {self.UnitPrice[row]}')
				return self.UnitPrice[row]
		return 0

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
	def GetUnitPrice(self):
		return self.UnitPrice

	####################################
	#
	def _ForceSetFloat(self, df, column):
		rows, columns = df.shape
		for row in range(rows):
			if (type(df.iat[row, df.columns.get_loc(column)]) is not float):
				df.iat[row, df.columns.get_loc(column)] = 0.0

	####################################
	# Conversion str to timedelta in Task Period Cell
	#
	def _ForceSetTimeDelta(self, df, column):

		total_delta = helper.TimeDeltaZero()
		rows, columns = df.shape
		for row in range(rows):
			text = df.iat[row, df.columns.get_loc(column)]
			delta = helper.HoursFloat2TimeDelta(text)
			df.iat[row, df.columns.get_loc(column)] = delta
			total_delta += delta

		return total_delta
