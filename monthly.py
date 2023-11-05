################################################################################
#
#
import pandas as pd
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
	def Do(self):
		rows = self.schTaskTitle.shape

		#
		for row in range(rows[0]):
#			self.schTaskTitle[row]
#			self.schTaskStartDate[row]
#			self.schTaskEndDate[row]

			print(type(self.schTaskStartDate[row]))
			print(f'Title: {self.schTaskTitle[row]} Start Date: {self.schTaskStartDate[row]} - End Date:{self.schTaskEndDate[row]}')
