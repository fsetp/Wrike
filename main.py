################################################################################
#
#
import schedule
from schedule import ScheduleLib

import journal
from journal import JournalLib

import progress
from progress import ProgressLib

import master
from master import TaskMasterLib

import monthly
from monthly import MonthlyLib

########################################
# main program
#
def main():
	###################################
	#
	sch = ScheduleLib()
	df = sch.Initialize()

#	sch.OutputTaskPeriod2Excel()

	###################################
	#
	mst = TaskMasterLib()
	mst.Initialize()

	###################################
	#
	jor = JournalLib()
	df = jor.Initialize()

	###################################
	#
	prg = ProgressLib()
	prg.Initialize()

	###################################
	#
	mon = MonthlyLib()
	mon.Initialize()

	###################################
	#
	prg.SetScheduleTaskTitle(sch.GetTaskTitle())
	prg.SetScheduleTaskStartDate(sch.GetTaskStartDate())
	prg.SetScheduleTaskEndDate(sch.GetTaskEndDate())

	prg.SetMasterTaskPeriod(mst.GetTaskPeriod())
	prg.SetMasterUnitPrice(mst.GetUnitPrice())

	mon.SetScheduleTaskTitle(sch.GetTaskTitle())
	mon.SetScheduleTaskStartDate(sch.GetTaskStartDate())
	mon.SetScheduleTaskEndDate(sch.GetTaskEndDate())
	mon.SetMasterTaskTitle(sch.GetTaskTitle())
	mon.SetMasterTaskPeriod(mst.GetTaskPeriod())
	mon.SetMasterUnitPrice(mst.GetUnitPrice())
	mon.SetJournalTaskTitle(jor.GetTaskTitle())
	mon.SetJournalWorkDate(jor.GetWorkDate())
	mon.SetJournalWorkerName(jor.GetWorkerName())
	mon.SetJournalWorkTime(jor.GetWorkTime())
	mon.SetJournalDescription(jor.GetDescription())

	prg.SetJournalTaskTitle(jor.GetTaskTitle())
	prg.SetJournalWorkDate(jor.GetWorkDate())
	prg.SetJournalWorkerName(jor.GetWorkerName())
	prg.SetJournalWorkTime(jor.GetWorkTime())
	prg.SetJournalDescription(jor.GetDescription())

	###################################
	#
	years = (2023, 2024)
	for year in years:
		for month in range(1, 13):
#			print(f'{year}/{month}')
			if (mon.GetMonthlyCount(year, month) > 0):
				mon.MakeMonthlyReport(year, month)

	mon.OutputMonthlyReport2Excel()


#	prg.MakeReport()

################################################################################
if __name__ == "__main__":
	main()
