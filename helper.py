################################################################################
#
import datetime
import math

####################################
#
def Delta2Hours(delta):
	return (delta.days * 24 + delta.seconds / 3600)

####################################
#
def TimeText2TimeDelta(time_text):
	text = time_text.split(':')
	delta = datetime.timedelta(hours = int(text[1]), minutes = int(text[2]))

	return delta

####################################
#
def DateText2Date(date_text):
	date = datetime.datetime.strptime(date_text, "%Y/%m/%d")

	return date

####################################
#
def TimeText2TimeDelta(hours_text):
	text = hours_text.split(':')
	hours_text = f'{text[1]}:{text[2]}:{text[3]}'
#	print(hours_text)
	delta = datetime.timedelta(days = int(text[0]), hours = int(text[1]), minutes = int(text[2]), seconds = int(text[3]))
	return delta


####################################
#
def DateText2DateIso(date_text):
	date = None
#	print(date_text)
	try:
		date = datetime.date.fromisoformat(date_text)
	except:
		pass
#	print(type(date))
	return date

####################################
#
def HoursText2TimeDelta(hours_text):
	if (type(hours_text) is str):
		try:
			value = float(hours_text)
		except:
			value = 0.0

		f, i = math.modf(value)
		delta = datetime.timedelta(hours = i, minutes = f * 60)

	else:
		delta = TimeDeltaZero()

	return delta

####################################
#
def HoursFloat2TimeDelta(hours):
	if (type(hours) is float):
		f, i = math.modf(hours)
		try:
			delta = datetime.timedelta(hours = i, minutes = f * 60)

		except:
			delta = TimeDeltaZero()

	else:
		delta = TimeDeltaZero()

	return delta

####################################
#
def TimeDeltaZero():
	return datetime.timedelta(hours = 0)

####################################
#
def DateZero():
	return datetime.date.today()

####################################
#
def Timestamp2Date(stamp):
	text = f'{stamp.year}-{stamp.month}-{stamp.day}'
	return text

####################################
#
def IsDateType(date):
	if (type(date) is datetime.date):
		return True
	return False
