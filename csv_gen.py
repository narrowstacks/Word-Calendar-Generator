import icalendar
import calendar
import recurring_ical_events
import urllib.request
import datetime
import os
from settings import CORRECTION, ROOMS
from settings import URL, HOLIDAYS


# to do:
# add more comments!

def calgen(first_day,last_day,month_or_week):
	''' generate calendar'''
 
	# create/clear temp csv file for processing and safety
	temp_file = "temp.csv"
	open(temp_file,"w").close()


	in_start_date = first_day.split("/")
	in_end_date = last_day.split("/")


	file_name_gen = f"{''.join(in_start_date)}-{''.join(in_end_date)}-calendar_table.csv"

	# get ical file and assign it to a calendar variable

	ical_string = urllib.request.urlopen(URL).read()
	gen_calendar = icalendar.Calendar.from_ical(ical_string)

	out_start_date, out_end_date = get_date_range(in_start_date, in_end_date,month_or_week)
	events = recurring_ical_events.of(gen_calendar).between(out_start_date, out_end_date)
	writeEventsToFile(events,temp_file)

	sort_csv(temp_file,file_name_gen,5)
	os.remove("temp.csv") 
	return file_name_gen

def get_date_range(raw_start_dates, raw_end_dates,month_or_week):

	''' '''

	# turn raw dates into readable variables
	start_year = int(raw_start_dates[2])
	start_month = int(raw_start_dates[0])
	start_day = int(raw_start_dates[1])

	end_year = int(raw_end_dates[2])
	end_month = int(raw_end_dates[0])
	end_day = int(raw_end_dates[1])
	
	days_in_month = calendar.monthrange(end_year,end_month)[1] 
	
	if month_or_week == "month":
		if end_month == 12:
			end_year += 1
			end_month = 1
			end_day = 1
		else:
			end_month += 1
		end_day = 1
		start_day = 1
	
	if month_or_week == "week":
		if end_day + 1 > days_in_month:
			if end_month == 13:
				end_year += 1
				end_month = 1
			else:
				end_month += 1

			end_day = end_day + 1 - days_in_month
		else:
			end_day += 1
		
	end_date = end_year, end_month, end_day
	start_date = start_year, start_month, start_day

	return start_date, end_date



def writeEventsToFile(events,file_name):
	holiday_seen = False
	last_day = []
	with open(file_name,"a") as table_gen:
		table_gen.writelines("\n")
		for event in events:
			# get event information
			name = event["SUMMARY"].replace(",","")
			start = convert_timezone(event["DTSTART"].dt)
			end = convert_timezone(event["DTEND"].dt)
			room = event["LOCATION"]

			# get strings for use in CSV
			if room in ROOMS:
				room = room
			else:
				room = CORRECTION[room]
			hour_started = datetime.datetime.strftime(start, '%H%M')
			hour_ended = datetime.datetime.strftime(end, '%H%M')
			day = str(start)[:-15]
			numDay = int(day[8:])
			monthDay = int(str(start)[5:7])			

			# if the day is a holiday, put it at the top of the day (even if events are occuring)
			if start in HOLIDAYS:
				if numDay in last_day:
					holiday_seen = True
				else:
					holiday_seen = False
				if holiday_seen == False:
					numDay = int(day[8:])
					row = f"{HOLIDAYS.get(start)}|HOLIDAY|0000|{2359}|{start}|{numDay}|{monthDay}\n"
					table_gen.write(row)
					last_day.append(numDay)
					holiday_seen = True
				else:
					pass					
			
			# detect and prevent specific dates excluded from reoccuring (because recurring_cal_events doesn't seem to want to)
			try:
				# if there are multiple EXDATE objects, treat as list and skip each item if it lands on a date to be skipped
				if type(event["EXDATE"]) is list:
					exdate = event["EXDATE"]
					for i in range(len(exdate)):
						unpacked_object_list = vars(exdate[i])['dts']
						unpacked_object = convert_timezone(vars(unpacked_object_list[0])['dt'])
						if start == unpacked_object:
							skip_event = True
							break
						else:
							skip_event = False			
				# treat EXDATE as single object
				else:	
					exdate = event["EXDATE"]
					exdate_object_init = vars(exdate)['dts']
					exdate_object_dt = vars(exdate_object_init[0])['dt']
					exdate_object_dt = convert_timezone(exdate_object_dt)
					if start == exdate_object_dt:
						skip_event = True
					else:
						skip_event = False
			except:
				# if there are no EXDATE objects, don't skip the event
				skip_event = False
			


			if skip_event == False:
				
				row = f"{name}|{room}|{hour_started}|{hour_ended}|{start}|{numDay}|{monthDay}\n"
				table_gen.write(row)
				
			


def sort_csv(csv_temp,csv_filename,list_position):
	''' sort CSV by calendar date ''' 
	
	open(csv_filename,"w").close()
	for num in range(35):
		with open(csv_filename,"a") as export:
			with open(csv_temp,'r') as source: 
				for line in source:
					if line == '\n':
						pass
					else:
						line_list = line.split("|")
						line_list[list_position] = int(line_list[list_position].replace("\n",""))
						if line_list[list_position] == num:
							export.write(line)

def convert_timezone(time):
    time_string = f'{time}'
    if time_string[-6:] == "+00:00":
        time = time.astimezone(timezone(SETTINGS.TIMEZONE))
    return time




