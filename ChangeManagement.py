'''
	Title:  Automated Change Mangement Agenda 0.6
	Author:  Joe Friedrich
	License:  MIT
'''

import csv
import datetime
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def get_input_file_name():
'''
	Write!
'''
	pause = input("Please hit the enter key when you are ready to select ITReport.CSV.")
	window = Tk()
	window.withdraw()
	file_name = askopenfilename()
	return file_name

def get_data(file_name):
'''
	Write!
'''
	open_file = open(file_name)
	data = list(csv.reader(open_file))
	open_file.close()
	return data

def get_user_date():
'''
	Write!
'''
	user_date = input('What is the date of the next meeting? (mm/dd/yy)  ')
	return user_date
	#Clean/check the input.

def write_ocr_block_to_file(csv_output, text, ocr_list):
'''
	Write!
'''
	header_line = "Start Date","End Date","Number","Request Title","Requestor"

	csv_output.writerow([text])
	csv_output.writerow(header_line)
	for ocr in ocr_list:
		csv_output.writerow(ocr)

#---------------------Begin Program--------------------

data = get_data(get_input_file_name())
garbage = data.pop(0) #removes headers
garbage = data.pop(0) #removes headers

user_date = get_user_date()

meeting_date_time = datetime.datetime.strptime(user_date + " 9:30 AM", "%m/%d/%y %I:%M %p")
print("\ncurrent meeting date/time " + meeting_date_time.isoformat())

previous_week = datetime.timedelta(days=7,hours=1)
previous_meeting_time = meeting_date_time - previous_week #one hour before the previous meeting
print("previous meeting date/time " + previous_meeting_time.isoformat())

next_week = datetime.timedelta(days=7)
next_meeting_time = meeting_date_time + next_week #the next meeting
print("next meeting date/time " + next_meeting_time.isoformat())

#this big-ass thing needs it's own function(s)==vvvvvvvvvvvvvvvvvvvv=
emergency_ocr = []
low_pre_approved = []
pre_approved = []
low = []
medium = []
high = []
open_5_days = []
previously_approved_changes = []

for ocr in data:
	emergency = ocr[7]
	priority = ocr[6]
	status = ocr[5]
	ocr_start_time = datetime.datetime.strptime(ocr[0], "%m/%d/20%y %I - %M %p")
	ocr_end_time = datetime.datetime.strptime(ocr[1], "%m/%d/20%y %I - %M %p")

	ocr = [ocr[0], ocr[1], ocr[4], ocr[2], ocr[16]]

	if emergency == 'Yes':
		if ocr_start_time > previous_meeting_time:
			emergency_ocr.append(ocr)
	elif status == "Awaiting Opp Approval":
		if ocr_start_time < meeting_date_time:
				if priority[0] == "3":
					low_pre_approved.append(ocr)
				else:
					pre_approved.append(ocr)
		elif priority[0] == "3":
			low.append(ocr)
		elif priority[0] == "2":
			medium.append(ocr)
		else:
			high.append(ocr)
	elif status == "Approved/Scheduled":
		if ocr_end_time < previous_meeting_time:
			open_5_days.append(ocr)
		elif meeting_date_time < ocr_start_time < next_meeting_time:
			previously_approved_changes.append(ocr)
#this big-ass thing needs it's own function(s)=^^^^^^^^^^^^^^^^^^^^=

if high != []:
	print("\nLast High Priority: " + high[-1][2])
	high_medium = high + medium
else:
	print("\nNo High Priority")
	high_medium = medium

#Windows only user profile shortcut
user_folder = os.environ['USERPROFILE']
output_csv = open(user_folder + r'\desktop\ChangeManagement.csv', 'w', newline='')
output_csv_writer = csv.writer(output_csv)

write_ocr_block_to_file(output_csv_writer, "OPEN 5 DAYS", open_5_days)
write_ocr_block_to_file(output_csv_writer, "EMERGENCY", emergency_ocr)
write_ocr_block_to_file(output_csv_writer, "PRE-APPROVED", pre_approved)
write_ocr_block_to_file(output_csv_writer, "LOW PRE-APPROVED", low_pre_approved)
write_ocr_block_to_file(output_csv_writer, "HIGH / MEDIUM", high_medium)
write_ocr_block_to_file(output_csv_writer, "LOW", low)
write_ocr_block_to_file(output_csv_writer, "PREVIOUS CHANGES", previously_approved_changes)

output_csv.close()