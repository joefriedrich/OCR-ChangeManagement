'''
	Title:  Emails to Names 1.1
	Author:  Joe Friedrich
	License:  MIT (except where explictly stated)
'''

import re
import os
import datetime

'''
	Copy and paste the attendees of a Lync/Skype meeting into
    a text file.  Specify the location of that text file in 
    the open_file variable.  For each name, this will remove 
	any text after (and including) @ and < until it reaches a ;.
	Outputs to a file specified in the file_location variable.
'''

user_folder = os.environ['USERPROFILE']

open_file = open(user_folder + r'\desktop\attendees.txt')
name_list = open_file.read().split('; ')
open_file.close()

output = ''
for name in name_list:
    separated_name = name.split(' <', 1)
    if len(separated_name) < 2:
        separated_name = separated_name[0].split('@')
    output += separated_name[0] + '; '
	
file_location = user_folder + "\desktop\CMMattendance" + ".txt"

write_file = open(file_location, 'w', newline='')
write_file.write(output)
write_file.close()