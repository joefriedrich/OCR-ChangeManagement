'''
	Title:  Emails to Names 1.0
	Author:  Joe Friedrich
	License:  MIT (except where explictly stated)
'''

import re

'''
	Write!
'''

open_file = open(r'c:\users\friedrichj\desktop\attendees.txt')

outside_carrots = re.compile(r'[\w\s;.]*,[\w\s.]*')

name_list = outside_carrots.findall(open_file.read())

output = ''
for name in name_list:
    output = output + name[:-1]

print (output)