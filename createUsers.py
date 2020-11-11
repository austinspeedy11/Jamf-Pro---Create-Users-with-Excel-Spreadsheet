#!/usr/bin/python

import xlrd
import requests

#Declare sheet path
#loc = "/Users/austin.north/Desktop/test.xlsx"

loc = input("Please enter the path to your .xlsx file: ")

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)

user = input("Please enter your Jamf Pro API user's username: ")
password = input("Please enter your Jamf Pro API user's password: ")
instanceName = input("Please enter your Jamf Pro instance name: ")

sheetIDs = []
for i in range(sheet.nrows -1):
		sheetIDs.append(int((sheet.cell_value(i + 1, 0))))
		
sheetNames = []
for i in range(sheet.nrows -1):
	sheetNames.append((sheet.cell_value(i + 1, 1)))
	
sheetFullNames = []
for i in range(sheet.nrows -1):
	sheetFullNames.append((sheet.cell_value(i + 1, 2)))

sheetEmails = []
for i in range(sheet.nrows -1):
	sheetEmails.append((sheet.cell_value(i + 1, 3)))
		
sheetPhoneNumbers = []
for i in range(sheet.nrows -1):
	sheetPhoneNumbers.append(int(sheet.cell_value(i + 1, 4)))
	
sheetPositions = []
for i in range(sheet.nrows -1):
	sheetPositions.append((sheet.cell_value(i + 1, 5)))

for i in range(sheet.nrows -1):
	resp = requests.post("https://" + instanceName + ".jamfcloud.com/JSSResource/users/id/0", auth=(user, password), headers={"content-Type": "text/xml"}, data="<user><id>0</id><name>" + sheetNames[i] + "</name><full_name>" + sheetFullNames[i] + "</full_name><email>" + sheetEmails[i] + "</email><email_address>aharrison@company.com</email_address><phone_number>" + str(sheetPhoneNumbers[i]) + "</phone_number><position>" + sheetPositions[i] + "</position><sites><site><id>-1</id><name>None</name></site></sites></user>")

	print("Creating user with Name value \"" + sheetNames[i] + "\", Full Name value \"" + sheetFullNames[i] + "\", Email value \"" + sheetEmails[i] + "\", Phone Number value \"" + str(sheetPhoneNumbers[i]) + "\", and Position value \"" + sheetPositions[i] + "\".")