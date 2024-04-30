# xlsxTest - EELE488 Prototype
# Sam Lovering
# Last Updated 4/22/2024
#
# This program receives a JSON string from either initData.py or the DB, then converts
# it into a spreadsheet.

import datetime
import xlsxwriter
import json
from reports import initData


def create_spreadsheet(jsonDict):
    #Initialize Spreadsheet

    #Create Excel file with meta data
    fileName = 'attendanceReport' + datetime.datetime.now().strftime("%m%d%H%M") + '.xlsx'

    print("Creating Symbolic Spreadsheet")
    workbook = xlsxwriter.Workbook('xlsx/'+fileName)
    worksheet = workbook.add_worksheet()

    #Turn Timestamp into a "P", create note of timestamp
    designation_format = workbook.add_format({'bold': True, 'bg_color': '#D47554', 'align': 'center'})

    #Create Header for file
    colIterator = 1  #Temp colIterator
    headerKeys = jsonDict['0']
    for key in headerKeys.keys():
        worksheet.write(0, colIterator, key)
        colIterator += 1

    #Start Iterators at 0 (Temp, rework with yields)
    rowIterator = 1
    colIterator = 1
    #Write Data to spreadsheet
    for key, body in jsonDict.items():
        worksheet.write(rowIterator, 0, int(key))
        for bodyKey, bodyVal in body.items():
            if bodyKey == 'Timestamp':
                worksheet.write(rowIterator, colIterator, 'P', designation_format)
                worksheet.write_comment(rowIterator, colIterator, str(bodyVal))

            elif bodyKey == "AttendeeID":
                worksheet.write(rowIterator, colIterator, int(bodyVal))
            else:
                worksheet.write(rowIterator, colIterator, str(bodyVal))
            colIterator += 1
        rowIterator += 1
        colIterator = 1
    worksheet.autofit()
    worksheet.set_column_pixels('E:E', 100)
    workbook.close()

    return fileName


def parse_attendance_events(tempDict):
    tempDict = json.loads(tempDict)
    for key in tempDict:
        timestamp = f"{tempDict[key]['Date']} {tempDict[key]['Time']}"
        tempDict[key]['Timestamp'] = timestamp
        del tempDict[key]['Date']
        del tempDict[key]['Time']
    return tempDict


def generate_spreadsheet():
    fileName = create_spreadsheet(parse_attendance_events(initData.create_random_data(50)))
    return json.dumps(fileName)


if __name__ == "__main__":
    print("xlsxTest called with __main__")
