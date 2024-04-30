# xlsxTest - EELE488 Prototype
# Sam Lovering
# Last Updated 4/15/2024
#
# This program receives a JSON string from either initData.py or the DB, then converts
# it into a spreadsheet.
#
# TODO:
# 2. Attempt Representing with a Symbol -> Verify w/ manual
# 3. Create "extremes" to model prototype
# 4. Make Unit Tests output into excel
#

import xlsxwriter
import json
import initData
import datetime


def create_spreadsheet(jsonDict):
    #Initialize Spreadseet
    print("Creating Spreadsheet")
    #Change with Meta Data
    workbook = xlsxwriter.Workbook('plain.xlsx')
    worksheet = workbook.add_worksheet()

    #Create Header for file
    colIterator = 1  #Temp colIterator
    headerKeys = jsonDict['0']
    for key in headerKeys.keys():
        worksheet.write(0, colIterator, key)
        colIterator += 1

    #Create format for times
    timeFormat = workbook.add_format({'num_format': 'hh:mm:ss AM/PM'})
    #Start Iterators at 0 (Temp, rework with yields)
    rowIterator = 1
    colIterator = 1
    for key, body in jsonDict.items():
        worksheet.write(rowIterator, 0, int(key))
        for bodyKey, bodyVal in body.items():
            if bodyKey == 'Time':
                worksheet.write(rowIterator, colIterator, bodyVal, timeFormat)
            elif bodyKey == "AttendeeID":
                worksheet.write(rowIterator, colIterator, int(bodyVal))
            else:
                worksheet.write(rowIterator, colIterator, bodyVal)
            colIterator += 1
        rowIterator += 1
        colIterator = 1
    worksheet.autofit()
    worksheet.set_column_pixels('E:E', 100)
    workbook.close()


def create_symbolic_spreadsheet(jsonDict):
    #Initialize Spreadsheet
    print("Creating Symbolic Spreadsheet")
    workbook = xlsxwriter.Workbook('symbolic.xlsx')
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


#This needs work, create header then apply formatting after?
def create_conditional_spreadsheet(jsonDict):
    #Initialize Spreadsheet
    print("Creating Conditional Spreadsheet")
    workbook = xlsxwriter.Workbook('conditional.xlsx')
    worksheet = workbook.add_worksheet()

    #Create Header for file
    colIterator = 1  #Temp colIterator
    headerKeys = jsonDict['0']
    for key in headerKeys.keys():
        worksheet.write(0, colIterator, key)
        colIterator += 1

    #Reset Iterators
    rowIterator = 1
    colIterator = 1
    #Create a conditional format for "out of range" times
    conditionalFormat = workbook.add_format({'num_format': 'hh:mm:ss AM/PM', 'bg_color': 'red'})
    timeFormat = workbook.add_format({'num_format': 'hh:mm:ss AM/PM'})

    for key, body in jsonDict.items():
        worksheet.write(rowIterator, 0, int(key))
        for bodyKey, bodyVal in body.items():
            if bodyKey == 'Time':
                #Intentionally create out of range time (Rework)
                tempTime = (datetime.datetime.combine(datetime.date(1, 1, 1), bodyVal) + datetime.timedelta(
                    minutes=7)).time()
                worksheet.conditional_format(rowIterator, colIterator, rowIterator, colIterator,
                                             {'type': 'time',
                                              'criteria': 'greater than',
                                              'value': bodyVal,
                                              'format': conditionalFormat})

                worksheet.write(rowIterator, colIterator, tempTime, timeFormat)
                colIterator += 1
            elif bodyKey == "AttendeeID":
                worksheet.write(rowIterator, colIterator, int(bodyVal))
            else:
                worksheet.write(rowIterator, colIterator, bodyVal)
            colIterator += 1
        rowIterator += 1
        colIterator = 1
    worksheet.autofit()
    worksheet.set_column_pixels('E:E', 100)
    workbook.close()


# This will likely need to be written out a little more to acquire times.
def parse_json(input):
    data = json.loads(input)
    for key in data:
        data[key]['Time'] = datetime.time.fromisoformat(data[key]['Time'])
    return data


#The following to methods are used for test UnitTests
def create_manual_spreadsheet():
    tempDict = parse_json(initData.create_manual_data())
    # Convert Time into datetime format for conditional spreadseet.

    create_spreadsheet(tempDict)
def create_con_spreadsheet():
    tempDict = parse_json(initData.create_manual_data())
    #Convert Time into datetime format for conditional spreadseet.
    create_conditional_spreadsheet(tempDict)

def create_sym_spreadsheet():
    tempDict = parse_json(initData.create_manual_data())
    #Take all times and dates in parsed json and turn it into a timestamp
    for key in tempDict:
        timestamp = f"{tempDict[key]['Date']} {tempDict[key]['Time']}"
        tempDict[key]['Timestamp'] = timestamp
        del tempDict[key]['Date']
        del tempDict[key]['Time']
    create_symbolic_spreadsheet(tempDict)

def create_auto_sym_spreadsheet(tempDict):
    for key in tempDict:
        timestamp = f"{tempDict[key]['Date']} {tempDict[key]['Time']}"
        tempDict[key]['Timestamp'] = timestamp
        del tempDict[key]['Date']
        del tempDict[key]['Time']
    create_symbolic_spreadsheet(tempDict)

def create_automatic_spreadsheet(max):
    create_spreadsheet(parse_json(initData.create_random_data(max)))


if __name__ == "__main__":
    print("xlsxTest called with __main__")
    #jsonTest = initData.create_manual_data()
    #jsonDict = parse_json(jsonTest)
    create_auto_sym_spreadsheet(parse_json(initData.create_random_data(10000)))