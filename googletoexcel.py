from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
import numpy as num
import pandas as pd
import copy

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'google_excel_key.json'

creds = None
creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)


# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1b4UOb4PrexdcJTyLv1ptjL3n8tfTYV4HAlX20x-Zc6Y'


service = build('sheets', 'v4', credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()
result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                            range="workinghours").execute()
values = result.get('values', [])
# print(result)

# print(len(values)) 

# creating a dictionary with location as key and name as value
# for example:
# 55_lentor_way = {"Name" : [Islam, Subrot]
# "1" : [8,4,8,0]
# "2" : [8,2,8,3]}
###################################################################################
# creating different locations for different sheets
locations = []
dict_locations = []
for row in values[1:]:
        try:
                if row[3] not in locations:
                        locations.append(row[3])
                if row[5] not in locations:
                        locations.append(row[5])
        except IndexError:
                continue

print(locations)
# create dictionary template for each worksite
template = {}
template['Names'] = []
for day in range(1,32):
        template[day] = []
template['Pay/hour'] = []
template['Hours']=[]
template['Pay'] = []
template['Total Pay'] = []
# print(dict_locations)

# data set for worker salay (can be moved to another file)
workerpay_dict = {      "Rana MD" : [3.375, 5.0625],
                        "Subrot" : [3,4.5],
                        "Nasim" : [2.75, 4.125],
                        "Shahabuddin" : [3.25, 4.875],
                        "Mofazzol" : [3, 4.5],
                        "Alam MD Mozibul" : [4.375, 6.5625],
                        "Hossen MD Monir" : [2.75, 4.125],
                        "Rahman Azizur" : [2.75, 4.125],
                        "Islam" : [5, 7.5],
                        "Hasan" : [3.625, 5.4375],
                        "Gourango" : [2.375, 3.5625],
                        "赵家军" : [11],
                        "王玉镇" : [14]}

def add_name(name, salary, worksite):
        worksite['Names'].append(name)
        for day in range(1, 32):
                worksite[day].append(0)
        worksite['Pay/hour'].append(salary[0])
        worksite['Hours'].append(0)
        worksite['Pay'].append(0)
        worksite['Total Pay'].append(0)
        if len(salary) == 2:
                worksite['Names'].append(name + " (OT)")
                for day in range(1, 32):
                        worksite[day].append(0)
                worksite['Pay/hour'].append(salary[1])
                worksite['Hours'].append(0)
                worksite['Pay'].append(0)
                worksite['Total Pay'].append(0)

# fill dictionaries with names and 0 hours first
column_headers = []
for key, value in workerpay_dict.items():
        # temp_dict = copy.deepcopy(template)
        add_name(key, value, template)

for i in range(len(locations)):
        temp_sheet = copy.deepcopy(template)
        dict_locations.append(temp_sheet)
        
def duplicates(lst, item):
        return [i for i, x in enumerate(lst) if x == item]

# filter each data row in google excel and add values to respective worksites
for row in values[1:]:
        # adjust to filter by name
        # if row[1] == 'Islam':
        #         print(row)
        if len(row) == 7:
                worksite_index =locations.index(row[3])
                worksite_OT_index = locations.index(row[5])

                # extract day from the string
                working_day_list = list(row[2])
                backslash_index = duplicates(working_day_list, "/")
                working_day = ""
                for i in range(backslash_index[0]+1, backslash_index[1]):
                        working_day += working_day_list[i]
                working_day = int(working_day)

                # print(worksite_index)
                # print(worksite_OT_index)
                # print(working_day)

                # data cleaning for that particular row (change "" to "0")
                if row[6] == "":
                        row[6] == "0"
                if row[4] == "":
                        row[4] == "0"

                # only china men no OT
                if row[1]== "赵家军" or row[1] == "王玉镇":
                        working_hours = int(row[4]) + int(row[6])
                        working_hours_OT = 0
                else:
                        working_hours = int(row[4])
                        working_hours_OT = int(row[6])

                # update dictionary
                name_index = dict_locations[worksite_index]["Names"].index(row[1])
                dict_locations[worksite_index][working_day][name_index] = working_hours
                if row[1]!= "赵家军" or row[1] != "王玉镇":
                        name_index = dict_locations[worksite_index]["Names"].index(row[1] + " (OT)")
                        dict_locations[worksite_index][working_day][name_index] = working_hours_OT
        else:
                # print(row)
                worksite_index =locations.index(row[3])

                # extract day from the string
                working_day_list = list(row[2])
                backslash_index = duplicates(working_day_list, "/")
                working_day = ""
                for i in range(backslash_index[0]+1, backslash_index[1]):
                        working_day += working_day_list[i]
                working_day = int(working_day)

                # print(worksite_index)
                # print(working_day)

                working_hours = int(row[4])

                # update dictionary
                name_index = dict_locations[worksite_index]["Names"].index(row[1])
                dict_locations[worksite_index][working_day][name_index] = working_hours

# df = pd.DataFrame(dict_locations[0], columns = dict_locations[0].keys())
# df.to_excel (r'55 Lentor Way.xlsx', index = False, header=True)

# df = pd.DataFrame(dict_locations[1], columns = dict_locations[0].keys())
# df.to_excel (r'142 Rangoon Road.xlsx', index = False, header=True)


# calculating total hours per worker
def calc_total_hours(dict_loc):
        for location in dict_loc:
                no_of_workers = len(location['Names'])
                for day in range(1,32):
                        for worker in range(no_of_workers):
                                location['Hours'][worker] += location[day][worker]
calc_total_hours(dict_locations)

# calculating pay per worker
def calc_pay(dict_loc):
        for location in dict_loc:
                no_of_workers = len(location['Names'])
                for worker in range(no_of_workers):
                        location['Pay'][worker] = location['Hours'][worker] * location['Pay/hour'][worker]
calc_pay(dict_locations)

# calculating total pay per worker
def calc_total_pay(dict_loc):
        for location in dict_loc:
                worker_list = workerpay_dict.keys()
                for worker in range(len(location['Names'])):
                        if location['Names'][worker] in worker_list:
                                # print(location['Names'][worker])
                                if location['Names'][worker] == "赵家军" or location['Names'][worker] == "王玉镇":
                                        location['Total Pay'][worker] = location['Pay'][worker] 
                                else:
                                        location['Total Pay'][worker] = location['Pay'][worker] + location['Pay'][worker+1]
calc_total_pay(dict_locations)

# creating individual summary 
def individual_summary(loc, dic_loc, loc_temp):
        temp = copy.deepcopy(loc_temp)
        dic_loc.append(temp)
        loc.append("Individual Summary")
        # print(dic_loc[6])
        for worker in range(len(dic_loc[0]['Names'])):
                for worksite in range(len(loc)-1):
                        for day in range(1,32):
                                dic_loc[len(loc)-1][day][worker] += dic_loc[worksite][day][worker]
                        dic_loc[len(loc)-1]['Hours'][worker] += dic_loc[worksite]['Hours'][worker]
                        dic_loc[len(loc)-1]['Pay'][worker] += dic_loc[worksite]['Pay'][worker]
                        dic_loc[len(loc)-1]['Total Pay'][worker] += dic_loc[worksite]['Total Pay'][worker]

individual_summary(locations, dict_locations, template)

# writing all workplaces to one excel
writer = pd.ExcelWriter('May 2021.xlsx', engine='xlsxwriter')
for location in range(len(locations)):
        df = pd.DataFrame(dict_locations[location], columns = dict_locations[0].keys())
        df.to_excel (writer, sheet_name=locations[location])
writer.save()
                
