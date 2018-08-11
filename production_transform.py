import xlrd
import os
import re
import csv
import re
from datetime import datetime
import xlsxwriter 
 
MONTH_DICTIONARY = {"січень":"01", "лютий":"02", "березень":"03", "квітень":"04", "травень":"05", "червень":"06", "липень":"07", "серпень":"08", "вересень":"09", "жовтень":"10", "листопад":"11", "грудень":"12"}
COMPANY_CODES =  {'ДП"ДержВуглеПостач"':"40225511", 'ДП "Волиньвугілля"':"32365965", 'ДП "Мирноградвугілля"':"32087941", 
                    'ДП "Первомайськвугілля"':"32320594", 'ДП "Селидіввугілля"': "33426253", 'ПАТ "Лисичанськвугілля"':"32359108", 
                    'ДП "Львіввугілля"': "32323256", 'ДП "Торецьквугілля"':"33839013", 'ДП "Добропiллявугiлля':"37014600", 
                    'ТОВ "ДТЕК Добропiллявугiлля"':"37014600", 'ТОВ "ДТЕК СА"': "37596090", 'ДП "Львiввугiлля"':"32323256", 
                    'ДП "Красноармійськвугілля"': "32087941", 'ДП "Шахта ім. М.С.Сургая"':"40695853", 'ДП "Ш/у "Південнодонбаське №1"':"34032208",
                    'ДП "ВК "Краснолиманська"': "31599557", 'ПАТ "ш.Надія"': "00178175"}
YEAR_RE = re.compile("\d{4}")
NOT_SPACE = re.compile("\S+")
FIELD_NAMES_FILE = "production_field_names.csv"
FOLDER_NAME = "production_input"
OUT_FILENAME = "coal_production_"
HEADERS = ['month', "company", "company_code", "output", "value", "ton_cost"]

def load_workbook(wb):
    global sheet, ncol, dp, date_
    sheet = wb.sheet_by_index(0)
    ncol = sheet.ncols
    nrows = sheet.nrows
    content_began = False
    for i in range(nrows):
        cell = sheet.cell(i,0)
        if content_began:
            if not is_blank(cell):
                    dump_row(i)
            else:
                break
        else:
            if date_ == "":
                date_text = date_in_cell(cell)
                if date_text:
                    date_ = date_text
            if cell.value == "у т.ч.":
                content_began = True

def date_in_cell(cell):
    month_in_string = [MONTH_DICTIONARY[k] for k in MONTH_DICTIONARY.keys() if k in str(cell.value)]
    if len(month_in_string) > 0:
        year_matched = YEAR_RE.search(cell.value)
        if year_matched:
            year = year_matched.group(0)
            date_string = month_in_string[0] + "." + year
            return date_string

def is_blank(cell):
    return NOT_SPACE.search(str(cell.value)) == None

def dump_row(row_number):
    global sheet_dict, date_
    for k in fields_dictionary.keys():
        new_value = sheet.cell(row_number, int(k) - 1).value
        if isinstance(new_value, str):
            sheet_dict[fields_dictionary[k]].append(new_value.strip())
        else:
            sheet_dict[fields_dictionary[k]].append(new_value)
    sheet_dict['month'].append(date_)

def dict_to_list(dict_, headers):
    l = []
    for i in range(len(dict_[headers[0]])):
        new_l = []
        for h in headers:
            new_l.append(dict_[h][i])
        l.append(new_l)
    return l

with open(FIELD_NAMES_FILE, 'r') as vf:
    var_reader = csv.reader(vf)
    fields_dictionary = {}
    for l in var_reader:
        fields_dictionary[l[0]] = l[2]
sheet_dict = {}
for k in fields_dictionary.keys():
    sheet_dict[fields_dictionary[k]] = []
sheet_dict['month'] = []

files = os.listdir(FOLDER_NAME)
files = [f for f in files if f.endswith(".xls")]
for f in files:
    date_ = ""
    wb = xlrd.open_workbook(os.path.join(FOLDER_NAME, f), formatting_info=True)
    load_workbook(wb)

codes = [COMPANY_CODES[c] for c in sheet_dict['company']]
sheet_dict['company_code'] = codes
sheet_dict['month'] = [datetime.strptime(m, '%m.%Y') for m in sheet_dict['month']]

coal_list = dict_to_list(sheet_dict, HEADERS)
coal_list = sorted(coal_list, key=lambda x: x[0],reverse=True)
month_to_filename = datetime.strftime(coal_list[0][0], '%Y_%m')

with open(OUT_FILENAME + month_to_filename + ".csv", "w") as cfile:
    csvwriter = csv.writer(cfile)
    csvwriter.writerow(HEADERS)
    for i in range(len(coal_list)):
        l = coal_list[i][:]
        l[0] = datetime.strftime(l[0],"%m.%Y")
        csvwriter.writerow(l)

out_wb = xlsxwriter.Workbook(OUT_FILENAME + month_to_filename + ".xlsx")
worksheet = out_wb.add_worksheet()
datef = out_wb.add_format({'num_format':"mm.yyyy"})
numf = out_wb.add_format({'num_format':"0.00"})
headerf = out_wb.add_format({'bold':True})
for i in range(len(HEADERS)):
    worksheet.write(0, i, HEADERS[i], headerf)
for i in range(len(coal_list)):
    for j in range(len(HEADERS)):
        if j == 0:
            worksheet.write(i+1, j, coal_list[i][j], datef)
        elif j >  2:
            worksheet.write(i+1, j, coal_list[i][j], numf)
        else:
            worksheet.write(i+1, j, coal_list[i][j])
out_wb.close()      