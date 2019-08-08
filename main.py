from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from testxls import  wb, ft, bord_hor, fill, al, create_table
import testxls
from report import patient_list, num_pat, latest_file
from other_const import table_symbol, tests_dick
import datetime
import os
import glob


wb = Workbook()
ws0 = wb.active

print('Введите имя врача: ')
doc_name = input()

date_test = datetime.datetime.today()

'''Заполнение первой страницы'''


ws0.merge_cells('A1:I1')
ws0['A1'] = 'Лабораторный отчет'
ws0['A1'].font = ft[1]
ws0['A1'].alignment = al[1]
ws0.merge_cells('H2:I2')
ws0['H2'] = date_test.strftime("%d/%m/%Y_ %H:%M")



'''Отрисовка первой страницы'''

count_pat = len(patient_list)
step_row = 0


for i in range(count_pat):
    step_row = step_row + 4
    step_row_str = str(step_row)
    step_row_str_res = str(step_row + 1 )
    step_row_str_name = str(step_row - 1)
    ws0['A' + step_row_str_name] = patient_list[i].name
    ws0['D' + step_row_str_name] = 'Prob ' + patient_list[i].num_prob
    yrow = len(patient_list[i].tests)
    ''' В одну строчку убирается  9 тестов '''
    if yrow <= 9: # Если количество тестов меньше 9
        for k in range(yrow):
            xcol_tests = str(table_symbol[k] + step_row_str)
            xcol_results = str(table_symbol[k] + step_row_str_res)
            ws0[xcol_tests] = patient_list[i].tests[k]
            ws0[xcol_tests].fill = fill
            ws0[xcol_results] = patient_list[i].results[k]
    else:  # Если количество тестов больше 9
        for k in range(9):
            xcol_tests = str(table_symbol[k] + step_row_str)
            xcol_results = str(table_symbol[k] + step_row_str_res)
            ws0[xcol_tests] = patient_list[i].tests[k]
            ws0[xcol_tests].fill = fill
            ws0[xcol_results] = patient_list[i].results[k]
        hi_yrow = yrow - 9
        step_row = step_row + 3
        step_row_str = str(step_row)
        step_row_str_res = str(step_row + 1)
        for k in range(hi_yrow):
            xcol_tests = str(table_symbol[k] + step_row_str)
            xcol_results = str(table_symbol[k] + step_row_str_res)
            ws0[xcol_tests] = patient_list[i].tests[k]
            ws0[xcol_tests].fill = fill
            ws0[xcol_results] = patient_list[i].results[k]


sheets = ['']

''' Цикл отрисовывающий книги под каждого пациента'''
for i in range(count_pat):
    k = str(patient_list[i].num_prob)
    
    patient_name_sheet = 0
    patient_name_sheet = wb.create_sheet('Prob ' + k)
    create_table(patient_name_sheet)
    patient_name_sheet['B9'] = patient_list[i].name
    stp_row = 14
    patient_name_sheet['D47'] = date_test.strftime("%d/%m/%Y")
    patient_name_sheet['H47'] = doc_name
    

    for j in range(len(patient_list[i].tests)):
        stp_row += 1
        stp_row_str1 = str(stp_row )
        stp_row_str2 = str(stp_row )
        patient_name_sheet['A' + stp_row_str1] = tests_dick[patient_list[i].tests[j]][0]
        patient_name_sheet['A' + stp_row_str1].font = ft[2]
        patient_name_sheet['E' + stp_row_str2] = patient_list[i].results[j]
        patient_name_sheet['E' + stp_row_str1].font = ft[2]
        patient_name_sheet['G' + stp_row_str1] = tests_dick[patient_list[i].tests[j]][1]
        patient_name_sheet['G' + stp_row_str1].font = ft[2]
        patient_name_sheet['H' + stp_row_str1] = tests_dick[patient_list[i].tests[j]][2]
        patient_name_sheet['H' + stp_row_str1].font = ft[2]
        if j % 2 != 0:
            patient_name_sheet['A' + stp_row_str1].fill = fill
            patient_name_sheet['E' + stp_row_str1].fill = fill
            patient_name_sheet['G' + stp_row_str1].fill = fill
            patient_name_sheet['H' + stp_row_str1].fill = fill
        else : pass

    sheets.append(patient_name_sheet)

file_name = latest_file[23:-3]
report_file_name = 'rep/ ' + file_name + 'xlsx'
print(report_file_name)
print('готово')

wb.save(report_file_name)
