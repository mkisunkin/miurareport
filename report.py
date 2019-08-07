import csv
import os
import glob

class Patient():
    def __init__(self, number, name, tests,results,num_prob,time):
        #self.tests = tests
        self.name = name
        self.tests = []
        self.number = number
        self.results = []
        self.num_prob = 1
        self.time = 'time'


testlist = [] 
new_row = -1
old_row = 0

list_of_files = glob.glob('C:/BCA/ExportedResults/*')
latest_file = max(list_of_files, key=os.path.getctime)

with open(latest_file) as csvfile:  # считываем фаил
     reader = csv.reader(csvfile, delimiter = ';')
     for row in reader:
        testlist.append(row)       # создаем список

patient_list = []

patient_list.append(Patient(0, testlist[0][2], ' ', ' ', ' ',' ')) #создаем нулевого пациента

num_pat = 0
for row in testlist:
        new_row += 1
        if new_row == 0:
                old_row = 0
        else:
            old_row = new_row - 1

        # сравнивание имя пациента с именем предыдущей строки
        if testlist[new_row][2] == testlist[old_row][2]:
                patient_list[num_pat].tests.append(
                    testlist[new_row][4])  # добавляем список тестов
                patient_list[num_pat].results.append(
                    testlist[new_row][7])  # добавляем список результатов
                patient_list[num_pat].num_prob = testlist[new_row][6] 
                # добавляем номер пробирки
                patient_list[num_pat].time = testlist[new_row][11]
        else:
                num_pat += 1
                # добавляем новый объект
                patient_list.append(
                    Patient(num_pat, testlist[new_row][2], ' ', ' ', ' ',' '))
                patient_list[num_pat].tests.append(
                    testlist[new_row][4])  # добавляем список тестов
                patient_list[num_pat].results.append(
                    testlist[new_row][7])  # добавляем список результатов
                patient_list[num_pat].num_prob = testlist[new_row][6]   # добавляем номер пробирки
                patient_list[num_pat].time = testlist[new_row][11]

print('Обработано пациентов -', num_pat)
#print(patient_list[treb].name, patient_list[treb].tests, patient_list[treb].results)

