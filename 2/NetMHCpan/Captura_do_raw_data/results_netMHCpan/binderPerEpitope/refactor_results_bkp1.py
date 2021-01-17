from openpyxl import load_workbook, Workbook
from os import getcwd
import csv
import pandas as pd
import numpy as np

directory = getcwd()
book_results = load_workbook( directory + '/result.xlsx')
sheet_results = book_results.active

# criando novo arquivo
# csv_file = open(directory + '/refactored_result.csv', 'w')
# csv_file.write('BLANK,\n')

hlas = dict()

for sheet in book_results:
    print('Sheet: ', sheet.title)
    hla_name = sheet.title
    # hlas[hla_name] = [hla_name]
    hlas[hla_name] = list()

    for index_sheet, linha in enumerate(sheet):
        protein, nb, sb = linha[0].value, linha[1].value, linha[2].value
        print(protein, nb, sb)
        hlas[hla_name].append([protein, nb, sb])

    # csv_file.write(hla_name + ',\n')
    hlas[hla_name].pop(0)

# csv_file.close()

# csv_file = open(directory + '/refactored_result.csv', 'r')
# lines = csv_file.readlines()
# csv_file.close()

csv_file = open(directory + '/refactored_result.csv', 'a')
csv_file.write('BLANK,\n')

index = 0
for hla, proteins in hlas.items():
    print(index, hla, proteins)
    index += 1

    line = hla + ','
    for protein in proteins:
        line += str(protein[1]) + ',' + str(protein[2]) + ','

    line = line[:-1] + '\n'
    csv_file.write(line)

csv_file.close()