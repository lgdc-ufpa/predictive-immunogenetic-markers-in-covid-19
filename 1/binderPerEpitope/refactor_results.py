from openpyxl import load_workbook, Workbook
from os import getcwd, listdir

directory = getcwd()
files = listdir(directory)

print('--> Arquivos')
for file in files:
    print(file)

book_results = None
sheet_results = None

opened = False

print('\n')
while not opened:
    try:
        result_file = input('Entre com o nome do arquivo sem a extensão .xlsx: ')

        book_results = load_workbook( directory + '/' + result_file + '.xlsx')
        sheet_results = book_results.active
        opened = True
    except:
        print('ERRO: nome inválido. Insira um nome de arquivo válido')


hlas = dict()
protein_columns = dict()

for sheet in book_results:

    print('Sheet: ', sheet.title)
    hla_name = sheet.title
    hlas[hla_name] = list()

    for index_sheet, linha in enumerate(sheet):
        protein, nb, sb = linha[0].value, linha[1].value, linha[2].value
        hlas[hla_name].append([protein, nb, sb])

        if not protein in protein_columns and not protein == 'Protein':
            protein_columns[protein] = protein

    hlas[hla_name].pop(0)

csv_file = open(directory + '/refactored_result.csv', 'a')

columns = ""

for protein in protein_columns:
    columns += 'NB_' + protein + ',' + 'SB_' + protein + ','

columns = ',' + columns[:-1] + '\n'
csv_file.write(columns)

index = 0
for hla, proteins in hlas.items():
    print(index, hla, proteins)
    index += 1

    line = hla.replace('_', ':') + ','
    for protein in proteins:
        line += str(protein[1]) + ',' + str(protein[2]) + ','

    line = line[:-1] + '\n'
    csv_file.write(line)

csv_file.close()
