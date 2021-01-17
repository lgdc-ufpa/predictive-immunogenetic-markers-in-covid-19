import os
from openpyxl import load_workbook

hla_directory = os.getcwd() + '/files/'
hla_files = sorted(os.listdir(hla_directory), reverse=False)

for elem in hla_files:
    if '.xls.xlsx' in elem:
        os.rename(str(hla_directory + elem), str(hla_directory + elem[:2] + 'xlsx'))

hla_directory = os.getcwd() + '/files/'
hla_files = sorted(os.listdir(hla_directory), reverse=False)

li_absent_rank_files = []
for index_hla, hla_file in enumerate(hla_files):
    # proteins[protein_name] = [NumBinders, NumStrongBinders]

    hla_book = load_workbook(hla_directory + hla_file)
    hla_sheet = hla_book.active
    hla_name = str(hla_sheet['D1'].value).replace(':', '_')
    print(hla_file, hla_name)

    index = 3
    while str(hla_sheet['C' + str(index)].value) != 'None':

        if str(hla_sheet['H' + str(index)].value) == 'None':

            li_absent_rank_files.append([hla_file, hla_name,
                                         str(hla_sheet['A' + str(index)].value),
                                         hla_sheet['B' + str(index)].value,
                                         hla_sheet['C' + str(index)].value,
                                         hla_sheet['D' + str(index)].value,
                                         hla_sheet['E' + str(index)].value,
                                         str(hla_sheet['F' + str(index)].value),
                                         str(hla_sheet['G' + str(index)].value),
                                         str(hla_sheet['H' + str(index)].value),
                                         str(hla_sheet['I' + str(index)].value),
                                         str(hla_sheet['J' + str(index)].value)])

        index += 1

    hla_book.close()

#TODO: storage the result into a csv file
with open(os.getcwd() + f'{os.sep}absent_ranks.csv', 'a') as f:
    header = 'File,hla_name,Pos,Peptide,ID,core,icore,l-log50k,nM,Rank,Ave,Nb\n'
    f.write(header)
    for elem in li_absent_rank_files:
        line = ['None' if v is None else v for v in elem]
        line = ",".join(line).replace('None', '') + '\n'
        f.write(line)
    f.close()
