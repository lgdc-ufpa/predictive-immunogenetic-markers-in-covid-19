from pathlib import Path
import os

# diretorio onde estão os xls bugados
dir_downloads = str(Path.home()) + '/Downloads'
print(dir_downloads)

file = []
for _, _, file in os.walk(dir_downloads):
    pass

n = 0 # n de xls

for elem in file:

    # checa se os 3 ultimos caracteres sao iguais a xls
    if 'xls' in elem[-3:]:
        print('xls: ', elem)

        # renomeia e troca a extensao
        old_file = os.path.join(dir_downloads, elem)
        new_file = os.path.join(dir_downloads, str(n + 1) + '.xlsx')

        os.rename(old_file, new_file)
        n += 1

print('total de arquivos renomeados e trocados de extensão: ', n)

