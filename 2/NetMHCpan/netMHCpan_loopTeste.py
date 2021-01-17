from openpyxl import load_workbook
#estado = [numero_analises, analise_corrente, n_alelos_remanescentes]

alelos_path = "C:\\Users\\mauro\\OneDrive\\Doutorado\\alelostestefinal.xlsx"
book_alelos = load_workbook(alelos_path)
sheet = book_alelos.active

alelos = []
numero_analises = 0
alelos_remanescentes = False
linha = 0

final = False
while(not final):
    linha = linha + 1
    # print(sheet["A"+str(linha)].value)
    if str(sheet["A"+str(linha)].value) == "None":
        print("Acabou")
        final = True
        linha = linha - 1
print(linha)
print("numero de analises de 10: ", int(linha/10))
print("numero de alelos da ultima analise: ", linha%10)

if(linha%10 == 0):
    numero_analises = int(linha/10)
else:
    numero_analises = int(linha/10) + 1
    n_alelos_remanescentes = linha%10

print(numero_analises, n_alelos_remanescentes)

for analise in range (numero_analises):
    print("analise: ", analise+1)
    if (analise + 1) != numero_analises:
        print("Pegando 10 alelos")
        for i in range(10):
            print(i+1,(analise*10 + i + 1) ,sheet["A"+str(analise*10 + i + 1)].value)
            if not i == 9:
                alelos.append(sheet["A"+str(analise*10 + i + 1)].value + ",")
            else:
                alelos.append(sheet["A"+str(analise*10 + i + 1)].value)
        print(alelos)
    else:
        for i in range(n_alelos_remanescentes):
            print(i+1,(analise*10 + i + 1) ,sheet["A"+str(analise*10 + i + 1)].value)
            if not i == 9:
                alelos.append(sheet["A"+str(analise*10 + i + 1)].value + ",")
            else:
                alelos.append(sheet["A"+str(analise*10 + i + 1)].value)
        print(alelos)
        print("Pegando alelos remanescentes :", n_alelos_remanescentes)
