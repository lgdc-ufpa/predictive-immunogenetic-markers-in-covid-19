from openpyxl import load_workbook
#estado = [numero_analises, analise_corrente, n_alelos_remanescentes]

alelos_path = "C:\\Users\\mauro\\OneDrive\\Doutorado\\alelostestefinal.xlsx"


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
