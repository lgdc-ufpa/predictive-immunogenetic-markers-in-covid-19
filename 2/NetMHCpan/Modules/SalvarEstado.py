global diretorio
global numero_analises
global ultima_analise
global n_alelos_remanescentes

global nome_epitopos
global epitopos_path

# numero_analises = str(numero_analises) #str(numero de analises)
# ultima_analise = str(ultima_analise)
# n_alelos_remanescentes = str(n_alelos_remanescentes)

# teste SalvarEstado
# nome_epitopos = "VARV epitopes"

def salvarEstado(diretorio, nome_epitopos, ultima_analise, numero_analises, n_alelos_remanescentes):

    # Apaga o registro
    file = open(diretorio + nome_epitopos + '_estado.txt', 'w')
    file.close()

    # checa se há registros
    file = open(diretorio + nome_epitopos + '_estado.txt', 'r')
    estados = file.readlines()
    print(estados)
    file.close()

    # registra
    # ultima analise e numeroamalises estao com valores trocados
    file = open(diretorio + nome_epitopos + '_estado.txt', 'w')
    file.write(str(numero_analises) + '\n' + str(ultima_analise) + '\n' + str(n_alelos_remanescentes))
    file.close()

    # checa se há registros
    file = open(diretorio + nome_epitopos + '_estado.txt', 'r')
    estados = file.readlines()
    print(estados)
    estado = file.readline()
    print(estado)
    file.close()

# numero_analises = str(104) #str(numero de analises)
# ultima_analise = str(3)
# n_alelos_remanescentes = str(6)
#
# nome_epitopos_txt = "VARV epitopes"
# diretorio = "C:\\Users\\mauro\\OneDrive\\Doutorado\\Epitopos\\"
# salvarEstado(numero_analises, ultima_analise, n_alelos_remanescentes)
#
# nome_epitopos = "VARV epitopes"
# file = open(diretorio + nome_epitopos + '_estado.txt', 'r')
#
# estados = file.readlines()
# print(estados)
# file.close()
# ultima_analise = estados[0]
# numero_analises = estados[1]
# n_alelos_remanescentes = int(estados[2])
#
# print(ultima_analise, numero_analises, n_alelos_remanescentes)