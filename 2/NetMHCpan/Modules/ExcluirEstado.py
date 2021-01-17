global diretorio
global nome_epitopos
import os



def excluirEstado(diretorio, nome_epitopos):

    # Apaga o registro
    os.remove(diretorio + nome_epitopos + '_estado.txt')


# nome_epitopos = "VARV epitopes"
# diretorio = "C:\\Users\\mauro\\OneDrive\\Doutorado\\Epitopos\\"
# excluirEstado(diretorio, nome_epitopos)
