# from netMHCpan.Modules.modules import putResultsIntoExcel as putResultsIntoExcel
from NetMHCpan.Captura_do_raw_data.Modules.modules import putResultsIntoExcel as putResultsIntoExcel

global sheet
global hla_list
global epitopo
global excel_path # precisa porque vai salvar no mesmo diretorio

def getMinimoRnaking_sb_wb(sheet, hla_list, epitopo, excel_path):
# def getMinimoRnaking_sb_wb(sheet, epitopo, excel_path):
    coluna_nm = str
    coluna_hla = str
    coluna_ranking = str
    binders = int
    sb = int
    wb = int

    final = False
    linha = 3 # porque é a partir da linha 3 que comeca o ranking

    for i in range(len(hla_list)):

        if (i + 1) == 1:
            coluna_hla = "D"
            coluna_ranking = "H"
            coluna_nm = "G"

        if (i + 1) == 2:
            coluna_hla = "I"
            coluna_ranking = "M"
            coluna_nm = "L"

        if (i + 1) == 3:
            coluna_hla = "N"
            coluna_ranking = "R"
            coluna_nm = "Q"

        if (i + 1) == 4:
            coluna_hla = "S"
            coluna_ranking = "W"
            coluna_nm = "V"

        if (i + 1) == 5:
            coluna_hla = "X"
            coluna_ranking = "AB"
            coluna_nm = "AA"

        if (i + 1) == 6:
            coluna_hla = "AC"
            coluna_ranking = "AG"
            coluna_nm = "AF"

        if (i + 1) == 7:
            coluna_hla = "AH"
            coluna_ranking = "AL"
            coluna_nm = "AK"

        if (i + 1) == 8:
            coluna_hla = "AM"
            coluna_ranking = "AQ"
            coluna_nm = "AP"

        if (i + 1) == 9:
            coluna_hla = "AR"
            coluna_ranking = "AV"
            coluna_nm = "AU"

        if (i + 1) == 10:
            coluna_hla = "AW"
            coluna_ranking = "BA"
            coluna_nm = "AZ"

        print("coluna_hla: ", coluna_hla)
        print("coluna_ranking: ", coluna_ranking )
        print("coluna_nm: ", coluna_nm)

        while not final:
            # print(sheet["H" + str(linha)].value)
            if str(sheet[coluna_ranking + str(linha)].value) != 'None':
                linha = linha + 1
            else:
                linha = (linha - 1)
                print("Ultimo Ranking: ", sheet[coluna_ranking + str(linha)].value)
                print("linha: ", linha)
                final = True
                print("Final")

        # Busca o menor ranking
        print("Buscando o menor ranking")

        #hla
        print("Hla de numero ", i + 1)
        # Inicialmente, o primeiro ranking é o menor
        binders = 0
        sb = 0
        wb = 0

        hla = str(sheet[coluna_hla + "1"].value)
        ranking_minimo = float(sheet[coluna_ranking + "3"].value)
        nm_minimo = float(sheet[coluna_nm + "3"])

        for index in range(linha - 4):

            ranking_corrente = float(sheet[coluna_ranking + str(index + 4)].value)
            print("linha: ", index + 4, ".Ranking corrente: ", ranking_corrente)

            if ranking_corrente <= 2:
                binders = binders + 1
                print("binder! binders: ", binders)
                if ranking_corrente < 0.5:
                    sb = sb + 1
                    print("strongBinder! sb = ", sb)

            if ranking_corrente < ranking_minimo:
                ranking_minimo = ranking_corrente
                print(">>>>>>>>>>ranking minimo atualizado para " + str(ranking_corrente))

            nm_corrente = float(sheet[coluna_nm + str(index + 4)].value)

            if nm_corrente < nm_minimo:
                nm_minimo = nm_corrente

        wb = binders - sb
        print("binders, strong binders e weak binders: ")
        print(binders, sb, wb)

        # print("ranking_minimo: ", ranking_minimo)
        # sheet["E" + str(1)] = ranking_minimo
        # print(sheet["E" + str(1)].value)

        print("binders: ", binders)
        print("sb: ", sb)
        print("wb: ", wb)

        results = [hla, ranking_minimo, sb, wb, nm_minimo]
        print("result: ", results)

        putResultsIntoExcel(epitopo, results, excel_path)