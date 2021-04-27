
from openpyxl.utils.cell import column_index_from_string
import table_python as tp
import openpyxl

log_errs = []
def genLogErrors():
    file = open("erros.txt", "w")
    file.write(log_errs.__repr__())
    file.close()
    
def write_tests(file_name="FP.xlsx", interval="E13:E54"):
    # Escreve os numeros de ensaios na lista
    # values_to_insert['0-> primeira lista da minha lista']['0-> primeiro elemento da minha lista']['corresponde aos elementos da lista querida ->0]
    book_xlsx = tp.loadBook(file_name)
    ws = book_xlsx.active
    listsheet_tup = ws[interval]
    values_to_insert = tp.getArcValues("elet",30,15,5)
    arc_size = len(values_to_insert)
    arc_index = 0
    error = None
    while(arc_index < arc_size+1):
        try:
            for tup_value in listsheet_tup:
                tup_value[0].value = values_to_insert[arc_index][0].value
                book_xlsx.save(file_name)
                arc_index += 1
        except Exception as error:
            log_errs.append(error)
            book_xlsx.close()
            break


def write_FP(interval, type_row, time_values, first_term, ratio, file_name="FP.xlsx", index_list=1):
    # Escreve no fp.xlsx
    # values_to_insert['0-> primeira lista da minha lista']['0-> primeiro elemento da minha lista']['corresponde aos elementos da lista querida ->0]
    book_xlsx = tp.loadBook(file_name)
    ws = book_xlsx.active
    listsheet_tup = ws[interval]
    arc_index = 0
    error = None
    if(type_row == "elet"):
        values_to_insert = tp.getArcValues("elet",time_values,first_term,ratio)
        arc_size = len(values_to_insert)
        while(arc_index < arc_size+1):
            try:
                for tup_value in listsheet_tup:
                    tup_value[0].value = values_to_insert[arc_index][1][index_list]
                    book_xlsx.save(file_name)
                    arc_index += 1
            except Exception as error:
                log_errs.append(error)
                book_xlsx.close()
                break
    else:
        values_to_insert = tp.getArcValues("amp", time_values,first_term,ratio)
        arc_size = len(values_to_insert)
        while(arc_index < arc_size+1):
            try:
                for tup_value in listsheet_tup:
                    if not values_to_insert[arc_index][1]:
                        arc_index += 1
                        break
                    else:
                        tup_value[0].value = values_to_insert[arc_index][1][index_list]
                        book_xlsx.save(file_name)
                        arc_index += 1
            except Exception as error:
                log_errs.append(error)
                book_xlsx.close()
                break


def init():
    try:
        insert_opt = input("Deseja inicializar com as config. padrões (S/N)?\t")
        if(insert_opt.lower() == 'n'):
            insert_test_interval = input("Insira o intervalo para preencher o num's de ensaios: \t")
            insert_fname = input("Insira o nome do arquivo. (Ex.:FP.xlsx):\t ")
            insert_interval = input("Escreva o intervalo que deseja preencher na tabela (Ex.:E13:E54):\t")
            insert_type = input("Insira o tipo que quer preencher na tabela (ELET ou AMP):\t")
            insert_time = int(input("Insira o valor de tempo que deseja (Em segundos):\t"))
            insert_fst_term = int(input("Insira o primeiro valor de tempo da tabela:\t"))
            insert_rat = int(input("Insira a razao dos tempos:\t"))
            insert_indl = int(input("Insira o indice de lista que deseja retornar(Para amp->(0,1); para elet->(0,2)):\t"))
            print("Registrando num's de ensaio no arquivo... Aguarde")
            write_tests(insert_fname, insert_test_interval)
            print("Registrando dados no arquivo... Aguarde")
            write_FP(insert_interval,insert_type,insert_time,insert_fst_term,insert_rat,insert_fname,insert_indl)
            print("Sucesso!")
        else:
            print("Registrando num's de ensaio no arquivo... Aguarde")
            write_tests("FP.xlsx")
            print("Registrando dados no arquivo... Aguarde")
            write_FP("I13:I54","elet",30,15,5,"FP.xlsx",2)
            write_FP("M13:M54","amp",30,15,5,"FP.xlsx")
            print("Sucesso!")
    except Exception as error:
        log_errs.append(error)
        genLogErrors()
        print("Falha! Consultar log de erros!")
        

    return 0


init()

