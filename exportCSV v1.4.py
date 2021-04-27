import pandas
import xmltodict as xd
import xml.etree.ElementTree as xmlet
from pandas import DataFrame
from pathlib import Path
import forallpeople as si
import time

def checkExists(data, name='dielectricDischarge'):
    try:
        df_dieletric = data.loc[name][0]
        SIZE_DD = len(df_dieletric['resultDD'])
        return df_dieletric, SIZE_DD
    except:
        return '-1'
def is_not_empty(data):
    if(data.empty == True):
        return 'Sem dados!'
    else:
        return data.to_string()

# Inicializacao dos parametros padroes
print("Aguarde... Conversão em processo")
time.sleep(3)
print("Ajustando paramêtros...")
time.sleep(3)
path = Path('./Modelos/')
test_files = list(path.glob('*.test'))
count_files = len(test_files)
units = si
units.environment('default', top_level=False)
contador = 0
for j in range(0, count_files):
    try:
        contador += 1
        print("\n", contador)
        xmlpsr = xmlet.parse(test_files[j]).getroot()
        dictxml = xd.parse(xmlet.tostring(
            xmlpsr, encoding='ISO-8859-1', method='xml'))
        dfxml = pandas.DataFrame(dictxml)
        df_results = dfxml.loc['results'][0]
        df_dieletric, SIZE_DD = checkExists(dfxml)
        # Funcoes de busca
        def getColumns(name, data=dfxml,): 
             print(data.loc['testNumber'][0])
             return data.loc[name][0]+';'

        def getColumnsByDataOrdDict(index, name_dict, name='result',
                                    data=df_results): 
            return float(data[name][index][name_dict])

        def getSize(data=df_results):
            try:
                return len(data['result'])
            except TypeError:
                return 0
        SIZE_R = getSize()

        # Modelos de dataset
        data_model = [
            {
                "Infos de Modelo": "",
                "Número de teste;": getColumns("testNumber"),
                "Data e hora;": getColumns("date"),
                "Modelo;": getColumns("model"),
                "Número de Serial;": getColumns("serialNumber"),
                "Firmware;": getColumns("firmware"),
                "Modo;": getColumns("mode"),
                "": ""
            }]

        data_comp = [
            {
                "Infos de Componente": "",
                "DaiTimeA;": getColumns('daiTimeA'),
                "DaiTimeB;": getColumns('daiTimeB'),
                "Dai Resultado;": getColumns("daiResult"),
                "piTimeA;": getColumns("piTimeA"),
                "piTimeB;": getColumns("piTimeB"),
                "piTime Resultado;": getColumns("piResult"),
                "SVT;": getColumns("svt"),
                "DD;": getColumns("dd"),
                "Capacitance;": getColumns("capacitance"),
                "": ""

            }]
        data_eletric = []
        for i in range(0, 100):
            try:
                data_eletric.append({"Infos elet.;": "",
                                     "mm:sscu;": str(getColumnsByDataOrdDict(i, 'seconds')*units.s)+';',
                                     "V;": str(getColumnsByDataOrdDict(i, 'voltage')*units.V)+';',
                                     "Ohm;": str(getColumnsByDataOrdDict(i, 'resistance')*units.Ohm)+';',
                                     "":"",})
            except:
                pass
        data_discharge = []
        for i in range(0, 100):
            try:
                data_discharge.append({"Infos de Disc.;": "",
                                       "mm:ss;": str(getColumnsByDataOrdDict(i, 'seconds')*units.s)+';',
                                       "Amperagem;":(str(getColumnsByDataOrdDict(i, 'voltage')/getColumnsByDataOrdDict(i, 'resistance')*units.A))+';',
                                       })
                #  "Amperagem": str(((getColumnsByDataOrdDict(i, 'voltage')/getColumnsByDataOrdDict(i, 'resistance'))*units.A))+';',
            except:
                pass
        # Conversao e salvamento dos datasets em um csv
        pandas.set_option('display.max_colwidth', 1000)
        df_data = pandas.DataFrame(data_model, index=['']).T
        df_component = pandas.DataFrame(data_comp, index=['']).T
        df_eletric = pandas.DataFrame(
            data_eletric, index=['']*SIZE_R).T
        file = open('./outputs/output'+str(dfxml.loc['testNumber'][0])+'.csv', 'w')
        file.write(is_not_empty(df_data))
        file.write(is_not_empty(df_component))
        file.write(is_not_empty(df_eletric))
        if data_discharge:
            df_dieletric = pandas.DataFrame(data_discharge, index=['']*len(data_discharge)).T.to_string().strip('\t')
            file.write(df_dieletric)
        file.close()
    except IndexError:
        raise("Erro de indice, estouro de pilha.")
else:
    print("Conversão completa!")
