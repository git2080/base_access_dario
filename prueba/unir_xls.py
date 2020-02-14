import pandas as pd
import urllib.request as ul
import zipfile as zf
import numpy as np
import shutil
import datetime
from GoogleDrive import authentication, update_file

def datetransform(df):
    try:
        data = df.strftime('%d-%m-%Y')
        return data
    except ValueError:
        return ""

def main():
    url = np.array([
        ["RMOEC17.zip", "http://10.232.73.202/pe/sred/planillas/Stgo/RMOEC17.zip"],
        ["RMOEC18.zip", "http://10.232.73.202/pe/sred/planillas/Stgo/RMOEC18.zip"],
        ["RMOEC19.zip", "http://10.232.73.202/pe/sred/planillas/Stgo/RMOEC19.zip"],
        ["RMOEC20.zip", "http://10.232.73.202/pe/sred/planillas/Stgo/RMOEC20.zip"],
        ["ZNOEC17.zip", "http://10.232.73.202/pe/sred/planillas/Zonas/ZNOEC17.zip"],
        ["ZNOEC18.zip", "http://10.232.73.202/pe/sred/planillas/Zonas/ZNOEC18.zip"],
        ["ZNOEC19.zip", "http://10.232.73.202/pe/sred/planillas/Zonas/ZNOEC19.zip"],
        ["ZNOEC20.zip", "http://10.232.73.202/pe/sred/planillas/Zonas/ZNOEC20.zip"],
        ["ordenes_ATP.XLS", "http://10.232.73.202/pe/atp/planillas/ordenes_ATP.XLS"]])
    datos = pd.DataFrame()
    for i in range (len(url)):
        str_url1 = str(url[i][0])
        str_url2 = str(url[i][1])
        ul.urlretrieve(str_url2, str_url1)
        if str_url1 == "ordenes_ATP.XLS":
            dato_aux = pd.read_excel(str_url1)
            dato_aux.rename(columns=lambda x: x.lower(), inplace=True)
            datos = pd.concat([datos, dato_aux], ignore_index = True)
        else:
            shutil.unpack_archive(str_url1)
            dato_aux = pd.read_excel("oec.xls", sheet_name = "ORDENES")
            dato_aux.rename(columns=lambda x: x.lower(), inplace=True)
            if len(datos) != 0:
                datos = pd.concat([datos, dato_aux], ignore_index = True)
            else:
                datos = pd.concat([dato_aux], ignore_index = True)
    datos.drop(["nameproyecto", "tipotrabajo", "clientefinal", "direccion", "altura", "cantidad", "ppto_ingenieria", "ppto_dibujo", 
    "ppto_canalizacion", "ppto_lineas", "ppto_cables", "ppto_fo", "ppto_reposicion", "ppto_iti", "ppto_total_baremos",
    "arf_ingenieria", "arf_dibujo", "arf_canalizacion", "arf_lineas", "arf_cables", "arf_fo", "arf_reposicion", "arf_iti",
    "arf_totalbaremos", "arf", "calificaciones", "foliosremedy", "zonaextrema"], axis = 1, inplace = True)
    datos[datos.columns[66]] = pd.to_datetime(datos[datos.columns[66]], format= '%d-%m-%Y')
    datos[datos.columns[67]] = pd.to_datetime(datos[datos.columns[67]], format= '%d-%m-%Y')
    datos[datos.columns[68]] = pd.to_datetime(datos[datos.columns[68]], format= '%d-%m-%Y') 
    for n in range(len(datos.columns)):
        if datos[datos.columns[n]].dtype == "datetime64[ns]": 
            datos[datos.columns[n]] = datos[datos.columns[n]].apply(datetransform)
        elif datos.columns[n][:3] == "arf":
            datos[datos.columns[n]] = datos[datos.columns[n]].map(str)
            datos[datos.columns[n]] = datos[datos.columns[n]].str.rstrip(".0")
        elif datos.columns[n][:4] == "ppto":
            datos[datos.columns[n]] = datos[datos.columns[n]].map(str)
            datos[datos.columns[n]] = datos[datos.columns[n]].str.rstrip(".0")
    datos.numinterno = datos.numinterno.astype(str)
    for n in datos:
        if n == "numinterno":
            p=0
            for i in datos[n]:
                datos.numinterno[p] = i.replace(",","-")
                p += 1
    datos.to_csv("OEC.csv", index= False)
    update_file(authentication(), 'OEC.csv', 'text/csv', "1Vc_3fVaVONRMQycE1n9kUTcKPVkbnfXV")
main()