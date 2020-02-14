import pandas as pd
import urllib.request as ul
import numpy as np
import shutil
import mechanicalsoup
from win32com.client import Dispatch
from GoogleDrive import authentication, update_file

SCOPES = ['https://www.googleapis.com/auth/drive']

def main():
    excel = "https://dclientes.telefonicachile.cl/saip/app/ejecutivo/export?filter%5B_sort_order%5D=ASC&filter%5B_sort_by%5D=nombres&filter%5B_page%5D=0&filter%5B_per_page%5D=25&filter%5BestadoContractual%5D%5Bvalue%5D=1&format=xls"
    url = mechanicalsoup.StatefulBrowser()
    url.open("https://dclientes.telefonicachile.cl/saip/login")
    url.select_form('form[action="/saip/login_check"]')
    url["_username"] = "135484725"
    url["_password"] = "135484725"
    url.submit_selected()
    url.follow_link("app/ejecutivo/list")
    resp = url.session.get(excel)
    resp.raise_for_status()
    """with open("plantel_activo.xls", "wb") as outf:
        outf.write(resp.content)
    xl = Dispatch('Excel.Application')
    wb = xl.Workbooks.Open("C:/Users/Usuario2080/Desktop/Saip/plantel_activo.xls")
    xl.DisplayAlerts = False
    wb.Save()
    xl.Quit()"""
    data = pd.read_excel("plantel_activo.xls")
    tabla = ["Nombre", "Rut", "Comuna", "Empresa", "Gerencia", "Subgerencia", "Sucursal", "Cargo", "Funcion", "Tipo Empresa",
    "Estado Contractual", "Tipo Contrato", "Estado Ejecutivo"]
    for n in data:
        if n not in tabla:
            data.drop(n, axis = 1, inplace = True)
    for n in data:
        if n == "Rut":
            p=0
            for i in data[n]:
                data.Rut[p] = i.replace(".","")
                p += 1
            p=0
            for i in data[n]:
                data.Rut[p] = i.replace("-","")
                p += 1
    data.to_csv("saip.csv", index = False)
    update_file(authentication(), "saip.csv", "text/csv", "11Kp8dBxbr2-E5SKeEVsdzJLMPXkDh3BA")
main()