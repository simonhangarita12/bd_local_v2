import requests
import pandas as pd
from datetime import datetime as dt
from datetime import timedelta
import gspread
from gspread_dataframe import set_with_dataframe
from google.oauth2.service_account import Credentials

# Ingresamos la informacion a Google Sheets

spreadsheet_id = "19homL4Jy986NCcy8I4QLI4www6TfKuM0JEqF1Yl7cyg"

worksheet_name = "Ventas"

scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
credentials = Credentials.from_service_account_file(
   r"D:\\Users\\siste\\bd_local\\sheets-ensayo-cc8278c28d36.json",
   scopes=scopes
)


#df = pd.read_excel(excel_file)


gc = gspread.authorize(credentials)
sh = gc.open_by_key(spreadsheet_id)
worksheet = sh.worksheet(worksheet_name)


all_values = worksheet.get_all_values()


df = pd.DataFrame(all_values[1:], columns=all_values[0])
df_respaldo=df.copy()
#Pasamos todas las columnas de df a texto para evitar errores
df_respaldo["FECHA"]=df_respaldo.apply(lambda x:str(x["FECHA"]),axis=1)
df_respaldo["VENDEDOR"]=df_respaldo.apply(lambda x:str(x["VENDEDOR"]),axis=1)
df_respaldo["CLIENTE"]=df_respaldo.apply(lambda x:str(x["CLIENTE"]),axis=1)
df_respaldo["VENTA"]=df_respaldo.apply(lambda x:int(x["VENTA"]) if x["VENTA"]!="" else 0,axis=1)
df_respaldo["PRODUCTO COMPLETO"]=df_respaldo.apply(lambda x:str(x["PRODUCTO COMPLETO"]),axis=1)
df_respaldo.to_excel("ventas_reporte.xlsx",index=False)


