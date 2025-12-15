import gspread
from gspread_dataframe import set_with_dataframe
from google.oauth2.service_account import Credentials
import pandas as pd
import mysql.connector as mariadb
import parametros as par
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
gc = gspread.authorize(credentials)
sh = gc.open_by_key(spreadsheet_id)
worksheet = sh.worksheet(worksheet_name)


all_values = worksheet.get_all_values()


df = pd.DataFrame(all_values[1:], columns=all_values[0])
print(df.head())
df = df.iloc[::-1].reset_index(drop=True)
df["FECHA"]=df["FECHA"].astype(str)
df["VENTA"]=df["VENTA"].replace("", "0")
df["VENTA"]=df["VENTA"].astype(int)
df["VALOR PROVEEDOR"]=df["VALOR PROVEEDOR"].replace("", "0")
df["VALOR PROVEEDOR"]=df["VALOR PROVEEDOR"].astype(int)
df["VALOR COMISIONABLE"]=df["VALOR COMISIONABLE"].replace("", "0")
df["VALOR COMISIONABLE"]=df["VALOR COMISIONABLE"].astype(int)
df=df.drop(columns=["VALOR COMISIONABLE"])
#definimos funciones para formatear el servicio y la factura a un entero
def formatear_servicio(servicio):
    if servicio =="Recurrencia":
        return 1
    elif servicio =="Anualidad":
        return 2
    elif servicio =="Otros productos":
        return 3
    elif servicio=="Anualidad RTE":
        return 4
    else:
        return 0
def formatear_factura(factura):
    if factura =="FACTURADO":
        return 1
    else:
        return 0
df["SERVICIO"]=df.apply(lambda x:x["SERVICIO"].strip(),axis=1
                        )
df["SERVICIO"]=df["SERVICIO"].apply(formatear_servicio)
df["FACTURA"]=df["FACTURA"].apply(formatear_factura)


try:
    # Establecemos la conexión a MariaDB
    mydb = mariadb.connect(
        host = par.host,
        user = par.usuario,
        password = par.password,
        database = par.bd,
        port = par.puerto,
    )

    if mydb.is_connected():
        cursor = mydb.cursor()
        for row in df.to_dict('records'):
            sql = "INSERT IGNORE INTO VENTAS (fecha, vendedor, cliente, venta,producto_completo,servicio,valor_proveedor,factura) VALUES (%(FECHA)s, %(VENDEDOR)s, %(CLIENTE)s, %(VENTA)s, %(PRODUCTO COMPLETO)s,%(SERVICIO)s,%(VALOR PROVEEDOR)s,%(FACTURA)s)"
            cursor.execute(sql, row)

        mydb.commit()
        

except mariadb.Error as err:
    print(f"Error al conectar o insertar datos: {err}")

finally:
    #Cierra la conexión
    if 'mydb' in locals() and mydb.is_connected():
        cursor.close()
        mydb.close()
        print("Conexión a MariaDB cerrada.")