import pandas as pd
import unicodedata
consolidado_andres=pd.read_excel(r"C:\Users\PCC\bd_local\MAESTRO_EMPRESAS_2025.xlsx")
consolidado_plataforma=pd.read_excel(r"C:\Users\PCC\bd_local\DB_Empresas.xlsx")
#Eliminamos las empresas cuyo Nombre no aparezca
consolidado_plataforma = consolidado_plataforma.dropna(subset=['Empresa'])
consolidado_plataforma=consolidado_plataforma.reset_index()

#Eliminamos el index del consolidado de la plataforma
#Eliminamos los espacios en blanco en el campo de Nit
consolidado_andres["Nit"]=consolidado_andres.apply(lambda x:str(x["Nit"]).replace(" ",""),axis=1)
#Con el objetivo de unificar nombres de las empresas y luego hacer la unión sin problemas, 
# vamos a colocar los nombres de las empresas en mayusculas para ambos dataframes
consolidado_andres["Razón Social"]=consolidado_andres.apply(lambda x:x["Razón Social"].upper(),axis=1)
consolidado_plataforma["Empresa"]=consolidado_plataforma.apply(lambda x:x["Empresa"].upper(),axis=1)
consolidado_andres["Ciudad"]=consolidado_andres.apply(lambda x:str(x["Ciudad"]).upper(),axis=1)
consolidado_plataforma["Ciudad"]=consolidado_plataforma.apply(lambda x:str(x["Ciudad"]).upper(),axis=1)
consolidado_andres["Departamento"]=consolidado_andres.apply(lambda x:str(x["Departamento"]).upper(),axis=1)
#Creamos una funcion para eliminar ciertos caracteres conflictivos en el número de contacto
def caracteres_conflictivos(numero:str):
    numero=numero.replace("/","")
    numero=numero.replace("(","")
    numero=numero.replace(")","")
    numero=numero.replace(".","")
    numero=numero.replace("-","")
    numero=numero.replace(" ","")
    return numero
consolidado_andres["Teléfono Contacto SST"]=consolidado_andres.apply(lambda x:caracteres_conflictivos(str(x["Teléfono Contacto SST"])),axis=1)

#Vamos a eliminar tildes de los nombres de la ciudad y de las empresas para poder hacer unificación de nombres
def quitar_tildes(texto):
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
consolidado_andres["Razón Social"]=consolidado_andres.apply(lambda x:quitar_tildes(x["Razón Social"]),axis=1)
consolidado_plataforma["Empresa"]=consolidado_plataforma.apply(lambda x:quitar_tildes(x["Empresa"]),axis=1)
consolidado_andres["Ciudad"]=consolidado_andres.apply(lambda x:quitar_tildes(x["Ciudad"]),axis=1)
consolidado_andres["Departamento"]=consolidado_andres.apply(lambda x:quitar_tildes(x["Departamento"]),axis=1)
consolidado_plataforma["Ciudad"]=consolidado_plataforma.apply(lambda x:quitar_tildes(x["Ciudad"]),axis=1)

#Como última limpieza para unificar los nombres quitamos la parte de D.C. cuando se utilice la ciudad de 
# Bógota
def quitar_cd(ciudad):
    ciudad=ciudad.replace("D.C.","")
    ciudad=ciudad.replace("D.C","")
#Renombramos el nombre de la columna Empresa a Razón Social con el objetivo de tener la llave en común con 
# la que podamos realizar la unión
consolidado_plataforma=consolidado_plataforma.rename(columns={"Empresa":"Razón Social"})
#esta función la usamos para verificar que columnas que queremos cruzar no tengan valores diferentes 
# en las dos bases de datos
def revisar_diferentes(empresa,valor1,valor2):
    if pd.isna(valor1) or pd.isna(valor2):
        pass
    elif valor1==valor2:
        pass
    elif valor1!=valor2:
        print(empresa,valor1,valor2)
#Está función la usaremos posteriormente a la anterior que creamos. Para verificar el caso donde una de las 
# dos es vacia y en este caso devolver la otra en las columnas que estamos cruzando la información
def revisar_vacios_andres(empresa,valor1,valor2):
    if pd.isna(valor1) and pd.isna(valor2):
        pass
    elif not pd.isna(valor1) and not pd.isna(valor2):
        pass
    elif not pd.isna(valor1) and pd.isna(valor2):
        pass
    elif pd.isna(valor1) and not pd.isna(valor2):
        print(empresa,valor1,valor2)

consolidado=pd.merge(consolidado_andres,consolidado_plataforma,on='Razón Social', how='outer').reset_index()
consolidado.to_excel("prueba.xlsx",index=False)
for i in range(len(consolidado)):
    empresa=consolidado["Razón Social"].loc[i]
    ciudad1=str(consolidado["Valor contrato_x"].loc[i])
    ciudad2=str(consolidado["Valor contrato_y"].loc[i])
    revisar_diferentes(empresa,ciudad1,ciudad2)