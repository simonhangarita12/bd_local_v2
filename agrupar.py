import pandas as pd

df = pd.read_excel("1. CRM_TU EMPRESA SEGURA 16_06_2025.xlsx", sheet_name="CRM ANALISTAS")

#print(df.loc[0])
si_conexiones=df[df["ASISTENCIA"]=="SI"]

no_conexiones=df[(df["ASISTENCIA"]=="NO")|(df["ASISTENCIA"]=="NO CONEXIÓN")]
conexion_grabada=df[df["ASISTENCIA"]=="CONEXIÓN AUTÓNOMA DE ANALISTA"]

si_conexiones["SI"]=1
no_conexiones["NO"]=1
conexion_grabada["NO CONEXION, PERO GRABADA"]=1

si_grupo=si_conexiones.groupby(["Nombre de la empresa"])["SI"].sum().reset_index()
no_grupo=no_conexiones.groupby(["Nombre de la empresa"])["NO"].sum().reset_index()
conexion_grabada_grupo=conexion_grabada.groupby(["Nombre de la empresa"])["NO CONEXION, PERO GRABADA"].sum().reset_index()

total_grupo=pd.concat([si_grupo,no_grupo,conexion_grabada_grupo])
total_grupo = total_grupo.groupby('Nombre de la empresa')[['SI','NO','NO CONEXION, PERO GRABADA']].sum().reset_index()
total_grupo.to_excel("total_grupo.xlsx",index=False)