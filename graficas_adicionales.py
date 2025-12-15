import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime as dt
import dateparser
informacion=pd.read_excel("GRAFICOS MENSUALES LAURA.XLSX")
#quitamos las filas que no tienen empresa
informacion=informacion[informacion["Nombre de la empresa"].isna()==False]
informacion=informacion.reset_index(drop=True)
empresas=informacion["Nombre de la empresa"].unique()
informacion["ASISTENCIA"]=informacion.apply(lambda x: str(x["ASISTENCIA"]).upper(), axis=1)
informacion["ASISTENCIA"]=informacion.apply(lambda x: "NO CONEXIÓN" if x["ASISTENCIA"]=="NO" else x["ASISTENCIA"], axis=1)
informacion=informacion.rename(columns={"PORCENTAJE AVANCE ":"PORCENTAJE AVANCE"})
def safe_datetime_conversion(date_str):
    try:
        return pd.to_datetime(date_str)
    except:
        return date_str

informacion["Conexión"] = informacion["Conexión"].apply(safe_datetime_conversion)
#Transformamos las fechas que falten por transformar
def parse_any_date(date_str):
    if pd.isna(date_str):
        return pd.NaT
    return dateparser.parse(
        str(date_str), 
        languages=['es', 'en'],
        settings={'DATE_ORDER': 'DMY'} 
    )
informacion["Conexión"] = informacion.apply(lambda x: parse_any_date(x["Conexión"]) if isinstance(x["Conexión"], str) else x["Conexión"], axis=1)
mes_dict={1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
          7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}
informacion["Mes"]=informacion.apply(lambda x: mes_dict[x["Conexión"].month] if x["Conexión"].month in mes_dict else "Mes no especificado", axis=1)
informacion["PORCENTAJE AVANCE"]=informacion["PORCENTAJE AVANCE"].fillna(0)
informacion["Conexiones"]=informacion.apply(lambda x: 1 if x["ASISTENCIA"]=="SI" else 0, axis=1)
informacion["Conexiones autónomas"]=informacion.apply(lambda x: 1 if x["ASISTENCIA"]=="CONEXIÓN AUTÓNOMA DE ANALISTA" else 0, axis=1)
informacion["No conexiones"]=informacion.apply(lambda x: 1 if x["ASISTENCIA"]=="NO CONEXIÓN" else 0, axis=1)
#informacion.to_excel("borrar_urgente.xlsx", index=False)
numeric_columns = ["PORCENTAJE AVANCE", "Conexiones", "Conexiones autónomas", "No conexiones"]
for col in numeric_columns:
    informacion[col] = pd.to_numeric(informacion[col], errors='coerce')
informacion.reset_index(drop=True, inplace=True)
def eliminar_repetidos(df):
    """Eliminamos las conexiones repetidas mirando la fecha de conexión y el nombre de la empresa, y la idea es eliminar la repetición de esta conexión"""
    dic_conexiones={empresa:[[],[]] for empresa in empresas}
    
    for i in range(len(df)):
        fecha_conexion = df.loc[i, "Conexión"]
        empresa = df.loc[i, "Nombre de la empresa"]
        asistencia = df.loc[i, "ASISTENCIA"]
        if fecha_conexion in dic_conexiones[empresa][0] and fecha_conexion is not pd.NaT:
            indice = dic_conexiones[empresa][0].index(fecha_conexion)
            if not pd.isna(dic_conexiones[empresa][1][indice]):
                df.loc[i,"Conexiones"]=0
                df.loc[i,"Conexiones autónomas"]=0
                df.loc[i,"No conexiones"]=0
            else:
                dic_conexiones[empresa][1][indice] = asistencia
            if dic_conexiones[empresa][1][indice] != asistencia and dic_conexiones[empresa][1][indice]!="":
                print(f"Revisión necesaria para {empresa} en la fecha {fecha_conexion}: asistencia anterior: {dic_conexiones[empresa][1][indice]}, asistencia nuevo: {asistencia}")
            

        else:
            dic_conexiones[empresa][0].append(fecha_conexion)
            dic_conexiones[empresa][1].append(asistencia)


    return df
informacion=eliminar_repetidos(informacion)

grupo_acumulado=informacion.groupby(["Nombre de la empresa"])[["PORCENTAJE AVANCE","Conexiones", "Conexiones autónomas", "No conexiones"]].sum().reset_index()
grupo_acumulado_reset = grupo_acumulado.reset_index(drop=True)
#vamos a añadir la informacion de todos los meses conjunta
for empresa in empresas:
    fila_nueva=["TU EMPRESA SEGURA","INDEFINIDO",empresa,"nit",dt(2025,1,1),"Hora","NO CONEXIÓN",
                "PLANEAR",float(grupo_acumulado[grupo_acumulado["Nombre de la empresa"]==empresa]["PORCENTAJE AVANCE"]),
                "Todos",float(grupo_acumulado[grupo_acumulado["Nombre de la empresa"]==empresa]["Conexiones"]),
                float(grupo_acumulado[grupo_acumulado["Nombre de la empresa"]==empresa]["Conexiones autónomas"]),
                float(grupo_acumulado[grupo_acumulado["Nombre de la empresa"]==empresa]["No conexiones"])]
    columns = ['Programa', 'Analista', 'Nombre de la empresa', 'NIT', 'Conexión','Hora','ASISTENCIA',
               'CICLO','PORCENTAJE AVANCE','Mes','Conexiones','Conexiones autónomas','No conexiones']
    nueva_fila_df = pd.DataFrame([fila_nueva], columns=columns)
    informacion = pd.concat([informacion, nueva_fila_df], ignore_index=True)


informacion_grupo=informacion.groupby(["Nombre de la empresa","Mes"])[["PORCENTAJE AVANCE", "Conexiones"]].sum().reset_index()
conexiones_grupo=informacion.groupby(["Nombre de la empresa","Mes"])[["Conexiones", "Conexiones autónomas", "No conexiones"]].sum().reset_index()
informacion_melt = informacion_grupo.melt(
    id_vars="Nombre de la empresa",
    value_vars=["Conexiones", "PORCENTAJE AVANCE"],
    var_name="Variable",
    value_name="Valor"
)
conexiones_melt= conexiones_grupo.melt(
    id_vars="Nombre de la empresa",
    value_vars=["Conexiones", "Conexiones autónomas","No conexiones"],
    var_name="Variable",
    value_name="Valor"
)
fig = go.Figure()

months = informacion_grupo["Mes"].unique()

for month in months:
    df_month = informacion_grupo[informacion_grupo["Mes"] == month]
    
    fig.add_trace(go.Bar(
        x=df_month["Nombre de la empresa"], y=df_month["Conexiones"],
        name=f"Conexiones - {month}", marker_color="steelblue", visible=(month==months[0]),
        offsetgroup="Conexiones",   
        legendgroup="Conexiones",
        text=df_month["Conexiones"],
        texttemplate="<span style='font-size:20px;color:black;'>%{text}</span>",  # force fixed font size
        textposition="outside",
        cliponaxis=False
        
    ))
    
    fig.add_trace(go.Bar(
        x=df_month["Nombre de la empresa"], y=df_month["PORCENTAJE AVANCE"],
        name=f"Porcentaje Avance - {month}", marker_color="tomato", visible=(month==months[0]),
        offsetgroup="Porcentaje Avance",  
        legendgroup="Porcentaje Avance",
        text=df_month["PORCENTAJE AVANCE"],
        texttemplate="<span style='font-size:20px;color:black;'>%{text}</span>",  # force fixed font size
        textposition="outside",
        cliponaxis=False
        
    ))


buttons = []
for i, month in enumerate(months):
    visible = [False] * (2 * len(months)) 
    visible[2*i] = True
    visible[2*i+1] = True
    
    buttons.append(
        dict(label=month,
             method="update",
             args=[{"visible": visible},
                   {"title": f"Datos para {month}", "barmode": "group"}
                   ])
    )



fig.update_traces(textposition="outside",textfont=dict(size=14, color='black'))
fig.update_layout(
    updatemenus=[dict(
        buttons=buttons,
        direction="down",
        showactive=True
    )],
    barmode="group",
    height=1200,
    width=3000,
    title="Conexiones y Porcentaje de Avance por Empresa",
    xaxis=dict(
        tickangle=-45,  
        tickmode='auto',
        tickfont=dict(size=10),  
        automargin=True  
    )
)



fig_2 = go.Figure()


for month in months:
    df_month_conexiones = conexiones_grupo[conexiones_grupo["Mes"] == month]
    
    fig_2.add_trace(go.Bar(
        x=df_month_conexiones["Nombre de la empresa"], y=df_month_conexiones["Conexiones"],
        name=f"Conexiones - {month}", marker_color="steelblue", visible=(month==months[0]),
        offsetgroup="Conexiones",   
        legendgroup="Conexiones",
        text=df_month_conexiones["Conexiones"],
    ))
    
    fig_2.add_trace(go.Bar(
        x=df_month_conexiones["Nombre de la empresa"], y=df_month_conexiones["Conexiones autónomas"],
        name=f"Conexiones autónomas - {month}", marker_color="tomato", visible=(month==months[0]),
        offsetgroup="Conexiones autónomas",  
        legendgroup="Conexiones autónomas",
        text=df_month_conexiones["Conexiones autónomas"],
    ))
    fig_2.add_trace(go.Bar(
        x=df_month_conexiones["Nombre de la empresa"], y=df_month_conexiones["No conexiones"],
        name=f"No conexiones - {month}", marker_color="green", visible=(month==months[0]),
        offsetgroup="No conexiones",  
        legendgroup="No conexiones",
        text=df_month_conexiones["No conexiones"],
    ))


buttons_2 = []
for i, month in enumerate(months):
    visible = [False] * (3 * len(months)) 
    visible[3*i] = True
    visible[3*i+1] = True
    visible[3*i+2] = True

    buttons_2.append(
        dict(label=month,
             method="update",
             args=[{"visible": visible},
                   {"title": f"Datos para {month}", "barmode": "group"}
                   ])
    )



fig_2.update_traces(textposition="outside",textfont=dict(size=14, color='black'))
fig_2.update_layout(
    updatemenus=[dict(
        buttons=buttons_2,
        direction="down",
        showactive=True
    )],
    barmode="group",
    height=1200,
    width=3000,
    xaxis=dict(
        tickangle=-45,  
        tickmode='auto',
        tickfont=dict(size=10),  
        automargin=True  
    ) 
)

figures= [fig, fig_2]

with open('adicionales_plot.html', 'w') as f:
    f.write('<html><head><title>Avance</title></head><body>')
    f.write('<h1 style="text-align:center">Graficas conexiones vs porcentaje de avance</h1>')
    f.write('<div style="display:flex; flex-direction:column; justify-content:center">')
    for i, fig in enumerate(figures):
        f.write(f'<div style="width:80%; margin:10px; border:1px solid #ddd; padding:10px">')
        f.write(fig.to_html(full_html=True, include_plotlyjs='cdn' if i==0 else False))
        f.write('</div>')
    f.write('</div></body></html>')