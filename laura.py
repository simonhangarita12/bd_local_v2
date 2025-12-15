import pandas as pd
import plotly.express as px
from datetime import datetime as dt

informacion_conexiones=pd.read_excel(r"D:\Users\siste\bd_local\TuEmpresaSegura.xlsx")
informacion_conexiones=informacion_conexiones.rename(columns={"hora de ingreso":"Fecha de conexión","tiempo conectado":"Tiempo de conexión asesor",
                                                              "tiempo conectado asistentes":"Tiempo de conexión asistentes","tiempo muerto inevitable":"Tiempo de retraso"})
informacion_conexiones["Fecha de conexión"]=informacion_conexiones.apply(lambda x: pd.to_datetime(str(x["Fecha de conexión"])),axis=1)
informacion_conexiones["Tiempo de conexión asesor"]=informacion_conexiones.apply(lambda x:pd.to_timedelta(str(x["Tiempo de conexión asesor"])),axis=1)
informacion_conexiones["Tiempo de conexión asistentes"]=informacion_conexiones.apply(lambda x:pd.to_timedelta(str(x["Tiempo de conexión asistentes"])) ,axis=1)
informacion_conexiones["Tiempo de retraso"]=informacion_conexiones.apply(lambda x:pd.to_timedelta(str(x["Tiempo de retraso"])),axis=1)
lista_empresas=list(informacion_conexiones["Empresa"].unique())
informacion_conexiones=informacion_conexiones.sort_values(["Fecha de conexión"])
#Pasamos el tiempo de conexión a formato float para que pueda graficarse con facilidad
informacion_conexiones["Tiempo de conexión del asesor"]=informacion_conexiones['Tiempo de conexión asesor'].dt.total_seconds() // 60 + (informacion_conexiones['Tiempo de conexión asesor'].dt.total_seconds() % 60)/100
informacion_conexiones["Tiempo de conexión de los asistentes"]=informacion_conexiones['Tiempo de conexión asistentes'].dt.total_seconds() // 60 + (informacion_conexiones['Tiempo de conexión asistentes'].dt.total_seconds() % 60)/100
informacion_conexiones["Tiempo de retraso"]=informacion_conexiones['Tiempo de retraso'].dt.total_seconds() // 60 + (informacion_conexiones['Tiempo de retraso'].dt.total_seconds() % 60)/100

fig1=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[0]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[0]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig1.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig2=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[1]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[1]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig2.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig3=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[2]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[2]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig3.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig4=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[3]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[3]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig4.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig5=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[4]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[4]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig5.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig6=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[5]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[5]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig6.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig7=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[6]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[6]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig7.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig8=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[7]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[7]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig8.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig9=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[8]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[8]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig9.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig10=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[9]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[9]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig10.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig11=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[10]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[10]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig11.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig12=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[11]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[11]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig12.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig13=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[12]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[12]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig13.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig14=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[13]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[13]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig14.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig15=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[14]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[14]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig15.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig16=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[15]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[15]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig16.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig17=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[16]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[16]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig17.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig18=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[17]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[17]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig18.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig19=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[18]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[18]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig19.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig20=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[19]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[19]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig20.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig21=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[20]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[20]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig21.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig22=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[21]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[21]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig22.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig23=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[22]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[22]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig23.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig24=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[23]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[23]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig24.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig25=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[24]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[24]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig25.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig26=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[25]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[25]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig26.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig27=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[26]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[26]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig27.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig28=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[27]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[27]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig28.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig29=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[28]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[28]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig29.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig30=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[29]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[29]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig30.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig31=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[30]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[30]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig31.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig32=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[31]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[31]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig32.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig33=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[32]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[32]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig33.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig34=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[33]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[33]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig34.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig35=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[34]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[34]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig35.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig36=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[35]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[35]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig36.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig37=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[36]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[36]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig37.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig38=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[37]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[37]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig38.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig39=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[38]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[38]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig39.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig40=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[39]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[39]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig40.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig41=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[40]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[40]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig41.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig42=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[41]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[41]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig42.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig43=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[42]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[42]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig43.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig44=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[43]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[43]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig44.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
fig45=px.line(informacion_conexiones[informacion_conexiones["Empresa"]==lista_empresas[44]],
             x="Fecha de conexión",
             y="Tiempo de conexión del asesor",
             title=f'Conexiones de {lista_empresas[44]}',
             hover_data=['Empresa', 'Tiempo de conexión de los asistentes','Nombre','Tiempo de retraso'],)
fig45.update_traces(
    hovertemplate=(
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Asesor:</b> %{customdata[2]}<br>" +
        "<b>Horario de conexión:</b> %{x}<br>" + 
        "<b>Tiempo de conexión del asesor (minutos):</b> %{y}<br>"+
        "<b>Tiempo de conexión de los asistentes (minutos) :</b> %{customdata[1]}<br>" +
        "<b>Tiempo de retraso del cliente (minutos):</b> %{customdata[3]}<br>" +
        "<extra></extra>"
    )
)
list_fig=[fig1,fig2,fig3,fig4,fig5,fig6,fig7,fig8,fig9,fig10,
          fig11,fig12,fig13,fig14,fig15,fig16,fig17,fig18,fig19,fig20,
          fig21,fig22,fig23,fig24,fig25,fig26,fig27,fig28,fig29,fig30,
          fig31,fig32,fig33,fig34,fig35,fig36,fig37,fig38,fig39,fig40,
          fig41,fig42,fig43,fig44,fig45]

with open('laura.html', 'w') as f:
    f.write('<html><head><title>Conexiones Comfama</title></head><body>')
    f.write('<h1 style="text-align:center">Registros de reuniones Tu empresa segura Comfama</h1>')
    f.write('<div style="display:flex; flex-wrap:wrap; justify-content:center">')

    for i, fig in enumerate(list_fig):
        f.write(f'<div style="width:45%; margin:10px; border:1px solid #ddd; padding:10px">')
        f.write(fig.to_html(full_html=True, include_plotlyjs="cdn" if i==0 else False))
        f.write('</div>')

    f.write('</div></body></html>')
