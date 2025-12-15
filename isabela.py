import pandas as pd
import plotly.express as px
from plotly.subplots import make_subplots
import plotly.graph_objects as go

# Ruta al archivo
archivo_excel = 'Suspensiones y Retiros 2023.xlsx'





enero_2024 = pd.read_excel(archivo_excel, sheet_name="ENERO 2024")
enero_2024=enero_2024[["VALOR","ANALISTA","EMPRESA","CAUSAL EXPLICACION","MOTIVO"]]
enero_2024["ANALISTA"]=enero_2024.apply(lambda x:str(x["ANALISTA"]).upper().strip(),axis=1)
enero_2024["Empresas Perdidas"]=1
#eliminamos la ultima fila que en este caso realiza una suma de los valores
enero_2024=enero_2024.iloc[:-1, :]

febrero_2024 = pd.read_excel(archivo_excel, sheet_name="FEBRERO 2024")
febrero_2024=febrero_2024[["VALOR","ANALISTA","EMPRESA","CAUSAL EXPLICACION","MOTIVO"]]
febrero_2024["ANALISTA"]=febrero_2024.apply(lambda x:str(x["ANALISTA"]).upper().strip(),axis=1)
febrero_2024["Empresas Perdidas"]=1
#eliminamos la ultima fila que en este caso realiza una suma de los valores
febrero_2024=febrero_2024.iloc[:-1, :]


marzo_2024 = pd.read_excel(archivo_excel, sheet_name="MARZO 2024")
marzo_2024=marzo_2024[["VALOR","ANALISTA","EMPRESA","CAUSAL EXPLICACION","MOTIVO"]]
marzo_2024["ANALISTA"]=marzo_2024.apply(lambda x:str(x["ANALISTA"]).upper().strip(),axis=1)
marzo_2024["Empresas Perdidas"]=1
#eliminamos la ultima fila que en este caso realiza una suma de los valores
marzo_2024=marzo_2024.iloc[:-1, :]

abril_2024 = pd.read_excel(archivo_excel, sheet_name="ABRIL 2024")
abril_2024=abril_2024[["VALOR","ANALISTA","EMPRESA","CAUSAL EXPLICACION","MOTIVO"]]
abril_2024["ANALISTA"]=abril_2024.apply(lambda x:str(x["ANALISTA"]).upper().strip(),axis=1)
abril_2024["Empresas Perdidas"]=1
#eliminamos la ultima fila que en este caso realiza una suma de los valores
abril_2024=abril_2024.iloc[:-1, :]

mayo_2024 = pd.read_excel(archivo_excel, sheet_name="MAYO 2024")
mayo_2024=mayo_2024[["VALOR","ANALISTA","EMPRESA","CAUSAL EXPLICACION","MOTIVO"]]
mayo_2024["ANALISTA"]=mayo_2024.apply(lambda x:str(x["ANALISTA"]).upper().strip(),axis=1)
mayo_2024["Empresas Perdidas"]=1
#eliminamos la ultima fila que en este caso realiza una suma de los valores
mayo_2024=mayo_2024.iloc[:-1, :]

junio_2024 = pd.read_excel(archivo_excel, sheet_name="JUNIO 2024")
junio_2024=junio_2024[["VALOR","ANALISTA","EMPRESA","CAUSAL EXPLICACION","MOTIVO"]]
junio_2024["ANALISTA"]=junio_2024.apply(lambda x:str(x["ANALISTA"]).upper().strip(),axis=1)
junio_2024["Empresas Perdidas"]=1
#eliminamos la ultima fila que en este caso realiza una suma de los valores
junio_2024=junio_2024.iloc[:-1, :]

julio_2024 = pd.read_excel(archivo_excel, sheet_name="JULIO 2024")
julio_2024=julio_2024[["VALOR","ANALISTA","EMPRESA","CAUSAL EXPLICACION","MOTIVO"]]
julio_2024["ANALISTA"]=julio_2024.apply(lambda x:str(x["ANALISTA"]).upper().strip(),axis=1)
julio_2024["Empresas Perdidas"]=1
#eliminamos la ultima fila que en este caso realiza una suma de los valores
julio_2024=julio_2024.iloc[:-1, :]

agosto_2024 = pd.read_excel(archivo_excel, sheet_name="AGOSTO 2024")
agosto_2024=agosto_2024[["VALOR","ANALISTA","EMPRESA","CAUSAL EXPLICACION","MOTIVO"]]
agosto_2024["ANALISTA"]=agosto_2024.apply(lambda x:str(x["ANALISTA"]).upper().strip(),axis=1)
agosto_2024["Empresas Perdidas"]=1
#eliminamos la ultima fila que en este caso realiza una suma de los valores
agosto_2024=agosto_2024.iloc[:-1, :]

septiembre_2024 = pd.read_excel(archivo_excel, sheet_name="SEPTIEMBRE 2024")
septiembre_2024=septiembre_2024[["VALOR","ANALISTA","EMPRESA","CAUSAL EXPLICACION","MOTIVO"]]
septiembre_2024["ANALISTA"]=septiembre_2024.apply(lambda x:str(x["ANALISTA"]).upper().strip(),axis=1)
septiembre_2024["Empresas Perdidas"]=1
#eliminamos la ultima fila que en este caso realiza una suma de los valores
septiembre_2024=septiembre_2024.iloc[:-1, :]

octubre_2024 = pd.read_excel(archivo_excel, sheet_name="OCTUBRE 2024")
octubre_2024=octubre_2024[["VALOR","ANALISTA","EMPRESA","CAUSAL EXPLICACION","MOTIVO"]]
octubre_2024["ANALISTA"]=octubre_2024.apply(lambda x:str(x["ANALISTA"]).upper().strip(),axis=1)
octubre_2024["Empresas Perdidas"]=1
#eliminamos la ultima fila que en este caso realiza una suma de los valores
octubre_2024=octubre_2024.iloc[:-1, :]

noviembre_2024 = pd.read_excel(archivo_excel, sheet_name="NOVIEMBRE 2024")
noviembre_2024=noviembre_2024[["VALOR","ANALISTA","EMPRESA","CAUSAL EXPLICACION","MOTIVO"]]
noviembre_2024["ANALISTA"]=noviembre_2024.apply(lambda x:str(x["ANALISTA"]).upper().strip(),axis=1)
noviembre_2024["Empresas Perdidas"]=1
#eliminamos la ultima fila que en este caso realiza una suma de los valores
noviembre_2024=noviembre_2024.iloc[:-1, :]


diciembre_2024 = pd.read_excel(archivo_excel, sheet_name="DICIEMBRE 2024")
diciembre_2024=diciembre_2024[["VALOR","ANALISTA","EMPRESA","CAUSAL EXPLICACION","MOTIVO"]]
diciembre_2024["ANALISTA"]=diciembre_2024.apply(lambda x:str(x["ANALISTA"]).upper().strip(),axis=1)
diciembre_2024["Empresas Perdidas"]=1
diciembre_2024=diciembre_2024.dropna()

enero_2025 = pd.read_excel(archivo_excel, sheet_name="ENERO 2025")
enero_2025=enero_2025[["VALOR","ANALISTA","EMPRESA","CAUSAL EXPLICACION","MOTIVO"]]
enero_2025["ANALISTA"]=enero_2025.apply(lambda x:str(x["ANALISTA"]).upper().strip(),axis=1)
enero_2025["Empresas Perdidas"]=1
enero_2025=enero_2025.dropna()

febrero_2025 = pd.read_excel(archivo_excel, sheet_name="FEBRERO 2025")
febrero_2025=febrero_2025[["VALOR","ANALISTA","EMPRESA","CAUSAL EXPLICACION","MOTIVO"]]
febrero_2025["ANALISTA"]=febrero_2025.apply(lambda x:str(x["ANALISTA"]).upper().strip(),axis=1)
febrero_2025["Empresas Perdidas"]=1
febrero_2025=febrero_2025.dropna()

marzo_2025 = pd.read_excel(archivo_excel, sheet_name="MARZO 2025")
marzo_2025=marzo_2025[["VALOR","ANALISTA","EMPRESA","CAUSAL EXPLICACION","MOTIVO"]]
marzo_2025["ANALISTA"]=marzo_2025.apply(lambda x:str(x["ANALISTA"]).upper().strip(),axis=1)
marzo_2025["Empresas Perdidas"]=1
marzo_2025=marzo_2025.dropna()

abril_2025 = pd.read_excel(archivo_excel, sheet_name="ABRIL 2025")
abril_2025=abril_2025[["VALOR","ANALISTA","EMPRESA","CAUSAL EXPLICACION","MOTIVO"]]
abril_2025["ANALISTA"]=abril_2025.apply(lambda x:str(x["ANALISTA"]).upper().strip(),axis=1)
abril_2025["Empresas Perdidas"]=1
abril_2025=abril_2025.dropna()

dic_2024={"Enero":enero_2024["VALOR"].sum(),
          "Febrero":febrero_2024["VALOR"].sum(),
          "Marzo":marzo_2024["VALOR"].sum(),
          "Abril":abril_2024["VALOR"].sum(),
          "Mayo":mayo_2024["VALOR"].sum(),
          "Junio":junio_2024["VALOR"].sum(),
          "Julio":julio_2024["VALOR"].sum(),
          "Agosto":agosto_2024["VALOR"].sum(),
          "Septiembre":septiembre_2024["VALOR"].sum(),
          "Octubre":octubre_2024["VALOR"].sum(),
          "Noviembre":noviembre_2024["VALOR"].sum(),
          "Diciembre":diciembre_2024["VALOR"].sum()
          }

dic_2025={"Enero":enero_2025["VALOR"].sum(),
          "Febrero":febrero_2025["VALOR"].sum(),
          "Marzo":marzo_2025["VALOR"].sum(),
          "Abril":abril_2025["VALOR"].sum()
          }

acumulado=pd.concat([enero_2024,febrero_2024,marzo_2024,
                     abril_2024,mayo_2024,junio_2024,
                     julio_2024,agosto_2024,septiembre_2024,
                     octubre_2024,noviembre_2024,diciembre_2024,
                     enero_2025,febrero_2025,marzo_2025,abril_2025])
acumulado_2024=pd.concat([enero_2024,febrero_2024,marzo_2024,
                     abril_2024,mayo_2024,junio_2024,
                     julio_2024,agosto_2024,septiembre_2024,
                     octubre_2024,noviembre_2024,diciembre_2024])
acumulado_2025=pd.concat([enero_2025,febrero_2025,marzo_2025,abril_2025])
acumulado_group=acumulado.groupby(["MOTIVO"])["Empresas Perdidas"].sum().reset_index()

fig1 = px.bar(enero_2024,
              x='ANALISTA',
              y='Empresas Perdidas', 
              title='Empresas que se han retirado por analista Enero 2024',
              hover_data=['EMPRESA', 'CAUSAL EXPLICACION',"VALOR"])

fig1.update_layout(
    hoverlabel=dict(
        bgcolor="white",
        font_size=10,
        font_family="Arial",
        namelength=-1,  
        align="left"    
    )
)
fig1.update_layout(
    xaxis=dict(
        tickangle=45,  
        tickfont=dict(size=10)  
    )
)
fig1.update_traces(
    hovertemplate=(
        "<b>Analista:</b> %{x}<br>" + 
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Causal:</b> %{customdata[1]}<br>" +
        "<b>Valor servicio:</b> %{customdata[2]}<br>" +
        "<extra></extra>"
    )
)
fig2 = px.bar(febrero_2024,
              x='ANALISTA',
              y='Empresas Perdidas', 
              title='Empresas que se han retirado por analista Febrero 2024',
              hover_data=['EMPRESA', 'CAUSAL EXPLICACION',"VALOR"])

fig2.update_layout(
    hoverlabel=dict(
        bgcolor="white",
        font_size=10,
        font_family="Arial",
        namelength=-1,  
        align="left"    
    )
)
fig2.update_layout(
    xaxis=dict(
        tickangle=45,  
        tickfont=dict(size=10)  
    )
)
fig2.update_traces(
    hovertemplate=(
        "<b>Analista:</b> %{x}<br>" + 
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Causal:</b> %{customdata[1]}<br>" +
        "<b>Valor servicio:</b> %{customdata[2]}<br>" +
        "<extra></extra>"
    )
)

fig3 = px.bar(marzo_2024,
              x='ANALISTA',
              y='Empresas Perdidas', 
              title='Empresas que se han retirado por analista Marzo 2024',
              hover_data=['EMPRESA', 'CAUSAL EXPLICACION',"VALOR"])

fig3.update_layout(
    hoverlabel=dict(
        bgcolor="white",
        font_size=10,
        font_family="Arial",
        namelength=-1,  
        align="left"    
    )
)
fig3.update_layout(
    xaxis=dict(
        tickangle=45,  
        tickfont=dict(size=10)  
    )
)
fig3.update_traces(
    hovertemplate=(
        "<b>Analista:</b> %{x}<br>" + 
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Causal:</b> %{customdata[1]}<br>" +
        "<b>Valor servicio:</b> %{customdata[2]}<br>" +
        "<extra></extra>"
    )
)

fig4 = px.bar(abril_2024,
              x='ANALISTA',
              y='Empresas Perdidas', 
              title='Empresas que se han retirado por analista Abril 2024',
              hover_data=['EMPRESA', 'CAUSAL EXPLICACION',"VALOR"])

fig4.update_layout(
    hoverlabel=dict(
        bgcolor="white",
        font_size=10,
        font_family="Arial",
        namelength=-1,  
        align="left"    
    )
)
fig4.update_layout(
    xaxis=dict(
        tickangle=45,  
        tickfont=dict(size=10)  
    )
)
fig4.update_traces(
    hovertemplate=(
        "<b>Analista:</b> %{x}<br>" + 
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Causal:</b> %{customdata[1]}<br>" +
        "<b>Valor servicio:</b> %{customdata[2]}<br>" +
        "<extra></extra>"
    )
)


fig5 = px.bar(mayo_2024,
              x='ANALISTA',
              y='Empresas Perdidas', 
              title='Empresas que se han retirado por analista Mayo 2024',
              hover_data=['EMPRESA', 'CAUSAL EXPLICACION',"VALOR"])

fig5.update_layout(
    hoverlabel=dict(
        bgcolor="white",
        font_size=10,
        font_family="Arial",
        namelength=-1,  
        align="left"    
    )
)
fig5.update_layout(
    xaxis=dict(
        tickangle=45,  
        tickfont=dict(size=10)  
    )
)
fig5.update_traces(
   hovertemplate=(
        "<b>Analista:</b> %{x}<br>" + 
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Causal:</b> %{customdata[1]}<br>" +
        "<b>Valor servicio:</b> %{customdata[2]}<br>" +
        "<extra></extra>"
    )
)

fig6 = px.bar(junio_2024,
              x='ANALISTA',
              y='Empresas Perdidas', 
              title='Empresas que se han retirado por analista Junio 2024',
              hover_data=['EMPRESA', 'CAUSAL EXPLICACION',"VALOR"])

fig6.update_layout(
    hoverlabel=dict(
        bgcolor="white",
        font_size=10,
        font_family="Arial",
        namelength=-1,  
        align="left"    
    )
)
fig6.update_layout(
    xaxis=dict(
        tickangle=45,  
        tickfont=dict(size=10)  
    )
)
fig6.update_traces(
    hovertemplate=(
        "<b>Analista:</b> %{x}<br>" + 
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Causal:</b> %{customdata[1]}<br>" +
        "<b>Valor servicio:</b> %{customdata[2]}<br>" +
        "<extra></extra>"
    )
)

fig7 = px.bar(julio_2024,
              x='ANALISTA',
              y='Empresas Perdidas', 
              title='Empresas que se han retirado por analista Julio 2024',
              hover_data=['EMPRESA', 'CAUSAL EXPLICACION',"VALOR"])

fig7.update_layout(
    hoverlabel=dict(
        bgcolor="white",
        font_size=10,
        font_family="Arial",
        namelength=-1,  
        align="left"    
    )
)
fig7.update_layout(
    xaxis=dict(
        tickangle=45,  
        tickfont=dict(size=10)  
    )
)
fig7.update_traces(
    hovertemplate=(
        "<b>Analista:</b> %{x}<br>" + 
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Causal:</b> %{customdata[1]}<br>" +
        "<b>Valor servicio:</b> %{customdata[2]}<br>" +
        "<extra></extra>"
    )
)

fig8 = px.bar(agosto_2024,
              x='ANALISTA',
              y='Empresas Perdidas', 
              title='Empresas que se han retirado por analista Agosto 2024',
              hover_data=['EMPRESA', 'CAUSAL EXPLICACION',"VALOR"])

fig8.update_layout(
    hoverlabel=dict(
        bgcolor="white",
        font_size=10,
        font_family="Arial",
        namelength=-1,  
        align="left"    
    )
)
fig8.update_layout(
    xaxis=dict(
        tickangle=45,  
        tickfont=dict(size=10)  
    )
)
fig8.update_traces(
    hovertemplate=(
        "<b>Analista:</b> %{x}<br>" + 
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Causal:</b> %{customdata[1]}<br>" +
        "<b>Valor servicio:</b> %{customdata[2]}<br>" +
        "<extra></extra>"
    )
)

fig9 = px.bar(septiembre_2024,
              x='ANALISTA',
              y='Empresas Perdidas', 
              title='Empresas que se han retirado por analista Septiembre 2024',
              hover_data=['EMPRESA', 'CAUSAL EXPLICACION',"VALOR"])

fig9.update_layout(
    hoverlabel=dict(
        bgcolor="white",
        font_size=10,
        font_family="Arial",
        namelength=-1,  
        align="left"    
    )
)
fig9.update_layout(
    xaxis=dict(
        tickangle=45,  
        tickfont=dict(size=10)  
    )
)
fig9.update_traces(
    hovertemplate=(
        "<b>Analista:</b> %{x}<br>" + 
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Causal:</b> %{customdata[1]}<br>" +
        "<b>Valor servicio:</b> %{customdata[2]}<br>" +
        "<extra></extra>"
    )
)

fig10 = px.bar(octubre_2024,
              x='ANALISTA',
              y='Empresas Perdidas', 
              title='Empresas que se han retirado por analista Octubre 2024',
              hover_data=['EMPRESA', 'CAUSAL EXPLICACION',"VALOR"])

fig10.update_layout(
    hoverlabel=dict(
        bgcolor="white",
        font_size=10,
        font_family="Arial",
        namelength=-1,  
        align="left"    
    )
)
fig10.update_layout(
    xaxis=dict(
        tickangle=45,  
        tickfont=dict(size=10)  
    )
)
fig10.update_traces(
    hovertemplate=(
        "<b>Analista:</b> %{x}<br>" + 
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Causal:</b> %{customdata[1]}<br>" +
        "<b>Valor servicio:</b> %{customdata[2]}<br>" +
        "<extra></extra>"
    )
)

fig11 = px.bar(noviembre_2024,
              x='ANALISTA',
              y='Empresas Perdidas', 
              title='Empresas que se han retirado por analista Noviembre 2024',
              hover_data=['EMPRESA', 'CAUSAL EXPLICACION',"VALOR"])

fig11.update_layout(
    hoverlabel=dict(
        bgcolor="white",
        font_size=10,
        font_family="Arial",
        namelength=-1,  
        align="left"    
    )
)
fig11.update_layout(
    xaxis=dict(
        tickangle=45,  
        tickfont=dict(size=10)  
    )
)
fig11.update_traces(
    hovertemplate=(
        "<b>Analista:</b> %{x}<br>" + 
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Causal:</b> %{customdata[1]}<br>" +
        "<b>Valor servicio:</b> %{customdata[2]}<br>" +
        "<extra></extra>"
    )
)
fig12 = px.bar(diciembre_2024,
              x='ANALISTA',
              y='Empresas Perdidas', 
              title='Empresas que se han retirado por analista Diciembre 2024',
              hover_data=['EMPRESA', 'CAUSAL EXPLICACION',"VALOR"])

fig12.update_layout(
    hoverlabel=dict(
        bgcolor="white",
        font_size=10,
        font_family="Arial",
        namelength=-1,  
        align="left"    
    )
)
fig12.update_layout(
    xaxis=dict(
        tickangle=45,  
        tickfont=dict(size=10)  
    )
)
fig12.update_traces(
    hovertemplate=(
        "<b>Analista:</b> %{x}<br>" + 
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Causal:</b> %{customdata[1]}<br>" +
        "<b>Valor servicio:</b> %{customdata[2]}<br>" +
        "<extra></extra>"
    )
)
fig13 = px.bar(enero_2025,
              x='ANALISTA',
              y='Empresas Perdidas', 
              title='Empresas que se han retirado por analista Enero 2025',
              hover_data=['EMPRESA', 'CAUSAL EXPLICACION',"VALOR"])

fig13.update_layout(
    hoverlabel=dict(
        bgcolor="white",
        font_size=10,
        font_family="Arial",
        namelength=-1,  
        align="left"    
    )
)
fig13.update_layout(
    xaxis=dict(
        tickangle=45,  
        tickfont=dict(size=10)  
    )
)
fig13.update_traces(
    hovertemplate=(
        "<b>Analista:</b> %{x}<br>" + 
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Causal:</b> %{customdata[1]}<br>" +
        "<b>Valor servicio:</b> %{customdata[2]}<br>" +
        "<extra></extra>"
    )
)
fig14 = px.bar(febrero_2025,
              x='ANALISTA',
              y='Empresas Perdidas', 
              title='Empresas que se han retirado por analista Febrero 2025',
              hover_data=['EMPRESA', 'CAUSAL EXPLICACION',"VALOR"])

fig14.update_layout(
    hoverlabel=dict(
        bgcolor="white",
        font_size=10,
        font_family="Arial",
        namelength=-1,  
        align="left"    
    )
)
fig14.update_layout(
    xaxis=dict(
        tickangle=45,  
        tickfont=dict(size=10)  
    )
)
fig14.update_traces(
    hovertemplate=(
        "<b>Analista:</b> %{x}<br>" + 
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Causal:</b> %{customdata[1]}<br>" +
        "<b>Valor servicio:</b> %{customdata[2]}<br>" +
        "<extra></extra>"
    )
)

fig15 = px.bar(marzo_2025,
              x='ANALISTA',
              y='Empresas Perdidas', 
              title='Empresas que se han retirado por analista Marzo 2025',
              hover_data=['EMPRESA', 'CAUSAL EXPLICACION',"VALOR"])

fig15.update_layout(
    hoverlabel=dict(
        bgcolor="white",
        font_size=10,
        font_family="Arial",
        namelength=-1,  
        align="left"    
    )
)
fig15.update_layout(
    xaxis=dict(
        tickangle=45,  
        tickfont=dict(size=10)  
    )
)
fig15.update_traces(
    hovertemplate=(
        "<b>Analista:</b> %{x}<br>" + 
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Causal:</b> %{customdata[1]}<br>" +
        "<b>Valor servicio:</b> %{customdata[2]}<br>" +
        "<extra></extra>"
    )
)
fig16 = px.bar(abril_2025,
              x='ANALISTA',
              y='Empresas Perdidas', 
              title='Empresas que se han retirado por analista Abril 2025',
              hover_data=['EMPRESA', 'CAUSAL EXPLICACION',"VALOR"])

fig16.update_layout(
    hoverlabel=dict(
        bgcolor="white",
        font_size=10,
        font_family="Arial",
        namelength=-1,  
        align="left"    
    )
)
fig16.update_layout(
    xaxis=dict(
        tickangle=45,  
        tickfont=dict(size=10)  
    )
)
fig16.update_traces(
    hovertemplate=(
        "<b>Analista:</b> %{x}<br>" + 
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Causal:</b> %{customdata[1]}<br>" +
        "<b>Valor servicio:</b> %{customdata[2]}<br>" +
        "<extra></extra>"
    )
)

fig17=px.pie(names=list(dic_2024.keys()),
             values=list(dic_2024.values()),
             title="Dinero perdido por suspensiones y retiros en 2024",
             labels={'names': 'Mes', 'values': 'Cantidad'})


fig18 = px.pie(
    names=list(dic_2025.keys()),
    values=list(dic_2025.values()),
    title="Dinero perdido por suspensiones y retiros en 2025",
    labels={'names': 'Mes', 'values': 'Cantidad'}  
)



fig19 = px.bar(acumulado_group,
              x='MOTIVO',
              y='Empresas Perdidas', 
              title='Motivos de suspensiones o retiros',
             )
fig19.update_layout(
    xaxis=dict(
        visible=False  
    )
)
fig19.update_traces(
    hovertemplate=(
        "<b>Motivo:</b> %{x}<br>" +
        "<b>Cantidad:</b> %{y}<br>" +
        "<extra></extra>"
    )
)
fig20 = px.pie(
    names=list(acumulado_group["MOTIVO"]),
    values=list(acumulado_group["Empresas Perdidas"]),
    title="Cantidad de retiros y suspensiones por motivo",
    labels={'names': 'Motivo de retiro', 'values': 'Cantidad'}  
)
fig20.update_layout(showlegend=False)
fig21 = px.bar(acumulado_2025,
              x='ANALISTA',
              y='Empresas Perdidas', 
              title='Acumulado de retiros y suspensiones por analista 2025',
              hover_data=['EMPRESA', 'CAUSAL EXPLICACION',"VALOR"])

fig21.update_layout(
    hoverlabel=dict(
        bgcolor="white",
        font_size=10,
        font_family="Arial",
        namelength=-1,  
        align="left"    
    )
)
fig21.update_layout(
    xaxis=dict(
        tickangle=45,  
        tickfont=dict(size=10)  
    )
)
fig21.update_traces(
    hovertemplate=(
        "<b>Analista:</b> %{x}<br>" + 
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Causal:</b> %{customdata[1]}<br>" +
        "<b>Valor servicio:</b> %{customdata[2]}<br>" +
        "<extra></extra>"
    )
)

list_fig=[fig1,fig2,fig3,fig4,fig5,fig6,fig7,fig8,fig9,
          fig10,fig11,fig12,fig13,fig14,fig15,fig16,fig17,
          fig18,fig19,fig20,fig21]
with open('multi_plot.html', 'w') as f:
    f.write('<html><head><title>Suspensiones y retiros</title></head><body>')
    f.write('<h1 style="text-align:center">Suspensiones y Retiros</h1>')
    f.write('<div style="display:flex; flex-wrap:wrap; justify-content:center">')
    
    for i, fig in enumerate(list_fig):
        f.write(f'<div style="width:45%; margin:10px; border:1px solid #ddd; padding:10px">')
        f.write(fig.to_html(full_html=True, include_plotlyjs='cdn' if i==0 else False))
        f.write('</div>')
    
    f.write('</div></body></html>')



