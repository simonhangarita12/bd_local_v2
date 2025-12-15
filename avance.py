import pandas as pd
from datetime import datetime as dt
from datetime import timedelta
import plotly.express as px
from plotly.subplots import make_subplots
import plotly.graph_objects as go
data_avance=pd.read_excel("result cia sgsst1.xlsx")
data_avance=data_avance.rename(columns={"razonsocial":"Cliente",
                                        "analista":"Analista",
                                        "fecha_inicio":"Fecha de inicio",
                                        "ultimo_ciclo":"Ciclo actual"})
data_avance["Nombre analista"]=data_avance["Analista"]+" "+data_avance["last_name"]
data_avance=data_avance.drop(columns=["last_name","Analista"])
data_avance["Fecha de inicio"]=data_avance.apply(lambda x:pd.to_datetime(x["Fecha de inicio"]),axis=1)
data_avance["Fecha"]=data_avance.apply(lambda x:dt.strftime(x["Fecha de inicio"],"%Y-%m-%d"),axis=1)
data_avance["Fecha"]=data_avance.apply(lambda x:pd.to_datetime(x["Fecha"]),axis=1)
fecha_entrega= dt.combine(dt.today(), dt.min.time())-timedelta(days=1)
data_avance["Dias transcurridos"]=data_avance.apply(lambda x:(fecha_entrega-x["Fecha"]).days,axis=1)
data_avance=data_avance.drop(columns=["Fecha"])
data_avance=data_avance.dropna()
data_avance=data_avance[data_avance["Nombre analista"]!="LORENZO CHIMENO"]
data_avance["Nombre analista"]=data_avance.apply(lambda x:x["Nombre analista"].upper(),axis=1)
#Vamos a crear una variable para definir el tamaño de los burbubujas en el gráfico de burbujas
data_avance["Size"]=data_avance["Dias transcurridos"]/data_avance["Ciclo actual"]
#Creamos otra variable para tener el numero de reuniones
data_avance["Numero de empresas en plataforma"]=1
group_analistas_ciclo=data_avance.groupby(["Nombre analista"])["Ciclo actual"].mean().reset_index()
group_analistas_numero=data_avance.groupby(["Nombre analista"])["Numero de empresas en plataforma"].sum().reset_index()
data_avance=data_avance.drop(columns="Numero de empresas en plataforma")
group_analistas=pd.merge(group_analistas_ciclo,group_analistas_numero,on="Nombre analista")
group_analistas["Numero de empresas en modulo 3"] = group_analistas.apply(
    lambda x: len(data_avance.query(
        '`Nombre analista` == @x["Nombre analista"] and `Ciclo actual` == 3'
    )),
    axis=1
)
group_analistas["Numero de empresas en modulo 2"] = group_analistas.apply(
    lambda x: len(data_avance.query(
        '`Nombre analista` == @x["Nombre analista"] and `Ciclo actual` == 2'
    )),
    axis=1
)
group_analistas["Numero de empresas en modulo 1"]= group_analistas.apply(
    lambda x: len(data_avance.query(
        '`Nombre analista` == @x["Nombre analista"] and `Ciclo actual` == 1'
    )),
    axis=1
)
fig_barras=px.bar(group_analistas,
                  x="Nombre analista",
                  y="Ciclo actual",
                  title="Promedio de avance en los ciclos para los analistas",
                  hover_data=["Numero de empresas en plataforma"]
                  )
                

fig_barras.update_layout(
    hoverlabel=dict(
        bgcolor="white",
        font_size=10,
        font_family="Arial",
        namelength=-1,  
        align="left"    
    )
)
fig_barras.update_layout(
    xaxis=dict(
        tickangle=45,  
        tickfont=dict(size=10)  
    )
)
fig_barras.update_traces(
    hovertemplate=(
        "<b>Analista:</b> %{x}<br>" + 
        "<b>Promedio de avance en los módulos:</b> %{y}<br>" +
        "<b>Número de empresas en plataforma:</b> %{customdata[0]}<br>" +
        "<extra></extra>"
    )
)

lista_analistas=list(data_avance["Nombre analista"].unique())

for i in range(1,len(lista_analistas)+1):
    var_name = f"item_counts{i}"
    persona_df = data_avance[data_avance['Nombre analista'] == lista_analistas[i-1]]
    globals()[var_name] = persona_df['Ciclo actual'].value_counts().reset_index()
    globals()[f"item_counts{i}"].columns=['Item', 'Count']



fig1=px.pie(item_counts1,
            values="Count",
            names="Item",
            title=f'Distribución del avance en los módulos de {lista_analistas[0]}',
            labels={'Item': 'Módulo', 'Count': 'Cantidad'} )
fig2=px.pie(item_counts2,
            values="Count",
            names="Item",
            title=f'Distribución del avance en los módulos de {lista_analistas[1]}',
            labels={'Item': 'Módulo', 'Count': 'Cantidad'} )
fig3=px.pie(item_counts3,
            values="Count",
            names="Item",
            title=f'Distribución del avance en los módulos de {lista_analistas[2]}',
            labels={'Item': 'Módulo', 'Count': 'Cantidad'} )
fig4=px.pie(item_counts4,
            values="Count",
            names="Item",
            title=f'Distribución del avance en los módulos de {lista_analistas[3]}',
            labels={'Item': 'Módulo', 'Count': 'Cantidad'} )
fig5=px.pie(item_counts5,
            values="Count",
            names="Item",
            title=f'Distribución del avance en los módulos de {lista_analistas[4]}',
            labels={'Item': 'Módulo', 'Count': 'Cantidad'} )
fig6=px.pie(item_counts6,
            values="Count",
            names="Item",
            title=f'Distribución del avance en los módulos de {lista_analistas[5]}',
            labels={'Item': 'Módulo', 'Count': 'Cantidad'} )
fig7=px.pie(item_counts7,
            values="Count",
            names="Item",
            title=f'Distribución del avance en los módulos de {lista_analistas[6]}',
            labels={'Item': 'Módulo', 'Count': 'Cantidad'} )
fig8=px.pie(item_counts8,
            values="Count",
            names="Item",
            title=f'Distribución del avance en los módulos de {lista_analistas[7]}',
            labels={'Item': 'Módulo', 'Count': 'Cantidad'} )
fig9=px.pie(item_counts9,
            values="Count",
            names="Item",
            title=f'Distribución del avance en los módulos de {lista_analistas[8]}',
            labels={'Item': 'Módulo', 'Count': 'Cantidad'} )
fig10=px.pie(item_counts10,
            values="Count",
            names="Item",
            title=f'Distribución del avance en los módulos de {lista_analistas[9]}',
            labels={'Item': 'Módulo', 'Count': 'Cantidad'} )
fig11=px.pie(item_counts11,
            values="Count",
            names="Item",
            title=f'Distribución del avance en los módulos de {lista_analistas[10]}',
            labels={'Item': 'Módulo', 'Count': 'Cantidad'} )
fig12=px.pie(item_counts12,
            values="Count",
            names="Item",
            title=f'Distribución del avance en los módulos de {lista_analistas[11]}',
            labels={'Item': 'Módulo', 'Count': 'Cantidad'} )
fig13=px.pie(item_counts13,
            values="Count",
            names="Item",
            title=f'Distribución del avance en los módulos de {lista_analistas[12]}',
            labels={'Item': 'Módulo', 'Count': 'Cantidad'} )
fig14=px.pie(item_counts14,
            values="Count",
            names="Item",
            title=f'Distribución del avance en los módulos de {lista_analistas[13]}',
            labels={'Item': 'Módulo', 'Count': 'Cantidad'} )
fig15=px.pie(item_counts15,
            values="Count",
            names="Item",
            title=f'Distribución del avance en los módulos de {lista_analistas[14]}',
            labels={'Item': 'Módulo', 'Count': 'Cantidad'} )
fig16=px.pie(item_counts16,
            values="Count",
            names="Item",
            title=f'Distribución del avance en los módulos de {lista_analistas[15]}',
            labels={'Item': 'Módulo', 'Count': 'Cantidad'} )
fig17=px.pie(item_counts17,
            values="Count",
            names="Item",
            title=f'Distribución del avance en los módulos de {lista_analistas[16]}',
            labels={'Item': 'Módulo', 'Count': 'Cantidad'} )
fig18=px.pie(item_counts18,
            values="Count",
            names="Item",
            title=f'Distribución del avance en los módulos de {lista_analistas[17]}',
            labels={'Item': 'Módulo', 'Count': 'Cantidad'} )
fig19=px.pie(item_counts19,
            values="Count",
            names="Item",
            title=f'Distribución del avance en los módulos de {lista_analistas[18]}',
            labels={'Item': 'Módulo', 'Count': 'Cantidad'} )
fig20=px.pie(item_counts20,
            values="Count",
            names="Item",
            title=f'Distribución del avance en los módulos de {lista_analistas[19]}',
            labels={'Item': 'Módulo', 'Count': 'Cantidad'} )
fig21=px.pie(item_counts21,
            values="Count",
            names="Item",
            title=f'Distribución del avance en los módulos de {lista_analistas[20]}',
            labels={'Item': 'Módulo', 'Count': 'Cantidad'} )
fig22=px.pie(item_counts22,
            values="Count",
            names="Item",
            title=f'Distribución del avance en los módulos de {lista_analistas[21]}',
            labels={'Item': 'Módulo', 'Count': 'Cantidad'} )
fig_scatter=px.scatter(data_avance,
                x="Ciclo actual",
                y="Dias transcurridos",
                hover_data=["Cliente","Nombre analista"],
                title="Distribución de los días transcurridos y el avance actual",
                size="Size",
                size_max=40,
            )
fig_scatter.update_traces(
    hovertemplate=(
        "<b>Ciclo actual:</b> %{x}<br>" + 
        "<b>Días transcurridos:</b> %{y}<br>" +
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<b>Analista:</b> %{customdata[1]}<br>" +
        "<extra></extra>"
    )
)
fig_scatter.update_traces(marker=dict(color='orange'))  
#fig_scatter.update_layout(width=1200, height=650)

list_fig=[fig_barras,fig1,fig2,fig3,fig4,fig5,fig6,fig7,fig8,fig9,
          fig10,fig11,fig12,fig13,fig14,fig15,fig16,fig17,fig18,
          fig19,fig20,fig21,fig22,fig_scatter]




with open('plataforma.html', 'w') as f:
    f.write('<html><head><title>Progreso en la plataforma</title></head><body>')
    f.write('<h1 style="text-align:center">Avance en la plataforma por analista</h1>')
    f.write('<div style="display:flex; flex-wrap:wrap; justify-content:center">')
    
    for i, fig in enumerate(list_fig):
        f.write(f'<div style="width:45%; margin:10px; border:1px solid #ddd; padding:10px">')
        f.write(fig.to_html(full_html=True, include_plotlyjs='cdn' if i==0 else False))
        f.write('</div>')
    
    f.write('</div></body></html>')
data_avance=data_avance.sort_values(by="Size",ascending=False)
group_analistas=group_analistas.rename(columns={"Ciclo actual":"Ciclo promedio de avance"})
group_analistas=group_analistas.sort_values(by="Ciclo promedio de avance",ascending=True)
informe=data_avance[["Cliente","Nombre analista","Fecha de inicio","Ciclo actual","Dias transcurridos"]]
"""with pd.ExcelWriter('Informe_plataforma.xlsx', engine='openpyxl') as writer:
    informe.to_excel(writer, sheet_name="PROGRESO POR EMPRESA", index=False)
    group_analistas.to_excel(writer, sheet_name='INFORMACION ANALISTAS', index=False)
"""