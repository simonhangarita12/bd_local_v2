import pandas as pd
import plotly.express as px

df_total=pd.read_excel('Retiros.xlsx')
enero=df_total[df_total["MES"]=="Enero"]
febrero=df_total[df_total["MES"]=="Febrero"]
marzo=df_total[df_total["MES"]=="Marzo"]
abril=df_total[df_total["MES"]=="Abril"]
mayo=df_total[df_total["MES"]=="Mayo"]
junio=df_total[df_total["MES"]=="Junio"]

fig_enero = px.bar(enero,
              x='Analista',
              y='Valor', 
              #log_y=True,
              barmode='group',
              color='Cliente',
              title='Empresas que se han retirado por analista Enero 2025',
              hover_data=['Cliente'])

fig_enero.update_layout(
    hoverlabel=dict(
        bgcolor="white",
        font_size=10,
        font_family="Arial",
        namelength=-1,  
        align="left"    
    )
)
fig_enero.update_layout(
    xaxis=dict(
        tickangle=45,  
        tickfont=dict(size=10)  
    )
)
fig_enero.update_traces(
    hovertemplate=(
        "<b>Analista:</b> %{x}<br>" + 
        "<b>Valor servicio:</b> %{y}<br>" +
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<extra></extra>"
    )
)

fig_febrero = px.bar(febrero,
              x='Analista',
              y='Valor', 
              #log_y=True,
              barmode='group',
              color='Cliente',
              title='Empresas que se han retirado por analista Febrero 2025',
              hover_data=['Cliente'])

fig_febrero.update_layout(
    hoverlabel=dict(
        bgcolor="white",
        font_size=10,
        font_family="Arial",
        namelength=-1,  
        align="left"    
    )
)
fig_febrero.update_layout(
    xaxis=dict(
        tickangle=45,  
        tickfont=dict(size=10)  
    )
)
fig_febrero.update_traces(
    hovertemplate=(
        "<b>Analista:</b> %{x}<br>" + 
        "<b>Valor servicio:</b> %{y}<br>" +
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<extra></extra>"
    )
)

fig_marzo = px.bar(marzo,
              x='Analista',
              y='Valor', 
              #log_y=True,
              barmode='group',
              color='Cliente',
              title='Empresas que se han retirado por analista Marzo 2025',
              hover_data=['Cliente'])

fig_marzo.update_layout(
    hoverlabel=dict(
        bgcolor="white",
        font_size=10,
        font_family="Arial",
        namelength=-1,  
        align="left"    
    )
)
fig_marzo.update_layout(
    xaxis=dict(
        tickangle=45,  
        tickfont=dict(size=10)  
    )
)
fig_marzo.update_traces(
    hovertemplate=(
        "<b>Analista:</b> %{x}<br>" + 
        "<b>Valor servicio:</b> %{y}<br>" +
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<extra></extra>"
    )
)

fig_abril = px.bar(abril,
              x='Analista',
              y='Valor', 
              #log_y=True,
              barmode='group',
              color='Cliente',
              title='Empresas que se han retirado por analista Abril 2025',
              hover_data=['Cliente'])

fig_abril.update_layout(
    hoverlabel=dict(
        bgcolor="white",
        font_size=10,
        font_family="Arial",
        namelength=-1,  
        align="left"    
    )
)
fig_abril.update_layout(
    xaxis=dict(
        tickangle=45,  
        tickfont=dict(size=10)  
    )
)
fig_abril.update_traces(
    hovertemplate=(
        "<b>Analista:</b> %{x}<br>" + 
        "<b>Valor servicio:</b> %{y}<br>" +
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<extra></extra>"
    )
)

fig_mayo = px.bar(mayo,
              x='Analista',
              y='Valor', 
              #log_y=True,
              barmode='group',
              color='Cliente',
              title='Empresas que se han retirado por analista Mayo 2025',
              hover_data=['Cliente'])

fig_mayo.update_layout(
    hoverlabel=dict(
        bgcolor="white",
        font_size=10,
        font_family="Arial",
        namelength=-1,  
        align="left"    
    )
)
fig_mayo.update_layout(
    xaxis=dict(
        tickangle=45,  
        tickfont=dict(size=10)  
    )
)
fig_mayo.update_traces(
    hovertemplate=(
        "<b>Analista:</b> %{x}<br>" + 
        "<b>Valor servicio:</b> %{y}<br>" +
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<extra></extra>"
    )
)

fig_junio = px.bar(junio,
              x='Analista',
              y='Valor', 
              #log_y=True,
              barmode='group',
              color='Cliente',
              title='Empresas que se han retirado por analista Junio 2025',
              hover_data=['Cliente'])

fig_junio.update_layout(
    hoverlabel=dict(
        bgcolor="white",
        font_size=10,
        font_family="Arial",
        namelength=-1,  
        align="left"    
    )
)
fig_junio.update_layout(
    xaxis=dict(
        tickangle=45,  
        tickfont=dict(size=10)  
    )
)
fig_junio.update_traces(
    hovertemplate=(
        "<b>Analista:</b> %{x}<br>" + 
        "<b>Valor servicio:</b> %{y}<br>" +
        "<b>Cliente:</b> %{customdata[0]}<br>" +
        "<extra></extra>"
    )
)

list_fig=[fig_enero,fig_febrero,fig_marzo,fig_abril,fig_mayo,fig_junio]
with open('isa_plot.html', 'w') as f:
    f.write('<html><head><title>Retiros</title></head><body>')
    f.write('<h1 style="text-align:center">Suspensiones y Retiros 2025</h1>')
    f.write('<div style="display:flex; flex-direction:column; justify-content:center">')
    
    for i, fig in enumerate(list_fig):
        f.write(f'<div style="width:80%; margin:10px; border:1px solid #ddd; padding:10px">')
        f.write(fig.to_html(full_html=True, include_plotlyjs='cdn' if i==0 else False))
        f.write('</div>')
    
    f.write('</div></body></html>')