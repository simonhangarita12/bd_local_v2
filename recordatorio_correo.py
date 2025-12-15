from win32com.client import Dispatch
import pandas as pd
from datetime import datetime as dt
from datetime import timedelta
import os

analistas_dict = {
    "Javier Jiménez": "analistapesv@talentoconsultores.com.co",
    "Claudia Pachón": "analistasst10@talentoconsultores.com.co",
    "Claudia Higuita": "analistasst11@talentoconsultores.com.co",
    "Gisella Ruz Gómez": "analistasst12@talentoconsultores.com.co",
    "Adriana del Carmen Romero": "analistasst13@talentoconsultores.com.co",
    "Juliana Marín": "analistasst14@talentoconsultores.com.co",
    "Claudia Carvajal": "analistasst15@talentoconsultores.com.co",
    "José Adolfo Agudelo Marín": "analistasst16@talentoconsultores.com.co",
    "Kelly Johana Posada Madrid": "analistasst17@talentoconsultores.com.co",
    "Luz Elena Estrada Perez": "analistasst21@talentoconsultores.com.co",
    "Yasmin Andrea Buitrago": "analistasst23@talentoconsultores.com.co",
    "Yuliana Zapata": "analistasst25@talentoconsultores.com.co",
    "Maria Elena Loaiza Gallego": "analistasst28@talentoconsultores.com.co",
    "David Silva": "analistasst29@talentoconsultores.com.co",
    "Daniela Acevedo": "analistasst32@talentoconsultores.com.co",
    "Daniela Preciado": "analistasst34@talentoconsultores.com.co",
    "Rossana Ávilez": "analistasst37@talentoconsultores.com.co",
    "Maribel Tangarife": "analistasst38@talentoconsultores.com.co",
    "Juan Camilo Urrego": "analistasst4@talentoconsultores.com.co",
    "Jose Alcides Gallego": "analistasst5@talentoconsultores.com.co",
    "Gillary Cortés": "analistasst6@talentoconsultores.com.co",
    "Cristhian Gomez Sanchez": "analistasst7@talentoconsultores.com.co",
    "Antony Linero": "analistasst9@talentoconsultores.com.co",
    "Juan Diego Lopez Rios": "analistasst19@talentoconsultores.com.co",
    "Simón Henao": "automatizacion@talentoconsultores.com.co"
}
recordatorio=pd.read_excel(r"D:\Users\siste\bd_local\Programacion_recordatorios.xlsx")
recordatorio["Fecha recordatorio"]=recordatorio.apply(lambda x:pd.to_datetime(x["Fecha recordatorio"]),axis=1)
for i in range(len(recordatorio)):
    if recordatorio.loc[i,"Se recordo"]=="no":
        nombre_cliente=recordatorio.loc[i,"Nombre_del_cliente"]
        correo_cliente=recordatorio.loc[i,"Correo del cliente"]
        #añadiños como una petición explicita de un cliente el correo de la gerente a los recordatorios
        if nombre_cliente=="COLEGIATURA" or nombre_cliente=="COLEGIATURA COLOMBIANA DE COSMETOLOGIA":
            correo_cliente+=";direccion@colegiaturadecosmetologia.edu.co;administrativo@colegiaturadecosmetologia.edu.co"
        #eliminamos correos en caso de ser necesario también como una petición expresa del cliente
        clientes_eliminar_correos={"FJM INVERSIONES S.A.S":["juliana.sandino@farmu.com.co"],"MIO GROUP AMERICA S.A.S":["lina.garcia@migroup.com","lgarcia@mio.on"],
                                   "SISTEGA GEO OIL ENERGY":["d.navarrete@geoilenergy.com"],"GEO OIL ENERGY":["d.navarrete@geoilenergy.com"],
                                   "TRANSPORTE SEGURO ESPECIALIZADO":["transporteseguroyespecializado@gmail.com"],"TRANSPORTE SEGURO ESPECIALIZADO S.A.S":["transporteseguroyespecializado@gmail.com"],
                                   "INTELLIGENT ELECTRONIC SOLUTIONS S.A.S":["psicologasandrahenao@gmail.com"]
                                   }
        if nombre_cliente in clientes_eliminar_correos:
            lista_correos=correo_cliente.split(";")
            texto_correo=""
            contador=0
            for correo in lista_correos:
                if correo not in clientes_eliminar_correos[nombre_cliente]:
                    contador+=1
                    texto_correo+=correo+";"
                else:
                    continue
            if contador==0:
                correo_cliente=correo_cliente
            elif contador==1:
                correo_cliente=texto_correo
            else:
                correo_cliente=texto_correo[:-1]

        nombre_analista=recordatorio.loc[i,"Nombre del analista"]
        correo_analista=analistas_dict[nombre_analista]
        fecha_recordatorio=str(recordatorio.loc[i,"Fecha recordatorio"])
        fecha_reunion=str(recordatorio.loc[i,"Fecha reunion"])
        fecha_recordatorio=dt.strptime(fecha_recordatorio, "%Y-%m-%d %H:%M:%S")
        fecha_reunion=dt.strptime(fecha_reunion, "%Y-%m-%d %H:%M:%S")
        formato=""
        if fecha_reunion.hour>=12:
            formato="pm"
            fecha_reunion_formateada=fecha_reunion - timedelta(hours=12)
        else:
            formato="am"
            fecha_reunion_formateada=fecha_reunion
        hora_reunion=fecha_reunion_formateada.strftime("%H:%M")
        dia_recordatorio=fecha_recordatorio.strftime("%Y-%m-%d")
        #fecha_reunion=fecha_recordatorio+ timedelta(days=1)
        dia_reunion=fecha_reunion_formateada.strftime("%Y-%m-%d")
        lista=dia_reunion.split("-")
        dic_meses={"01":"Enero","02":"Febrero","03":"Marzo","04":"Abril","05":"Mayo","06":"Junio",
                "07":"Julio","08":"Agosto","09":"Septiembre","10":"Octubre","11":"Noviembre","12":"Diciembre"}
        dic_dias={"01":"1","02":"2","03":"3","04":"4","05":"5","06":"6",
                "07":"7","08":"8","09":"9","10":"10","11":"11","12":"12",
                "13":"13","14":"14","15":"15","16":"16","17":"17","18":"18",
                "19":"19","20":"20","21":"21","22":"22","23":"23","24":"24",
                "25":"25","26":"26","27":"27","28":"28","29":"29","30":"30","31":"31"}
        texto = f"""
<b>Cordial saludo</b>, equipo de <b>{nombre_cliente}</b>:
📅 Le recordamos que tiene agendada una <b>reunión de asesoría</b> con nuestro equipo, según los siguientes detalles:<br>
&nbsp;&nbsp;&nbsp;<b>• Fecha:</b> {dic_dias[lista[2]]} de {dic_meses[lista[1]]} del {lista[0]}
&nbsp;&nbsp;&nbsp;<b>• Hora:</b> {hora_reunion} {formato}
&nbsp;&nbsp;&nbsp;<b>• Analista asignado:</b> {nombre_analista}
&nbsp;&nbsp;&nbsp;<b>• Objetivo:</b> Apoyar en:
"""
        if dt.now()>=fecha_recordatorio and dt.now()<=fecha_reunion:
            outlook = Dispatch("Outlook.Application")
            draft = outlook.CreateItem(0)
            draft.BodyFormat = 2 
            draft.Subject =  f"Recordatorio asesoria Sistegra - {dic_dias[lista[2]]} de {dic_meses[lista[1]]}, {hora_reunion} {formato}"
            #draft.Body = texto
            
            image_dir = r"D:\Users\siste\bd_local"
            image_cid = "MyEmbeddedImage"
            if correo_analista=="analistapesv@talentoconsultores.com.co":
                image_file = "asesoria pesv_Mesa de trabajo 1.png"
                image_url="https://sistegrasst-my.sharepoint.com/personal/automatizacion_talentoconsultores_com_co/_layouts/15/download.aspx?share=EY0OY4Nz2DhPi7PyXmAHwS0BFnzYlnR3jWTmqpemsP6xaQ"
                texto+="""&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>1)</b> Cumplimiento normativo del PESV - Sistema Seguro
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>2)</b> Fortalecimiento de la cultura de seguridad vial en la organización
                Atentamente, <b>Equipo Talento Consultores 🤝</b>"""
            else:
                image_file = "asesoria sst_Mesa de trabajo 1.png"
                image_url="https://sistegrasst-my.sharepoint.com/personal/automatizacion_talentoconsultores_com_co/_layouts/15/download.aspx?share=ER9cr0tP_jFEgKMd3D9q7-oBkWBSPJYaUx4YuN0KUvyS_Q"
                texto+="""&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>1)</b> Cumplimiento normativo vigente
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>2)</b> Mejora de condiciones laborales
                Atentamente, <b>Equipo Talento Consultores 🤝</b>"""
            #vamos a añadir el link de la reunión en caso de que esté disponible
            link=recordatorio.loc[i,"Link reunión"]
            if not pd.isna(link):
                if link.endswith(">"):
                    link=link[:-1]
                texto += f'<br><b>Link de la reunión:</b> <a href="{link.strip()}" style="color: blue; text-decoration: underline;">Haga clic aquí para unirse</a>'
                
            image_path = os.path.join(image_dir, image_file)
            """attachment = draft.Attachments.Add(image_path)
            attachment.PropertyAccessor.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                f"<{image_cid}>"
            )"""
            texto_html = texto.replace("\n", "<br>")
            html_body = f"""
            <html>
            <body>
            {texto_html}
            <br><br>
            <img src="{image_url}" width="500">
            </body>
            </html>
            """


            draft.HTMLBody = html_body
            draft.To = correo_cliente
            #draft.CC = correo_analista
            draft.Send()
            recordatorio.loc[i,"Se recordo"]="si"
  

print("Recordatorios creados")
#with open(r'C:\Users\PCC\bd_local\logs.txt', 'a', encoding='utf-8') as file:
#    file.write(str(dt.now()))
recordatorio.to_excel(r"D:\Users\siste\bd_local\Programacion_recordatorios.xlsx",index=False)

