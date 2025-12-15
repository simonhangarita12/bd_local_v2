import win32com.client
import pandas as pd
from datetime import timedelta, datetime as dt
import re
import unicodedata
import holidays

def get_shared_calendar_events(shared_name):
    print("Analista:", shared_name)
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    recipient = namespace.CreateRecipient(shared_name)
    recipient.Resolve()

    if not recipient.Resolved:
        print("No se pudo resolver el nombre del usuario:", shared_name)
        return pd.DataFrame()

    calendar_folder = namespace.GetSharedDefaultFolder(recipient, 9)  

    items = calendar_folder.Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")
   
    


    
    eventos = []
    start_date = dt.now() -timedelta(days=60)
    end_date = dt.now() + timedelta(days=3)


    

    for item in items:
        try:
            start = pd.to_datetime(str(item.Start)).tz_localize(None)
            end = pd.to_datetime(str(item.End)).tz_localize(None)
        
            if end>end_date:
                break
            if start < start_date or end > end_date:
                continue
    
            subject = item.Subject 
            attendees_adresses = []

            teams_link = None
            try:
                body = item.Body
                match = re.search(r"https://teams\.microsoft\.com/l/meetup-join/[^\s]+", body)
                if match:
                    teams_link = match.group(0)
            except Exception as e:
                print("No se pudo leer el cuerpo:", e)

            for i in range(item.Recipients.Count):
                try:
                    recipient = item.Recipients[i]
                    address_entry = recipient.AddressEntry
                    smtp_address = ""


                    if address_entry.Type == "EX":  
                        exch_user = address_entry.GetExchangeUser()
                        if exch_user is not None:
                            smtp_address = exch_user.PrimarySmtpAddress
                        else:
                            smtp_address = address_entry.Address
                         
                    else:
                        smtp_address = address_entry.Address
                        

                    pattern = r'\b[A-Za-z0-9._%+-]+@(sistegra\.com|talentoconsultores\.com\.co)\b'
                    if smtp_address and not re.match(pattern=pattern, string=smtp_address):
                        attendees_adresses.append(smtp_address)
                        
                except Exception as e:
                    print(f"Error retrieving attendee at index {i}:", e)
            

            eventos.append([subject, start, end,attendees_adresses,teams_link
                            ])
        except Exception as e:
            print("Error analista:", shared_name)
            print("Error:", e)

    eventos_df = pd.DataFrame(eventos, columns=["Subject", "Start", "End","Attendees_adresses","link"])
    return eventos_df
diccionario_analistas={
                      
                       
                       "David Silva":"analistasst29@talentoconsultores.com.co",
                       "José Adolfo Agudelo Marín":"analistasst16@talentoconsultores.com.co",
                       "Luz Elena Estrada Perez":"analistasst21@talentoconsultores.com.co",
                       "Jose Alcides Gallego":"analistasst5@talentoconsultores.com.co",
                       "Claudia Pachón":"analistasst10@talentoconsultores.com.co",
                       "Kelly Johana Posada Madrid":"analistasst17@talentoconsultores.com.co",
                       "Gisella Ruz Gómez":"analistasst12@talentoconsultores.com.co",
                       "Cristhian Gomez Sanchez":"analistasst7@talentoconsultores.com.co",
                       "Juan Diego Lopez Rios":"analistasst19@talentoconsultores.com.co",
                       "Maria Elena Loaiza Gallego":"analistasst28@talentoconsultores.com.co",
                       "Claudia Carvajal":"analistasst15@talentoconsultores.com.co",
                       "Yuliana Zapata":"analistasst25@talentoconsultores.com.co",
                       "Gillary Cortés":"analistasst6@talentoconsultores.com.co",
                       "Adriana del Carmen Romero":"analistasst13@talentoconsultores.com.co",
                       "Maribel Tangarife":"analistasst38@talentoconsultores.com.co",
                       "Juliana Marín":"analistasst14@talentoconsultores.com.co",
                       "Yasmin Andrea Buitrago":"analistasst23@talentoconsultores.com.co",
                       "Javier Jiménez":"analistapesv@talentoconsultores.com.co",
                       "Antony Linero":"analistasst9@talentoconsultores.com.co",
                       "Claudia Higuita":"analistasst11@talentoconsultores.com.co",
                       "Daniela Acevedo":"analistasst32@talentoconsultores.com.co"
                       
                       
                       
                       }
dataframes=[]

for analista, correo in diccionario_analistas.items():
    df = get_shared_calendar_events(correo)
    df['Analista'] = analista
    dataframes.append(df)


df_final = pd.concat(dataframes, ignore_index=True)
df_final["Subject"]=df_final.apply(lambda x:str(x["Subject"]).upper(),axis=1)
#df_final.to_excel("revision.xlsx", index=False)


diccionario_cambios_nombres={"(":"",")":"","_":"","[":"","]":"","{":"","}":"",":":"",
                           "/":"","!":"","¡":"",
                            "EN PERSONA":"","PRESENCIAL":"","REUNION SG-SST":"","REUNION SEMANAL":"",
                            "REUNION SGSST":"","REUNION SG SST":"","REUNION SEMANAL":"","ASESORIA A-":"",
                            "ASESORIA A":"","ASESORIA EN SEGURIDAD Y SALUD EN EL TRABAJO -":"",
                            "ASESORIA EN SEGURIDAD Y SALUD EN EL TRABAJO-":"",
                            "ASESORIA EN SEGURIDAD Y SALUD EN EL TRABAJO":"",
                            "REUNION SEMANAL":"","ASESORIA COMFAMA":"","ASESORIA DE SISTEGRA -":"",
                            "ASESORIA SISTEGRA -":"","ASESORIA SISTEGRA":"","ASESORIA DE SISTEGRA":"",
                            "TU EMPRESA SEGURA COMFAMA":"","ASESORIA SST":"","ASESORIAS":"","ASESORIA A ":"","EN SEGURIDAD Y SALUD EN EL TRABAJO -":"",
                            "EN SEGURIDAD Y SALUD EN EL TRABAJO-":"","EN SEGURIDAD Y SALUD EN EL TRABAJO":"","TALENTO CONSULTORES":"",
                            "ENTREGA FINAL PLAN CHOQUE -":"","ENTREGA FINAL PLAN CHOQUE-":"","ENTREGA FINAL PLAN CHOQUE":"",
                            "SEGUIMIENTO SST":"","REUNION":"","REUNION ARL":"","REPROGRAMADA -":"","REPROGRAMADA-":"","REPROGRAMADA":"","MATRIZ LEGAL":"",
                            "SITEGRA":'',"2025":"","REPROGRAMACION":"","CAPACITACION GENERAL DE COMFAMA":"","ASASORIA":"",
                            "ASESORIIA":"","ASESORIA PESV SISTEGRA":"","ASESORIA PESV":"","SG- SST":"","SG -SST":"","SG - SST":"",
                           "SGSST":"","SG SST":"","SG-SST":"","SG-SSG":"","SISTEGRA -":"","SISTEGA -":"","SISTEGRA":"","SST -":"","SST-":"","SST":"","SG.":"","PESV":"","PEVS-":"",
                           "STT-":"","COMFAMA":""}

def manejo_enie(texto):
    return texto.replace("Ñ", "NNN")
def quitar_tildes(texto):
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
def recuperar_enie(texto):
    texto=texto.replace("NNN", "Ñ")
    texto = " ".join(texto.split())
    return texto


def eliminar_asesoria_inicio(nombre_empresa):
    """Eliminamos la palabra asesoria unicamente cuando aparezca al inicio del título de la reunión

    Args: nombre_empresa(str) nombre de la empresa a limpiar

    Returns: nombre_empresa(str) nombre de la empresa luego de eliminar la palabra asesoría únicamente si aparece al inicio"""
     
    pattern = rf"^{"ASESORIA"}\b"
    return re.sub(pattern, '', nombre_empresa, count=1, flags=re.IGNORECASE)

def eliminar_asesoria_final(nombre_empresa):
    """Eliminamos la palabra asesoria unicamente cuando aparezca al final del título de la reunión

    Args: nombre_empresa(str) nombre de la empresa a limpiar

    Returns: nombre_empresa(str) nombre de la empresa luego de eliminar la palabra asesoría únicamente si aparece al final"""
     
    pattern = rf"{"ASESORIA"}\b$"
    return re.sub(pattern, '', nombre_empresa, count=1, flags=re.IGNORECASE)
def eliminar_coma_inicio(nombre_empresa):
    """Eliminamos la coma unicamente cuando aparezca al inicio del título de la reunión

    Args: nombre_empresa(str) nombre de la empresa a limpiar

    Returns: nombre_empresa(str) nombre de la empresa luego de eliminar la coma únicamente si aparece al inicio"""
    if not isinstance(nombre_empresa, str):
        return "" if nombre_empresa is None else str(nombre_empresa) 
    pattern = r"^,\s*"  
    if re.match(r"^,", nombre_empresa):
        return re.sub(pattern, '', nombre_empresa, count=1)
    return nombre_empresa
    
def eliminar_guion_inicio(nombre_empresa: str) -> str:
    """Elimina el guion solo si aparece al inicio, junto con espacios opcionales"""
    if not isinstance(nombre_empresa, str):
        return "" if nombre_empresa is None else str(nombre_empresa)
    pattern = r"^-\s*"  
    if re.match(r"^-", str(nombre_empresa)):
        return re.sub(pattern, '', nombre_empresa, count=1)
    return nombre_empresa
def eliminar_punto_inicio(nombre_empresa):
    """Eliminamos el punto unicamente cuando aparezca al inicio del título de la reunión

    Args: nombre_empresa(str) nombre de la empresa a limpiar

    Returns: nombre_empresa(str) nombre de la empresa luego de eliminar el punto únicamente si aparece al inicio"""
    if not isinstance(nombre_empresa, str):
        return "" if nombre_empresa is None else str(nombre_empresa) 
    pattern = r"^\.\s*"  
    if re.match(r"^\.", nombre_empresa): 
        return re.sub(pattern, '', nombre_empresa, count=1)
    return nombre_empresa


def eliminar_palabras(nombre_empresa):

    """Eliminamos palabras que no aportan mucho al nombre de la empresa, o que no 

    permiten la unificación por nombre de empresa

    Args: nombre_empresa(str) nombre de la empresa a limpiar

  

    Returns: nombre_empresa(str) nombre de la empresa luego de realizar los cambios"""

    terminos_a_eliminar=diccionario_cambios_nombres



    nombre_empresa = " ".join(nombre_empresa.split())

    for palabra in terminos_a_eliminar:

       nombre_empresa=nombre_empresa.replace(palabra,terminos_a_eliminar[palabra])

       nombre_empresa = " ".join(nombre_empresa.split())
    #Al final volvemos a hacer split debido a todos los espacios en blanco que se generaron
    nombre_empresa = " ".join(nombre_empresa.split())
    return nombre_empresa
"""def capitalizar(nombre_empresa):
    #Vamos a obtener el nombre de la empresa en mayúsculas, y la idea es capitalizar el nombre y devolverlo más formalmente

    #Args: nombre_empresa(str) nombre de la empresa en mayúsculas

  

    #Returns: nombre_capitalizado(str) nombre de la empresa capitalizado
    lista=nombre_empresa.split()
    nombre_capitalizado=''
    for e in lista:
        if e!='':
            nombre_capitalizado+=str(e).capitalize()+' '
    nombre_capitalizado=nombre_capitalizado[:-1]
    return nombre_capitalizado"""
def quitar_canceladas(nombre_empresa):
    """Filtramos las reuniones que tengan la palabra cancelada para no tenerlas en cuenta

    Args: nombre_empresa(str) nombre de la empresa a limpiar

    Returns: nombre_empresa(str) nombre de la empresa luego de eliminar la palabra cancelada únicamente si aparece al inicio"""
    eliminar=False
    pattern_1 = r"\bcancelada\b"
    pattern_2 = r"\bcancelado\b"
    if re.search(pattern_1, nombre_empresa, flags=re.IGNORECASE) or re.search(pattern_2, nombre_empresa, flags=re.IGNORECASE):
        eliminar=True
    return eliminar
def clean_subject(text):
    """Aplicamos todas las funciones para limpiar el nombre y regresamos el texto final
    Args: text(str) nombre de la empresa antes de aplicarle la limpieza
    Returns: text(str) nombre de la empresa después de aplicarle todos los procesos de limpieza"""
    text = manejo_enie(text)
    text = quitar_tildes(text)
    text = eliminar_palabras(text)
    text = eliminar_asesoria_inicio(text)
    text = eliminar_coma_inicio(text)
    text = eliminar_guion_inicio(text)
    text = eliminar_punto_inicio(text)
    text = recuperar_enie(text)
    return text
df_final["Subject"] = df_final["Subject"].fillna("").apply(clean_subject)
df_final["Cancelada"] = df_final["Subject"].apply(quitar_canceladas)
df_final=df_final[df_final["Cancelada"]==False]
df_final=df_final.drop(columns=["Cancelada"])
#eliminamos las reuniones cuyo nombre no es indicativo del cliente
df_final=df_final[df_final["Subject"]!=""]
#eliminamos igualmente las reuniones a las cuales no se invitó a ningún cliente
df_final=df_final[df_final["Attendees_adresses"].map(len)>0]

def seleccionar_fecha_recordatorio(fecha_reunion):
    """Seleccionamos la fecha en la que se enviará el correo recordatorio
        Args: fecha_reunion(datetime) fecha en la que se tiene programada la reunión
        Returns: fecha_recordatorio(datetime) fecha en la que se enviará el recordatorio
    """
    fecha_probable=fecha_reunion - timedelta(days=1)
    co_holidays = holidays.CountryHoliday('CO')
    seleccionada=False
    while not seleccionada:
        if fecha_probable in co_holidays or fecha_probable.weekday()>=5:
            fecha_probable=fecha_probable-timedelta(days=1)

        else:
            seleccionada=True
    return fecha_probable
def recordatorio_inmediato(fecha_reunion):
    """Seleccionamos la fecha de envío del recordatorio en caso de ser un recordatorio cercano a la hora de la reunión
        Args: fecha_reunion(datetime) fecha en la que se tiene programada la reunión
        Returns: fecha_recordatorio(datetime) fecha en la que se enviará el recordatorio inmediato a la hora de la reunión"""
    return fecha_reunion - timedelta(hours=1)

df_final["Mes"]=df_final.apply(lambda x:x["Start"].month,axis=1)
df_final=df_final.drop(columns=["Start","End","link","Attendees_adresses"])
df_final["Numero de reuniones"]=1
pd_grouped= df_final.groupby(["Subject","Mes","Analista"])["Numero de reuniones"].sum().reset_index()

#df_final.to_excel("Revision agenda.xlsx", index=False)

pd_grouped.to_excel("Agenda Clientes.xlsx", index=False)
