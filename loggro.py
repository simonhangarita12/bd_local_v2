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
   
    def seleccionar_fecha():
        day_inicial=1
        day_final=2
        fecha_probable=dt.now()+timedelta(days=day_final)
        co_holidays = holidays.CountryHoliday('CO')
        seleccionada=False
        #Tenemos un caso por si es el último día de la semana para que tome la información del segundo día de la siguiente
        ultimo_dia=False
        fecha_actual=dt.now()

        
    
        for i in range(1,4):
            if fecha_actual.weekday() not in [2,3,4]:
                ultimo_dia= False
            else:
                if fecha_actual.weekday()==2:
                    fecha_siguiente=fecha_actual+timedelta(days=1)
                    fecha_siguiente_2=fecha_actual+timedelta(days=2)
                    if fecha_siguiente and fecha_siguiente_2 in co_holidays:
                        ultimo_dia=True
                elif fecha_actual.weekday()==3:
                    fecha_siguiente=fecha_actual+timedelta(days=1)
                    if fecha_siguiente in co_holidays:
                        ultimo_dia=True
                else:
                    ultimo_dia=True
        #definimos una parada para que no siga buscando de una semana en adelante
        pare=7
        if ultimo_dia:
            while pare>0:
                pare-=1
                if 0<fecha_probable.weekday()<=5 and fecha_probable not in co_holidays:
                    return day_final,day_final-1
                    
                else:
                    fecha_probable=fecha_probable+timedelta(days=1)
                    if fecha_probable in co_holidays:
                        day_final+=1
                    day_final+=1
                
        while not seleccionada:
            if fecha_probable in co_holidays or fecha_probable.weekday()>=5:
                fecha_probable=fecha_probable+timedelta(days=1)
                day_final+=1
            else:
                seleccionada=True
        return day_final,day_inicial
    day_final,day_inicial=seleccionar_fecha()


    
    eventos = []
    start_date = dt.now() +timedelta(days=day_inicial,hours=10)
    end_date = dt.now() + timedelta(days=day_final,hours=10)


    

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
diccionario_analistas={"David Silva":"analistasst29@talentoconsultores.com.co",
                       "José Adolfo Agudelo Marín":"analistasst16@talentoconsultores.com.co",
                       "Luz Elena Alzate":"analistasst21@talentoconsultores.com.co",
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
                           "STT-":"","COMFAMA":"","SISTEGA":"","TU EMPRESA SEGURA DE":"","TU EMPRESA SEGURA":"","ASESRIA":"","SISEGRA":"","ASEOSRIA":"",
                           "ASSESORIA":""}

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

def eliminar_punto_final(nombre_empresa):
    """
    Elimina el punto únicamente cuando aparece al final del nombre de la empresa.

    Args:
        nombre_empresa (str): Nombre de la empresa a limpiar.

    Returns:
        str: Nombre de la empresa luego de eliminar el punto únicamente si aparece al final.
    """
    if not isinstance(nombre_empresa, str):
        return "" if nombre_empresa is None else str(nombre_empresa)

    pattern = r"\.\s*$"  # dot optionally followed by spaces, at end of string
    return re.sub(pattern, '', nombre_empresa)


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
def cambiar_sas(nombre_empresa):
    """
    Elimina el punto únicamente cuando aparece al final del nombre de la empresa.

    Args:
        nombre_empresa (str): Nombre de la empresa a limpiar.

    Returns:
        str: Nombre de la empresa luego de eliminar el punto únicamente si aparece al final.
    """
    if not isinstance(nombre_empresa, str):
        return "" if nombre_empresa is None else str(nombre_empresa)

    pattern = r"(?<![A-Za-z0-9])SAS(?![A-Za-z0-9.?\-])"  
    return re.sub(pattern, 'S.A.S', nombre_empresa)
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
def reuniones_festivo(fecha_reunion):
    """Marcamos las reuniones que están programadas en un día festivo para posteriormente eliminarlas con un filtro
    Returns: festivo (boolean)"""
    co_holidays = holidays.CountryHoliday('CO')
    fecha=dt.combine(fecha_reunion,dt.min.time())
    festivo=False
    if fecha in co_holidays or fecha.weekday()==6:
        festivo=True
    return festivo
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
    text=eliminar_punto_final(text)
    text=cambiar_sas(text)
    text = recuperar_enie(text)
    return text
df_final["Festivo"]=df_final["End"].apply(reuniones_festivo)
df_final=df_final[df_final["Festivo"]==False]
df_final=df_final.drop(columns=["Festivo"])
df_final["Subject"] = df_final["Subject"].fillna("").apply(clean_subject)
df_final["Cancelada"] = df_final["Subject"].apply(quitar_canceladas)
df_final=df_final[df_final["Cancelada"]==False]
df_final=df_final.drop(columns=["Cancelada"])
df_final["Subject"]=df_final.apply(lambda x:"" if x["Subject"]=="AUDITORIAS" else x["Subject"],axis=1)
#eliminamos las reuniones cuyo nombre no es indicativo del cliente
df_final=df_final[df_final["Subject"]!=""]
#eliminamos igualmente las reuniones a las cuales no se invitó a ningún cliente
df_final=df_final[df_final["Attendees_adresses"].map(len)>0]
#Eliminamos aquellas reuniones que son capacitaciones en lugar de asesorías normales
df_final=df_final[df_final["Subject"].str.contains("CAPACITACION")==False]
df_final=df_final[df_final["Subject"].str.contains("CAPACITACIÓN")==False]
#Corregimos un error en el nombre de una empresa puesto por una analista
df_final["Subject"]=df_final.apply(lambda x:"COLEGIATURA COLOMBIANA DE COSMETOLOGIA" if x["Subject"]=="COELGIATURA" else x["Subject"],axis=1)
df_final["Subject"]=df_final.apply(lambda x:"ALUMINIOS Y VIDRIOS DE ANTIOQUIA" if x["Subject"]=="LUMINIOS" else x["Subject"],axis=1)
df_final["Subject"]=df_final.apply(lambda x:"ANDRES GOMEZ" if x["Subject"]=="NDRES GOMEZ" else x["Subject"],axis=1)
df_final["Subject"]=df_final.apply(lambda x:"ABB CONSULTING" if x["Subject"]=="BB CONSULTING" else x["Subject"],axis=1)
df_final["Subject"]=df_final.apply(lambda x:"CONSOLCARGO" if x["Subject"]=="COSOLCARGO" else x["Subject"],axis=1)
df_final["Subject"]=df_final.apply(lambda x:"ASESORIA MINERA S.A.S" if x["Subject"]=="MINERA S.A.S" else x["Subject"],axis=1)
df_final["Subject"]=df_final.apply(lambda x:"TRANSPORTANDO AM" if x["Subject"]=="TAM" else x["Subject"],axis=1)
df_final=df_final.reset_index()

#Ahora eliminaremos las reuniones que se crucen para un analista.Es decir, si hay 2 reuniones a la misma hora eliminamos ambas porque no hay forma de saber cual es la correcta.
analistas=df_final["Analista"].unique()
analistas_reuniones_dia={analista:{} for analista in analistas}
reuniones_eliminar=[]
reuniones_repetidas=[]
for i in range(len(df_final)):
    analista=df_final.loc[i,"Analista"]
    fecha_reu=df_final.loc[i,"Start"]
    empresa=df_final.loc[i,"Subject"]
    if fecha_reu not in analistas_reuniones_dia[analista]:
        analistas_reuniones_dia[analista][fecha_reu]=empresa
    elif empresa != analistas_reuniones_dia[analista][fecha_reu]:
        reuniones_eliminar.append((fecha_reu,analista))
for j in range(len(reuniones_eliminar)):
    fecha,nombre=reuniones_eliminar[j]
    df_final=df_final[(df_final["Analista"]!=nombre) | (df_final["Start"]!=fecha)]
#Ahora si una reunión con la misma empresa se repite varias veces vamos a dejar sólo una
df_final=df_final.drop_duplicates(subset=["Subject","Start","Analista"])

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


programar=pd.read_excel(r"D:\Users\siste\bd_local\Programacion_recordatorios.xlsx")
programar["Fecha reunion"]=programar.apply(lambda x:pd.to_datetime(x["Fecha reunion"]),axis=1)
if len(programar)>0:
    programar=programar[programar["Fecha reunion"]>pd.to_datetime(dt.now())]
programar["Fecha reunion"]=programar["Fecha reunion"].astype(str)
df_final = df_final.reset_index(drop=True)
for i in range(len(df_final)):
    fecha_recordatorio=seleccionar_fecha_recordatorio(df_final.loc[i,"Start"])
    fecha_recordatorio_inmediato=recordatorio_inmediato(df_final.loc[i,"Start"])
    fecha_reunion=df_final.loc[i,"Start"]  
    attendees=""
    for attendee in df_final.loc[i,"Attendees_adresses"]:
        attendees+=attendee+";"
    if len(df_final.loc[i,"Attendees_adresses"])!=1:
        attendees=attendees[:-1]
    nueva_fila=[df_final.loc[i,"Subject"],attendees,df_final.loc[i,"Analista"],str(fecha_recordatorio),"no",str(fecha_reunion),"no",str(df_final.loc[i,"link"])]
    nueva_fila_series = pd.Series(nueva_fila, index=programar.columns)
    #algunas empresas me solicitan sólo un mensaje recordatorio en lugar de dos por tanto debo añadir la lista de empresas y un condicional para evitar este caso
    lista_empresas=["AGENCIA DE SEGUROS MCALLISTER E HIJOS ASOCIADOS LTDA"]
    valor_verdad=True
    if df_final.loc[i,"Subject"] in lista_empresas:
        valor_verdad=False
    if valor_verdad:
        nueva_fila_inmediato=[df_final.loc[i,"Subject"],attendees,df_final.loc[i,"Analista"],str(fecha_recordatorio_inmediato),"no",str(fecha_reunion),"si",str(df_final.loc[i,"link"])]
        nueva_fila_series_inmediato = pd.Series(nueva_fila_inmediato, index=programar.columns)
    programar = pd.concat([programar, nueva_fila_series.to_frame().T], ignore_index=True)
    if valor_verdad:
        programar = pd.concat([programar, nueva_fila_series_inmediato.to_frame().T], ignore_index=True)
programar.to_excel(r"D:\Users\siste\bd_local\Programacion_recordatorios.xlsx",index=False)










