import mysql.connector as mariadb
import pandas as pd
import re
from datetime import datetime as dt
from datetime import timedelta
import unicodedata
import parametros as par
import warnings

warnings.filterwarnings("ignore")
fecha_inicio_logs = dt.now().strftime('%Y-%m-%d %H:%M:%S')
fuentes_logs = []
registros_cargados_logs = []
destinos_logs = []

def salidaLog():
    """
    Este metodo se utiliza para organizar el log, es decir mostrar la informacion mas relevante del codigo.
    como es el el tiempo de ejecucion, los indices donde se ingesta y donde se obtienen los registros y 
    finalemente si se obtiene una ingesta exitosa o no.
    """
    print(f"Fecha_inicio: {fecha_inicio_logs}")
    print(f"Fecha_fin: {fecha_fin_logs}")
    print("Fuentes:", ' | '.join(map(str, fuentes_logs)))
    print("Registros cargados:", ' | '.join(map(str, registros_cargados_logs)))
    print("Destino:", ' | '.join(map(str, destinos_logs)))


def extraer_datos( user=par.usuario, password=par.password, host=par.host, port=par.puerto, database=par.bd):
    """Traemos los datos generados el día de hoy en el data lake de la base de datos que tenemos
    Args:
        user(str): usuario de la base de datos
        password(str): contraseña de la base de datos
        host(str): host de la base de datos
        port(int): puerto de la base de datos
        database(str): nombre de la base de datos
    Returns: data_horas(DataFrame): DataFrame con los datos obtenidos de las reuniones en columnas claves con las que vamos a trabajar"""
    mariadb_connection = mariadb.connect(user=user, 
                                        password=password, 
                                        host=host, 
                                        port=port,
                                        database=database)
    create_cursor = mariadb_connection.cursor()
    sql_statement = "SELECT * FROM DATALAKE_REUNIONES WHERE DATE(`fecha_extraccion`)= CURDATE()"
    create_cursor.execute(sql_statement)
    result = create_cursor.fetchall()
    columns=['MeetingId', 'Numero de participantes', 'Empresa', 'Email', 'Nombre',
            "inicio asistencia","fin asistencia", "id participante", "Rol",
            "hora de ingreso", "hora de salida", "segundos asistencia",
            "tiempo conectado", "tenant id", "tenant","Match Meeting Id",
            "duracion planeada", "inicio agendado","fecha de extraccion"]
    data_horas= pd.DataFrame(result, columns=columns)
    fuentes_logs.append("DATALAKE_REUNIONES")
    return data_horas

def data_a_trabajar(df:pd.DataFrame):
    """Aquí seleccionamos las columnas más relevantes y hacemos un pequeño preprosesamiento de los datos,
    eliminando los duplicados y reemplazando valornes nulos por un formato que no nos genere problemas a 
    la hora de migrar los datos limpios nuevamente a nuestra base de datos"""
    #seleccionamos las columnas mas relevantes para el analisis
    filt=df[["Nombre","Numero de participantes","Email","Rol","Empresa","hora de ingreso",
                    "hora de salida","duracion planeada","inicio agendado","tiempo conectado","MeetingId","id participante"]]
    filt=filt.fillna("")
    filt=filt.drop_duplicates().reset_index(drop=True)
    filt.index=filt.index+1
    return filt
def error_lectura_strings_numeros(filt:pd.DataFrame):
    """Arreglamos un pequeño error que nos presenta problemas para tomar los correos y los nombres
      como strings y es que en este caso algunos correos y nombres quizas por su corta longitud,
      o inexistencia los toma como enteros o float. Por tanto debemos hacer una pequeña corrección 
      inicial y luego podemos tratar estas dos columnas como cadenas de texto"""

    for i in filt.index:  
        if isinstance(filt.loc[i, "Email"], float):
            filt.loc[i, "Email"] = "None"
    for i in filt.index:  
        if isinstance(filt.loc[i, "Nombre"], int):
            filt.loc[i, "Nombre"] = "None"
    return filt


def es_analista(email:str):
    """Esta función nos sirve para ver si el participante de la reunión es un analista
      o no buscando una expresión regular que coincida con los correos de los analistas"""
    patron  = r"analista[a-zA-Z0-9_.+-]*@talentoconsultores"
    if re.search(patron, email, re.IGNORECASE):
        return 1
    else:
        return 0
    
#Ademas en las conexiones aveces hay presentes ai notetaker de alguna forma por tanto 
# intentaremos a su vez ponerlos como analistas, ya que en realidad no cuentan como un 
def seleccion_analista_ia(filt:pd.DataFrame):
    """Haremos un proceso parecido al que se realiza para identificar los analistas, 
    en el caso de que se use una ia durante la reunión para grabar la misma, ya que esta 
    ia no debería contar como un participante más y esto nos puede dañar los futuros algoritmos"""
    patron_ai  = r"AI Notetaker"

    #creamos una nueva columna temporal asignando el valor en binario de si es el caso de un analista 
    # o no
    filt["Es_analista"]= filt.apply(lambda x:es_analista(x["Email"]),axis=1)
    filt["Es_ia"]=filt.apply(lambda x:1 if re.search(patron_ai,x["Nombre"]) else 0,axis=1)
    filt["Es_ia"]=filt.apply(lambda x:1 if x["Nombre"]=="read.ai meeting notes" else x["Es_ia"],axis=1)
    return filt


def eliminar_repetidos_por_fallos_de_extraccion(filt:pd.DataFrame):
    """Vamos a arreglar un  error que sucede aveces en la extracción de los datos
    y es que la información en ocasiones se extrae mal o con datos faltantes en algunas columnas,
    no podemos eliminarla del todo cuando eliminamos repetidos. Para esto utilizamos el siguiente
    algoritmo que elimina completamente los repetidos que llegaron por mala extracción de los datos."""
    

    reunion_actual=filt["MeetingId"].iloc[0]
    lista_eliminar=[]
    for i in range(filt.shape[0]):
        if filt.loc[i+1,"MeetingId"]!=reunion_actual:
            reunion_actual=filt.loc[i+1,"MeetingId"]
            numero_asistentes=filt.loc[i+1,"Numero de participantes"]
            for n in range(1,numero_asistentes):
                if filt.loc[i+1+n,"MeetingId"]!=reunion_actual:
                    for j in range(0,n):
                        lista_eliminar.append(i+j+1)
                    break
    for e in lista_eliminar:
        filt=filt.drop(e)
    filt=filt.reset_index(drop=True)
    filt.index=filt.index+1
    return filt


def error_multi_organizador(filt:pd.DataFrame):
    """A continuacion se corrige el problema de que haya mas de un organizador en la reunion. 
    Lo cual tiene cierto sentido de forma interna, pero que nos daña los algoritmos siguientes.
    Entonces es mejor hacer la corrección, así sea un caso que tenga sentido"""
    

    reunion_actual=filt["MeetingId"].iloc[0]
    for i in range(filt.shape[0]):
        if filt.loc[i+1,"MeetingId"]!=reunion_actual:
            reunion_actual=filt.loc[i+1,"MeetingId"]
            numero_asistentes=filt.loc[i+1,"Numero de participantes"]
            for n in range(1,numero_asistentes):
                if filt.loc[i+1+n,"Rol"]!="Presenter":
                    filt.loc[i+1+n,"Rol"]="Presenter"
    return filt

def diccionario_nombres(filt:pd.DataFrame):
    """Esta función nos sirve para devolver un diccionario con los correos de los analistas y 
    su respetivo valor que es el nombre del analista más común asignado a su correo. 
    Lo cual nos va a servir en algunos algoritmos más adelante y taambién para unificar los nombres.
    
    Args:filt(DataFrame): DataFrame con el preprocesamiento de los datos realizado hasta el momento
    Returns: dict_nombres_moda(dict): diccionario con los correos de los analistas y su respectivo nombre"""
    analistas= filt[filt["Es_analista"]==1]["Email"].unique()
    dict_nombres={analista:[] for analista in analistas}
    for i in range(filt.shape[0]):
        if filt.loc[i+1,"Es_analista"]==1:
            dict_nombres[filt.loc[i+1,"Email"]].append(filt.loc[i+1,"Nombre"])
    from statistics import mode
    llaves=list(dict_nombres.keys())
    dict_nombres_moda={correo:mode(dict_nombres[correo]) for correo in llaves}
    return dict_nombres_moda

def nombre_erroneo_analistas(df:pd.DataFrame):
    """Vamos a arreglar un error que nos encontramos y es que aveces a los analistas
    se les asigna un nombre erroneo. Posiblemente de un asistente a la reunión y por tanto esto
    nos puede evitar obtener información correcta de los analistas."""
    df=df.copy()
    dict_nombres_moda=diccionario_nombres(df)
    lista_revisar=[]
    
    for i in range(df.shape[0]):
        if df.loc[i+1,"Rol"]=="Organizer" :
            numero_asistentes=df.loc[i+1,"Numero de participantes"]
            for j in range(1,numero_asistentes):
                if df.loc[i+j+1,"Es_analista"]==1:
                    if dict_nombres_moda[df.loc[i+j+1,"Email"]] != df.loc[i+j+1,"Nombre"]:
                        df.loc[i+j+1,"Es_analista"]=0
                    else:
                        lista_revisar.append((i,i+j+1))
    """Finalmente asignamos a las apariciones de inteligencias artificiales en las reunionesel estatus de analista con el objetivo de que en las conexiones y en los futuros algoritmosno cuente estas apariciones como un participante mas"""                    
    df["Es_analista"] = df.apply(lambda x: 1 if x.get("Es_ia", 0) == 1 else x.get("Es_analista", 0), axis=1)
    #df=df.drop(columns=["Es_ia"]) 
    return df 

def error_desconexion_reconexion(filt:pd.DataFrame):
    """Se arregla el error de que una persona se desconecta y se vuelta a conectar 
    varias veces en la reunion. Lo cual dañaba los tiempos para el analisis"""

    #vamos a usar en cambio de formato inicial en varios columnas,
    #  para tenerlas como fecha y asi poder hacer la manipulacion que queremos 
    # hacer para que se le sumen los tiempos de conexión a la persona que pudo
    #  haberse salido y vuelto a ingresar varias veces a la reunion.
    filt["tiempo conectado"]=filt.apply(lambda x: pd.to_timedelta(x["tiempo conectado"]),axis=1)
    filt["hora de ingreso"]=filt.apply(lambda x: pd.to_datetime(x["hora de ingreso"])-timedelta(hours=5) if str(x["hora de ingreso"])!="System.Object[]" else  dt(1990,1,1),axis=1)
    filt["hora de salida"]=filt.apply(lambda x: pd.to_datetime(x["hora de salida"])-timedelta(hours=5) if str(x["hora de salida"])!="System.Object[]" else  dt(2050,1,1),axis=1)
    filt["Nombre"]=filt.apply(lambda x: str(x["Nombre"]).strip(),axis=1)
    lista_cambios=[]
    for i in range(filt.shape[0]):
        if filt.loc[i+1,"Rol"]=="Organizer":
            lista_asistentes=[]
            numero_asistentes=filt.loc[i+1,"Numero de participantes"]
            for j in range(1,numero_asistentes):
                if filt.loc[i+j+1,"Nombre"] in lista_asistentes and filt.loc[i+j+1,"Nombre"]!="None":
                    numero_cambio=lista_asistentes.index(filt.loc[i+j+1,"Nombre"])+1
                    #vamos a considerar los dos casos en los que los intervalos de tiempo no se cruzan, lo cuál representa verdaderamente una desconexión y retorno a la reunión
                    if filt.loc[i+j+1,"hora de ingreso"]<filt.loc[i+numero_cambio+1,"hora de ingreso"]:
                        #vamos a dar un pequeño margen de error de 1 minuto para considerar los cruces de tiempo
                        if filt.loc[i+j+1,"hora de salida"]-timedelta(minutes=1)<filt.loc[i+numero_cambio+1,"hora de ingreso"]:
                            #vamos a ubicar la suma de los tiempos de conexion en la primera aparicion del asistente
                            
                            filt["tiempo conectado"].iloc[i+numero_cambio]=filt.loc[i+1+numero_cambio,"tiempo conectado"]+filt.loc[i+1+j,"tiempo conectado"]
                            filt["hora de ingreso"].iloc[i+numero_cambio]=min(filt.loc[i+1+numero_cambio,"hora de ingreso"],filt.loc[i+1+j,"hora de ingreso"])
                            lista_asistentes.append(filt.loc[i+j+1,"Nombre"])
                            lista_cambios.append((i+numero_cambio,i+j))
                        else:
                            lista_asistentes.append(filt.loc[i+j+1,"Nombre"])
                    elif filt.loc[i+j+1,"hora de ingreso"]>filt.loc[i+numero_cambio+1,"hora de ingreso"]:
                        if filt.loc[i+numero_cambio+1,"hora de salida"]-timedelta(minutes=1)<filt.loc[i+j+1,"hora de ingreso"]:
                            #vamos a ubicar la suma de los tiempos de conexion en la primera aparicion del asistente
                            
                            filt["tiempo conectado"].iloc[i+numero_cambio]=filt.loc[i+1+numero_cambio,"tiempo conectado"]+filt.loc[i+1+j,"tiempo conectado"]
                            filt["hora de ingreso"].iloc[i+numero_cambio]=min(filt.loc[i+1+numero_cambio,"hora de ingreso"],filt.loc[i+1+j,"hora de ingreso"])
                            lista_asistentes.append(filt.loc[i+j+1,"Nombre"])
                            lista_cambios.append((i+numero_cambio,i+j))
                        else:
                            lista_asistentes.append(filt.loc[i+j+1,"Nombre"])
                    else:
                        lista_asistentes.append(filt.loc[i+j+1,"Nombre"])
                else:
                    lista_asistentes.append(filt.loc[i+j+1,"Nombre"])
    return filt
       
def desconexion_reconexion_analista(filt:pd.DataFrame):
       
    """Vamos a arreglar el mismo error para los analistas ya que se utiliza un algoritmo 
    distinto y mucho mas sencillo para encontrar sus desconexiones"""
    for i in range(filt.shape[0]):
        if filt.loc[i+1,"Rol"]=="Organizer":
            nombre_comparar=filt.loc[i+1,"Nombre"]
            correo_comparar=filt.loc[i+1,"Email"]
            numero_asistentes=filt.loc[i+1,"Numero de participantes"]
            for j in range(1,numero_asistentes):
                if filt.loc[i+j+1,"Nombre"]==nombre_comparar and filt.loc[i+j+1,"Email"]==correo_comparar:
                    #vamos a ubicar la suma de los tiempos de conexion en la primera aparicion del asistente
                    
                    if filt.loc[i+j+1,"hora de ingreso"]<filt.loc[i+1,"hora de ingreso"]:
                    #vamos a dar un pequeño margen de error de 1 minuto para considerar los cruces de tiempo
                        if filt.loc[i+j+1,"hora de salida"]-timedelta(minutes=1)<filt.loc[i+1,"hora de ingreso"]:
                            #vamos a ubicar la suma de los tiempos de conexion en la primera aparicion del asistente
                            print(i+j,"     ",i)
                            filt["tiempo conectado"].iloc[i]=filt.loc[i+1,"tiempo conectado"]+filt.loc[i+j+1,"tiempo conectado"]
                            filt["hora de ingreso"].iloc[i]=min(filt.loc[i+1,"hora de ingreso"],filt.loc[i+j+1,"hora de ingreso"])
                    else:
                        #vamos a ubicar la suma de los tiempos de conexion en la primera aparicion del asistente
                        if filt.loc[i+1,"hora de salida"]-timedelta(minutes=1)<filt.loc[i+j+1,"hora de ingreso"]:
                            #vamos a ubicar la suma de los tiempos de conexion en la primera aparicion del asistente
                            print(i+j,"     ",i)
                            filt["tiempo conectado"].iloc[i]=filt.loc[i+1,"tiempo conectado"]+filt.loc[i+j+1,"tiempo conectado"]
                            filt["hora de ingreso"].iloc[i]=min(filt.loc[i+1,"hora de ingreso"],filt.loc[i+j+1,"hora de ingreso"])
        return filt

def error_analista_no_organizador(filt:pd.DataFrame):
    """Arreglamos el error de que el analista no aparezca como el organizador de la reunión. 
    Muy posiblemente porque este mismo no fue el que programó la reunión y 
    se la agendaron desde gerencia, pero queremos que aparezca como organizador para facilitar la 
    implementación de los algoritmos siguientes"""
    #Llamamos a esta función que nos devolvera el dataframe con los nombres corregidos y 
    #el diccionario con los nombres de los analistas basado en sus correos.        
    filt=nombre_erroneo_analistas(filt)
    dict_nombres_moda=diccionario_nombres(filt)

    actual_reunion=filt.loc[1,"MeetingId"]
    lista_actuales=[]
    posicion_ubicar=0
    filt.iloc[0], filt.iloc[1] = filt.iloc[1].copy(), filt.iloc[0].copy()
    for i in range(filt.shape[0]):
        if filt.loc[i+1,"MeetingId"]!=actual_reunion:
            actual_reunion=filt.loc[i+1,"MeetingId"]
            posible_organizador=""
            posicion_organizador=0
            for ele in lista_actuales:
                if filt.loc[ele,"Rol"]=="Organizer":
                    break
                else:
                    if filt.loc[ele,"Email"] in dict_nombres_moda.keys():
                        if dict_nombres_moda[filt.loc[ele,"Email"]] == filt.loc[ele,"Nombre"]:
                            
                            posible_organizador=filt.loc[ele,"Email"]
                            posicion_organizador=ele
                            break
            if posible_organizador!="":
                filt["Rol"].iloc[posicion_organizador-1]="Organizer"
                filt.iloc[posicion_organizador-1],filt.iloc[posicion_ubicar]=filt.iloc[posicion_ubicar].copy(),filt.iloc[posicion_organizador-1].copy()
            actual_reunion=filt.loc[i+1,"MeetingId"]
            lista_actuales=[]
            lista_actuales.append(i+1)
            posicion_ubicar=i
        else:
            lista_actuales.append(i+1)
    return filt


def obtener_tiempo_conexion_asistentes(filt:pd.DataFrame):
    """Se realiza la implementación del algoritmo diseñado para obtener las variables que 
    luego nos ayuden a analizar los tiempos de conexión tanto del analista como de los asistentes y 
    también los tiempos muertos y de retraso de los asistentes que se presentan en la reunión."""
    #creamos nuevas columnas para obtener posteriormente los tiempos muertos por reunion
    filt["tiempos"]=filt.apply(lambda x: [], axis=1)
    filt["tiempos inevitables"]=filt.apply(lambda x: [], axis=1)

    #en este proceso vamos a almacenar los tiempos que duraron los asistentes
    #  en la reunion junto con su hora de entrada
    filt["tiempos"]=filt.apply(lambda x: [], axis=1)
    filt["tiempos inevitables"]=filt.apply(lambda x: [], axis=1)

    for i in range(filt.shape[0]):
        if filt.loc[i+1,"Rol"]=="Organizer":
            numero_asistentes=filt.loc[i+1,"Numero de participantes"]
            if numero_asistentes>1:
                lista=[]
                inevitables=[]
                for j in range(1,numero_asistentes):
                    if filt.loc[i+j+1,"Es_analista"]==0:
                        lista.append(filt.loc[i+j+1,"tiempo conectado"])
                        inevitables.append(filt.loc[i+j+1,"hora de ingreso"])
                
                filt.at[i+1,"tiempos"]=lista
                filt.at[i+1,"tiempos inevitables"]=inevitables
    #seleccionamos como un aproximado del tiempo que duraron los asistentes 
    # a la persona que mas tiempo duro dentro de la reunion y seleccionamos
    #  su hora de llegada a la reunion
    filt["tiempo conectado asistentes"]=filt.apply(lambda x: max(x["tiempos"])if len(x["tiempos"])>0 else 0, axis=1)
    filt["tiempo de entrada inevitable"]=filt.apply(lambda x:x["tiempos inevitables"][x["tiempos"].index(max(x["tiempos"]))]if len(x["tiempos"])>0 else x["hora de ingreso"],axis=1)
    return filt


def tiempos_system_object(filt:pd.DataFrame):
    """En esta función vamos a realizar una manipulación pequeña para poder obtener un estimado
    del tiempo en el que se realizo la reunión, para los casos donde el analista no tenía tiempo
    de conexión debido a que en los datos se obtuvo System Object en lugar de obtener la fecha
    real en que se dio la conexión. Esto con el objetivo de no perder estos fragmentos de información"""

    filt["manipular hora"]=1
    for i in range(filt.shape[0]):
        if filt.loc[i+1,"Rol"]=="Organizer" and filt.loc[i+1,"hora de ingreso"]<dt(1991,1,1):
            numero_asistentes=filt.loc[i+1,"Numero de participantes"]
            for j in range(1,numero_asistentes):
                if filt.loc[i+j+1,"hora de ingreso"]>dt(1991,1,1):
                    filt["hora de ingreso"].iloc[i]=filt["hora de ingreso"].iloc[i+j]
                    filt["manipular hora"].iloc[i]=0
                    break
    return filt


def seleccion_organizadores(filt:pd.DataFrame):
    """Realizaremos un filtrado para seleccionar únicamente la información de los organizadores,
    ya que la información relevante de los asistentes ya está contenida en la información de los
    organizadores"""

    filt=filt[filt["Rol"]=="Organizer"]
    filt=filt.drop(columns=["Rol"])
    return filt


def filtrado_por_correo(filt:pd.DataFrame):
    """Realizaremos un filtrado de los correos utilizando expresiones regulares para obtener 
    unicamente los correos de los analistas. Los cuales siempre inician por analista"""
    
    patron  = r"analista[a-zA-Z0-9_.+-]*@talentoconsultores"
    #utilizamos esta funcion auxiliar para separar los analistas
    def aux_fun(email):
        if re.search(patron, email, re.IGNORECASE):
            return email
        else:
            return "None"
    filt.Email=filt.apply(lambda x: aux_fun(x.Email),axis=1) 

    #Dejamos unicamente los datos de los analistas
    filt=filt[filt.Email!="None"]
    return filt

def formateo_fechas(filt:pd.DataFrame):
    """Hacemos una transformación de todas las columnas que sean referentes
    a fechas o intervalos de tiempo a un formato donde las podamos manipular
    y considerar más facilmente como tal"""
    filt=filt.drop(columns=["tiempos","tiempos inevitables"])
    def fecha_ing(fecha):
        spanish_months = {
        'ene.': 'Jan', 'feb.': 'Feb', 'mar.': 'Mar', 'abr.': 'Apr',
        'may.': 'May', 'jun.': 'Jun', 'jul.': 'Jul', 'ago.': 'Aug',
        'sep.': 'Sep', 'oct.': 'Oct', 'nov.': 'Nov', 'dic.': 'Dec'
    }
        for spa in spanish_months:
            if spa in fecha:
                fecha = fecha.replace(spa, spanish_months[spa])
            else:
                pass
        return fecha
    filt["inicio agendado"]=filt.apply(lambda x: fecha_ing(x["inicio agendado"]),axis=1) 

    filt["inicio agendado"]=filt.apply(lambda x: pd.to_datetime(x["inicio agendado"]),axis=1)
    filt["duracion planeada"]=filt.apply(lambda x: pd.to_timedelta(x["duracion planeada"]),axis=1)

    filt["tiempo conectado asistentes"]=filt.apply(lambda x: pd.to_timedelta(x["tiempo conectado asistentes"]),axis=1)
    return filt


def filtrar_reuniones_ya_transcurridas(filt:pd.DataFrame):
    """Revisamos que no se introduzcan datos de reuniones ya agendadas, 
    pero que aún no se han presentado con el objetivo de evitar información 
    que ensucie los datos por tener información falsa"""
    year_actual,mes_actual,dia_actual=dt.now().year,dt.now().month,dt.now().day
    fecha=dt(year=year_actual,month=mes_actual,day=dia_actual)
    filt=filt[filt["inicio agendado"]<fecha]
    return filt


def creacion_nuevas_variables(filt:pd.DataFrame):
    """Creamos nuevas variables a partir de las variables que ya tenemos, 
    las cuales nos aportan información adicional relevante a la hora de visualizar y de 
    analizar la información que tenemos. Como el tiempo muerto total de la reunión, 
    el tiempo de impuntualidad de los asistentes, la diferencia de tiempo entre el tiempo de 
    reunión planeada y el tiempo real que tomó la reunión."""

    filt["hora de inicio"]=0
    filt["tiempo muerto inevitable"]=filt.apply(lambda x: x["tiempo de entrada inevitable"]-x["hora de ingreso"] if (x["tiempo de entrada inevitable"]!=timedelta(0)and x["manipular hora"]==1 and x["hora de ingreso"]<x["tiempo de entrada inevitable"]) else timedelta(0),axis=1)
    filt["hora de inicio"]=filt.apply(lambda x: x["inicio agendado"].hour,axis=1)
    filt["tiempo muerto"]=filt.apply(lambda x: x["tiempo conectado"]-x["tiempo conectado asistentes"] if (int(x["tiempo conectado asistentes"].total_seconds() // 60 % 60)>0  and x["tiempo conectado"]>x["tiempo conectado asistentes"] )else timedelta(0),axis=1)
    filt["diferencia tiempo"]=filt.apply(lambda x: int((x["duracion planeada"]-x["tiempo conectado"]).total_seconds() // 60 % 60) if x["duracion planeada"]>x["tiempo conectado"] else -int((x["tiempo conectado"]-x["duracion planeada"]).total_seconds() // 60 % 60),axis=1)
    filt=filt.drop(columns=["tiempo de entrada inevitable","manipular hora"])
    filt["tiempo muerto real"]=filt.apply(lambda x: x["tiempo muerto"]-x["tiempo muerto inevitable"] if x["tiempo muerto"]>x["tiempo muerto inevitable"] else timedelta(0),axis=1)
    return filt

def limpieza_general(filt:pd.DataFrame):
    """"Realizamos el proceso de limpieza más general que consiste en eliminar los duplicados que nos 
    hayan quedado (lo cual es precaución porque este proceso ya se hizo en detalle) y eliminar las 
    filas con valores vacios"""
    filt=filt.drop_duplicates()
    filt=filt.dropna()
    return filt

def quitar_tildes(texto):
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')

def eliminar_tildes(filt):
    """En esta función eliminamos las tildes de los nombres y de las empresas, 
    con el objetivo de unificar con mayor facilidad los nombres""" 
    filt["Nombre"]=filt.apply(lambda x:quitar_tildes(x["Nombre"]),axis=1)
    filt["Empresa"]=filt.apply(lambda x: quitar_tildes(x["Empresa"]),axis=1)
    filt["Empresa"]=filt.apply(lambda x: x["Empresa"].upper(), axis=1)
    #tratamos por aparte una empresa que nos puede traer problemas 
    # a la hora de eliminar palabras que no aportan al nombre de la empresa
    filt["Empresa"]=filt.apply(lambda x:"ASESORíA MINERA" if x["Empresa"]=='ASESORIA SST, ASESORIA MINERA' else x["Empresa"],axis=1)
    return filt


def eliminar_palabras(nombre_empresa):
    """Eliminamos palabras que no a´prtan mucho al nombre de la empresa, o que no 
    permiten la unificación por nombre de empresa
    
    Args: nombre_empresa(str) nombre de la empresa a limpiar
    
    Returns: nombre_empresa(str) nombre de la empresa luego de realizar los cambios"""
    terminos_a_eliminar=par.diccionario_cambios_nombres

    nombre_empresa = " ".join(nombre_empresa.split())
    for palabra in terminos_a_eliminar:
        nombre_empresa=nombre_empresa.replace(palabra,terminos_a_eliminar[palabra])
    nombre_empresa=nombre_empresa.replace("ASESORIAS","")
    #Al final volvemos a hacer split debido a todos los espacios en blanco que se generaron
    nombre_empresa = " ".join(nombre_empresa.split())
    return nombre_empresa


def reemplazar_nombre(filt:pd.DataFrame):
    """Esta función también cumple el objetivo de ayudarnos a unificar los nombre de las empresas.
    Sin embargo,esta función se utiliza para los casos donde no podemos utilizar la función replace.
    Porque son términos tan generales que el reemplazarlos puede afectar a otras empresas."""
    terminos_a_eliminar=par.diccionario_reemplazos_nombres
    for palabra in terminos_a_eliminar:
        filt["Empresa"]=filt.apply(lambda x:palabra if x["Empresa"]==terminos_a_eliminar[palabra] else x["Empresa"],axis=1)
    return filt

def unificar_nombres(filt:pd.DataFrame):
    """En esta función vamos a intentar unificar y limpiar los nombres utilizando las dos funciones 
    construidas con anterioridad y colocando los nombre en mayusculas.
    """
    filt["Empresa"]=filt.apply(lambda x: x["Empresa"].upper(), axis=1)
    filt["Empresa"]=filt.apply(lambda x:eliminar_palabras(x["Empresa"]),axis=1)
    filt=reemplazar_nombre(filt)
    return filt

def eliminar_errores(filt:pd.DataFrame):
    """A continuación se repararán todos los errores de los datos que se han presentado de 
    momento utilizando las funciones especificas para cada error que ya se han construido con 
    anterioridad y devolvemos el dataframe con estos errores corregidos, listo para aplicar 
    los algoritmos posteriores y el analisis requerido"""
    filt=error_lectura_strings_numeros(filt)
    filt=seleccion_analista_ia(filt)
    filt=eliminar_repetidos_por_fallos_de_extraccion(filt)
    filt=error_multi_organizador(filt)
    filt=nombre_erroneo_analistas(filt)
    filt=error_desconexion_reconexion(filt)
    filt=desconexion_reconexion_analista(filt)
    filt=error_analista_no_organizador(filt)
    return filt
def filtrado_informacion(filt:pd.DataFrame):
    """Realizamos todos los filtrados de la información necesarios para obtener la información 
    necesaria para el analisis, y para concentrarnos únicamente en la porción de información 
    relevante a nuestro problema"""
    filt=obtener_tiempo_conexion_asistentes(filt)
    filt=tiempos_system_object(filt)
    filt=seleccion_organizadores(filt)
    filt=filtrado_por_correo(filt)
    filt=formateo_fechas(filt)
    filt=filtrar_reuniones_ya_transcurridas(filt)
    return filt

def limpieza_conjunto(filt:pd.DataFrame):
    """Realizamos la limpieza total de la información utilizando todas las funciones 
    creadas con el objetivo de hacer limpieza de los datos y retornamos el dataframe ya limpio
    """
    filt=limpieza_general(filt)
    filt=eliminar_tildes(filt)
    filt=unificar_nombres(filt)
    #Finalmente las columnas que no enviaremos a la base de datos
    filt=filt.drop(columns=["Email","Numero de participantes","hora de salida","inicio agendado","hora de inicio","Es_analista","Es_ia"])

    return filt

def creacion_variables(filt:pd.DataFrame):
    filt['Fecha_procesamiento'] = dt.now().strftime("%Y-%m-%d %H:%M:%S")
    #filt['Fecha_procesamiento'] = filt.apply(lambda x: pd.to_datetime(x['Fecha_procesamiento']), axis=1)
    filt['reuniones'] = 1
    return filt
def error_varias_reuniones(filt:pd.DataFrame):
    """Arreglaremos un error que se presenta cuando aparecen varias reuniones de la misma empresa en 
    el mismo día. Lo cual puede estar ensuciando los datos al hacer parecer que la empresa tiene más 
    conexiones de las que en verdad tiene. Además, en el caso de que se presenten dos reuniones de la 
    misma empresa y en la reunión que se va a desechar hay un tiempo de conexión significativo, 
    se lo vamos a añadir a la reunión que si se dio"""
    #Creamos una nueva columna para la fecha de la reunión en formato datetime
    filt["fecha"]=filt.apply(lambda x:dt.combine(x["hora de ingreso"],dt.min.time()),axis=1)
    #Reseteamos el indice para poder recorrer más facilmente el dataframe
    filt=filt.reset_index(drop=True)
    #Utilizamos el diccionario siguiente para almacenar las renuiones del día que dió el 
    # analista con la misma empresa, con el objetivo de luego comparar y dejar unicamente 
    # la reunión que si se dio, o en caso de que ninguna se diera, unicamente asignar una falta
    dic_reuniones_dia={}
    for i in range(filt.shape[0]):
        clave=str(filt["fecha"].iloc[i])+filt["Empresa"].iloc[i]+filt["Nombre"].iloc[i]
        valor=[filt["tiempo conectado"].iloc[i],filt["tiempo conectado asistentes"].iloc[i]]
        if clave not in dic_reuniones_dia.keys():
            dic_reuniones_dia[clave]=[valor]
        else:
            dic_reuniones_dia[clave].append(valor)
    filt["reuniones"]=1
    for i in range(filt.shape[0]):
        clave=str(filt["fecha"].iloc[i])+filt["Empresa"].iloc[i]+filt["Nombre"].iloc[i]
        maximo_asistentes=max(dic_reuniones_dia[clave], key=lambda x: x[1])
        maximo_analista=max(dic_reuniones_dia[clave], key=lambda x: x[0])
        tiempo_asistentes= filt["tiempo conectado asistentes"].iloc[i]
        tiempo_analista=filt["tiempo conectado"].iloc[i]
        if maximo_asistentes[1]==timedelta(0):
            if maximo_analista[0]!=tiempo_analista:
                filt["reuniones"].iloc[i]=0
        else:
            if maximo_asistentes[1]!=tiempo_asistentes:
                filt["reuniones"].iloc[i]=0
            else:
                lista_conexiones=dic_reuniones_dia[clave]
                for conexion in lista_conexiones:
                    if conexion[1]>timedelta(minutes=3):
                        if conexion[0]!=tiempo_analista or conexion[1]!=tiempo_asistentes:
                            filt["tiempo conectado asistentes"].iloc[i]+=conexion[1]
                            filt["tiempo conectado"].iloc[i]+=conexion[0]
    #borramos la columna que creamos para este algoritmo
    filt=filt.drop(columns="fecha")
    return filt

def enviar_datos(filt:pd.DataFrame,user=par.usuario, password=par.password, host=par.host, port=par.puerto, database=par.bd):
    """En esta función vamos a enviar los datos despues de pasar por el proceso de limpieza y 
    preprocesamiento a la base de datos, en nuestra tabla donde contenemos la información ya limpia
    
    Args:
        filt(DataFrame): DataFrame con la información ya limpia y lista para ser enviada a la base de datos
        user(str): usuario de la base de datos
        password(str): contraseña de la base de datos
        host(str): host de la base de datos
        port(int): puerto de la base de datos
        database(str): nombre de la base de datos"""

    mariadb_connection = mariadb.connect(user=user, 
                                         password=password, 
                                         host=host, 
                                         port=port, 
                                         database=database)
    create_cursor = mariadb_connection.cursor()
    """create_cursor.execute(
            CREATE TABLE IF NOT EXISTS REUNIONES (
                Nombre VARCHAR(255),
                Empresa VARCHAR(255),
                `hora de ingreso` VARCHAR(255),
                `duracion planeada` VARCHAR(255),
                `tiempo conectado` VARCHAR(255),
                MeetingId VARCHAR(255),
                `id participante` VARCHAR(255),
                `tiempo conectado asistentes` VARCHAR(255),
                `tiempo muerto inevitable` VARCHAR(255),
                `tiempo muerto` VARCHAR(255),
                `diferencia tiempo` INT,
                `tiempo muerto real` VARCHAR(255),
                `reuniones` INT,
                PRIMARY KEY (MeetingId, `id participante`,`tiempo conectado`)            
                        ))
    """

    cantidad_registros = 0
    for index, row in filt.iterrows():
        try:
            create_cursor.execute("""
                INSERT IGNORE INTO REUNIONES (
                    Nombre, Empresa, `Hora_ingreso`,`Duracion_planeada`, `Tiempo_conectado`,`Meeting_Id`,
                    `Id_participante`,`Tiempo_conectado_asistente`, `Tiempo_muerto_inevitable`,
                    `Tiempo_muerto`,`Diferencia_tiempo`, `Tiempo_muerto_real`, `Reuniones`, `Fecha_procesamiento`
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,%s,%s,%s)
            """, (
                str(row['Nombre']), str(row['Empresa']), str(row['hora de ingreso']), str(row['duracion planeada']), 
                str(row['tiempo conectado']),str(row['MeetingId']),str(row["id participante"]),
                str(row['tiempo conectado asistentes']), str(row['tiempo muerto inevitable']),  
                str(row['tiempo muerto']), int(row['diferencia tiempo']), 
                str(row['tiempo muerto real']), int(row['reuniones']),
                row['Fecha_procesamiento'],

            ))
            cantidad_registros += create_cursor.rowcount
        except mariadb.Error as e:
            print("el error es: ", e)
        
    try:
        mariadb_connection.commit()
        print("Datos insertados correctamente")
        destinos_logs.append("REUNIONES")
    except mariadb.Error as e:
        print("el error es: ", e)
    finally:
        create_cursor.close()
        mariadb_connection.close()
        registros_cargados_logs.append(cantidad_registros)


if __name__=="__main__":
    filt=extraer_datos()
    filt=data_a_trabajar(filt)
    filt=eliminar_errores(filt)
    filt=filtrado_informacion(filt)
    filt=creacion_nuevas_variables(filt)
    filt=limpieza_conjunto(filt)
    filt=creacion_variables(filt)
    #print(filt)
    enviar_datos(filt)
    fecha_fin_logs = dt.now().strftime('%Y-%m-%d %H:%M:%S')
    salidaLog()




