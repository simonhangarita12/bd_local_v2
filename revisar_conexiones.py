import mysql.connector as mariadb
import pandas as pd
from datetime import timedelta
from datetime import datetime as dt


mariadb_connection = mariadb.connect(user="automatizacion", 
                                     password="LMUW7qczz4JRbJAuV6nmv8HbW", 
                                     host="172.16.100.10", 
                                     port=3306,
                                     database="atencion_clientes")
create_cursor = mariadb_connection.cursor()
sql_statement="""SELECT * FROM REUNIONES WHERE Reuniones=1 AND Empresa LIKE '%BLACK%';"""
create_cursor.execute(sql_statement)


result = create_cursor.fetchall()


column_names = [desc[0] for desc in create_cursor.description]


df1 = pd.DataFrame(result, columns=column_names)

create_cursor.close()
mariadb_connection.close()

mariadb_connection = mariadb.connect(user="automatizacion", 
                                     password="LMUW7qczz4JRbJAuV6nmv8HbW", 
                                     host="172.16.100.10", 
                                     port=3306,
                                     database="atencion_clientes")
create_cursor = mariadb_connection.cursor()
sql_statement="""SELECT * FROM REUNIONES WHERE Reuniones=1 AND Empresa LIKE '%COLORES%';"""
create_cursor.execute(sql_statement)


result = create_cursor.fetchall()


column_names = [desc[0] for desc in create_cursor.description]


df2 = pd.DataFrame(result, columns=column_names)

create_cursor.close()
mariadb_connection.close()

mariadb_connection = mariadb.connect(user="automatizacion", 
                                     password="LMUW7qczz4JRbJAuV6nmv8HbW", 
                                     host="172.16.100.10", 
                                     port=3306,
                                     database="atencion_clientes")
create_cursor = mariadb_connection.cursor()
sql_statement="""SELECT * FROM REUNIONES WHERE Reuniones=1 AND Empresa LIKE '%MONTECARLO%';"""
create_cursor.execute(sql_statement)


result = create_cursor.fetchall()


column_names = [desc[0] for desc in create_cursor.description]


df3 = pd.DataFrame(result, columns=column_names)

create_cursor.close()
mariadb_connection.close()

mariadb_connection = mariadb.connect(user="automatizacion", 
                                     password="LMUW7qczz4JRbJAuV6nmv8HbW", 
                                     host="172.16.100.10", 
                                     port=3306,
                                     database="atencion_clientes")
create_cursor = mariadb_connection.cursor()
sql_statement="""SELECT * FROM REUNIONES WHERE Reuniones=1 AND Empresa LIKE '%ARENAS%';"""
create_cursor.execute(sql_statement)


result = create_cursor.fetchall()


column_names = [desc[0] for desc in create_cursor.description]


df4 = pd.DataFrame(result, columns=column_names)

create_cursor.close()
mariadb_connection.close()




#df = pd.concat([df1, df2])  
df = pd.concat([df1, df2])
df=pd.concat([df,df3])
df=pd.concat([df,df4])
print(df.columns)
df=df[['Nombre','Empresa', 'Hora_ingreso','Tiempo_conectado','Tiempo_conectado_asistente', 'Tiempo_muerto_inevitable']]

df=df.rename(columns={'Nombre':'Analista','Hora_ingreso':"Hora de inicio de la reunión",'Tiempo_conectado':'Tiempo de conexión analista',
                      'Tiempo_conectado_asistente':'Tiempo de conexión cliente','Tiempo_muerto_inevitable':'Tiempo de impuntualidad cliente'})
df=df.sort_values(by="Hora de inicio de la reunión",ascending=True)
df["Hora de inicio de la reunión"]=df.apply(lambda x:str(x["Hora de inicio de la reunión"]),axis=1)
df["Tiempo de conexión analista"]=df.apply(lambda x:str(x["Tiempo de conexión analista"]),axis=1)
df["Tiempo de conexión cliente"]=df.apply(lambda x:str(x["Tiempo de conexión cliente"]),axis=1)
df["Tiempo de impuntualidad cliente"]=df.apply(lambda x:str(x["Tiempo de impuntualidad cliente"]),axis=1)
df.to_excel("Conexiones_revision.xlsx",index=False)
