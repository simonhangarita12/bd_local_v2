import mysql.connector as mariadb
import pandas as pd
from datetime import timedelta
from datetime import datetime as dt
mariadb_connection = mariadb.connect(user="test", 
                                     password="password", 
                                     host="localhost", 
                                     port=3306,
                                     database="testdb")
create_cursor = mariadb_connection.cursor()
sql_statement = "SELECT * FROM data_cleaned_v5"
create_cursor.execute(sql_statement)
result = create_cursor.fetchall()
columns=['Nombre', 'Empresa', 'hora de ingreso', 'duracion planeada', 'tiempo conectado','MeetingId',
         'id participante',"tiempo conectado asistentes","tiempo muerto inevitable", "tiempo muerto", 
         "diferencia tiempo", "tiempo muerto real", "reuniones"]
filt= pd.DataFrame(result, columns=columns)

create_cursor.close()
mariadb_connection.close()
filt["hora de ingreso"]=filt.apply(lambda x: pd.to_datetime(x["hora de ingreso"]) ,axis=1)
filt["duracion planeada"]=filt.apply(lambda x: pd.to_timedelta(x["duracion planeada"]),axis=1)
filt["tiempo conectado"]=filt.apply(lambda x: pd.to_timedelta(x["tiempo conectado"]),axis=1)
filt["tiempo conectado asistentes"]=filt.apply(lambda x: pd.to_timedelta(x["tiempo conectado asistentes"]),axis=1)
filt["tiempo muerto inevitable"]=filt.apply(lambda x: pd.to_timedelta(x["tiempo muerto inevitable"]),axis=1)
filt["tiempo muerto"]=filt.apply(lambda x: pd.to_timedelta(x["tiempo muerto"]),axis=1)
filt["tiempo muerto real"]=filt.apply(lambda x: pd.to_timedelta(x["tiempo muerto real"]),axis=1)



print(filt["id participante"].tail())

