import dashboard
import pandas as pd
from sqlalchemy import create_engine
filt=dashboard.filt.copy()



usuario = 'root'
contraseña = '1234'
host = 'localhost'
puerto = 3306
nombre_bd = 'testdb'

engine = create_engine(f'mariadb+mariadbconnector://{usuario}:{contraseña}@{host}:{puerto}/{nombre_bd}')

filt.to_sql("data_limpia", con=engine, if_exists="replace", index=False)
print("DataFrame guardado en la base de datos correctamente.")