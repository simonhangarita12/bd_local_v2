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
#sql_statement="""SELECT * FROM REUNIONES ORDER BY Fecha_procesamiento DESC LIMIT 7;"""
#sql_statement = """SELECT * FROM REUNIONES WHERE EMPRESA='INVERSIONES ORREGO Y CIA SA' ORDER BY Hora_ingreso;"""
#sql_statement = "UPDATE IGNORE REUNIONES SET Empresa='ANDRES CARMONA CASTILLO AL CARGO LOGISTICA' WHERE EMPRESA='ANDRES CARMONA CASTILLO AL CARGO LOGISTICA';"
#sql_statement = "UPDATE IGNORE REUNIONES SET Reuniones=0 WHERE EMPRESA='TRANSPORTES CAPETRANS SAS PESV' AND Hora_ingreso='2025-12-13 10:17:12.851000';"
#sql_statement = "SELECT * FROM REUNIONES WHERE EMPRESA='VELAS CQP PESV' AND Hora_ingreso='2025-07-07 09:05:00.340000';"
#sql_statement = """DELETE  FROM REUNIONES WHERE EMPRESA='DE PANIFICADORA MIPAN COLOMBIA';"""
#sql_statement = "UPDATE IGNORE REUNIONES SET EMPRESA='MARIA VELEZ VIANA' WHERE EMPRESA='SIN ESPECIFICAR' AND Hora_ingreso='2025-07-07 11:17:04.229000';"
sql_statement = "UPDATE REUNIONES SET Tiempo_conectado_asistente='0 days 01:06:22',Tiempo_muerto_inevitable='0 days 00:00:06.875000',Tiempo_muerto='0 days 00:00:01',Tiempo_muerto_real='0 days 00:00:00' WHERE EMPRESA ='CONSTRUCCION CON CONCIENCIA SAS PESV' AND Hora_ingreso='2025-12-12 15:29:32.141000';"
#sql_statement = """UPDATE IGNORE REUNIONES 
#                  SET Empresa='LUNA SAS' 
#                  WHERE EMPRESA='LUNA SAS"';"""
#sql_statement = """UPDATE IGNORE REUNIONES SET Reuniones=0 WHERE EMPRESA='Sin Especificar' AND Hora_ingreso='2025-07-10 12:53:30.265000';"""
#sql_statement= "SELECT * FROM REUNIONES WHERE EMPRESA LIKE '%EDELUX%';"
#sql_statement= "DESCRIBE VENTAS;"
#sql_statement="TRUNCATE TABLE VENTAS;"
#sql_statement="DELETE FROM VENTAS WHERE FECHA='11/13/2025';"
#sql_statement="DELETE FROM VENTAS WHERE ID>1000;"
#sql_statement="DROP TABLE VENTAS;"
#sql_statement="CREATE TABLE VENTAS (id INT PRIMARY KEY AUTO_INCREMENT,fecha VARCHAR(12) NOT NULL, vendedor VARCHAR(60) NOT NULL,cliente VARCHAR(220) NOT NULL,venta INT, producto_completo VARCHAR(70) NOT NULL, servicio INT,valor_proveedor INT,factura INT);"
#sql_statement="SELECT * FROM VENTAS ORDER BY FECHA;"
#sql_statement="UPDATE VENTAS SET vendedor='MARIA CAMILA GARCIA RESTREPO' WHERE CLIENTE='PROTOCOLOS DE SALUD MENTAL GESTION TECNOLOGICA & CONTABLE SAS' AND vendedor='LAURA DIAZ';"
#sql_statement="UPDATE VENTAS SET servicio=3 WHERE ID=681;"
sql_statement="SELECT * FROM PRESUPUESTO_ASESOR"
create_cursor.execute(sql_statement)
#mariadb_connection.commit()
result = create_cursor.fetchall()

create_cursor.close()
mariadb_connection.close()
print(result)
"""df=pd.DataFrame(result, columns=[i[0] for i in create_cursor.description])
df.to_excel("info_ventas_2.xlsx", index=False)"""