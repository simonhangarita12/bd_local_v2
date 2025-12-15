import pandas as pd
from datetime import timedelta
reuniones_recordar=pd.read_excel(r"C:\Users\PCC\bd_local\Plantilla_recordatorios.xlsx")
programar=pd.read_excel(r"C:\Users\PCC\bd_local\Programacion_recordatorios.xlsx")
#print(reuniones_recordar.dtypes)
for i in range(len(reuniones_recordar)):
    fecha_recordatorio=reuniones_recordar.loc[i,"Fecha de la reunión"]
    hora_recordatorio=reuniones_recordar.loc[i,"Hora de la reunion"]
    minuto_recordatorio=reuniones_recordar.loc[i,"Minutos de la reunion"]
    formato=reuniones_recordar.loc[i,"Formato"]
    if formato=="pm":
        hora_recordatorio=int(hora_recordatorio)+12
    fecha=fecha_recordatorio-timedelta(days=1)+timedelta(hours=int(hora_recordatorio),minutes=int(minuto_recordatorio))
    nueva_fila=[reuniones_recordar.loc[i,"Nombre_del_cliente"],reuniones_recordar.loc[i,"Correo del cliente"],reuniones_recordar.loc[i,"Nombre del analista"],str(fecha),"no"]
    nueva_fila_series = pd.Series(nueva_fila, index=programar.columns)
    programar = pd.concat([programar, nueva_fila_series.to_frame().T], ignore_index=True)
programar.to_excel(r"C:\Users\PCC\bd_local\Programacion_recordatorios.xlsx",index=False)