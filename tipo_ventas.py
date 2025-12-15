import pandas as pd
df=pd.read_excel("VENTAS 2025 (1).xlsx")
df=df.drop(columns=["Unnamed: 8"])
df["VALOR FACTURADO"]=df["VALOR FACTURADO"].astype(int)
df["NUMERO"]=1
df["ASESOR"]=df.apply(lambda x:str(x["ASESOR"]).strip(),axis=1)
df["TIPO"]=df["TIPO"].fillna("NO ESPECIFICA")
df["CANAL"]=df["CANAL"].fillna("NO ESPECIFICA")
df_tipo=df.groupby(["ASESOR","TIPO"])[["VALOR FACTURADO","NUMERO"]].sum().reset_index()
df_canal=df.groupby(["ASESOR","CANAL"])[["VALOR FACTURADO","NUMERO"]].sum().reset_index()
asesores=df["ASESOR"].unique()
with pd.ExcelWriter('asesor_ventas_tipo.xlsx', engine='openpyxl') as writer:
    for asesor in asesores:
        df_aux=df_tipo[df_tipo["ASESOR"]==asesor]
        df_aux.to_excel(writer, sheet_name=asesor, index=False)
with pd.ExcelWriter('asesor_ventas_canal.xlsx', engine='openpyxl') as writer:
    for asesor in asesores:
        df_aux=df_canal[df_canal["ASESOR"]==asesor]
        df_aux.to_excel(writer, sheet_name=asesor, index=False)