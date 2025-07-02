from datetime import date, timedelta
import pandas as pd
from helpers import generate_files, date_columns, reset_filters

# función Tipo 1
def type1(df, send_separately, executives, EXCEL_FOLDER, prefix, files):
    df = reset_filters(df)
    
    if df.empty:
        print("El DataFrame está vacío. No se pueden generar archivos.")
        return files

    df = date_columns(df, ['Anticipo', 'Fiel cumplimiento'])

    wanted_columns = ['Código', 'Nombre Ejecutivo Técnico', 'Estado Proyecto', 'Anticipo', 'Fiel cumplimiento']
    deadline = pd.Timestamp(date.today() + timedelta(weeks=12))
    current_date = pd.Timestamp(date.today())

    df_base_filtering = df[
        (df['Estado Proyecto'] == "VIGENTE") &
        (df['Anticipo'] <= deadline) &
        (df['Anticipo'] >= current_date)
    ].copy()

    df_base_filtering = df_base_filtering.sort_values(by='Anticipo', ascending=True)

    if df_base_filtering.empty:
        print("No hay datos que cumplan con las condiciones para generar archivos.")
        return files

    generate_files(send_separately, executives, df_base_filtering, EXCEL_FOLDER, prefix, files, wanted_columns, type="T1")

    df = reset_filters(df)
    return files

# función Tipo 2
def type2(df, send_separately, executives, EXCEL_FOLDER, prefix, files):
    df = reset_filters(df)
    if df.empty:
        print("El DataFrame está vacío. No se pueden generar archivos.")
        return files

    wanted_columns = ['Código', 'Nombre Ejecutivo Técnico',  'Fecha Resolucion', 'Rut Beneficiario', 'Estado de informe final', 'Anticipo','Fiel cumplimiento']

    df_base_filtering = df[
        (df['Anticipo'] == "No se encuentra anticipo") & 
        ((df['Estado Proyecto'] == "VIGENTE") | (df['Estado Proyecto'] == "VIGENTE(Reprogramación Enviada)"))
    ].copy()

    if df_base_filtering.empty:
        print("No hay datos que cumplan con las condiciones para generar archivos.")
        return files

    df_base_filtering = date_columns(df_base_filtering, ['Fecha Resolucion'])
    
    generate_files(send_separately, executives, df_base_filtering, EXCEL_FOLDER, prefix, files, wanted_columns, type="T2")
   
    df = reset_filters(df)
    return files

# función Tipo 3
def type3(df, send_separately, executives, EXCEL_FOLDER, prefix, files):
    df = reset_filters(df)
    if df.empty:
        print("El DataFrame está vacío. No se pueden generar archivos.")
        return files

    wanted_columns = ['Código de Proyecto', 'Ejecutivo Técnico Proyecto', 'Tipo de Informe', 'Fecha Entrega Programada', 'Fecha Entrega Real', 'Estado Proyecto']
    
    
    df = date_columns(df, ['Fecha Entrega Real', 'Fecha Entrega Programada'])

    df_base_filtering = df[
        (df['Fecha Entrega Real'].notna()) &
        (df['Estado Proyecto'] == "VIGENTE") &
        (df['Fecha cierre técnico'].isna()) &
        (df['Estado de revisión técnica'] == "PENDIENTE")
    ].copy()

    df_base_filtering = df_base_filtering.sort_values(by='Fecha Entrega Programada', ascending=True)

    if df_base_filtering.empty:
        print("No hay datos que cumplan con las condiciones para generar archivos.")
        return files

    generate_files(send_separately, executives, df_base_filtering, EXCEL_FOLDER, prefix, files, wanted_columns, type="T3")    
    
    df = reset_filters(df)
    return files

# función Tipo 4
def type4(df, send_separately, executives, EXCEL_FOLDER, prefix, files):
    df = reset_filters(df)
    if df.empty:
        print("El DataFrame está vacío. No se pueden generar archivos.")
        return files

    wanted_columns = ['Código de Proyecto', 'Ejecutivo Técnico Proyecto', 'Fecha Entrega Programada', 'Estado Proyecto']
    deadline = pd.Timestamp(date.today() + timedelta(weeks=12))

    df_base_filtering = df[
        (df['Fecha Entrega Real'].isna()) &
        (df['Estado Proyecto'] == "VIGENTE") &
        (df['Fecha Entrega Programada'] <= deadline) &
        (df['Estado de informe'] == "PENDIENTE")
    ].copy()

    if df_base_filtering.empty:
        print("No hay datos que cumplan con las condiciones para generar archivos.")
        return files

    df_base_filtering = date_columns(df_base_filtering, ['Fecha Entrega Programada'])

    generate_files(send_separately, executives, df_base_filtering, EXCEL_FOLDER, prefix, files, wanted_columns, type="T4")
    
    df = reset_filters(df)
    return files

# función Tipo 5
def type5(df, send_separately, executives, EXCEL_FOLDER, prefix, files):
    df = reset_filters(df)
    if df.empty:
        print("El DataFrame está vacío. No se pueden generar archivos.")
        return files

    wanted_columns = ['Código', 'Nombre Ejecutivo Técnico', 'Estado Proyecto', 'Fecha Resolucion', 'Rut Beneficiario', 'Estado de informe final']

    df_base_filtering = df[df['Estado Proyecto'] != "VIGENTE"].copy()

    if df_base_filtering.empty:
        print("No hay datos que cumplan con las condiciones para generar archivos.")
        return files

    if 'Fecha Resolucion' in df_base_filtering.columns:
        df_base_filtering = date_columns(df_base_filtering, ['Fecha Resolucion'])

    generate_files(send_separately, executives, df_base_filtering, EXCEL_FOLDER, prefix, files, wanted_columns, type="T5")

    df = reset_filters(df)
    return files

