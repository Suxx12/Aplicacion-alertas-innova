import pandas as pd

def executives_list(df):
    executives = []
    executives_v2 = df['Ejecutivo Técnico Proyecto'].unique()
    executives.extend(executives_v2)

    executives = [e for e in executives if pd.notna(e)]

    return executives


def get_email_by_name(name, df):
    executives_emails = []

    wanted_column = ["Ejecutivo tecnico", "Correo Ejecutivo técnico"]

    missing_columns = [col for col in wanted_column if col not in df.columns]
    if missing_columns:
        raise KeyError(f"Las columnas faltantes en df son: {missing_columns}")

    executives_emails = df[wanted_column]

    name_upper = name.upper().rstrip()
    email_row = executives_emails[executives_emails["Ejecutivo tecnico"].str.upper() == name_upper]
    
    if not email_row.empty:
        return email_row["Correo Ejecutivo técnico"].values[0]
    else:
        return None
