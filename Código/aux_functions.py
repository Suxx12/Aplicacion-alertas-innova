import os
import shutil
from filters import type1, type2, type3, type4, type5, type6 #Nuevo type 6
import webbrowser
import urllib.parse

def cleanup(EXCEL_FOLDER, selected_file_path):
    if os.path.exists(EXCEL_FOLDER):
        shutil.rmtree(EXCEL_FOLDER)

    if os.path.exists(selected_file_path):
        with open(selected_file_path, "w") as file:
            file.write("")    


def open_excel_folder(EXCEL_FOLDER):
    if os.name == 'nt':  
        os.startfile(EXCEL_FOLDER)
    elif os.name == 'posix':  
        subprocess.run(['open' if sys.platform == 'darwin' else 'xdg-open', EXCEL_FOLDER])
    else:
        print(f"No se puede abrir la carpeta en este sistema operativo: {os.name}")             

        
def generate_excel(executives, send_separately, alert_type, df_snapshot, df_projects_advance, prefix, EXCEL_FOLDER):
    files = []

    if alert_type == "Tipo 1": 
        files = type1(df_snapshot, send_separately, executives, EXCEL_FOLDER, prefix, files)
        
    elif alert_type == "Tipo 2": 
        files = type2(df_snapshot, send_separately, executives, EXCEL_FOLDER, prefix, files)
    
    elif alert_type == "Tipo 3":
        files = type3(df_projects_advance, send_separately, executives, EXCEL_FOLDER, prefix, files)
        
    elif alert_type == "Tipo 4": 
        files = type4(df_projects_advance, send_separately, executives, EXCEL_FOLDER, prefix, files)
        
    elif alert_type == "Tipo 5":
        files = type5(df_snapshot, send_separately, executives, EXCEL_FOLDER, prefix, files)

    #Nuevo excel para informes técnicos pendientes de revisión ejecutiva
    elif alert_type == "Tipo 6":
        files = type6(df_snapshot, send_separately, executives, EXCEL_FOLDER, prefix, files)  

    return files        


def generate_email(to, cc, subject, body):
    params = {
        "subject": subject,        
        "body": body              
    }

    if cc:
        params["cc"] = cc
    encoded_params = urllib.parse.urlencode(params)

    encoded_params = encoded_params.replace("+", "%20")

    mailto_url = f"mailto:{to}?{encoded_params}"
    webbrowser.open(mailto_url)
    
def show_notice(message, label_aviso, app):
    label_aviso.configure(text=message)  
    label_aviso.update_idletasks()  
    app.after(4000, lambda: label_aviso.configure(text=""))        
    
def send_emails(occupation_option_menu, multi_select, input1, input2, checkbox_1, checkbox_2, df_snapshot, df_projects_advance, prefix, EXCEL_FOLDER, label_aviso, app):
    if occupation_option_menu.get() == "Seleccionar Tipo Aviso":
        show_notice("Por favor, seleccione un tipo de aviso.", label_aviso, app)
        return

    selected_executives = multi_select.selected_with_emails
    if not selected_executives:
        show_notice("Por favor, seleccione uno o más ejecutivos.", label_aviso, app)
        return

    subject = input1.get()
    body = input2.get("1.0", "end").strip()
    send_separately = checkbox_1.get()
    include_cc = checkbox_2.get()

    if not subject or not body:
        print("Por favor, llena los campos de Asunto y Mensaje.")
        return

    names = [name for name, _ in selected_executives]
    cc = "sofia.ahumada@corfo.cl" if include_cc else ""
    
    selected_type = occupation_option_menu.get().split(":")[0].strip()  

    files = generate_excel(names, send_separately, selected_type, df_snapshot, df_projects_advance, prefix, EXCEL_FOLDER)

    if not files:
        print("No se generaron archivos Excel.")
        return

    open_excel_folder(EXCEL_FOLDER)

    if send_separately:
        for (name, email), file in zip(selected_executives, files):
            name = name.split(" ")[0].lower().capitalize()

            custom_message = body.replace("{Nombre ejecutivo}", name)
            generate_email(email, cc, subject, custom_message)
    else:
        recipients = ", ".join([str(email) for _, email in selected_executives])
        generate_email(recipients, cc, subject, body)    
        
        
     
    
    
    