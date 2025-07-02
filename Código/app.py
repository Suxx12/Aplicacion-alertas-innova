import customtkinter as ctk
from tkinter import filedialog
import os
from extract_info import executives_list
import pandas as pd
import atexit
from data import messages_types, file_types
from data import file_types
import MultiSelectClass as msc
from aux_functions import cleanup, send_emails
import platform

selected_file_path = "utils.py"
path_file = ""
executives = []

EXCEL_FOLDER = "excel_alertas"

prefix = ""

if not os.path.exists(EXCEL_FOLDER):
    os.makedirs(EXCEL_FOLDER)
            
atexit.register(lambda: cleanup(EXCEL_FOLDER, selected_file_path))

ctk.set_appearance_mode("dark")  
ctk.set_default_color_theme("dark-blue")            


def button_select():
    file_path = filedialog.askopenfilename(
        title="Seleccionar Archivo",
        filetypes=(("Archivos de texto", "*.xlsx"), ("Todos los archivos", "*.*"))
    )
    if file_path:  
        global path_file
        global df_snapshot, df_db_projects, df_projects_advance
        global executives

        path_file = r"{}".format(file_path)

        df_projects_advance = pd.read_excel(path_file, sheet_name="Avance de proyectos (tipos)")
        df_db_projects = pd.read_excel(path_file, sheet_name="DB Proyectos Vigentes")
        df_snapshot = pd.read_excel(path_file, sheet_name="SNAPSHOT_PROYECTOS")

        executives = executives_list(df_projects_advance)

        multi_select.update_options(executives)
        multi_select.update_dataframe(df_db_projects)

        file_name = os.path.basename(file_path) 
        archivo_label.configure(text=f"Archivo Seleccionado: \"{file_name}\"")           


def update_inputs(*args):
    formatted_type = occupation_option_menu.get()  
    
    if formatted_type == "Seleccionar Tipo Aviso":
        return

    tipo = formatted_type.split(":")[0].strip()  

    global prefix

    data = messages_types.get(tipo, {"subject": "", "separate_message": "", "group_message": ""})
    subject = data.get("subject", "")
    separate_message = data.get("separate_message", "")
    group_message = data.get("group_message", "")

    prefix = file_types.get(tipo, {""})
    prefix = prefix.get("prefix", "")

    input1.delete(0, "end")
    input1.insert(0, subject)

    if checkbox_1.get():  
        message = separate_message.format(executive_name="{Nombre ejecutivo}")
    else:
        message = group_message

    input2.delete("1.0", "end")
    input2.insert("1.0", message)
    
     
def send():
    send_emails(occupation_option_menu, multi_select, input1, input2, checkbox_1, checkbox_2, df_snapshot, df_projects_advance, prefix, EXCEL_FOLDER, label_aviso, app)

values_list = [f"{tipo}: {datos['prefix'].replace('_', ' ')}" for tipo, datos in file_types.items()]


### == App diseño == ###
app = ctk.CTk()
app.title("Alertas Innova")
app.geometry("400x660")  

# Forzar que la ventana aparezca en primer plano en macOS
if platform.system() == 'Darwin':  # macOS
    app.lift()
    app.attributes('-topmost', True)
    app.after_idle(app.attributes, '-topmost', False)

app.wm_minsize(width=400, height=660)  
app.wm_maxsize(width=400, height=660)

app.grid_columnconfigure(0, weight=1) 

bold_font = ctk.CTkFont(family="Arial", size=25, weight="bold")
bold_text = ctk.CTkFont(family="Arial", size=14, weight="bold")

Title = ctk.CTkLabel(app, text="Alertas Innova", font=bold_font)
Title.grid(row=0, column=0, padx=20, pady=(25,5), sticky="w")

message1 = ctk.CTkLabel(app, text="Automatizar la generación de correos con avisos de proyectos.")
message1.grid(row=1, column=0, padx=20, pady=1, sticky="w", columnspan=2)

archivo = ctk.CTkButton(app, text="Seleccionar Archivo", command=button_select, fg_color="transparent", border_width=2)
archivo.grid(row=2, column=0, padx=20, pady=(20,0), sticky="ew", columnspan=2)

archivo_label = ctk.CTkLabel(app, text="", font=("Arial", 12), anchor="w", wraplength=360)
archivo_label.grid(row=3, column=0, padx=20, pady=0, sticky="ew", columnspan=2)

type_alert = ctk.CTkLabel(app, text="Tipo Aviso", font=bold_text)
type_alert.grid(row=4, column=0, padx=20, pady=(5,0), sticky="w")

occupation_option_menu = ctk.CTkOptionMenu(
    app, 
    values=values_list,
    anchor="w",
    command=update_inputs
)
occupation_option_menu.set("Seleccionar Tipo Aviso")
occupation_option_menu.grid(row=5, column=0, padx=20, pady=(5, 20), columnspan=2, sticky="ew")

type_executive = ctk.CTkLabel(app, text="Ejecutivos", font=bold_text)
type_executive.grid(row=6, column=0, padx=20, pady=0, sticky="w")

multi_select = msc.MultiSelectOptionMenu(app, executives)
multi_select.grid(row=7, column=0, padx=20, pady=(5,10), columnspan=2, sticky="ew") 

checkbox_1 = ctk.CTkCheckBox(app, text="Enviar por separado.", command=lambda: update_inputs())
checkbox_1.grid(row=8, column=0, padx=20, pady=(10, 20), sticky="w")

checkbox_2 = ctk.CTkCheckBox(app, text="cc: Sofía Ahumada")
checkbox_2.grid(row=8, column=1, padx=20, pady=(10, 20), sticky="w")

label1 = ctk.CTkLabel(app, text="Asunto")
label1.grid(row=9, column=0, padx=20, pady=0, sticky="w")

input1 = ctk.CTkEntry(app, placeholder_text="Aviso Correos")
input1.grid(row=10, column=0, padx=20, pady=(0,10), columnspan=2, sticky="ew")

label2 = ctk.CTkLabel(app, text="Mensaje")
label2.grid(row=11, column=0, padx=20, pady=(10,0), sticky="w")

input2 = ctk.CTkTextbox(app, height=100)
input2.grid(row=12, column=0, padx=20, pady=0, columnspan=4, sticky="nsew")

label_aviso = ctk.CTkLabel(app, text="", font=("Arial", 11), anchor="w", wraplength=360)
label_aviso.grid(row=13, column=0, padx=20, pady=(5, 0), sticky="w", columnspan=2)

button = ctk.CTkButton(app, text="Generar", command=lambda: send())
button.grid(row=14, column=0, padx=20, pady=(5,20), sticky="ew", columnspan=2)

app.mainloop()