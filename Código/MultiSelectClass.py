import customtkinter as ctk
from tkinter import Toplevel  
from extract_info import get_email_by_name

class MultiSelectOptionMenu:
    def __init__(self, parent, options, title="Seleccionar Ejecutivos", df_db_proyectos = None):
        self.parent = parent
        self.options = options
        self.selected_options = {option: False for option in options}
        self.title = title
        self.df_db_proyectos = df_db_proyectos

        self.button = ctk.CTkButton(parent, text=self.get_selected_text(), command=self.open_menu)
        self.selected_with_emails = []

    def grid(self, **kwargs):
        self.button.grid(**kwargs)  
        
    def update_dataframe(self, df):
        self.df_db_proyectos = df    

    def open_menu(self):
        self.top = Toplevel(self.parent)
        self.top.title(self.title)
        self.top.geometry("330x400")
        self.top.wm_minsize(width=330, height=400)
        self.top.wm_maxsize(width=330, height=600)
        self.top.configure(bg="#2c2c2c")
        self.top.grab_set()

        self.top.protocol("WM_DELETE_WINDOW", self.close_menu)

        title_label = ctk.CTkLabel(self.top, text=self.title, font=("Arial", 14, "bold"))
        title_label.pack(pady=10)

        container = ctk.CTkFrame(self.top)
        container.pack(fill="both", expand=True, padx=10, pady=5)

        self.canvas = ctk.CTkCanvas(container, bg="#2c2c2c", highlightthickness=0)
        self.scrollbar = ctk.CTkScrollbar(container, orientation="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.options_frame = ctk.CTkFrame(self.canvas)
        self.canvas.create_window((0, 0), window=self.options_frame, anchor="nw")

        self.all_selected_var = ctk.BooleanVar(value=all(self.selected_options.values()))
        chk_all = ctk.CTkCheckBox(
            self.options_frame, text="Seleccionar Tod@s", variable=self.all_selected_var, command=self.toggle_all
        )
        chk_all.pack(anchor="w", padx=10, pady=5)

        self.check_vars = {}
        for option in self.options:
            var = ctk.BooleanVar(value=self.selected_options[option])
            chk = ctk.CTkCheckBox(
                self.options_frame,
                text=option,
                variable=var,
                command=self.update_all_selected,
            )
            chk.pack(anchor="w", padx=10, pady=5)
            self.check_vars[option] = var

        self.options_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        self.canvas.bind_all("<MouseWheel>", lambda event: self._on_mouse_wheel(event, self.canvas))
        self.canvas.bind_all("<Button-4>", lambda event: self._on_mouse_wheel(event, self.canvas))  
        self.canvas.bind_all("<Button-5>", lambda event: self._on_mouse_wheel(event, self.canvas))  

        close_button = ctk.CTkButton(self.top, text="Cerrar", command=self.close_menu)
        close_button.pack(pady=10)

    def update_options(self, new_options):
        self.options = new_options
        self.selected_options = {option: False for option in new_options}
        self.selected_with_emails = []
        self.button.configure(text=self.get_selected_text())

    def close_menu(self):
        if hasattr(self, 'canvas'):
            self.canvas.unbind_all("<MouseWheel>")
            self.canvas.unbind_all("<Button-4>")
            self.canvas.unbind_all("<Button-5>")

        for option, var in self.check_vars.items():
            self.selected_options[option] = var.get()

        self.selected_with_emails = []
        for option, is_selected in self.selected_options.items():
            if is_selected:
                email = get_email_by_name(option, self.df_db_proyectos)
                if email:
                    self.selected_with_emails.append((option, email))

        self.button.configure(text=self.get_selected_text(), )
        self.top.destroy()

    def _on_mouse_wheel(self, event, canvas):
        if event.num == 4:  
            canvas.yview_scroll(-1, "units")
        elif event.num == 5:  
            canvas.yview_scroll(1, "units")
        else: 
            canvas.yview_scroll(-1 * int(event.delta / 120), "units")

    def toggle_all(self):
        value = self.all_selected_var.get()
        for option, var in self.check_vars.items():
            var.set(value)

    def update_all_selected(self):
        if all(var.get() for var in self.check_vars.values()):
            self.all_selected_var.set(True)
        else:
            self.all_selected_var.set(False)

    def get_selected_text(self):
        selected = [option for option, is_selected in self.selected_options.items() if is_selected]
        if not selected:
            return "Seleccionar Ejecutivos"
        elif len(selected) == len(self.options):
            return "Tod@s seleccionados"
        else:
            return ", ".join(selected)
