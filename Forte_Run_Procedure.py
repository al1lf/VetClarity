import customtkinter as ctk
from customtkinter import *
from sqlalchemy import create_engine, text
from sqlalchemy.exc import SQLAlchemyError
from tkinter import messagebox, filedialog
from PIL import Image
from dotenv import load_dotenv
import os

# Charger les variables d'environnement à partir du fichier .env
load_dotenv()

# Variable globale pour stocker le chemin du fichier sélectionné
selected_file_path = None
defaut_path = os.getenv('SQL_SERVER') + '/' + os.getenv('SQL_DB_FOR')

# Fonction pour exécuter la procédure stockée
def execute_procedure(proc_name):
    print(f"Exécution de la procédure stockée : {proc_name}")
    global selected_file_path
    if not selected_file_path:
        # messagebox.showerror("Erreur", "Veuillez d'abord sélectionner un fichier.")
        label_file.configure(text="Veuillez d'abord sélectionner un fichier", text_color='red')
        return

    try:
        # Créer l'URL de connexion SQLAlchemy pour SQL Server
        server = os.getenv('SQL_SERVER')
        username = os.getenv('SQL_USER')
        password = os.getenv('SQL_PASSWORD')
        database = os.getenv('SQL_DB_FOR')
        # Formater l'URL de connexion pour SQLAlchemy (via pymssql)
        connection_url = f'mssql+pymssql://{username}:{password}@{server}/{database}'    
        # Connexion à la base de données
        engine = create_engine(connection_url)
        with engine.connect() as connection:
            # Passer le nom du fichier en paramètre à la procédure stockée
            file_name = os.path.basename(selected_file_path)
            connection.execute(text(f"EXEC {proc_name} @fileName='{file_name}'"))
            connection.commit()  # Commit si nécessaire

        # messagebox.showinfo("Succès", "Importation exécutée avec succès !")
        label_execute.configure(text="Importation exécutée avec succès !", text_color='green')
        print("Importation exécutée avec succès !")

    except SQLAlchemyError as e:
        # messagebox.showerror("Erreur", f"Une erreur s'est produite : {str(e)}")
        # messagebox.showerror("Erreur", f"Une erreur s'est produite.\nLe fichier sélectionné n'est pas correct.")
        label_execute.configure(text="Une erreur s'est produite.\nLe fichier sélectionné n'est pas correct.", text_color='red')
            
# Fonction pour importer un fichier
def import_file(defaut_name):
    global selected_file_path
    selected_file_path = filedialog.askopenfilename(
            initialdir=defaut_path, 
            title="Sélectionner un fichier", 
            # filetypes=[("Fichiers Excel", filetypes: {defaut_name})])
            filetypes=[("Fichiers Excel", defaut_name)])
    if selected_file_path:
        print(f"Fichier sélectionné : {selected_file_path}")
        label_file.configure(text=os.path.basename(selected_file_path), text_color='white')

# Fonction pour mettre à jour les boutons et les labels en fonction de la sélection dans la liste déroulante
def update_selection(event):
    selected_project = combo_box.get()
    print(f"Projet sélectionné : {selected_project}")
    label_file.configure(text="Aucun fichier sélectionné", text_color='white')
    label_execute.configure(text="")
    for item in data:
        if item["project_name"] == selected_project:
            btn_import.configure(command=lambda: import_file(item["defaut_name"]))
            btn_execute.configure(command=lambda: execute_procedure(item["proc_name"]))
            break
        
# Interface graphique avec CustomTkinter
window = ctk.CTk()
window.title("Importer fichiers")
window.geometry("600x400")

set_appearance_mode("dark")
set_default_color_theme("blue")

window.resizable(False, False)

window.bind('<Escape>', lambda event: window.quit())

# title
logo_img_data = Image.open("Forte.png")
logo_img = CTkImage(dark_image=logo_img_data, light_image=logo_img_data, size=(128,55))
title_label = ctk.CTkLabel(master=window, image=logo_img, text='', text_color='white', font=('Helvetica', 15, 'bold'))
title_label.pack(pady=20)

data = [
    {"project_name": "CUSTOMER DISCOUNTS", "proc_name": "IMPORT_DISCOUNT_CUSTOMERS", "defaut_name": "*CUSTOMER*DISCOUNT*.xls*"},
    {"project_name": "REBATE DISCOUNTS", "proc_name": "IMPORT_DISCOUNT_GROUPS", "defaut_name": "*MASTER*REBATE*DISCOUNTS*.xls*"},
    {"project_name": "MODIFICATIONS", "proc_name": "IMPORT_MODIFICATIONS", "defaut_name": "Modifications*.xls*"}
]
# Liste déroulante pour sélectionner le projet
project_names = [item["project_name"] for item in data]
combo_box = ctk.CTkComboBox(window, values=project_names, command=update_selection, width=300)
combo_box.set("Sélectionner un projet")
combo_box.pack(pady=10)

frame1 = ctk.CTkFrame(master=window)
frame1.pack(pady=5, padx=10, anchor='w')
btn_import = ctk.CTkButton(frame1, text="Sélectionner le fichier")
btn_import.pack(side='left', padx=50)
label_file = ctk.CTkLabel(frame1, text="Aucun fichier sélectionné", text_color='white')
label_file.pack(side='left', padx=5, pady=10)

frame2 = ctk.CTkFrame(master=window)
frame2.pack(pady=5, padx=10, anchor='w')
btn_execute = ctk.CTkButton(frame2, text="Importer les données")
btn_execute.pack(side='left', padx=50)
label_execute = ctk.CTkLabel(frame2, text="")
label_execute.pack(side='left', padx=5, pady=10)

# Add copyright text at the base of the window
copyright_label = ctk.CTkLabel(master=window, text="© ALF DATA, February 2025", text_color='white')
copyright_label.pack(side='bottom', pady=10, padx=10, anchor='e')

window.mainloop()
