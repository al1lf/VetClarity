import customtkinter as ctk
from customtkinter import *
import time
from datetime import datetime, timedelta
from sqlalchemy import create_engine, text
import pandas as pd
import zipfile
import os
from dotenv import load_dotenv
import paramiko
from tkinter import messagebox
from PIL import Image

# Charger les variables d'environnement à partir du fichier .env / Load environment variables from the .env file
load_dotenv()

def f_duration_text(total_seconds):
    if total_seconds < 1:
        duration_text = f"Exported successfully in {total_seconds * 1000:.0f} msec"
    elif total_seconds < 60:
        duration_text = f"Exported successfully in {total_seconds:.0f} sec"
    elif total_seconds < 3600:
        minutes, seconds = divmod(total_seconds, 60)
        duration_text = f"Exported successfully in {int(minutes)} min {int(seconds)} sec"
    else:
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        duration_text = f"Exported successfully in {int(hours)} hr {int(minutes)} min {int(seconds)} sec"
    return duration_text

def download_file_from_sftp(host, port, username, password, remote_file_path, local_dir):
    """
    Télécharge un fichier depuis un serveur SFTP vers un répertoire local.

    :param host: Adresse du serveur SFTP
    :param port: Port du serveur SFTP (par défaut : 22)
    :param username: Nom d'utilisateur SFTP
    :param password: Mot de passe SFTP
    :param remote_file_path: Chemin complet du fichier sur le serveur SFTP
    :param local_dir: Répertoire local où stocker le fichier téléchargé
    """
    try:
        # Établir une connexion SSH
        transport = paramiko.Transport((host, port))
        transport.connect(username=username, password=password)

        # Ouvrir un client SFTP
        sftp = paramiko.SFTPClient.from_transport(transport)

        # Déterminer le nom du fichier à partir du chemin distant
        filename = os.path.basename(remote_file_path)
        local_file_path = os.path.join(local_dir, filename)

        # Télécharger le fichier
        sftp.get(remote_file_path, local_file_path)
        print(f"✅ Fichier téléchargé avec succès : {local_file_path}")
        label_sftp.configure(text=f"✅ Fichier téléchargé avec succès sur \n{local_file_path}", text_color='green')
            
        # Fermer la connexion
        sftp.close()
        transport.close()

    except Exception as e:
        print(f"❌ Erreur lors du téléchargement de {remote_file_path}: {e}")
        label_sftp.configure(text=f"❌ Erreur lors du téléchargement de {remote_file_path}: {e}", text_color='red')
        
def button_sftp_upload_func(file_name):
    sftp_host = os.getenv('SFTP_HOST')
    sftp_port = int(os.getenv('SFTP_PORT'))
    sftp_user = os.getenv('SFTP_USER')
    sftp_password = os.getenv('SFTP_PASSWORD')
    
    sftp_folder = os.getenv('SFTP_FOLDER_VTQ')
    
    database = os.getenv('SQL_DB_VTQ')
    server = os.getenv('SQL_SERVER')
    local_dir = f'//{server}/datafiles/{database}'
    
    remote_file_path = f'/{sftp_folder}/{file_name}'
    
    if not file_name:
        messagebox.showinfo("Information", "Aucun fichier à envoyer")
    else:
        download_file_from_sftp(sftp_host, sftp_port, sftp_user, sftp_password, remote_file_path, local_dir)

# def f_run_stored_proc(procedure, file_name):
#     # Paramètres de connexion à la base de données
#     database = os.getenv('SQL_DB_VTQ')
#     server = os.getenv('SQL_SERVER')
#     username = os.getenv('SQL_USER')
#     password = os.getenv('SQL_PASSWORD')    

#     # Paramètres de connexion à la base de données
#     connection_string = f'mssql+pymssql://{username}:{password}@{server}/{database}'
#     print(connection_string)
#     engine = create_engine(connection_string)
    
#     # Connexion à la base de données
#     try:
#         with engine.connect() as connection:
#             query = text(f"{procedure} '{file_name}'")
#             label_sp0.configure(text=f"{query}", text_color='orange')
#             connection.execute(query)
#             connection.commit()
            
#     except Exception as e:
#         # print(f"Erreur lors de l'exécution de la requête: {e}")
#         label_sp.configure(text=f"❌ Erreur lors de l'exécution de la procédure: {e}", text_color='red')

#     finally:
#         # Fermeture de la connexion
#         connection.close()
#         label_sp.configure(text=f"✅ Procédure exécutée avec succés", text_color='green')

def f_run_stored_proc(procedure: str, file_name: str):
    """Exécute une procédure stockée SQL avec un paramètre file_name."""

    # 🔹 Récupération des variables d'environnement pour la connexion
    database = os.getenv('SQL_DB_VTQ')
    server = os.getenv('SQL_SERVER')
    username = os.getenv('SQL_USER')
    password = os.getenv('SQL_PASSWORD')

    # 🔹 Vérification que les variables sont bien définies
    if not all([database, server, username, password]):
        print("❌ Erreur : Les paramètres de connexion ne sont pas définis.")
        return

    # 🔹 Construction de la chaîne de connexion SQLAlchemy
    connection_string = f"mssql+pyodbc://{username}:{password}@{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server"
    # connection_string = f'mssql+pymssql://{username}:{password}@{server}/{database}'
    
    try:
        # 🔹 Création de l'engine SQLAlchemy
        engine = create_engine(connection_string, echo=True)  # echo=True pour voir les requêtes exécutées

        # 🔹 Connexion à la base et exécution de la procédure stockée
        with engine.connect() as connection:
            query = text("EXEC {} :file_name".format(procedure))  # Paramétrisation pour éviter l'injection SQL
            
            print(f"🔍 Exécution de la procédure : {procedure} avec file_name = {file_name}")

            result = connection.execute(query, {"file_name": file_name})
            connection.commit()  # Validation si la procédure modifie des données
            
            # 🔹 Affichage des résultats si la procédure retourne des données
            rows = result.fetchall() if result.returns_rows else []
            for row in rows:
                print('row: ',row)

            print("✅ Procédure exécutée avec succès")

    except Exception as e:
        print(f"❌ Erreur lors de l'exécution : {e}")

                      
# Fenêtre principale / Main window    
window = ctk.CTk()
window.title('Export files')
window.geometry('500x500')

set_appearance_mode("dark")
set_default_color_theme("green")
color_grey_dark ='#555656'
color_grey_light ='#a9a9a9'
color_red_light = '#FF5733'
color_red_dark = '#A42408'

# window.minsize(500,500)
# window.maxsize(500,500)
window.resizable(False, False)

window.bind('<Escape>', lambda event: window.quit())

# title
logo_img_data = Image.open("Vetoquinol.png")
logo_img = CTkImage(dark_image=logo_img_data, light_image=logo_img_data, size=(128,55))
# title_text = '\nVetoquinol Data Export'
title_label = ctk.CTkLabel(master=window, image=logo_img, text='', text_color='white', font=('Helvetica', 15, 'bold'))
title_label.pack(pady=20)

# work directory by default
directory_frame = ctk.CTkFrame(master=window, width=800)
directory_frame.pack(pady=5, padx=50, anchor='w')
directory_label = ctk.CTkLabel(master=directory_frame, text='Export from:', font=('Helvetica', 12, 'bold'))
directory_label.pack(side='left', padx=5)
defaut_path = os.getenv('SFTP_HOST') + '/' + os.getenv('SFTP_FOLDER_VTQ')
work_path = ctk.CTkLabel(directory_frame, text=str(defaut_path), wraplength=400)
work_path.pack(side='left', padx=5)

file_frame = ctk.CTkFrame(master=window, width=800)
file_frame.pack(pady=5, padx=50, anchor='w')
file_label = ctk.CTkLabel(master=file_frame, text='File name:', font=('Helvetica', 12, 'bold'))
file_label.pack(side='left', padx=5)
date_du_jour = datetime.today().strftime("%d.%m.%y")
defaut_file = f'Extract Accounts Creatio - {date_du_jour}.xlsx'
file_to_download = ctk.CTkEntry(master=file_frame, width=300, height=35)
file_to_download.pack(pady=20)
file_to_download.insert(0, defaut_file)

# téléchargement depuis le SFTP
button_sftp_string = ctk.StringVar(value='Download from SFTP')
button_sftp = ctk.CTkButton(master=window, textvariable=button_sftp_string, command=lambda: button_sftp_upload_func(str(file_to_download.get())), fg_color=color_red_light, hover_color=color_red_dark, corner_radius=15, border_color='#A42408', border_width=1, text_color='white', width=120)
button_sftp.pack(pady=20, anchor='center')
label_sftp = ctk.CTkLabel(master=window, text='', wraplength=400)
label_sftp.pack(pady=10)

# lancement de la procédure stockée
button_sp_string = ctk.StringVar(value='Run Import')
button_sp = ctk.CTkButton(master=window, textvariable=button_sp_string, command=lambda: f_run_stored_proc('IMPORT_EXTRACT_ACCOUNTS',str(file_to_download.get())), text_color='black', border_color='black', border_width=1, corner_radius=15, width=120)
button_sp.pack(pady=20, anchor='center')
label_sp0 = ctk.CTkLabel(master=window, text='', wraplength=400)
label_sp0.pack(pady=10)
label_sp = ctk.CTkLabel(master=window, text='', wraplength=400)
label_sp.pack(pady=10)

# Add copyright text at the base of the window
copyright_label = ctk.CTkLabel(master=window, text="© ALF DATA, February 2025, v1.0", text_color='white')
copyright_label.pack(side='bottom', pady=10, padx=10, anchor='e')

# run
window.mainloop()
