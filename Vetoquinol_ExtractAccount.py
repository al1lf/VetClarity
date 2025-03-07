import customtkinter as ctk
from customtkinter import *
from tkinter import messagebox
from datetime import datetime, timedelta
from sqlalchemy import create_engine, text
import pandas as pd
import os
from dotenv import load_dotenv
import paramiko
from PIL import Image
import io

# Charger les variables d'environnement à partir du fichier .env / Load environment variables from the .env file
load_dotenv()

def update_table_with_sftp_file(host, port, username, password, remote_file_path):
    """
    Lit un fichier directement depuis un serveur SFTP et le charge dans un DataFrame pandas,
    avant de l'exporter dans la base de données pour mettre à jour la table CUSTOMERS
    """
    try:
        # print('🔹 Lecture du fichier distant depuis:', remote_file_path)

        # Établir une connexion SSH
        transport = paramiko.Transport((host, port))
        transport.connect(username=username, password=password)

        # Ouvrir un client SFTP
        sftp = paramiko.SFTPClient.from_transport(transport)

        # Ouvrir le fichier distant en mode binaire
        with sftp.file(remote_file_path, 'rb') as remote_file:
            remote_file.prefetch()  # Optimisation de la lecture
            file_data = remote_file.read()  # Lire le fichier en mémoire

        # Fermer la connexion SFTP
        sftp.close()
        transport.close()

        # Charger le fichier dans un DataFrame Pandas
        file_buffer = io.BytesIO(file_data)  # Stocker en mémoire sous forme de buffer

        # Déterminer le type de fichier (Excel ou CSV)
        if remote_file_path.endswith('.xlsx'):
            df = pd.read_excel(file_buffer)
        elif remote_file_path.endswith('.csv'):
            df = pd.read_csv(file_buffer)
        else:
            raise ValueError("Format de fichier non supporté !")

        print(f"✅ Fichier {remote_file_path} chargé dans un DataFrame avec succès !")
        # label_update.configure(text=f"✅ Fichier {remote_file_path} chargé dans un DataFrame avec succès !", text_color='white')

        # nettoyage des données
        df_cleaned = df.dropna(subset=['Account No.'])
        df_cleaned = df_cleaned[~df_cleaned['Account No.'].isin(['000015','217417'])]
        df_no_duplicates = df_cleaned[~df_cleaned.duplicated(subset='Account No.', keep=False)]
        columns = ['Account No.','Name','Address','City','ZIP/postal code','State/province','Country','Code brick','Migration Id']
        df_extract_accounts = df_no_duplicates.loc[:, columns]
        df_extract_accounts['Account No.'] = df_extract_accounts['Account No.'].astype("string")
        df_extract_accounts['date'] = datetime.now()
        df_extract_accounts['file'] = os.path.basename(remote_file_path)
        
        # Paramètres de connexion à la base de données
        database = os.getenv('SQL_DB_VTQ')
        server = os.getenv('SQL_SERVER')
        username = os.getenv('SQL_USER')
        password = os.getenv('SQL_PASSWORD')
        
        # Paramètres de connexion à la base de données
        connection_string = f'mssql+pymssql://{username}:{password}@{server}/{database}'
        engine = create_engine(connection_string)
        
        try:
            # Connexion à la base de données    
            with engine.connect() as connection:
                # transfert le dataframe dans une table de la base de données
                df_extract_accounts.to_sql("tmp_extract_accounts", connection, index=False, if_exists="replace")
                connection.commit()
                # mise à jour de CUSTOMERS
                merge_query = """
                            MERGE CUSTOMERS AS TARGET
                            USING tmp_extract_accounts AS SOURCE
                            ON (TARGET.idCustomer = SOURCE.[Account No.] COLLATE Latin1_General_100_CS_AS)
                            WHEN MATCHED THEN
                                UPDATE SET
                                    TARGET.customerName = SOURCE.[Name],
                                    TARGET.add1 = SOURCE.[Address],
                                    TARGET.city = SOURCE.[City],
                                    TARGET.postcode = SOURCE.[ZIP/postal code],
                                    TARGET.province = SOURCE.[State/province],
                                    TARGET.country = SOURCE.[Country],
                                    TARGET.brick = SOURCE.[Code brick],
                                    TARGET.grade = 'VRS',
                                    TARGET.type2 = SOURCE.[Migration Id],
                                    TARGET.isActive = 1,
                                    TARGET.updateDate = SOURCE.[date],
                                    TARGET.updateSource = SOURCE.[file]
                            WHEN NOT MATCHED BY TARGET AND SOURCE.[Account No.] IS NOT NULL THEN
                                INSERT (idCustomer, customerName, add1, city, postcode, province, country, brick, type2, isActive, updateDate, updateSource)
                                VALUES ([Account No.], [Name], [Address], [City], [ZIP/postal code], [State/province], [Country], [Code brick], [Migration Id], 1, [date], [file])
                            WHEN NOT MATCHED BY SOURCE THEN
                                UPDATE SET TARGET.isActive = 0, TARGET.grade = NULL, TARGET.updateDate = GETDATE();
                            """
                try:
                    connection.exec_driver_sql(merge_query)
                    connection.commit()
                    label_update.configure(text=f"✅ Table CUSTOMERS mise à jour avec succès avec le fichier {remote_file_path} !", text_color='white')
                    
                except Exception as e:
                    label_update.configure(text=f"❌ Erreur lors de la mise à jour de la table CUSTOMERS avec le fichier {remote_file_path}", text_color='red')
                
                
        except Exception as e:
            # print(f"❌ Erreur lors de la connexion à la base: {e}")
            label_update.configure(text=f"❌ Erreur lors de la connexion à la base: {e}", text_color='red')
                
        finally:
            # Fermeture de la connexion
            connection.close()
    
    except Exception as e:
        # print(f"❌ Erreur lors de la récupération du fichier {remote_file_path} sur le SFTP: {e}")
        label_update.configure(text=f"❌ Erreur lors de la récupération du fichier {remote_file_path} sur le SFTP: {e}", text_color='red')
        return None  # Retourner None en cas d'erreur

# def download_file_from_sftp(host, port, username, password, remote_file_path, local_dir):
#     """
#     Télécharge un fichier depuis un serveur SFTP vers un répertoire local.

#     :param host: Adresse du serveur SFTP
#     :param port: Port du serveur SFTP (par défaut : 22)
#     :param username: Nom d'utilisateur SFTP
#     :param password: Mot de passe SFTP
#     :param remote_file_path: Chemin complet du fichier sur le serveur SFTP
#     :param local_dir: Répertoire local où stocker le fichier téléchargé
#     """
#     try:
#         # charger le fichier dans un dataframe
#         print('remote_file_path: ', local_file_path)
#         df = pd.read_excel(local_file_path) if filename.endswith('.xlsx') else pd.read_csv(local_file_path)
#         print("✅ Fichier chargé dans un DataFrame avec succès !\n", df.head())
        
#         # Établir une connexion SSH
#         transport = paramiko.Transport((host, port))
#         transport.connect(username=username, password=password)

#         # Ouvrir un client SFTP
#         sftp = paramiko.SFTPClient.from_transport(transport)

#         # Déterminer le nom du fichier à partir du chemin distant
#         filename = os.path.basename(remote_file_path)
#         local_file_path = os.path.join(local_dir, filename)

#         # Télécharger le fichier
#         sftp.get(remote_file_path, local_file_path)
#         print(f"✅ Fichier téléchargé avec succès : {local_file_path}")
#         label_sftp.configure(text=f"✅ Fichier téléchargé avec succès sur \n{local_file_path}", text_color='green')
            
#         # Fermer la connexion
#         sftp.close()
#         transport.close()

#     except Exception as e:
#         print(f"❌ Erreur lors du téléchargement de {remote_file_path}: {e}")
#         label_sftp.configure(text=f"❌ Erreur lors du téléchargement de {remote_file_path}: {e}", text_color='red')
    
def button_execute(file_name):
    sftp_host = os.getenv('SFTP_HOST')
    sftp_port = int(os.getenv('SFTP_PORT'))
    sftp_user = os.getenv('SFTP_USER')
    sftp_password = os.getenv('SFTP_PASSWORD')
    sftp_folder = os.getenv('SFTP_FOLDER_VTQ')
    remote_file_path = f'/{sftp_folder}/{file_name}'
    if not file_name:
        messagebox.showinfo("Information", "Aucun fichier à envoyer")
    else:
        update_table_with_sftp_file(sftp_host, sftp_port, sftp_user, sftp_password, remote_file_path)
        
# def button_download(file_name):
#     sftp_host = os.getenv('SFTP_HOST')
#     sftp_port = int(os.getenv('SFTP_PORT'))
#     sftp_user = os.getenv('SFTP_USER')
#     sftp_password = os.getenv('SFTP_PASSWORD')
#     sftp_folder = os.getenv('SFTP_FOLDER_VTQ')
    
#     database = os.getenv('SQL_DB_VTQ')
#     server = os.getenv('SQL_SERVER')
#     local_dir = f'//{server}/datafiles/{database}'
    
#     remote_file_path = f'/{sftp_folder}/{file_name}'
    
#     if not file_name:
#         messagebox.showinfo("Information", "Aucun fichier à envoyer")
#     else:
#         download_file_from_sftp(sftp_host, sftp_port, sftp_user, sftp_password, remote_file_path, local_dir)

# Fenêtre principale / Main window    
window = ctk.CTk()
window.title('Update CUSTOMERS with ExtractAccounts')
window.geometry('500x600')

set_appearance_mode("dark")
set_default_color_theme("green")
color_grey_dark ='#555656'
color_grey_light ='#a9a9a9'
color_red_light = '#FF5733'
color_red_dark = '#A42408'
color_blue_light = '#50b8fe'
color_blue_dark = '#2596be'

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

# téléchargement depuis le SFTP et intégration des données du fichier ExtractAccounts
button_update_string = ctk.StringVar(value='Execute')
button_update = ctk.CTkButton(master=window, textvariable=button_update_string, command=lambda: button_execute(str(file_to_download.get())), fg_color=color_blue_light, hover_color=color_blue_dark, corner_radius=15, border_color=color_blue_dark, border_width=1, text_color='white', width=120)
button_update.pack(pady=20, anchor='center')
label_update = ctk.CTkLabel(master=window, text='', wraplength=400)
label_update.pack(pady=10)

# # téléchargement depuis le SFTP
# button_sftp_string = ctk.StringVar(value='Download from SFTP')
# button_sftp = ctk.CTkButton(master=window, textvariable=button_sftp_string, command=lambda: button_download(str(file_to_download.get())), fg_color=color_red_light, hover_color=color_red_dark, corner_radius=15, border_color='#A42408', border_width=1, text_color='white', width=120)
# button_sftp.pack(pady=20, anchor='center')
# label_sftp = ctk.CTkLabel(master=window, text='', wraplength=400)
# label_sftp.pack(pady=10)

# Add copyright text at the base of the window
copyright_label = ctk.CTkLabel(master=window, text="© ALF DATA, March 2025, v1.1", text_color='white')
copyright_label.pack(side='bottom', pady=10, padx=10, anchor='e')

# run
window.mainloop()
