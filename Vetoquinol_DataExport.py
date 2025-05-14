import customtkinter as ctk
from customtkinter import *
import time
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
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

def f_generate_dates(end_date):
    """
    Génère une liste de dates au format '01/MM/AAAA' pour les 11 mois précédant la date donnée.   
    :param end_date: La date de fin sous forme de chaîne 'AAAA-MM-JJ' ou d'objet datetime.
    :return: Liste des dates sous le format '01/MM/AAAA'.
    
    Generates a list of dates in ‘01/MM/YYYY’ format for the 11 months preceding the given date.   
    :param end_date: The end date in the form of a string ‘YYYY-MM-DD’ or a datetime object.
    :return: List of dates in the format ‘01/MM/YYYY’.
    """
    # Convertir la date de fin en objet datetime si nécessaire / Convert the end date into a datetime object if necessary
    if isinstance(end_date, str):
        end_date = datetime.strptime(end_date, "%Y-%m-%d")
    
    dates = []
    for i in range(12):  # Inclure la date de fin et les 11 mois précédents / Include end date and previous 11 months
        date = end_date - relativedelta(months=i)
        formatted_date = date.strftime("01/%m/%Y")
        dates.append(formatted_date)
    
    return dates#[::-1]  # Inverser la liste pour avoir l'ordre croissant / Reverse the list to get ascending order

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

# create the final directory with timestamp included
def f_create_directory(date_str):
    period = datetime.strptime(date_str, "%d/%m/%Y").strftime("%y%m")
    export_path = work_path.cget('text') + '/' + period + '/' + datetime.now().strftime('%y%m%d%H%M%S')
    # vérifier si le répertoire mensuel à créer existe déjà
    if os.path.exists(export_path):
        # si le répertoire existe, générer un nouveau nom en ajoutant un suffixe numérique
        new_directory_path = export_path + "_1"
        index = 1
        while os.path.exists(new_directory_path):
            index += 1
            new_directory_path = export_path + "_" + str(index)
        # renommer le répertoire existant
        os.rename(export_path, new_directory_path)
        print(f"Le répertoire '{export_path}' a été renommé en '{new_directory_path}'.")
    # créer le répertoire s'il n'existe pas
    os.makedirs(export_path)
    print(f"Répertoire '{export_path}' créé avec succès.")
    return export_path

# Connexion au serveur SFTP et transfert du fichier
def f_upload_to_sftp(local_file, remote_path, sftp_host, sftp_port, sftp_user, sftp_password):
    try:
        # Créer la connexion SFTP
        transport = paramiko.Transport((sftp_host, sftp_port))
        transport.connect(username=sftp_user, password=sftp_password)
        sftp = paramiko.SFTPClient.from_transport(transport)

        # Séparer le répertoire distant et le nom du fichier
        remote_dir = os.path.dirname(remote_path)
        remote_filename = os.path.basename(remote_path)
    
        # Fonction pour créer le répertoire distant récursivement
        def ensure_remote_dir(sftp, remote_dir):
            dirs = remote_dir.split("/")
            current_dir = ""
            for dir in dirs:
                if dir:  # Éviter les chaînes vides dues à des "//"
                    current_dir = f"{current_dir}/{dir}" if current_dir else dir
                    try:
                        sftp.stat(current_dir)  # Vérifier si le dossier existe
                    except FileNotFoundError:
                        sftp.mkdir(current_dir)  # Créer le dossier s'il n'existe pas

        # Vérifier et créer le répertoire distant si nécessaire
        ensure_remote_dir(sftp, remote_dir)

        # Transférer le fichier vers le serveur SFTP
        sftp.put(local_file, remote_path)

        print(f"Fichier {local_file} transféré avec succès vers {remote_path}.")

        # Fermer la connexion SFTP
        sftp.close()
        transport.close()

    except Exception as e:
        print(f"Erreur lors du transfert SFTP : {e}")

# FONCTIONS D'EXPORT
# VERS CSV
def f_export_table_to_text(server, database, username, password, table_name, output_file, col_separator):
    connection_string = f'mssql+pymssql://{username}:{password}@{server}/{database}'
    engine = create_engine(connection_string)
    # Connexion à la base de données
    connection = engine.connect()
    
    try:
        # Exécutez une requête SQL pour sélectionner toutes les lignes de la table spécifiée
        query = f'SELECT * FROM {table_name}'
        dataframe = pd.read_sql(query, engine)
        # Affichez le DataFrame
        # print(dataframe.head())
        
        # Exportez le DataFrame dans un fichier csv
        dataframe.to_csv(output_file, sep=col_separator, index=False)
        
    except Exception as e:
        print(f"Erreur lors de l'exportation de la table {table_name}: {e}")

    finally:
        # Fermeture de la connexion
        connection.close()
        print(f"Exportation de {table_name} réussi")


# VERS EXCEL
def f_export_table_to_excel(server, database, username, password, table_name, output_file, sheetName):
    connection_string = f'mssql+pymssql://{username}:{password}@{server}/{database}'
    engine = create_engine(connection_string)
    # Connexion à la base de données
    connection = engine.connect()

    try:
        # Exécutez une requête SQL pour sélectionner toutes les lignes de la table spécifiée
        query = f'SELECT * FROM {table_name}'
        dataframe = pd.read_sql(query, engine)
        # Affichez le DataFrame
        # print(dataframe.head())
        
        # Exportez le DataFrame dans un fichier MS Excel
        dataframe.to_excel(output_file, sheet_name=sheetName, index=False)
        
    except Exception as e:
        print(f"Erreur lors de l'exportation de la table {table_name}: {e}")

    finally:
        # Fermeture de la connexion
        connection.close()
        print(f"Exportation de {table_name} réussi")


list_export_file_sftp = []

# functions checkboxes
def f_check1(export_path):
    project = '00011_File'
        
    # Paramètres de connexion à la base de données
    database = os.getenv('SQL_DB_VTQ')
    server = os.getenv('SQL_SERVER')
    username = os.getenv('SQL_USER')
    password = os.getenv('SQL_PASSWORD')    

    # mois du traitement
    period = datetime.strptime(droplist_var.get(), "%d/%m/%Y").strftime("%y%m")
    
    # # dataframe log
    # df_files = pd.DataFrame(columns=['file','time'])
    
    # Paramètres de connexion à la base de données
    connection_string = f'mssql+pymssql://{username}:{password}@{server}/{database}'
    engine = create_engine(connection_string)
    # Connexion à la base de données
    connection = engine.connect()

    # Connexion à la base de données
    try:
        with engine.connect() as connection:
            # fichier 00011_File multi feuilles
            # requêtes SQL pour extraire les données des trois tables
            date_start = datetime.now()
            query1 = "SELECT * FROM VIEW_PRODUCTS_VRS"
            query2 = "SELECT * FROM EXPORT_CRM_VAL"
            query3 = "SELECT * FROM EXPORT_CRM_Q"
            # exécution des requêtes et stockage des résultats dans des DataFrames
            df1 = pd.read_sql(query1, connection)
            df2 = pd.read_sql(query2, connection)
            df3 = pd.read_sql(query3, connection)
            # définition des noms des feuilles
            nom_feuille1 = "UK Products"
            nom_feuille2 = "UK Value"
            nom_feuille3 = "UK Units"
            # création d'un objet ExcelWriter pour écrire dans un fichier Excel
            file_name = project
            full_file_name = export_path + '/' + file_name + '_' + period + '.xlsx'
            print("full_file_name :", full_file_name)
            writer = pd.ExcelWriter(full_file_name, engine='xlsxwriter')
            # écriture des DataFrames dans différentes feuilles du fichier Excel avec les noms spécifiés
            df1.to_excel(writer, sheet_name=nom_feuille1, index=False)
            df2.to_excel(writer, sheet_name=nom_feuille2, index=False)
            df3.to_excel(writer, sheet_name=nom_feuille3, index=False)
            # sauvegarde du fichier Excel
            writer.close()
                
    except Exception as e:
        # print(f"Erreur lors de l'exécution de la requête: {e}")
        check1_label.configure(text=f"Erreur lors de l'exécution de la requête: {e}")

    finally:
        # Fermeture de la connexion
        connection.close()
        print(f"Exportation de {project} réussi")
        
        # Calculer la durée de l'exportation
        date_end = datetime.now()
        duration = date_end - date_start
        total_seconds = duration.total_seconds()
        duration_text = f_duration_text(total_seconds)
        # new_entry = pd.DataFrame([{'file': file_name, 'time': duration_text}])
        # if not new_entry.dropna(how='all').empty:
        #     df_files = pd.concat([df_files, new_entry], ignore_index=True)
        # # Afficher le dataframe
        # print(df_files)
        
        check1_label.configure(text=duration_text, text_color='white')
        list_export_file_sftp.append(full_file_name)
    
def f_check2(export_path):
    project = 'creatio_missing_customers'
    
    # Paramètres de connexion à la base de données
    database = os.getenv('SQL_DB_VTQ')
    server = os.getenv('SQL_SERVER')
    username = os.getenv('SQL_USER')
    password = os.getenv('SQL_PASSWORD')    

    # mois du traitement
    period = datetime.strptime(droplist_var.get(), "%d/%m/%Y").strftime("%y%m")
    
    # # dataframe log
    # df_files = pd.DataFrame(columns=['file','time'])
    
    # Paramètres de connexion à la base de données
    connection_string = f'mssql+pymssql://{username}:{password}@{server}/{database}'
    engine = create_engine(connection_string)
    # Connexion à la base de données
    connection = engine.connect()

    # Connexion à la base de données
    try:
        with engine.connect() as connection:
            date_start = datetime.now()
            file_name = project
            full_file_name = export_path + '/' + file_name + '_' + period + '.xlsx'
            f_export_table_to_excel(server, database, username, password, 'EXPORT_CREATIO_MISSING_CUSTOMERS', full_file_name,'missing customers')
            
    except Exception as e:
        # print(f"Erreur lors de l'exécution de la requête: {e}")
        check2_label.configure(text=f"Erreur lors de l'exécution de la requête: {e}")

    finally:
        # Fermeture de la connexion
        connection.close()
        print(f"Exportation de {project} réussi")

        # Calculer la durée de l'exportation
        date_end = datetime.now()
        duration = date_end - date_start
        total_seconds = duration.total_seconds()
        duration_text = f_duration_text(total_seconds)
        # new_entry = pd.DataFrame([{'file': file_name, 'time': duration_text}])
        # if not new_entry.dropna(how='all').empty:
        #     df_files = pd.concat([df_files, new_entry], ignore_index=True)
        # # Afficher le dataframe
        # print(df_files)
        
        check2_label.configure(text=duration_text, text_color='white')
        list_export_file_sftp.append(full_file_name)
    
def f_check3(export_path):
    project = 'GROUPS_VTQ'
    
    # Paramètres de connexion à la base de données
    database = os.getenv('SQL_DB_VTQ')
    server = os.getenv('SQL_SERVER')
    username = os.getenv('SQL_USER')
    password = os.getenv('SQL_PASSWORD')    

    # mois du traitement
    period = datetime.strptime(droplist_var.get(), "%d/%m/%Y").strftime("%y%m")
    
    # # dataframe log
    # df_files = pd.DataFrame(columns=['file','time'])
    
    # Paramètres de connexion à la base de données
    connection_string = f'mssql+pymssql://{username}:{password}@{server}/{database}'
    engine = create_engine(connection_string)
    # Connexion à la base de données
    connection = engine.connect()

    # Connexion à la base de données
    try:
        with engine.connect() as connection:
            # fichier GROUPS_VTQ multi feuilles
            # requêtes SQL pour extraire les données des trois tables
            date_start = datetime.now()
            query4 = "SELECT * FROM VIEW_GROUPS_DETAILS"
            query5 = "SELECT idCustomer, customerName, add1, city, postcode, updateSource, FORMAT(dateFrom,'dd/MM/yyyy') AS dateFrom, FORMAT(dateTo,'dd/MM/yyyy') AS dateTo, idGroup, groupName, mat FROM EXPORT_GROUPS_DETAILS ORDER BY idCustomer, dateFrom"
            # exécution des requêtes et stockage des résultats dans des DataFrames
            df4 = pd.read_sql(query4, connection)
            df5 = pd.read_sql(query5, connection)
            # définition des noms des feuilles
            nom_feuille4 = "groupslinks"
            nom_feuille5 = "details"
            # création d'un objet ExcelWriter pour écrire dans un fichier Excel
            file_name = project
            full_file_name = export_path + '/' + file_name + '_' + period + '.xlsx'
            # print(full_file_name)
            writer = pd.ExcelWriter(full_file_name, engine='xlsxwriter')
            # écriture des DataFrames dans différentes feuilles du fichier Excel avec les noms spécifiés
            df4.to_excel(writer, sheet_name=nom_feuille4, index=False)
            df5.to_excel(writer, sheet_name=nom_feuille5, index=False)
            # sauvegarde du fichier Excel
            writer.close()
            
    except Exception as e:
        # print(f"Erreur lors de l'exécution de la requête: {e}")
        check3_label.configure(text=f"Erreur lors de l'exécution de la requête: {e}")

    finally:
        # Fermeture de la connexion
        connection.close()
        print(f"Exportation de {project} réussi")

        # Calculer la durée de l'exportation
        date_end = datetime.now()
        duration = date_end - date_start
        total_seconds = duration.total_seconds()
        duration_text = f_duration_text(total_seconds)
        # new_entry = pd.DataFrame([{'file': file_name, 'time': duration_text}])
        # if not new_entry.dropna(how='all').empty:
        #     df_files = pd.concat([df_files, new_entry], ignore_index=True)
        # # Afficher le dataframe
        # print(df_files)      
        
        check3_label.configure(text=duration_text, text_color='white')
        list_export_file_sftp.append(full_file_name)
    
def f_check4(export_path):
    project = 'Monthly_VC_Data'
    
    # Paramètres de connexion à la base de données
    database = os.getenv('SQL_DB_VTQ')
    server = os.getenv('SQL_SERVER')
    username = os.getenv('SQL_USER')
    password = os.getenv('SQL_PASSWORD')    

    # mois du traitement
    period = datetime.strptime(droplist_var.get(), "%d/%m/%Y").strftime("%y%m")
    
    # # dataframe log
    # df_files = pd.DataFrame(columns=['file','time'])
    
    # Paramètres de connexion à la base de données
    connection_string = f'mssql+pymssql://{username}:{password}@{server}/{database}'
    engine = create_engine(connection_string)
    # Connexion à la base de données
    connection = engine.connect()

    # Connexion à la base de données
    try:
        with engine.connect() as connection:
            date_start = datetime.now()
            file_name = project
            full_file_name = export_path + '/' + file_name + '_' + period +'.csv'
            f_export_table_to_text(server, database, username, password, 'VIEW_Monthly_VC_Data', full_file_name,',')
            
    except Exception as e:
        # print(f"Erreur lors de l'exécution de la requête: {e}")
        check4_label.configure(text=f"Erreur lors de l'exécution de la requête: {e}")

    finally:
        # Fermeture de la connexion
        connection.close()
        print(f"Exportation de {project} réussi")

        # Calculer la durée de l'exportation
        date_end = datetime.now()
        duration = date_end - date_start
        total_seconds = duration.total_seconds()
        duration_text = f_duration_text(total_seconds)
        # new_entry = pd.DataFrame([{'file': file_name, 'time': duration_text}])
        # if not new_entry.dropna(how='all').empty:
        #     df_files = pd.concat([df_files, new_entry], ignore_index=True)
        # Afficher le dataframe
        # print(df_files)
        
        check4_label.configure(text=duration_text, text_color='white')
        list_export_file_sftp.append(full_file_name)
    
def f_check5(export_path):
    project = 'pivot_table'
    
    # Paramètres de connexion à la base de données
    database = os.getenv('SQL_DB_VTQ')
    server = os.getenv('SQL_SERVER')
    username = os.getenv('SQL_USER')
    password = os.getenv('SQL_PASSWORD')    

    # mois du traitement
    period = datetime.strptime(droplist_var.get(), "%d/%m/%Y").strftime("%y%m")
    
    # # dataframe log
    # df_files = pd.DataFrame(columns=['file','time'])
    
    # Paramètres de connexion à la base de données
    connection_string = f'mssql+pymssql://{username}:{password}@{server}/{database}'
    engine = create_engine(connection_string)
    # Connexion à la base de données
    connection = engine.connect()

    # Connexion à la base de données
    try:
        with engine.connect() as connection:
            date_start = datetime.now()
            file_name = project
            full_file_name = export_path + '/' + file_name + '_' + period + '.xlsx'
            # print(full_file_name)
            f_export_table_to_excel(server, database, username, password, 'SALES_PIVOT', full_file_name,'pivot table')
                        
    except Exception as e:
        # print(f"Erreur lors de l'exécution de la requête: {e}")
        check5_label.configure(text=f"Erreur lors de l'exécution de la requête: {e}")
        
    finally:
        # Fermeture de la connexion
        connection.close()
        print(f"Exportation de {project} réussi")

        # Calculer la durée de l'exportation
        date_end = datetime.now()
        duration = date_end - date_start
        total_seconds = duration.total_seconds()
        duration_text = f_duration_text(total_seconds)
        # new_entry = pd.DataFrame([{'file': file_name, 'time': duration_text}])
        # if not new_entry.dropna(how='all').empty:
        #     df_files = pd.concat([df_files, new_entry], ignore_index=True)
        # # Afficher le dataframe
        # print(df_files)
    
        check5_label.configure(text=duration_text, text_color='white')
        list_export_file_sftp.append(full_file_name) 
    
def f_check6(export_path):
    project = 'ProductInOrder_GB'
    
    # Paramètres de connexion à la base de données
    database = os.getenv('SQL_DB_VTQ')
    server = os.getenv('SQL_SERVER')
    username = os.getenv('SQL_USER')
    password = os.getenv('SQL_PASSWORD')    

    # mois du traitement
    period = datetime.strptime(droplist_var.get(), "%d/%m/%Y").strftime("%y%m")
    
    # # dataframe log
    # df_files = pd.DataFrame(columns=['file','time'])
    
    # Paramètres de connexion à la base de données
    connection_string = f'mssql+pymssql://{username}:{password}@{server}/{database}'
    engine = create_engine(connection_string)
    # Connexion à la base de données
    connection = engine.connect()

    # Connexion à la base de données
    try:
        with engine.connect() as connection:
            _year = datetime.now().year
            _month = datetime.now().month
            _day = datetime.now().day
            _hour = datetime.now().hour
            _minute = datetime.now().minute
            _ms = datetime.now().microsecond
            timestamp = str(_year)+str(_month)+str(_day)+str(_hour)+str(_minute)+str(_ms)
            date_start = datetime.now()
            file_name = project
            full_file_name = export_path + '/' + file_name + '_' + timestamp + '.csv'
            f_export_table_to_text(server, database, username, password, 'EXPORT_CREATIO', full_file_name,';')                    
            
    except Exception as e:
        # print(f"Erreur lors de l'exécution de la requête: {e}")
        check6_label.configure(text=f"Erreur lors de l'exécution de la requête: {e}")

    finally:
        # Fermeture de la connexion
        connection.close()
        print(f"Exportation de {project} réussi")

        # Calculer la durée de l'exportation
        date_end = datetime.now()
        duration = date_end - date_start
        total_seconds = duration.total_seconds()
        duration_text = f_duration_text(total_seconds)
        # new_entry = pd.DataFrame([{'file': file_name, 'time': duration_text}])
        # if not new_entry.dropna(how='all').empty:
        #     df_files = pd.concat([df_files, new_entry], ignore_index=True)
        # # Afficher le dataframe
        # print(df_files)   
    
        check6_label.configure(text=duration_text, text_color='white')
        list_export_file_sftp.append(full_file_name)
        
def f_check7(export_path):
    project = 'VTQ New Accounts week'
    # check7_label.configure(text=title_running)
    
    # Paramètres de connexion à la base de données
    database = os.getenv('SQL_DB_VTQ')
    server = os.getenv('SQL_SERVER')
    username = os.getenv('SQL_USER')
    password = os.getenv('SQL_PASSWORD')    

    # mois du traitement
    period = datetime.strptime(droplist_var.get(), "%d/%m/%Y").strftime("%y%m")
    
    # # dataframe log
    # df_files = pd.DataFrame(columns=['file','time'])
    
    # Paramètres de connexion à la base de données
    connection_string = f'mssql+pymssql://{username}:{password}@{server}/{database}'
    engine = create_engine(connection_string)
    # Connexion à la base de données
    connection = engine.connect()

    # Connexion à la base de données
    try:
        with engine.connect() as connection:
            _day = datetime.now().day
            date_start = datetime.now()
            file_name = project
            full_file_name = export_path + '/' + file_name + ' ' + period + str(_day) + '.xlsx'
            # print(full_file_name)
            f_export_table_to_excel(server, database, username, password, 'EXPORT_NEW_ACCOUNTS_WEEK', full_file_name,'New customers')
                        
    except Exception as e:
        # print(f"Erreur lors de l'exécution de la requête: {e}")
        check7_label.configure(text=f"Erreur lors de l'exécution de la requête: {e}")
        
    finally:
        # Fermeture de la connexion
        connection.close()
        print(f"Exportation de {project} réussi")

        # Calculer la durée de l'exportation
        date_end = datetime.now()
        duration = date_end - date_start
        total_seconds = duration.total_seconds()
        duration_text = f_duration_text(total_seconds)
        # new_entry = pd.DataFrame([{'file': file_name, 'time': duration_text}])
        # if not new_entry.dropna(how='all').empty:
        #     df_files = pd.concat([df_files, new_entry], ignore_index=True)
        # # Afficher le dataframe
        # print(df_files)
    
        check7_label.configure(text=duration_text, text_color='white')
        list_export_file_sftp.append(full_file_name)
            
def button_run_func():
    export_path = f_create_directory(droplist_var.get())
    if check1_var.get() == True:
        f_check1(export_path)
    if check2_var.get() == True:
        f_check2(export_path)
    if check3_var.get() == True:
        f_check3(export_path)
    if check4_var.get() == True:
        f_check4(export_path)
    if check5_var.get() == True:
        f_check5(export_path)
    if check6_var.get() == True:
        f_check6(export_path)
    if check7_var.get() == True:
        f_check7(export_path)
    # print("list_export_file_sftp: ", list_export_file_sftp)            
    
def button_all_func():
    check1_var.set('True')
    check2_var.set('True')
    check3_var.set('True')
    check4_var.set('True')
    check5_var.set('True')
    check6_var.set('True')
    check7_var.set('True')
    check1_label.configure(text='')
    check2_label.configure(text='')
    check3_label.configure(text='')
    check4_label.configure(text='')
    check5_label.configure(text='')
    check6_label.configure(text='')
    check7_label.configure(text='')

def button_clear_func():
    check1_var.set('False')
    check2_var.set('False')
    check3_var.set('False')
    check4_var.set('False')
    check5_var.set('False')
    check6_var.set('False')
    check7_var.set('False')
    check1_label.configure(text='')
    check2_label.configure(text='')
    check3_label.configure(text='')
    check4_label.configure(text='')
    check5_label.configure(text='')
    check6_label.configure(text='')
    check7_label.configure(text='')
    
def button_sftp_func(list_export_file_sftp):
    sftp_host = os.getenv('SFTP_HOST')
    sftp_port = int(os.getenv('SFTP_PORT'))
    sftp_user = os.getenv('SFTP_USER')
    sftp_password = os.getenv('SFTP_PASSWORD')
    sftp_folder = os.getenv('SFTP_FOLDER_VTQ')
    
    period = datetime.strptime(droplist_var.get(), "%d/%m/%Y").strftime("%y%m")
    sftp_location = sftp_folder + '/' + period
    
    if not list_export_file_sftp:
        messagebox.showinfo("Information", "Aucun fichier à envoyer")
    else:   
        for export_file in list_export_file_sftp:
            remote_path = f'{sftp_location}/{os.path.basename(export_file)}'  # Chemin distant sur le serveur SFTP
            # Transférer le fichier texte vers le serveur SFTP
            f_upload_to_sftp(export_file, remote_path, sftp_host, sftp_port, sftp_user, sftp_password)
    
def f_select_folder():
    dossier = filedialog.askdirectory()
    if dossier:
        work_path.configure(text=dossier)
            
# Fenêtre principale / Main window    
window = ctk.CTk()
window.title('Export files')
window.geometry('500x700')

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
directory_label = ctk.CTkLabel(master=directory_frame, text='Export to:', font=('Helvetica', 12, 'bold'))
directory_label.pack(side='left', padx=5)
defaut_path = os.getenv('CLIENT_PATH_VTQ')
work_path = ctk.CTkLabel(directory_frame, text=str(defaut_path), wraplength=400)
work_path.pack(side='left', padx=5)
# Bouton pour ouvrir le dialogue de sélection de dossier
button_folder = ctk.CTkButton(master=window, text='Modify', command=f_select_folder, fg_color=color_grey_light, hover_color=color_grey_dark, corner_radius=15, border_color=color_grey_dark, border_width=1, text_color='white', width=100)
button_folder.pack(pady=5)

# dropdown list
# generate the list of the last 12 months (01/MM/YYY)
# 1er jour du mois / 1st day of this month
today = datetime.now()
last_date = today.replace(day=1)   
items = f_generate_dates(last_date)
droplist_var = ctk.StringVar(value=items[0])

droplist_frame = ctk.CTkFrame(master=window)
droplist_frame.pack(pady=5, padx=50, anchor='w')
droplist_label = ctk.CTkLabel(master=droplist_frame, text="Select a month:", font=('Helvetica', 12, 'bold'))
droplist_label.pack(side='left', padx=5)
droplist = ctk.CTkComboBox(master=droplist_frame, variable=droplist_var, values=items)#, state='readonly')
droplist.pack(pady=5)
    
# Selection des fichiers à exporter
data_frame = ctk.CTkFrame(master=window, width=800, border_width=1, corner_radius=10)
data_frame.pack(pady=5, padx=50, anchor='center')
checks_label = ctk.CTkLabel(master=data_frame, text="Select the data to be created:", font=('Helvetica', 12, 'bold'))
checks_label.pack(pady=5, padx=5, anchor='w')

# Frame for buttons
button_frame = ctk.CTkFrame(master=data_frame)
button_frame.pack(pady=5, anchor='w')

# Selectionne tous les fichiers
button_all_string = ctk.StringVar(value='Select all')
button_all = ctk.CTkButton(master=button_frame, textvariable=button_all_string, command=button_all_func, fg_color=color_grey_light, hover_color=color_grey_dark, corner_radius=15, border_color=color_grey_dark, border_width=1, text_color='white')
button_all.pack(side='left', padx=5)

# Annuler la sélection
button_clear_string = ctk.StringVar(value='Clear selection')
button_clear = ctk.CTkButton(master=button_frame, textvariable=button_clear_string, command=button_clear_func, fg_color=color_grey_light, hover_color=color_grey_dark, corner_radius=15, border_color=color_grey_dark, border_width=1, text_color='white')
button_clear.pack(side='left', padx=5)

# checkboxes
def create_checkbox_with_label(frame, text, variable, label_text):
    checkbox_frame = ctk.CTkFrame(master=frame)
    checkbox_frame.pack(pady=2, anchor='w')
    checkbox = ctk.CTkCheckBox(checkbox_frame, text=text, variable=variable)
    checkbox.pack(side='left', padx=20)
    label = ctk.CTkLabel(checkbox_frame, text=label_text, font=('Helvetica', 10, 'italic'))
    label.pack(side='left', padx=5)
    return checkbox, label

check1_text = '00011 File'
check2_text = 'Creatio missing customers'
check3_text = 'Groups VTQ'
check4_text = 'Monthly VC data'
check5_text = 'Pivot spreadsheet'
check6_text = 'Product in Order'
check7_text = 'VTQ New Accounts week'

check1_var = ctk.BooleanVar(value=True)
check1, check1_label = create_checkbox_with_label(data_frame, check1_text, check1_var, '')

check2_var = ctk.BooleanVar(value=True)
check2, check2_label = create_checkbox_with_label(data_frame, check2_text, check2_var, '')

check3_var = ctk.BooleanVar(value=True)
check3, check3_label = create_checkbox_with_label(data_frame, check3_text, check3_var, '')

check4_var = ctk.BooleanVar(value=True)
check4, check4_label = create_checkbox_with_label(data_frame, check4_text, check4_var, '')

check5_var = ctk.BooleanVar(value=True)
check5, check5_label = create_checkbox_with_label(data_frame, check5_text, check5_var, '')

check6_var = ctk.BooleanVar(value=True)
check6, check6_label = create_checkbox_with_label(data_frame, check6_text, check6_var, '')

check7_var = ctk.BooleanVar(value=True)
check7, check7_label = create_checkbox_with_label(data_frame, check7_text, check7_var, '')

# crée les fichiers
button_run_string = ctk.StringVar(value='Export')
button_run = ctk.CTkButton(master=data_frame, textvariable=button_run_string, command=button_run_func, text_color='black', border_color='black', border_width=1, corner_radius=15, width=120)
button_run.pack(pady=10)

# envoi vers le SFTP
button_sftp_string = ctk.StringVar(value='Send to SFTP')
button_sftp = ctk.CTkButton(master=window, textvariable=button_sftp_string, command=lambda: button_sftp_func(list_export_file_sftp), fg_color=color_red_light, hover_color=color_red_dark, corner_radius=15, border_color='#A42408', border_width=1, text_color='white', width=120)
button_sftp.pack(pady=10, anchor='center')

# Add copyright text at the base of the window
copyright_label = ctk.CTkLabel(master=window, text="© ALF DATA, February 2025, v1.0", text_color='white')
copyright_label.pack(side='bottom', pady=10, padx=10, anchor='e')

# run
window.mainloop()
