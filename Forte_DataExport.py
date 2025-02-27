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
from tkinter import messagebox, filedialog
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
        date = end_date - timedelta(days=i * 30)  # Approximativement 1 mois / Approximately 1 month
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

# Créer un fichier .zip / Create a .zip file
def f_zip_files(file_path, archive_name):
    with zipfile.ZipFile(archive_name, 'w') as zipf:
        # Parcourir les fichiers à zipper / Browse files to zip
        for fichier in file_path:
            # Ajouter chaque fichier à l'archive zip / Add each file to the zip archive
            zipf.write(fichier, os.path.basename(fichier))
            
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
        
        
list_export_file_sftp = []
list_export_file_https = []

# functions checkboxes
def f_check1(export_path):
    project = 'CorporateRebates'
    
    # Paramètres de connexion à la base de données
    database = os.getenv('SQL_DB_FOR')
    server = os.getenv('SQL_SERVER')
    username = os.getenv('SQL_USER')
    password = os.getenv('SQL_PASSWORD')    

    # mois du traitement
    period = droplist_var.get()
    
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
            # liste les fichiers créé pour le zip final
            files_to_zip = []
            # Exécutez une requête SQL pour sélectionner toutes les lignes de la table spécifiée
            query_groups = text("SELECT DISTINCT groupName FROM DISCOUNTS_CORPORATE WHERE rebate IS NOT NULL ORDER BY 1")
            df_groups = pd.read_sql(query_groups, connection)
            if df_groups.empty:
                # print("No Groups")
                check1_label.configure(text="No group")
            else:
                i = 0
                for group in df_groups["groupName"]:
                    date_start = datetime.now()
                    
                    # print('group :',group)
                    clean_value = group.replace(' ','_') # remplacer " " (espace) par un "_" pour le nom du fichier
                    file_name = project + datetime.strptime(droplist_var.get(), "%d/%m/%Y").strftime("%Y%m") + '_' + clean_value + '.xlsx' # année sur 4 caractères
                    full_file_name = os.path.join(export_path, f'{file_name}')
                    
                    query = f"\
                    SELECT FORMAT(dates,'dd/MM/yyyy') AS dates, idCustomer AS CCODE, customerName AS Name, postcode AS [Post code],\
                        businessClinicCode, productName AS [Product Name], price, q AS Quantity, net AS [NET Sales],\
                        rebate AS [NET REBATE], discount AS [Rebate %]\
                    FROM DISCOUNTS_CORPORATE\
                    WHERE groupName='"+group+"'"
                    dataframe = pd.read_sql(query, connection)
                                    
                    # Modification de format de certains champs
                    # dataframe['NET Sales'] = dataframe['NET Sales'].astype(float)
                    # dataframe['NET REBATE'] = dataframe['NET REBATE'].astype(float)
                    # dataframe['Rebate %'] = dataframe['Rebate %'].astype(float)
                    # dataframe['price'] = round(dataframe['price'],4)
                    
                    # Affichez le DataFrame
                    # print(dataframe.head())
                    
                    # Exportez les données dans un fichier Excel
                    if group == 'VETS4PETS': # Gestion du complément totaux par clients pour VETS4PETS
                        
                        # ne garder que les sous-totaux pour le "summary"
                        dataframe_summary = dataframe[dataframe['Product Name'].isnull()]
                        # supprimer les colonnes liées aux produits
                        dataframe_summary = dataframe_summary.drop(['Product Name','price'], axis=1)
                        
                        # supprimer les sous-totaux
                        dataframe = dataframe[dataframe['Product Name'].notnull()]
                        
                        # avec 2 feuilles
                        with pd.ExcelWriter(full_file_name) as writer:
                            dataframe.to_excel(writer, sheet_name=group, index=False)
                            dataframe_summary.to_excel(writer, sheet_name=group+'_Summary', index=False)
                        
                    else:
                        dataframe.to_excel(full_file_name, sheet_name=clean_value, index=False)
                    
                    # Ajouter le fichier Excel dans une liste pour la création d'un fichier zip final
                    files_to_zip.append(full_file_name)
                    
                    i += 1
        
                # créer un fichier zip avec tous les fichiers Excel
                archive_zip = export_path + '/' + project + datetime.strptime(droplist_var.get(), "%d/%m/%Y").strftime("%y%m") + '.zip'
                
                if os.path.exists(archive_zip):
                    os.remove(archive_zip)

                f_zip_files(files_to_zip, archive_zip)

                # puis supprimer ces fichiers zippés
                for files in files_to_zip:
                    if os.path.exists(files):
                        os.remove(files)
                        # print(f"Le fichier {files} a été supprimé.")
                        check1_label.configure(text=f"Le fichier {files} a été supprimé.")
                    else:
                        # print(f"Le fichier {files} n'existe pas.")
                        check1_label.configure(text=f"Le fichier {files} n'existe pas.")
                    
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
        list_export_file_https.append(archive_zip)
    
def f_check2(export_path):
    project = 'CRMGroupWholesaler'
    
    # Paramètres de connexion à la base de données
    database = os.getenv('SQL_DB_FOR')
    server = os.getenv('SQL_SERVER')
    username = os.getenv('SQL_USER')
    password = os.getenv('SQL_PASSWORD')    

    # mois du traitement
    period = droplist_var.get()
    
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
            period = datetime.strptime(droplist_var.get(), "%d/%m/%Y").strftime("%y%m")
            # Exécutez une requête SQL pour sélectionner toutes les lignes de la table spécifiée
            query = f"SELECT c.id_customer AS 'CCODE', c.name1 AS 'Customer Name', c.postcode AS 'Postcode', c.rep AS 'TM Area', c.buyingGroup, c.FTVet AS 'Full time vets', practiceCategory AS 'LA /CA etc' \
                        , ISNULL(ROUND(c.sMATValue,2),0) AS 'Turnover before discounts MAT' \
                        , ISNULL(ROUND(s.reconcileQtr,2),0) AS 'Reconcile before discounts last 3 months' \
                        , t.last_date, t.colonne \
                    FROM VIEW_CUSTOMERS_DETAILS_LIVE c \
                    LEFT JOIN ( \
                        SELECT id_customer, last_date, 's - ' + wholesalerName AS 'colonne' FROM VIEW_CRM_ORDERS_LAST_SALES_BY_WS WHERE wholesalerName<>'supplement' \
                        UNION SELECT id_customer, last_date, 's - Last Sales' FROM VIEW_CRM_ORDERS_LAST_SALES \
                        UNION SELECT id_customer, last_date, 'o - Last Order' FROM VIEW_CRM_ORDERS_LAST_ORDER \
                        UNION SELECT id_customer, MAX(last_date), 'a - last activity' FROM VIEW_CRM_ORDERS_LAST_ACTIVITY GROUP BY id_customer \
                        UNION SELECT id_customer, last_date, 'a - ' + activityType COLLATE Latin1_General_100_CS_AS FROM VIEW_CRM_ORDERS_LAST_ACTIVITY WHERE activityType IS NOT NULL \
                        ) t ON t.id_customer=c.id_customer \
                    LEFT JOIN ( \
                        SELECT idCustomer, SUM(priceValQtr) AS 'reconcileQtr' \
                        FROM SALES s \
                        INNER JOIN PRODUCTS p ON p.idProduct=s.idProduct \
                        WHERE FORMAT(dates,'yyMM') = " + period + " \
                        AND p.brand='RECONCILE' \
                        GROUP BY idCustomer \
                        ) s ON s.idCustomer=c.id_customer COLLATE Latin1_General_100_CS_AS \
                    WHERE c.isActive=1 \
                    ORDER BY 1"
            
            dataframe = pd.read_sql(query, connection)
            # Affichez le DataFrame
            # print(dataframe.head())

            # Export du DataFrame dans un fichier csv
            date_start = datetime.now()

            pivoted_df = dataframe.pivot(index=['CCODE','Postcode','Customer Name','LA /CA etc','Full time vets','TM Area','buyingGroup','Turnover before discounts MAT','Reconcile before discounts last 3 months'], columns='colonne', values='last_date')
            pivoted_df.columns = [col[4:] for col in pivoted_df.columns] # renomme les colonnes pour supprimer le préfixe permettant le tri des colonnes pivotées ('a - ', 'o - ' ou 's - ')
            pivoted_df.reset_index(inplace=True) # supprime l'index...
            # column_value = pivoted_df.pop('value') # ...pour pouvoir supprimer la colonne 'Value'...
            # pivoted_df['value'] = column_value # ... et la replacer à la fin du dataframe
            # fixer l'ordre des colonnes
            ordered_columns = ['CCODE','Postcode','Customer Name','LA /CA etc','Full time vets','TM Area','buyingGroup','1-2-1','Account Review Meeting','Cake/Fruit & Learn','Cold Call', 'Email Activity','External Call','Intern. Technical Enquiry','L & L','Pharmacovigilance','Phone Meeting','Technical Enquiry','Technical Visit','Text Message','Video Meeting','last activity','Turnover before discounts MAT','Reconcile before discounts last 3 months','Last Order','Last Sales','ARK','Acravet','Agrihealth','Broomhall','C&M','Centaur','Chanelle','Direct Sales Ireland','Direct Sales UK','Henry Schein','Merlin','NVS','Uniphar','VSSCo']
            ordered_df = pivoted_df[ordered_columns]
            
            file_name = project
            full_file_name = os.path.join(export_path, f'{file_name}_{period}.txt')
            
            # exporter vers un fichier texte...
            ordered_df.to_csv(full_file_name, sep='\t', quoting=1, index=False)
            # ... ou Excel
            # pivoted_df.to_excel(full_file_name+'.xlsx',index=False)
            
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
    project = 'DiscountOutput'
    
    # Paramètres de connexion à la base de données
    database = os.getenv('SQL_DB_FOR')
    server = os.getenv('SQL_SERVER')
    username = os.getenv('SQL_USER')
    password = os.getenv('SQL_PASSWORD')    

    # mois du traitement
    period = droplist_var.get()
    
    # dataframe log
    df_files = pd.DataFrame(columns=['file','time'])
    
    # Paramètres de connexion à la base de données
    connection_string = f'mssql+pymssql://{username}:{password}@{server}/{database}'
    engine = create_engine(connection_string)
    # Connexion à la base de données
    connection = engine.connect()

    # Connexion à la base de données
    try:
        with engine.connect() as connection:
            # Exécutez une requête SQL pour sélectionner toutes les lignes de la table spécifiée
            query = f"SELECT [Geo],[TM],[Customer Code],[Customer Name],[Product Code],[Forte Discount] FROM TBL_DISCOUNT_OUTPUT"
            
            dataframe = pd.read_sql(query, connection)
            dataframe['Forte Discount'] = dataframe['Forte Discount'].apply(lambda x: f"{x:.6f}")
            # # Affichez le DataFrame
            # print(dataframe.head())

            # Export du DataFrame dans un fichier csv
            date_start = datetime.now()
            file_name = project
            period = datetime.strptime(droplist_var.get(), "%d/%m/%Y").strftime("%y%m")
            full_file_name = os.path.join(export_path, f'{file_name}_{period}.txt')
            dataframe.to_csv(full_file_name, sep='\t', quoting=1, index=False)
            
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
    project = 'Pharmacovigilance'
    
    # Paramètres de connexion à la base de données
    database = os.getenv('SQL_DB_FOR')
    server = os.getenv('SQL_SERVER')
    username = os.getenv('SQL_USER')
    password = os.getenv('SQL_PASSWORD')    

    # mois du traitement
    period = droplist_var.get()
    
    # dataframe log
    df_files = pd.DataFrame(columns=['file','time'])
    
    # Paramètres de connexion à la base de données
    connection_string = f'mssql+pymssql://{username}:{password}@{server}/{database}'
    engine = create_engine(connection_string)
    # Connexion à la base de données
    connection = engine.connect()

    # Connexion à la base de données
    try:
        with engine.connect() as connection:
            # Exécutez une requête SQL pour sélectionner toutes les lignes de la table spécifiée
            period = datetime.strptime(droplist_var.get(), "%d/%m/%Y").strftime("%y%m")
            query = f"SELECT * FROM VIEW_PHARMACOVIGILANCE WHERE FORMAT(CAST(date_time AS DATETIME2),'yyMM')='{period}'"
            
            dataframe = pd.read_sql(query, connection)
            # Affichez le DataFrame
            # print(dataframe.head())

            # Export du DataFrame dans un fichier csv
            date_start = datetime.now()
            file_name = project
            full_file_name = os.path.join(export_path, f'{file_name}_{period}.txt')
            dataframe.to_csv(full_file_name, sep='\t', quoting=1, index=False)
            
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
    project = 'PivotByTerritory'
    
    # Paramètres de connexion à la base de données
    database = os.getenv('SQL_DB_FOR')
    server = os.getenv('SQL_SERVER')
    username = os.getenv('SQL_USER')
    password = os.getenv('SQL_PASSWORD')    

    # mois du traitement
    period = droplist_var.get()
    
    # dataframe log
    df_files = pd.DataFrame(columns=['file','time'])
    
    # Paramètres de connexion à la base de données
    connection_string = f'mssql+pymssql://{username}:{password}@{server}/{database}'
    engine = create_engine(connection_string)
    # Connexion à la base de données
    connection = engine.connect()

    # Connexion à la base de données
    try:
        with engine.connect() as connection:
            # liste les fichiers créé pour le zip final
            files_to_zip = []
            # Exécutez une requête SQL pour sélectionner toutes les lignes de la table spécifiée
            query_territories = text("SELECT DISTINCT territoryName FROM SALES_PIVOT WHERE cie='FOR' AND territoryName IS NOT NULL ORDER BY 1")
            df_territories  = pd.read_sql(query_territories, connection)#connection.execute(query0)
            if df_territories.empty:
                # print("No Territories")
                check5_label.configure(text=f"No territory")
            else:
                i = 0
                for territory in df_territories["territoryName"]:
                    date_start = datetime.now()
                    
                    clean_value = territory.replace(' ','_') # remplacer " " (espace) par un "_" pour le nom du fichier
                    file_name = 'My_Data.xlsx'
                    period = datetime.strptime(droplist_var.get(), "%d/%m/%Y").strftime("%y%m")
                    full_file_name = os.path.join(export_path, f'{file_name}')                
                    file_zip = export_path + '/My_Data_' + period + '_' + clean_value + '.zip'
            
                    if os.path.exists(full_file_name):
                        os.remove(full_file_name)
                        
                    query = f"SELECT * FROM VIEW_SALES_PIVOT WHERE REPname='"+territory+"'"
                    dataframe = pd.read_sql(query, connection)
                    # Affichez le DataFrame
                    # print(dataframe)
                
                    dataframe.to_excel(full_file_name, sheet_name=territory, index=False)
                    
                    # Ajouter le fichier zip dans une liste pour la création d'un fichier zip final
                    files_to_zip.append(file_zip)
                    
                    # créer un fichier zip 
                    list_zip = []
                    list_zip.append(full_file_name)
                    f_zip_files(list_zip, file_zip)
                    
                    # date_end = datetime.now()
                    # duration = date_end - date_start
                    # new_entry = pd.DataFrame([{'file': clean_value,'time': date_end - date_start}])
                    # if not new_entry.dropna(how='all').empty:
                    #     df_files = pd.concat([df_files, new_entry], ignore_index=True)
                    i += 1
                    
            # supprimer le fichier 'My_Data.xlsx' restant
            if os.path.exists(full_file_name):
                os.remove(full_file_name)
        
            # créer un fichier zip avec tous les fichiers Excel
            archive_zip = export_path + '/' + project + period + '.zip'
            if os.path.exists(archive_zip):
                os.remove(archive_zip)

            f_zip_files(files_to_zip, archive_zip)

            # puis supprimer ces fichiers zippés
            for files in files_to_zip:
                if os.path.exists(files):
                    os.remove(files)
                    # print(f"Le fichier {files} a été supprimé.")
                else:
                    print(f"Le fichier {files} n'existe pas.")
                        
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
        list_export_file_https.append(archive_zip)
    
def f_check6(export_path):
    project = 'WholesalersSales'
    
    # Paramètres de connexion à la base de données
    database = os.getenv('SQL_DB_FOR')
    server = os.getenv('SQL_SERVER')
    username = os.getenv('SQL_USER')
    password = os.getenv('SQL_PASSWORD')    

    # mois du traitement
    period = droplist_var.get()
    
    # dataframe log
    df_files = pd.DataFrame(columns=['file','time'])
    
    # Paramètres de connexion à la base de données
    connection_string = f'mssql+pymssql://{username}:{password}@{server}/{database}'
    engine = create_engine(connection_string)
    # Connexion à la base de données
    connection = engine.connect()

    # Connexion à la base de données
    try:
        with engine.connect() as connection:
            # Exécutez une requête SQL pour sélectionner toutes les lignes de la table spécifiée
            query = f"SELECT dates AS 'DATES', YEAR(dates) AS 'year', MONTH(dates) AS 'month', \
                        idWholesaler AS 'DIS', wholesalerName AS 'DISname', \
                        groupName AS 'GName', \
                        idCustomer AS 'CCODE', customerName AS 'cname', add1, add2, city, postcode, province, country AS 'region', \
                        territoryName AS 'REPname',  customerCategory AS 'CustomerCategory', \
                        idProduct AS 'PCODE', productName AS 'PNAME', brand AS 'BRAND', category AS 'Category', species AS 'Species', hierarchy1 AS 'HIERARCHY1', \
                        foc AS 'FOC', \
                        q, priceVal AS 'Val', wsVal AS 'ws_val', valDiscByWs, valDiscByGroup, netSales,  \
                        cie AS 'Cie' \
                    FROM SALES_PIVOT"
            
            dataframe = pd.read_sql(query, connection)
            # Affichez le DataFrame
            # print(dataframe.head())

            # Export du DataFrame dans un fichier csv
            date_start = datetime.now()
            file_name = 'wholesalers_sales' #project
            full_file_name = os.path.join(export_path, f'{file_name}.txt')
            dataframe.to_csv(full_file_name, sep='\t', quoting=1, index=False)
            
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

def button_all_func():
    check1_var.set('True')
    check2_var.set('True')
    check3_var.set('True')
    check4_var.set('True')
    check5_var.set('True')
    check6_var.set('True')
    check1_label.configure(text='')
    check2_label.configure(text='')
    check3_label.configure(text='')
    check4_label.configure(text='')
    check5_label.configure(text='')
    check6_label.configure(text='')
            
def button_clear_func():
    check1_var.set('False')
    check2_var.set('False')
    check3_var.set('False')
    check4_var.set('False')
    check5_var.set('False')
    check6_var.set('False')
    check1_label.configure(text='')
    check2_label.configure(text='')
    check3_label.configure(text='')
    check4_label.configure(text='')
    check5_label.configure(text='')
    check6_label.configure(text='')

def button_sftp_https_func(list_export_file, export_type):
    sftp_host = os.getenv('SFTP_HOST')
    sftp_port = int(os.getenv('SFTP_PORT'))
    if export_type == 'sftp':
        sftp_user = os.getenv('SFTP_USER')
    else:
        sftp_user = os.getenv('HTTPS_USER')        
    sftp_password = os.getenv('SFTP_PASSWORD')
    sftp_folder = os.getenv('SFTP_FOLDER_FOR')
    
    # period = datetime.strptime(droplist_var.get(), "%d/%m/%Y").strftime("%y%m")
    sftp_location = sftp_folder # + '/' + period
    
    if not list_export_file:
        messagebox.showinfo("Information", "Aucun fichier à envoyer")
    else:   
        for export_file in list_export_file:
            remote_path = f'{sftp_location}/{os.path.basename(export_file)}'  # Chemin distant sur le serveur SFTP
            # Transférer le fichier texte vers le serveur SFTP
            # print(f"Transfert du fichier {export_file} vers {remote_path} avec {sftp_user}")
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
set_default_color_theme("blue")
color_grey_dark ='#555656'
color_grey_light ='#a9a9a9'
color_red_light = '#FF5733'
color_red_dark = '#A42408'

# window.minsize(400,500)
# window.maxsize(800,800)
window.resizable(False, False)

window.bind('<Escape>', lambda event: window.quit())

# title
logo_img_data = Image.open("Forte.png")
logo_img = CTkImage(dark_image=logo_img_data, light_image=logo_img_data, size=(128,55))
# title_text = '\nForte Data Export'
title_label = ctk.CTkLabel(master=window, image=logo_img, text='', text_color='white', font=('Helvetica', 15, 'bold'))
title_label.pack(pady=20)

# work directory by default
directory_frame = ctk.CTkFrame(master=window, width=800)
directory_frame.pack(pady=5, padx=50, anchor='w')
directory_label = ctk.CTkLabel(master=directory_frame, text='Export to:', font=('Helvetica', 12, 'bold'))
directory_label.pack(side='left', padx=5)
defaut_path = os.getenv('CLIENT_PATH_FOR')
work_path = ctk.CTkLabel(directory_frame, text=str(defaut_path), wraplength=400)
work_path.pack(side='left', padx=5)
# Bouton pour ouvrir le dialogue de sélection de dossier
button_folder = ctk.CTkButton(directory_frame, text='Modify', command=f_select_folder, fg_color=color_grey_light, hover_color=color_grey_dark, corner_radius=15, border_color=color_grey_dark, border_width=1, text_color='white', width=100)
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
droplist = ctk.CTkComboBox(master=droplist_frame, variable=droplist_var, values=items, state='readonly')
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

check1_text = 'Corporate rebates **' # https
check2_text = 'CRM group wholesaler *' # sftp
check3_text = 'Discount output *' # sftp
check4_text = 'Pharmacovigilance *' # sftp
check5_text = 'Pivot by territory **' # https
check6_text = 'Wholesaler sales *' # sftp

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

# crée les fichiers
button_run_string = ctk.StringVar(value='Export')
button_run = ctk.CTkButton(master=data_frame, textvariable=button_run_string, command=button_run_func, text_color='white', border_color='black', border_width=1, corner_radius=15, width=120)
button_run.pack(pady=10)

# Frame for buttons
button_frame_2 = ctk.CTkFrame(master=data_frame)
button_frame_2.pack(pady=5, anchor='w')

# envoi vers le SFTP
button_sftp_string = ctk.StringVar(value='* Send to SFTP')
button_sftp = ctk.CTkButton(master=button_frame_2, textvariable=button_sftp_string, command=lambda: button_sftp_https_func(list_export_file_sftp,'sftp'), fg_color=color_red_light, hover_color=color_red_dark, corner_radius=15, border_color='#A42408', border_width=1, text_color='white', width=120)
button_sftp.pack(side='left', padx=20)

# envoi vers le HTTPS
button_https_string = ctk.StringVar(value='** Send to HTTPS')
button_https = ctk.CTkButton(master=button_frame_2, textvariable=button_https_string, command=lambda: button_sftp_https_func(list_export_file_https,'https'), fg_color=color_red_light, hover_color=color_red_dark, corner_radius=15, border_color='#A42408', border_width=1, text_color='white', width=120)
button_https.pack(side='right', padx=20)

# Add copyright text at the base of the window
copyright_label = ctk.CTkLabel(master=window, text="© ALF DATA, February 2025, v1.0", text_color='white')
copyright_label.pack(side='bottom', pady=10, padx=10, anchor='e')

# run
window.mainloop()
