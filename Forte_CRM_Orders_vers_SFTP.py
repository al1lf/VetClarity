import pandas as pd
from sqlalchemy import create_engine
import paramiko
import os
from dotenv import load_dotenv

# Charger les variables d'environnement à partir du fichier .env
load_dotenv()

path = os.getenv('CLIENT_PATH_FOR')

# Connexion à SQL Server et exécution de la requête
def get_sql_data():
    # Créer l'URL de connexion SQLAlchemy pour SQL Server
    server = os.getenv('SQL_SERVER')
    username = os.getenv('SQL_USER')
    password = os.getenv('SQL_PASSWORD')
    # Connexion à la base
    database = os.getenv('SQL_DB_FOR')

    # Formater l'URL de connexion pour SQLAlchemy (via pymssql)
    connection_url = f'mssql+pymssql://{username}:{password}@{server}/{database}'
    
    # Créer l'engine SQLAlchemy
    engine = create_engine(connection_url)

    # Définir la requête SQL
    query = "SELECT * FROM VIEW_CRM_ORDERS_ACTIVITIES WHERE order_date_yyyy_mm_dd >= (SELECT DATEADD(MONTH,+1,MAX(dates)) FROM SALES)"
    # Exécuter la requête et charger les résultats dans un DataFrame pandas
    df = pd.read_sql(query, engine)

    # Exporter les résultats dans un fichier texte
    export_file = path + '/crm_orders.txt'
    df.to_csv(export_file, index=False)
    return export_file

# Connexion au serveur SFTP et transfert du fichier
def upload_to_sftp(local_file, remote_path, sftp_host, sftp_port, sftp_user, sftp_password):
    try:
        # Créer la connexion SFTP
        transport = paramiko.Transport((sftp_host, sftp_port))
        transport.connect(username=sftp_user, password=sftp_password)

        sftp = paramiko.SFTPClient.from_transport(transport)

        # Transférer le fichier vers le serveur SFTP
        sftp.put(local_file, remote_path)

        print(f"Fichier {local_file} transféré avec succès vers {remote_path}.")

        # Fermer la connexion SFTP
        sftp.close()
        transport.close()

    except Exception as e:
        print(f"Erreur lors du transfert SFTP : {e}")
        
if __name__ == "__main__":
    # Étape 1: Exécuter la requête SQL et créer le fichier texte
    export_file = get_sql_data()

    # Étape 2: Paramètres du serveur SFTP/HTTPS
    sftp_host = os.getenv('SFTP_HOST')
    sftp_port = int(os.getenv('SFTP_PORT'))
    sftp_user = os.getenv('SFTP_USER')
    sftp_password = os.getenv('SFTP_PASSWORD')
    
    sftp_folder = os.getenv('SFTP_FOLDER_FOR')
    sftp_location = sftp_folder
    remote_path = f'{sftp_location}/{os.path.basename(export_file)}'  # Chemin distant sur le serveur SFTP
    # Étape 3: Transférer le fichier texte vers le serveur SFTP
    print(f"Transfert du fichier {export_file} vers {remote_path} avec {sftp_user}")
    upload_to_sftp(export_file, remote_path, sftp_host, sftp_port, sftp_user, sftp_password)
