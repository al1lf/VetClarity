# Client X Data Export

Ce projet est une application GUI (interface graphique utilisateur) développée en Python utilisant `customtkinter` pour exporter des données à partir d'une base de données SQL vers des fichiers Excel ou texte.  
L'application permet de sélectionner les données à exporter et éventuellement de les envoyer à un serveur SFTP.

## Prérequis

- Python 3.x
- Les bibliothèques Python suivantes :
  - `customtkinter`
  - `datetime`
  - `sqlalchemy`
  - `pandas`
  - `zipfile`
  - `os`
  - `dotenv`
  - `paramiko`
  - `tkinter`
  - `PIL`

## Installation

1. Clonez ce dépôt sur votre machine locale.

2. Installez les dépendances requises en utilisant pip :

pip install -r requirements.txt

3. Créez un fichier .env dans le répertoire racine du projet et ajoutez les variables d'environnement suivantes :

  - SQL_SERVER = your_server_name
  - SQL_USER = your_username
  - SQL_PASSWORD = your_password
  - SQL_DB_xxx = your_database_name
  - CLIENT_PATH_xxx = your_default_export_path
  - SFTP_HOST = your_sftp_directory
  - SFTP_PORT = your_sftp_port
  - SFTP_USER = your_sftp_user
  - SFTP_PASSWORD = your_sftp_password
  - SFTP_FOLDER = your_sftp_folder

## Utilisation

Exécutez le script .py  
L'application GUI s'ouvrira.  
Vous pouvez sélectionner le répertoire d'exportation, choisir le mois et les données à exporter  
Cliquer sur le bouton Export pour lancer l'exportation.  
Eventuellement SFTP ou HTTPS pour exporter les fichiers vers un répertoire SFTP/HTTPS.

## Fonctionnalités

Sélection du répertoire d'exportation : Choisissez le répertoire où les fichiers exportés seront enregistrés. Par défaut, le répertoire indiqué est CLIENT_PATH.  
Sélection du mois : Sélectionnez le mois pour lequel vous souhaitez exporter les données.  
Sélection des données à exporter : Cochez les cases pour sélectionner les données à exporter.  Exportation des données : Cliquez sur le bouton Export pour exporter les données sélectionnées vers des fichiers Excel ou texte.  
Envoi des fichiers à un serveur SFTP ou HTTPS : Les fichiers exportés peuvent être envoyés à un serveur SFTP ou HTTPS.

## Structure du projet

  - .py : Le script principal contenant l'application GUI et les fonctions d'exportation des données.
  - .env : Fichier de configuration contenant les variables d'environnement pour la connexion à la base de données et le chemin d'exportation par défaut.
  - .bat: fichier d'exécution des .py
  - .png: fichier imgae
  - .ico: fichier d'icône
  - .gitignore: répertoires à ignorer pour Github

## Auteurs

ALF DATA, Février 2025


# Client X Data Export

This project is a GUI (Graphical User Interface) application developed in Python using customtkinter to export data from an SQL database to Excel or text files.  
The application allows selecting the data to export and optionally sending them to an SFTP server.

## Prerequisites

- Python 3.x
- Python librairies :
  - `customtkinter`
  - `datetime`
  - `sqlalchemy`
  - `pandas`
  - `zipfile`
  - `os`
  - `dotenv`
  - `paramiko`
  - `tkinter`
  - `PIL`

## Installation

1. Clone this repository to your local machine.

2. Install the required dependencies using pip:

pip install -r requirements.txt

3. Create a .env file in the root directory of the project and add the following environment variables:

  - SQL_SERVER = your_server_name
  - SQL_USER = your_username
  - SQL_PASSWORD = your_password
  - SQL_DB_xxx = your_database_name
  - CLIENT_PATH_xxx = your_default_export_path
  - SFTP_HOST = your_sftp_directory
  - SFTP_PORT = your_sftp_port
  - SFTP_USER = your_sftp_user
  - SFTP_PASSWORD = your_sftp_password
  - SFTP_FOLDER = your_sftp_folder

## Usage

Run the .py script The GUI application will open.  
You can select the export directory, choose the month and data to export.  
Click the Export button to start the export.  
Click the SFTP or HTTPS button to export the files to an SFTP/HTTPS directory.

## Features

Export directory selection: Choose the directory where the exported files will be saved. By default, the directory specified is CLIENT_PATH.  
Month selection: Select the month for which you want to export the data.  
Data selection: Check the boxes to select the data to export.  
Data export: Click the Export button to export the selected data to Excel or text files.  
Sending files to an SFTP or HTTPS server: The exported files can be sent to an SFTP or HTTPS server.

## Project Structure

  - .py: The main script containing the GUI application and data export functions.
  - .env: Configuration file containing environment variables for database connection and default export path.
  - .bat: execution file for .py
  - .png: image file
  - .ico: icon file
  - .gitignore: directories to ignore for Github

## Authors

ALF DATA, February 2025 ```