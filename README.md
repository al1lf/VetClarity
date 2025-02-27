# Client X Data Export

Ce projet est une application GUI (interface graphique utilisateur) développée en Python utilisant `customtkinter` pour exporter des données à partir d'une base de données SQL vers des fichiers Excel ou texte. L'application permet de sélectionner les données à exporter et éventuellement de les envoyer à un serveur SFTP.

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

SQL_DB=your_database_name
SQL_SERVER=your_server_name
SQL_USER=your_username
SQL_PASSWORD=your_password
CLIENT_PATH=your_default_export_path

## Utilisation
Exécutez le script .py
L'application GUI s'ouvrira. Vous pouvez sélectionner le répertoire d'exportation, choisir le mois et les données à exporter, puis cliquer sur le bouton Export pour lancer l'exportation et éventuellement SFTP pour exporter les fichiers vers un répertoire SFTP.

## Fonctionnalités
Sélection du répertoire d'exportation : Choisissez le répertoire où les fichiers exportés seront enregistrés. Par défaut, le répertoire indiqué est CLIENT_PATH.
Sélection du mois : Sélectionnez le mois pour lequel vous souhaitez exporter les données.
Sélection des données à exporter : Cochez les cases pour sélectionner les données à exporter.
Exportation des données : Cliquez sur le bouton Export pour exporter les données sélectionnées vers des fichiers Excel ou texte.
Envoi des fichiers à un serveur SFTP ou HTTPS : Les fichiers exportés peuvent être envoyés à un serveur SFTP ou HTTPS.

## Structure du projet
.py : Le script principal contenant l'application GUI et les fonctions d'exportation des données.
.env : Fichier de configuration contenant les variables d'environnement pour la connexion à la base de données et le chemin d'exportation par défaut.
.bat: fichier d'exécution des .py
.png: fichier imgae
.ico: fichier d'icône
.gitignore: répertoires à ignorer pour Github

## Auteurs
ALF DATA, Février 2025
