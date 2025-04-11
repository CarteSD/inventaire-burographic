# # # # # # # # # # # #
# But : Ce fichier contient l'ensemble des méthodes relatives à la base de données
# Par : Estéban DESESSARD - e.desessard@burographic.fr
# Date : 11/04/2025
# # # # # # # # # # # #

import pyodbc
from constantes import *

# But : Créer une connexion avec la base de données.
#       Retourne la connexion si celle réussie, None dans le cas contraire
def database_connection():
    try:
        if DB_USER and DB_PASSWORD:
            connection = pyodbc.connect(
                f"Driver={DB_DRIVER};"
                f"Server={DB_SERVER};"
                f"Database={DB_NAME};"
                f"UID={DB_USER};"
                f"PWD={DB_PASSWORD};"
            )
        else:
            connection = pyodbc.connect(
                f"Driver={DB_DRIVER};"
                f"Server={DB_SERVER};"
                f"Database={DB_NAME};"
                "Trusted_Connection=yes;"
            )

        if connection:
            print("Connexion réussie à la base de données SQL Server")
            return connection
    except pyodbc.Error as e:
        print(f"Erreur lors de la connexion à SQL Server : {e}")
        return None