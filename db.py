# # # # # # # # # # # #
# But : Ce fichier contient l'ensemble des méthodes relatives à la base de données
# Par : Estéban DESESSARD - e.desessard@burographic.fr
# Date : 11/04/2025
# # # # # # # # # # # #

import pyodbc
from utils import *

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

# But : Vérifier si un article existe dans la base de données
def article_exists(connection, item):
    try:
        cursor = connection.cursor()
        query = "SELECT COUNT(*) FROM ElementDef WHERE Code = ?"
        cursor.execute(query, item)
        result = cursor.fetchone()
        if result[0] > 0:
            return True
        else:
            return False

    except pyodbc.Error as e:
        write_log(f"[ERREUR] {str(e)}")
        return False

# But : Vérifier si une famille existe dans la base de données
def famille_exists(connection, famille):
    try:
        cursor = connection.cursor()
        query = "SELECT COUNT(*) FROM FamilleArticle WHERE Code = ?"
        cursor.execute(query, famille + ".")
        result = cursor.fetchone()
        if result[0] > 0:
            return True
        else:
            return False

    except pyodbc.Error as e:
        write_log(f"[ERREUR] {str(e)}")
        return False

# But : Récupérer le code de la famille d'un article
def get_famille(connection, item):
    try:
        cursor = connection.cursor()
        query = "SELECT Famille FROM ElementDef WHERE Code = ?"
        cursor.execute(query, item)
        result = cursor.fetchone()
        if result:
            return result[0].rstrip('.')
        else:
            return None

    except pyodbc.Error as e:
        write_log(f"[ERREUR] {str(e)}")
        return None

# But : Récupérer tous les articles de stock en base de données
def get_all_articles(connection):
    try:
        cursor = connection.cursor()
        query = "SELECT * FROM ElementStock"
        cursor.execute(query)
        result = cursor.fetchall()
        if result:
            return result
        else:
            return None

    except pyodbc.Error as e:
        write_log(f"[ERREUR] {str(e)}")
        return None

# But : Récupérer le stock d'un article
def get_article_stock(connection, id):
    try:
        cursor = connection.cursor()
        query = "SELECT * FROM ElementStock WHERE CodeElem = ?"
        cursor.execute(query, id)
        result = cursor.fetchone()
        if result:
            return result
        else:
            return None

    except pyodbc.Error as e:
        write_log(f"[ERREUR] {str(e)}")
        return None
    
# But : Récupérer la définition d'un article
def get_article_def(connection, id):
    try:
        cursor = connection.cursor()
        query = "SELECT * FROM ElementDef WHERE Code = ?"
        cursor.execute(query, id)
        result = cursor.fetchone()
        if result:
            return result
        else:
            return None

    except pyodbc.Error as e:
        write_log(f"[ERREUR] {str(e)}")
        return None

# But : Créer un mouvement de stock sur un article défini
#
#       Afin d'assurer la cohérence des données, met à jour la table ElementStock
#       pour garantir que les quantités restent cohérentes entre elles
def create_mvt(connection, typeMvt, article, quantite):
    try:
        cursor = connection.cursor()
        # Insertion du mouvement de stock
        query = "INSERT INTO ElementMvtStock (CodeElem, TypeMvt, Provenance, Date, Quantite, PA, Info) VALUES (?, ?, ?, ?, ?, ?, ?)"
        info = f"Inventaire manuel du {datetime.today().strftime('%d/%m/%Y')}"
        cursor.execute(query, [article[0], typeMvt, 'M', datetime.today(), quantite, article[5], info])
        connection.commit()

        # Mise à jour du stock de l'élément
        if typeMvt == 'E':
            query = "UPDATE ElementStock SET QttAppro = QttAppro + ? WHERE CodeElem = ?"
            cursor.execute(query, [quantite, article[0]])

        elif typeMvt == 'S':
            query = "UPDATE ElementStock SET QttConso = QttConso + ? WHERE CodeElem = ?"
            cursor.execute(query, [quantite, article[0]])
        connection.commit()

        return True
    
    except pyodbc.Error as e:
        write_log(f"[ERREUR] {str(e)}")
        return False