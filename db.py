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
            connection.autocommit = False
            print("Connexion réussie à la base de données SQL Server")
            return connection
    except pyodbc.Error as e:
        print(f"Erreur lors de la connexion à SQL Server : {e}")
        return None

# But : Vérifier si un article existe dans la base de données
def article_exists(connection, num_commercial):
    try:
        cursor = connection.cursor()
        query = "SELECT COUNT(*) FROM ElementDef WHERE NumCommercialGlobal = ?"
        cursor.execute(query, num_commercial)
        result = cursor.fetchone()
        return result[0] > 0

    except pyodbc.Error as e:
        write_log(f"[ERREUR] {str(e)}")
        return False

# But : Récupérer le code de la famille d'un article
def get_family(connection, num_commercial):
    try:
        cursor = connection.cursor()
        query = "SELECT FA.Code, FA.Libelle FROM FamilleArticle FA JOIN ElementDef ED ON FA.Code = ED.Famille WHERE ED.NumCommercialGlobal = ?"
        cursor.execute(query, num_commercial)
        result = cursor.fetchone()
        if result:
            return result
        else:
            return None

    except pyodbc.Error as e:
        write_log(f"[ERREUR] {str(e)}")
        return None

# But : Récupérer tous les articles de stock en base de données
def get_all_articles(connection):
    try:
        cursor = connection.cursor()
        query = "SELECT ES.*, ED.NumCommercialGlobal FROM ElementStock ES JOIN ElementDef ED ON ES.CodeElem = ED.Code"
        cursor.execute(query)
        result = cursor.fetchall()
        if result:
            return result
        else:
            return None

    except pyodbc.Error as e:
        write_log(f"[ERREUR] {str(e)}")
        return None

# But : Récupérer le stock d'un article à partir de son numéro commercial
def get_article_stock(connection, commercial_num):
    try:
        cursor = connection.cursor()
        query = "SELECT ES.*, ED.NumCommercialGlobal, ED.LibelleStd FROM ElementStock ES JOIN ElementDef ED ON ES.CodeElem = ED.Code WHERE ED.NumCommercialGlobal = ?"
        cursor.execute(query, commercial_num)
        result = cursor.fetchone()
        if result:
            return result
        else:
            return None

    except pyodbc.Error as e:
        write_log(f"[ERREUR] {str(e)}")
        return None
    
# But : Récupérer la définition d'un article
def get_article_name(connection, commercial_num):
    try:
        cursor = connection.cursor()
        query = "SELECT LibelleStd FROM ElementDef WHERE NumCommercialGlobal = ?"
        cursor.execute(query, commercial_num)
        result = cursor.fetchone()
        if result:
            return result[0]
        else:
            return None

    except pyodbc.Error as e:
        write_log(f"[ERREUR] {str(e)}")
        return None

# But : Créer un mouvement de stock sur un article défini
#
#       Afin d'assurer la cohérence des données, met à jour la table ElementStock
#       pour garantir que les quantités restent cohérentes entre elles
def create_movement(connection, movement_type, article, quantite):
    try:
        cursor = connection.cursor()
        # Insertion du mouvement de stock
        query = "INSERT INTO ElementMvtStock (CodeElem, TypeMvt, Provenance, Date, Quantite, PA, Info) VALUES (?, ?, ?, ?, ?, ?, ?)"
        info = f"Inventaire manuel du {datetime.today().strftime('%d/%m/%Y')}"
        cursor.execute(query, [article[0], movement_type, 'M', datetime.today(), quantite, article[5], info])

        # Mise à jour du stock de l'élément
        if movement_type == 'E':
            query = "UPDATE ElementStock SET QttAppro = QttAppro + ? WHERE CodeElem = ?"
            cursor.execute(query, [quantite, article[0]])

        elif movement_type == 'S':
            query = "UPDATE ElementStock SET QttConso = QttConso + ? WHERE CodeElem = ?"
            cursor.execute(query, [quantite, article[0]])

        return True
    
    except pyodbc.Error as e:
        write_log(f"[ERREUR] {str(e)}")
        return False