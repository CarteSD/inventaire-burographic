"""
    But : Ce fichier contient l'ensemble des méthodes relatives à la base de données
    Par : Estéban DESESSARD - e.desessard@burographic.fr
    Date : 11/04/2025
"""

import pyodbc
from utils import *

def database_connection():
    """
    Crée une connexion à la base de données.

    Args:
        Aucun
    
    Returns:
        connection (pyodbc.Connection): La connexion à la base de données si réussie, None sinon.

    Raises:
        pyodbc.Error: Si une erreur se produit lors de la connexion à la base de données.
    """
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

def article_exists(connection, num_commercial):
    """
    Vérifie si un article existe dans la base de données.
    
    Args:
        connection (pyodbc.Connection): La connexion à la base de données.
        num_commercial (str): Le numéro commercial de l'article à vérifier.

    Returns:
        bool: True si l'article existe, False sinon.

    Raises:
        pyodbc.Error: Si une erreur se produit lors de la requête.
    """
    try:
        cursor = connection.cursor()
        query = "SELECT COUNT(*) FROM ElementDef WHERE NumCommercialGlobal = ?"
        cursor.execute(query, num_commercial)
        result = cursor.fetchone()
        return result[0] > 0

    except pyodbc.Error as e:
        write_log(f"[ERREUR] {str(e)}")
        return False

def get_family(connection, num_commercial):
    """
    Récupère les informations d'une famille à partir du numéro commercial d'un article.

    Args:
        connection (pyodbc.Connection): La connexion à la base de données.
        num_commercial (str): Le numéro commercial de l'article.

    Returns:
        tuple: Un tuple contenant le code de la famille et son libellé, ou None si l'article n'existe pas.

    Raises:
        pyodbc.Error: Si une erreur se produit lors de la requête.
    """
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

def get_family_name(connection, family_code):
    """
    Récupère le nom de la famille à partir de son code.
    
    Args:
        connection (pyodbc.Connection): La connexion à la base de données.
        family_code (str): Le code de la famille.
        
    Returns:
        str: Le nom de la famille si trouvé, None sinon.
        
    Raises:
        pyodbc.Error: Si une erreur se produit lors de la requête.
    """
    try:
        cursor = connection.cursor()
        query = "SELECT Libelle FROM FamilleArticle WHERE Code = ?"
        cursor.execute(query, family_code + '.')
        result = cursor.fetchone()
        if result:
            return result[0]
        else:
            return None

    except pyodbc.Error as e:
        write_log(f"[ERREUR] {str(e)}")
        return None

def get_all_articles(connection):
    """
    Récupère tous les articles de stock à partir de la base de données.
    
    Args:
        connection (pyodbc.Connection): La connexion à la base de données.
        
    Returns:
        list: Une liste de tuples contenant les informations des articles, ou None si aucun article n'est trouvé.

    Raises:
        pyodbc.Error: Si une erreur se produit lors de la requête.
    """
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

def get_article_stock(connection, commercial_num):
    """
    Récupère le stock d'un article à partir de son numéro commercial.
    
    Args:
        connection (pyodbc.Connection): La connexion à la base de données.
        commercial_num (str): Le numéro commercial de l'article.
        
    Returns:
        tuple: Un tuple contenant les informations de stock de l'article, ou None si l'article n'existe pas.
        
    Raises:
        pyodbc.Error: Si une erreur se produit lors de la requête.
    """ 
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
    
def get_article_name(connection, commercial_num):
    """
    Récupère le nom d'un article à partir de son numéro commercial.
    
    Args:
        connection (pyodbc.Connection): La connexion à la base de données.
        commercial_num (str): Le numéro commercial de l'article.
        
    Returns:
        str: Le nom de l'article si trouvé, None sinon.
        
    Raises:
        pyodbc.Error: Si une erreur se produit lors de la requête.
    """
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

def create_movement(connection, movement_type, article, quantite):
    """
    Crée un mouvement de stock pour un article donné.
    Afin de garantir la cohérence des données, met à jour la table ElementStock
    pour s'assurer que les quantités restent cohérentes entre elles.

    Args:
        connection (pyodbc.Connection): La connexion à la base de données.
        movement_type (str): Le type de mouvement ('E' pour entrée, 'S' pour sortie).
        article (tuple): Les informations de l'article.
        quantite (int): La quantité à ajouter ou à soustraire.

    Returns:
        bool: True si le mouvement a été créé avec succès, False sinon.

    Raises:
        pyodbc.Error: Si une erreur se produit lors de la requête.
    """
    try:
        cursor = connection.cursor()
        # Insertion du mouvement de stock
        query = "INSERT INTO ElementMvtStock (CodeElem, TypeMvt, Provenance, Date, Quantite, PA, Info) VALUES (?, ?, ?, ?, ?, ?, ?)"
        info = f"Inventaire manuel du {find_closest_date().strftime('%d/%m/%Y')}"
        cursor.execute(query, [article[0], movement_type, 'M', find_closest_date(), quantite, article[5], info])

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