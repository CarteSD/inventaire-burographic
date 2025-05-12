"""
    But : Ce fichier contient l'ensemble des constantes et paramètres nécessaires
    au fonctionnement du module. Il doit être renommé en 'constantes.py' et
    complété avec les valeurs spécifiques à votre environnement.
    Par : Estéban DESESSARD - e.desessard@burographic.fr
    Date : 11/04/2025
"""

# Paramètres de base de données
DB_DRIVER = ""
DB_SERVER = ""
DB_NAME = ""
DB_USER = "" # Facultatif dans le cas de l'authentification Windows
DB_PASSWORD = "" # Facultatif dans le cas de l'authentification Windows

# Fichier de log
LOG_FILE = ""

# Version de l'application
VERSION = "v1.0.0"

# Codes d'erreur
ERROR_CODES = {
    # Erreurs fichiers (F)
    "F001": "Fichier non sélectionné",
    "F002": "Fichier inexistant",
    "F003": "Format de fichier invalide",
    "F004": "Erreur lecture fichier",
    
    # Erreurs articles (A)
    "A001": "Article inexistant",
    "A002": "Article sans famille",
    
    # Erreurs système fichiers (S)
    "S001": "Dossier existant",
    "S002": "Fichiers verrouillés",
    "S003": "Échec suppression",
    "S004": "Tentatives max atteintes",
    
    # Erreurs base de données (D)
    "D001": "Échec mise à jour stock",
    "D002": "Transaction annulée"
}