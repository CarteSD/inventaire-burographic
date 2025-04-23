# # # # # # # # # # # #
# But : Ce fichier contient l'ensemble des méthodes utilitaires autres
# Par : Estéban DESESSARD - e.desessard@burographic.fr
# Date : 11/04/2025
# # # # # # # # # # # #

import tkinter as tk
import time
from datetime import datetime
from constantes import *
import os
import sys

# But : Écrire un message dans le fichier de log
def write_log(message):
    log_file = LOG_FILE
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f"{datetime.now()} - {message}\n")

# But : Écrire un message sur l'interface utilisateur et dans le fichier de log
def log_and_display(message, text_box, root, delay=0):
    if delay:
        time.sleep(delay)
    text_box.insert(tk.END, message + "\n")
    text_box.see(tk.END)
    root.update()
    write_log(message)

# But : Générer un rapport d'exécution au format HTML
def generate_report(report_data):
    # Vérification que report_data est bien un dictionnaire
    if not isinstance(report_data, dict):
        raise TypeError("Les données du rapport doivent être un dictionnaire")

    # Génération d'un nom de fichier unique basé sur la date et l'heure
    now = datetime.now()
    timestamp = now.strftime("%Y-%m-%d")
    report_id = f"{timestamp}"
    report_filename = f"rapport_execution_{report_id}.html"

    # Date et heure formatées pour l'affichage
    date_str = now.strftime("%d/%m/%Y")
    time_str = now.strftime("%H:%M")

    # Extraction des statistiques
    stats = report_data.get("stats", {})
    total_articles = stats.get("total_articles", 0)
    different_articles = stats.get("different_articles", 0)
    familles_count = stats.get("familles_count", 0)
    errors_count = stats.get("errors_count", 0)

    # Extraction des erreurs et familles avec vérification du type
    errors = report_data.get("errors", {})
    details = report_data.get("details", {})

    # Génération du contenu HTML pour les erreurs
    errors_html = ""
    if errors:
        for error_code, error_details in errors.items():
            if isinstance(error_details, dict):
                message = error_details.get('message', 'N/A')
            else:
                message = str(error_details)

            errors_html += f"""
            <div class="error-card">
                <h3>Erreur {error_code}</h3>
                <p><strong>Message:</strong> {message}</p>
            </div>
            """
    else:
        errors_html = '<p>Aucune erreur n\'a été enregistrée durant l\'exécution.</p>'
    
    # Génération du contenu HTML pour les détails des articles
    details_html = ""
    valeur_totale = 0
    if details:
        for num_commercial, article_detail in details.items():
            code = article_detail.get("code")
            nom = article_detail.get("nom")
            ancien_stock = article_detail.get("ancien_stock")
            nouveau_stock = article_detail.get("nouveau_stock")
            difference = article_detail.get("difference")
            type_mvt = article_detail.get("type_mvt")
            pamp = round(article_detail.get("pamp"), 2)
            valeur = round(nouveau_stock * pamp, 2)

            # Ajouter la valeur à la valeur totale de l'inventaire
            valeur_totale += valeur
            
            # Déterminer l'action et la classe CSS pour le style
            if type_mvt == 'E':
                action = "Ajout"
                row_class = "added"
            elif type_mvt == 'S':
                action = "Retrait"
                row_class = "removed"
            else:
                action = "Inchangé"
                row_class = "unchanged"
            
            details_html += f"""
            <tr class="{row_class}">
                <td>{code}</td>
                <td>{num_commercial}</td>
                <td>{nom}</td>
                <td>{ancien_stock}</td>
                <td>{nouveau_stock}</td>
                <td>{difference}</td>
                <td>{action}</td>
                <td>{pamp}</td>
                <td>{valeur}</td>
            </tr>
            """

        # Ajouter la ligne de total à la fin du tableau
        details_html += f"""
        <tr class="total-row">
            <td colspan="8" style="text-align: right;">Valeur totale de l'inventaire :</td>
            <td>{round(valeur_totale, 2)}</td>
        </tr>
        """

    else:
        details_html = '<tr><td colspan="6">Aucun détail d\'article disponible.</td></tr>'

    # Charger le template
    template_path = resource_path("report_template.html")
    with open(template_path, "r", encoding="utf-8") as f:
        template = f.read()

    # Remplacer les variables dans le template
    html_content = template.replace("{{date_str}}", date_str)
    html_content = html_content.replace("{{time_str}}", time_str)
    html_content = html_content.replace("{{report_id}}", report_id)
    html_content = html_content.replace("{{total_articles}}", str(total_articles))
    html_content = html_content.replace("{{different_articles}}", str(different_articles))
    html_content = html_content.replace("{{familles_count}}", str(familles_count))
    html_content = html_content.replace("{{errors_count}}", str(errors_count))
    html_content = html_content.replace("{{errors_html}}", errors_html)
    html_content = html_content.replace("{{details_html}}", details_html)

    # Enregistrer le fichier
    output_dir = f"inventaires/inventaire_{report_id}"
    os.makedirs(output_dir, exist_ok=True)
    report_path = os.path.join(output_dir, report_filename)

    with open(report_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    return report_path

# But : Obtient le chemin absolu du fichier ou dossier spécifié
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)