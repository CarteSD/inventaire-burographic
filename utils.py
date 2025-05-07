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
import pdfkit

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

    # Tableau des familles dont on gère le stock
    stock_families = ["CCOUL.", "CI.", "CINT.", "COCC.", "CONSC.", "COP.", "INFOC.", "LIC.", "LICINT.", "LOG.", "MB.", "MC.", "MO.", "MOB.", "MULT.", "OR.", "ORINT.", "PA.", "PD.", "PDINT.", "PE.", "PEINT."]

    # Génération d'un nom de fichier unique basé sur la date et l'heure
    now = datetime.now()
    inventory_date = find_closest_date()
    inventory_date_str_ymd = inventory_date.strftime("%Y-%m-%d")
    report_id = f"{inventory_date_str_ymd}"
    report_filename = f"rapport_execution_{report_id}.pdf"

    # Dates formatées pour l'affichage
    inventory_date_str = inventory_date.strftime("%d/%m/%Y")
    execution_date_str = now.strftime("%d/%m/%Y")

    # Extraction des erreurs et familles
    errors = report_data.get("errors", {})
    errors_count = len(errors)
    families_values = report_data.get("families_values", {})

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
    details_families_html = ""
    families_values["COCC."] = {
        "libelle": "Copieur d'occasion",
        "value": "",
    }
    families_values["MO."] = {
        "libelle": "Matériel d\'occasion",
        "value": "",
    }
    families_values["INFOC."] = {
        "libelle": "Informatique d\'occasion",
        "value": "",
    }
    families_values["MOB."] = {
        "libelle": "Mobilier",
        "value": "",
    }
    sorted_items = sorted(families_values.items(), key=lambda x: x[0])
    filtered_items = []

    for code, datas in sorted_items:
        if code in stock_families:
            filtered_items.append((code, datas))

    for code, datas in filtered_items:

        # Formatage avec deux décimales fixes
        if not isinstance(datas["value"], str):
            value_fmt = f"{round(datas['value'], 2):.2f}".replace('.', ',')
        else :
            value_fmt = datas["value"]
        
        details_families_html += f"""
        <tr>
            <td>{code}</td>
            <td>{datas["libelle"]}</td>
            <td class="right-align">{value_fmt}</td>
        </tr>
        """
    
    # Ajouter la ligne de total à la fin du tableau
    details_families_html += f"""
    <tr class="total-row">
        <td colspan="2" style="text-align: right;">Valeur totale de l'inventaire :</td>
        <td class="right-align"></td>
    </tr>
    """

    # Charger le template
    template_path = resource_path("report_template.html")
    with open(template_path, "r", encoding="utf-8") as f:
        template = f.read()

    # Remplacer les variables dans le template
    html_content = template.replace("{{inventory_date_str}}", inventory_date_str)
    html_content = html_content.replace("{{execution_date_str}}", execution_date_str)
    html_content = html_content.replace("{{report_id}}", report_id)
    html_content = html_content.replace("{{errors_count}}", str(errors_count))
    html_content = html_content.replace("{{errors_html}}", errors_html)
    html_content = html_content.replace("{{details_families}}", details_families_html)

    # Enregistrer le fichier
    output_dir = f"inventaires/inventaire_{inventory_date_str_ymd}"
    os.makedirs(output_dir, exist_ok=True)
    report_path = os.path.join(output_dir, report_filename)
    
    config = pdfkit.configuration(wkhtmltopdf=r'./wkhtmltopdf.exe')
    options = {
        'margin-bottom': '1.5cm',
        'footer-right': '[page]/[topage]',
        'footer-font-size': '8',
    }
    pdfkit.from_string(html_content, report_path, options=options, configuration=config)
    
    return report_path

# But : Générer un rapport de stock HTML pour une famille d'articles
def generate_family_report(family_code, family_name, articles_data):
    # Préparation des données
    inventory_date = find_closest_date()
    date_str = inventory_date.strftime("%d/%m/%Y")
    date_path = inventory_date.strftime("%Y-%m-%d")
    
    # Génération du contenu HTML pour les détails des articles
    details_html = ""
    total_value = 0
    
    # Formatage des montants avec deux décimales fixes
    for code, article_data in articles_data.items():
        quantite = article_data.get("quantite", 0)
        unit_price = round(article_data.get("prix", 0), 2)
        total_price = round(quantite * unit_price, 2)
        total_value += total_price
        
        # Formatage avec deux décimales fixes
        unit_price_fmt = f"{unit_price:.2f}".replace('.', ',')
        total_price_fmt = f"{total_price:.2f}".replace('.', ',')
        
        details_html += f"""
        <tr>
            <td>{code}</td>
            <td>{article_data.get("nom", "")}</td>
            <td class="right-align">{quantite}</td>
            <td class="right-align">{unit_price_fmt}</td>
            <td class="right-align">{total_price_fmt}</td>
        </tr>
        """

    # Formatage du total avec deux décimales fixes
    total_value_fmt = f"{round(total_value, 2):.2f}".replace('.', ',')

    # Ajout de la ligne de total
    details_html += f"""
    <tr>
        <td colspan="4" class="right-align" style="font-weight: bold; padding-top: 10px; border-top: 1px solid #000;">TOTAL:</td>
        <td class="right-align" style="font-weight: bold; padding-top: 10px; border-top: 1px solid #000;">{total_value_fmt}</td>
    </tr>
    """
    
    # Charger le template
    template_path = resource_path("family_inventory_template.html")
    with open(template_path, "r", encoding="utf-8") as f:
        template = f.read()
    
    # Remplacer les variables dans le template
    html_content = template.replace("{{date_str}}", date_str)
    html_content = html_content.replace("{{famille}}", f"{family_code} - {family_name}")
    html_content = html_content.replace("{{details_html}}", details_html)

    output_dir = f"inventaires/inventaire_{date_path}/familles_rapports"
    os.makedirs(output_dir, exist_ok=True)
    report_path = os.path.join(output_dir, f"{family_code}.pdf")
    
    config = pdfkit.configuration(wkhtmltopdf=r'./wkhtmltopdf.exe')
    options = {
        'margin-left': '1.5cm',
        'footer-right': '[page]/[topage]',
        'footer-font-size': '10',
    }
    pdfkit.from_string(html_content, report_path, options=options, configuration=config)
    
    return report_path

# But : Obtient le chemin absolu du fichier ou dossier spécifié
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# But : Retourne la date précédente la plus proche entre le 30 avril et le 31 octobre
def find_closest_date():
    # Récupérer la date actuelle
    today = datetime.now()
    current_year = today.year
    current_month = today.month
    current_day = today.day

    # Déclaration des dates de référence
    april_30 = datetime(current_year, 4, 30)
    october_31 = datetime(current_year, 10, 31)

    # Vérification que les années des dates de référence sont correctes
    if current_month < 4 or (current_month == 4 and current_day < 30):
        april_30 = datetime(current_year - 1, 4, 30)

    if current_month < 10 or (current_month == 10 and current_day < 31):
        october_31 = datetime(current_year - 1, 10, 31)

    today_date = today.date()

    # Calculer la différence de jours entre aujourd'hui et les dates de référence
    diff_april_30 = abs((april_30.date() - today_date).days)
    diff_october_31 = abs((october_31.date() - today_date).days)

    # Comparer les différences
    if diff_april_30 <= diff_october_31:
        return april_30
    else:
        return october_31