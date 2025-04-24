# # # # # # # # # # # #
# But : Ce fichier contient l'ensemble des méthode relatives à l'interface utilisateur
# Par : Estéban DESESSARD - e.desessard@burographic.fr
# Date : 11/04/2025
# # # # # # # # # # # #

import tkinter as tk
import os
import shutil
import time
from datetime import datetime
from tkinter import filedialog, messagebox
from constantes import *
from db import *
from utils import *
import webbrowser

# But : Classe de l'application métier, elle permet de faire
#       le pont entre l'utilisateur et l'interface
class Interface:

    # CONSTRUCTEUR
    def __init__(self, root):
        # Création de la fenêtre utilisateur
        self.root = root
        self.root.title("BUROGRAPHIC - Inventaire")
        self.root.geometry("800x600")
        self.root.iconbitmap(os.path.join(os.path.dirname(__file__), 'icone.ico'))
        self.connection = database_connection()

        # Variables pour le chemin du fichier d'inventaire
        self.inventory_file_path = tk.StringVar()

        # Création de la fenêtre principale
        self.main_frame = tk.Frame(self.root, padx=10, pady=10)
        self.main_frame.pack(fill=tk.BOTH, expand=True, anchor="w")

        # Création du message d'arrivée sur l'application
        self.welcome_message = tk.Label(
            self.main_frame, text="Bienvenue dans le module d'inventaire de BUROGRAPHIC",
           font=("Arial", 16, "bold")
        )
        self.description = tk.Label(
            self.main_frame,
            text="Ce module vous permet de créer un inventaire pour votre entreprise",
            font=("Arial", 12)
        )
        self.welcome_message.pack(pady=20)
        self.description.pack()

        # Cadre pour la sélection de fichier
        self.file_selection_frame = tk.Frame(self.main_frame)
        self.file_selection_frame.pack(fill=tk.X, pady=15)

        self.file_label = tk.Label(self.file_selection_frame, text="Fichier sélectionné:", font=("Arial", 10))
        self.file_label.pack(anchor=tk.W)

        # Création du bouton pour choisir le fichier d'inventaire
        self.file_path_entry = tk.Entry(self.file_selection_frame, textvariable=self.inventory_file_path, width=50)
        self.file_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, pady=5)

        self.browse_button = tk.Button(self.file_selection_frame, text="Parcourir...", command=self.select_file)
        self.browse_button.pack(side=tk.RIGHT, padx=5, pady=5)

        # Bouton pour lancer l'inventaire
        self.launch_inventory_button = tk.Button(
            self.main_frame,
            text="Lancer l'inventaire",
            command=self.launch_inventory,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 12, "bold"),
            padx=10,
            pady=5
        )
        self.launch_inventory_button.pack(pady=20)

        # Cadre pour les résultats
        self.results_frame = tk.Frame(self.main_frame)
        self.results_frame.pack(fill=tk.BOTH, expand=True)

        self.text_scrollbar = tk.Scrollbar(self.results_frame)
        self.text_box = tk.Text(self.results_frame, height=15, width=80, yscrollcommand=self.text_scrollbar.set,
                                state="normal")
        self.text_scrollbar.config(command=self.text_box.yview)

        # Placement de la boîte de texte et de la barre de défilement
        self.text_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.text_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Initialisation du tableau de données pour le rapport d'exécution
        self.report_datas = {
            "errors": {},
            "families_values": {},
            "stats": {
                "total_articles": 0,
                "different_articles": 0,
                "familles_count": 0,
                "errors_count": 0
            }
        }

    # But : Méthode permettant de sélectionner un fichier dans
    #       l'explorateur de fichiers
    def select_file(self):
        filename = filedialog.askopenfilename(
            title="Sélectionnez un fichier texte",
            filetypes=(("Fichiers texte", "*.txt"), ("Tous les fichiers", "*.*"))
        )
        if filename:
            self.inventory_file_path.set(filename)

    # But : Lancer la création de l'inventaire
    def launch_inventory(self):
        # Vider la zone d'informations
        self.text_box.delete(1.0, tk.END)

        # Réinitialiser le rapport d'exécution
        self.report_datas = {
            "errors": {},
            "families_values": {},
            "stats": {
                "total_articles": 0,
                "different_articles": 0,
                "familles_count": 0,
                "errors_count": 0
            }
        }

        # Affichage du message de récupération du fichier d'inventaire
        log_and_display("Récupération du fichier d'inventaire...", self.text_box, self.root)

        # Récupération du fichier sélectionné
        file_path = self.inventory_file_path.get()
        if not file_path:
            messagebox.showerror("Erreur", "Veuillez sélectionner un fichier d'inventaire.")
            return
        if not os.path.exists(file_path):
            messagebox.showerror("Erreur", "Le fichier sélectionné n'existe pas.")
            return
        if not file_path.endswith(".txt"):
            messagebox.showerror("Erreur", "Le fichier sélectionné n'est pas un fichier texte.")
            return

        # Affichage du message de lecture du fichier d'inventaire
        log_and_display("Lecture du fichier d'inventaire...", self.text_box, self.root, 0.5)

        # Lecture du fichier d'inventaire
        try:
            inventories_dicrectory = ".\\inventaires"
            if not os.path.exists(inventories_dicrectory):
                log_and_display(f"Création du dossier {inventories_dicrectory}...", self.text_box, self.root, 0.5)
                os.makedirs(inventories_dicrectory)

            # Affichage du message de création du fichier d'inventaire
            log_and_display("Création du dossier d'inventaire à la date du jour", self.text_box, self.root, 0.5)

            current_date = datetime.now().strftime("%Y-%m-%d")
            log_and_display(f"Date d'inventaire : {current_date}", self.text_box, self.root)
            this_inventory_directory = os.path.join(inventories_dicrectory, f"inventaire_{current_date}")
            temp_inventory_directory = os.path.join(inventories_dicrectory, f"temp_inventaire_{current_date}")

            inventory_exists = os.path.exists(this_inventory_directory)
            
            # Création du dossier temporaire pour préparer le nouvel inventaire
            log_and_display(f"Création du dossier temporaire {temp_inventory_directory}...", self.text_box, self.root, 0.5)
            os.makedirs(temp_inventory_directory)

            if inventory_exists:
                log_and_display(f"Le dossier {this_inventory_directory} existe déjà", self.text_box, self.root, 0.5)
                time.sleep(2)
                overwrite = messagebox.askyesno("Dossier déjà existant", f"Le dossier {this_inventory_directory} existe déjà. Voulez-vous l'écraser ?")
                if not overwrite:
                    log_and_display("Annulation de l'opération.", self.text_box, self.root, 0.5)

                    # Nettoyage du dossier temporaire
                    shutil.rmtree(temp_inventory_directory)
                    return
    
            with open(file_path, 'r') as file:
                raw_datas = file.readlines()

            # Affichage du message de récupération des articles
            log_and_display("Récupération des articles...", self.text_box, self.root, 1)

            # Ajout du nombre d'article total au rapport d'exécution
            self.report_datas["stats"]["total_articles"] = len(raw_datas)

            # Création d'un dictionnaire pour transformer le fichier en code => quantité
            articles_dictionnary = {}
            undefined_articles = []
            for code in raw_datas:
                code = code.replace("\n", "").strip()
                if code == "":
                    continue
                # Vérification de l'existence de l'article dans la base de données
                if not article_exists(self.connection, code):
                    if code not in undefined_articles:
                        log_and_display(f"L'article {code} n'existe pas dans la base de données", self.text_box, self.root)
                        error_name = f"Article {code} inexistant"
                        skip = messagebox.askyesno("Article inexistant", f"L'article {code} n'existe pas dans la base de données.\n\n Voulez-vous l'ignorer et continuer ?")
                        if skip:
                            undefined_articles.append(code)
                            log_and_display(f"Article {code} ignoré.", self.text_box, self.root, 0.5)
                            index_actuel = raw_datas.index(code + '\n')
                            if index_actuel == 0:
                                next_code = raw_datas[index_actuel + 1].replace("\n", "").strip()
                                error_msg = self.format_article_error_message(code, "first", next_code=next_code)
                            elif index_actuel == len(raw_datas) - 1:
                                prev_code = raw_datas[index_actuel - 1].replace("\n", "").strip()
                                error_msg = self.format_article_error_message(code, "last", prev_code=prev_code)
                            else:
                                prev_code = raw_datas[index_actuel - 1].replace("\n", "").strip()
                                next_code = raw_datas[index_actuel + 1].replace("\n", "").strip()
                                error_msg = self.format_article_error_message(code, index_actuel, prev_code, next_code)
                            
                            self.report_datas["errors"][error_name] = error_msg
                            continue
                        else:
                            log_and_display("Annulation de l'opération.", self.text_box, self.root, 0.5)

                            # Nettoyage du dossier temporaire
                            shutil.rmtree(temp_inventory_directory)
                            return
                else:
                    if code in articles_dictionnary:
                        articles_dictionnary[code] += 1
                    else:
                        articles_dictionnary[code] = 1

            # Copie du fichier brut pour en garder une trace
            raw_file = os.path.join(temp_inventory_directory, f"inventaire_brut_{current_date}.txt")
            shutil.copyfile(file_path, raw_file)

            # Création du fichier code;quantite dans le dossier temporaire
            output_file = os.path.join(temp_inventory_directory, f"inventaire_trie_{current_date}.txt")
            with open(output_file, 'w', encoding='utf-8') as file:
                for key, value in articles_dictionnary.items():
                    file.write(f"{key};{value}\n")
                    self.report_datas["stats"]["different_articles"] += 1

            # Création du tableau des familles scannées
            familles = []
            for key in articles_dictionnary.keys():
                # Récupération de la famille de l'article
                famille = get_famille(self.connection, key)[0].replace(".", "")
                if famille is None and key not in undefined_articles:
                    log_and_display(f"La récupération de la famille pour l'article {key} a échoué", self.text_box, self.root)
                    error_name = f"Erreur de récupération de famille pour l'article {key}"
                    skip = messagebox.askyesno("Échec de récupération",
                                            f"La récupération de la famille pour l'article {key} a échoué.\n\n Voulez-vous l'ignorer et continuer ?")
                    if skip:
                        log_and_display(f"Article {key} ignoré.", self.text_box, self.root, 0.5)
                        self.report_datas["errors"][error_name] = f"Article {key} ignoré, dû a une erreur de récupération de famille en base de données. Ignoré, opération reprise."
                        continue
                    else:
                        log_and_display("Annulation de l'opération.", self.text_box, self.root, 0.5)

                        # Nettoyage du dossier temporaire
                        shutil.rmtree(temp_inventory_directory)
                        return
                if famille not in familles:
                    if not famille_exists(self.connection, famille):
                        log_and_display(f"La famille {famille} n'existe pas dans la base de données", self.text_box, self.root)
                        error_name = f"Famille {famille} inexistante"
                        skip = messagebox.askyesno("Famille inexistante",
                                                f"La famille {famille} n'existe pas dans la base de données.\n\n Voulez-vous l'ignorer et continuer ?")
                        if skip:
                            log_and_display(f"Famille {famille} ignorée.", self.text_box, self.root, 0.5)
                            self.report_datas["errors"][error_name] = f"Famille {famille} absente en base de données. Ignoré, opération reprise"
                            continue
                        else:
                            log_and_display("Annulation de l'opération.", self.text_box, self.root, 0.5)

                            # Nettoyage du dossier temporaire
                            shutil.rmtree(temp_inventory_directory)
                            return
                    familles.append(famille)

            # Ajout du nombre de familles au rapport d'exécution
            self.report_datas["stats"]["familles_count"] = len(familles)

            # Création du dossier pour les familles dans le dossier temporaire
            families_directory = os.path.join(temp_inventory_directory, "familles")
            if not os.path.exists(families_directory):
                log_and_display(f"Création du dossier {families_directory}...", self.text_box, self.root, 0.5)
                os.makedirs(families_directory)

            # Création de chaque fichier d'inventaire par famille
            for famille in familles:
                log_and_display(f"Création du fichier d'inventaire pour la famille {famille}...", self.text_box, self.root, 1)
                family_file = os.path.join(families_directory, f"{famille}.txt")
                with open(family_file, 'w', encoding='utf-8') as file:
                    for key in articles_dictionnary.keys():
                        if get_famille(self.connection, key)[0].replace(".", "") == famille:
                            file.write(f"{key};{articles_dictionnary[key]}\n")

            # Exécution de la fonction update_stock
            log_and_display("Lancement de la mise à jour des stocks", self.text_box, self.root, 3)
            self.update_stock(articles_dictionnary)

            # Remplacer l'ancien inventaire si nécessaire
            if inventory_exists:
                log_and_display("Tout s'est bien passé, remplacement de l'ancien inventaire...", self.text_box, self.root, 1)
                
                # Supprimer l'ancien dossier d'inventaire
                log_and_display(f"Suppression de l'ancien inventaire {this_inventory_directory}...", self.text_box, self.root, 0.5)
                shutil.rmtree(this_inventory_directory)
            
            # Renommer le dossier temporaire en dossier final
            log_and_display(f"Finalisation de l'inventaire...", self.text_box, self.root, 0.5)
            shutil.move(temp_inventory_directory, this_inventory_directory)

            # Après avoir finalisé l'inventaire
            log_and_display("Inventaire terminé.", self.text_box, self.root, 1)

            # Génération des rapports HTML par famille
            log_and_display("Génération des rapports par famille...", self.text_box, self.root, 1)
            current_date_ymd = datetime.now().strftime("%Y-%m-%d")  # Format pour le chemin du dossier

            # Pour chaque famille, collecter les articles et générer un rapport
            for famille in familles:
                log_and_display(f"Génération du rapport pour la famille {famille}...", self.text_box, self.root, 0.5)
                # Récupérer tous les articles de cette famille
                famille_articles = {}
                for key, value in articles_dictionnary.items():
                    try:
                        article_famille = get_famille(self.connection, key)[0].replace(".", "")
                        if article_famille == famille:
                            # Récupérer les détails de l'article
                            article_data = get_article_stock(self.connection, key)
                            if article_data:
                                # Créer une entrée dans le dictionnaire
                                famille_articles[key] = {
                                    "nom": article_data[7],       # Le nom/libellé de l'article
                                    "quantite": value,            # La quantité scannée
                                    "prix": article_data[5]       # Le prix moyen pondéré (PAMP)
                                }
                    except Exception as e:
                        write_log(f"[ERREUR] Impossible de récupérer les détails de l'article {key}: {str(e)}")
                
                # Générer le rapport pour cette famille si des articles sont présents
                if famille_articles:
                    famille_name = ""
                    try:
                        # Récupérer le nom complet de la famille
                        cursor = self.connection.cursor()
                        query = "SELECT Libelle FROM FamilleArticle WHERE Code = ?"
                        cursor.execute(query, [famille + '.'])
                        result = cursor.fetchone()
                        if result:
                            famille_name = result[0]
                    except Exception as e:
                        write_log(f"[ERREUR] Impossible de récupérer le nom de la famille {famille}: {str(e)}")
                        famille_name = famille
                    
                    # Générer le rapport HTML
                    family_report = generate_family_report(famille, famille_name, famille_articles)
                    log_and_display(f"Rapport généré pour la famille {famille}: {os.path.basename(family_report)}", self.text_box, self.root)

            # Génération du rapport d'exécution global
            log_and_display("Génération du rapport d'exécution...", self.text_box, self.root, 1)
            self.report_datas["stats"]["errors_count"] = len(self.report_datas["errors"])
            report = generate_report(self.report_datas)

            log_and_display(f"Rapport d'exécution généré : {report}", self.text_box, self.root, 1)
            user_wants_open = messagebox.askyesno("Rapport généré", f"Le rapport d'exécution d'inventaire a été généré à l'emplacement {report}.\n\n Souhaitez-vous l'ouvrir ?")
            if user_wants_open:
                log_and_display(f"Ouverture du rapport d'exécution...", self.text_box, self.root)
                webbrowser.open(f"file:///{os.path.abspath(report)}")

        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors du traitement du fichier: {str(e)}")
            write_log(f"[ERREUR] {str(e)}")
            # En cas d'erreur, nettoyer le dossier temporaire s'il existe
            if os.path.exists(temp_inventory_directory):
                try:
                    shutil.rmtree(temp_inventory_directory)
                    log_and_display(f"Nettoyage du dossier temporaire suite à une erreur", self.text_box, self.root)
                except:
                    pass
            return

    # But : permet de mettre à jour le stock en se basant sur un dictionnaire code => quantité
    def update_stock(self, correct_stock):
        # Récupérer tous les articles de la base de données
        all_articles = get_all_articles(self.connection)

        if all_articles is not None:
            for article in all_articles:
                num_commercial = article[6]
                if num_commercial in correct_stock:
                    # Mettre à jour l'article dans la base de données
                    self.compare_and_update_article_stock(num_commercial, correct_stock[num_commercial])
                else:
                    # Mettre la valeur du stock à 0
                    self.compare_and_update_article_stock(num_commercial, 0)

    # But : Comparer la quantité théorique et la réelle afin de réaliser un mouvement de stock
    def compare_and_update_article_stock(self, num_commercial, real_quantity):
        # Récupérer la quantité en stock théorique
        bd_article = get_article_stock(self.connection, num_commercial)

        code = bd_article[0].replace(".", "")
        pamp = bd_article[5]

        qte_appro = bd_article[2]
        qte_conso = bd_article[3]
        qte_stock = qte_appro - qte_conso

        # Créer un mouvement de stock pour corriger la différence
        diff = abs(qte_stock - real_quantity)
        if qte_stock > real_quantity:
            type_mvt = 'S'
        elif qte_stock < real_quantity:
            type_mvt = 'E'
        else:
            type_mvt = None

        # Mettre à jour la somme de l'inventaire de la famille
        famille_article = get_famille(self.connection, num_commercial)

        # Vérifier si la famille existe
        if famille_article is None:
            # Gérer le cas d'une famille inexistante
            log_and_display(f"L'article {code} n'a pas de famille valide associée", self.text_box, self.root, 0.05)
            self.report_datas["errors"][f"Famille inexistante pour l'article {code}"] = f"L'article {code} n'a pas de famille valide associée. Mis à jour, mais ne figurera dans aucun inventaire par famille, opération reprise."
            return
        else:
            code_famille = famille_article[0]
            famille_libelle = famille_article[1]

        # Mettre à jour les valeurs dans le rapport
        if self.report_datas["families_values"].get(code_famille, None) is None:
            self.report_datas["families_values"][code_famille] = {}
            self.report_datas["families_values"][code_famille]["libelle"] = famille_libelle
            self.report_datas["families_values"][code_famille]["value"] = 0
        self.report_datas["families_values"][code_famille]["value"] += pamp * real_quantity

        # Gestion d'une potentielle erreur lors de la mise à jour
        if type_mvt is not None:
            log_and_display(f"Mise à jour de l'article {code} à sa nouvelle quantité : {real_quantity}", self.text_box, self.root, 0.02)
            if not create_mvt(self.connection, type_mvt, bd_article, diff) :
                log_and_display(f"La mise à jour de l'article {code} a échoué", self.text_box, self.root, 0.02)
            else:
                write_log(f"Mise à jour de l'article {code} réussie")
        else:
            log_and_display(f"Aucune mise à jour nécessaire pour l'article {code}", self.text_box, self.root, 0.02)

    # But : Récupérer le nom / libellé d'un article
    def get_article_name(self, code):
        article = get_article_def(self.connection, code)
        if article is None:
            return "inconnu"
        return article[1]

    # But : Formater le message d'erreur pour l'affichage
    def format_article_error_message(self, code, position, prev_code=None, next_code=None):
        if position == "first":
            article_suivant = self.get_article_name(next_code)
            return f"Article {code} absent en base de données. Situé en première position, avant {next_code} ({article_suivant}). Ignoré, opération reprise"
        elif position == "last":
            article_precedent = self.get_article_name(prev_code)
            return f"Article {code} absent en base de données. Situé en dernière position, après {prev_code} ({article_precedent}). Ignoré, opération reprise"
        else:
            article_precedent = self.get_article_name(prev_code)
            article_suivant = self.get_article_name(next_code)
            line_number = position + 1
            return f"Article {code} absent en base de données. Situé entre {prev_code} ({article_precedent}) et {next_code} ({article_suivant}) à la ligne {line_number}. Ignoré, opération reprise"