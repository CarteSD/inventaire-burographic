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
        self.root.title(f"BUROGRAPHIC - Inventaire {VERSION}")
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
            text="Ce module vous permet de rééquilibrer les stocks de votre base de données.\n"
            "Sélectionnez un fichier d'inventaire et cliquez sur 'Lancer l'inventaire'.\n\n",
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
        self.report_data = {
            "errors": {},
            "families_values": {},
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
        # Désactiver les boutons
        self.launch_inventory_button.config(state=tk.DISABLED)
        self.browse_button.config(state=tk.DISABLED)
        self.root.update()

        # Vider la zone d'informations
        self.text_box.delete(1.0, tk.END)

        # Réinitialiser le rapport d'exécution
        self.report_data = {
            "errors": {},
            "families_values": {},
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
            inventories_directory = ".\\inventaires"
            if not os.path.exists(inventories_directory):
                log_and_display(f"Création du dossier {inventories_directory}...", self.text_box, self.root, 0.5)
                os.makedirs(inventories_directory)

            # Affichage du message de création du fichier d'inventaire
            log_and_display("Création du dossier d'inventaire à la date correcte", self.text_box, self.root, 0.5)

            current_date = datetime.now().strftime("%Y-%m-%d")
            inventory_date = find_closest_date().strftime("%Y-%m-%d")
            log_and_display(f"Date du jour : {current_date}", self.text_box, self.root, 0.5)
            log_and_display(f"Date d'inventaire : {inventory_date}", self.text_box, self.root)
            this_inventory_directory = os.path.join(inventories_directory, f"inventaire_{inventory_date}")
            temp_inventory_directory = os.path.join(inventories_directory, f"temp_inventaire_{inventory_date}")

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

            # Création d'un dictionnaire pour transformer le fichier en code => quantité
            articles_dictionnary = {}
            undefined_articles = []
            families = []
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
                            
                            self.report_data["errors"][error_name] = error_msg
                            continue
                        else:
                            log_and_display("Annulation de l'opération.", self.text_box, self.root, 0.5)

                            # Nettoyage du dossier temporaire
                            shutil.rmtree(temp_inventory_directory)
                            return
                else:
                    if code not in articles_dictionnary:
                        family = get_family(self.connection, code)
                        if family is None :
                            if code not in undefined_articles:
                                log_and_display(f"L'article {code} n'a pas de famille valide associée", self.text_box, self.root, 0.05)
                                error_name = f"Famille invalide pour l'article {code}"
                                skip = messagebox.askyesno("Famille invalide", f"L'article {code} n'a pas de famille valide associée.\n\n Voulez-vous l'ignorer et continuer ?")
                                if skip:
                                    undefined_articles.append(code)
                                    log_and_display(f"Article {code} ignoré.", self.text_box, self.root, 0.5)
                                    self.report_data["errors"][error_name] = f"L'article {code} n'a pas de famille valide associée. Ignoré, opération reprise."
                                    continue
                                else:
                                    log_and_display("Annulation de l'opération.", self.text_box, self.root, 0.5)

                                    # Nettoyage du dossier temporaire
                                    shutil.rmtree(temp_inventory_directory)
                                    return
                        else:
                            family = family[0].replace(".", "")
                            if family not in families:
                                families.append(family)
                            articles_dictionnary[code] = 1
                    else:
                        articles_dictionnary[code] += 1

            # Copie du fichier brut pour en garder une trace
            raw_file = os.path.join(temp_inventory_directory, f"inventaire_brut_{inventory_date}.txt")
            shutil.copyfile(file_path, raw_file)

            # Création du fichier code;quantite dans le dossier temporaire
            output_file = os.path.join(temp_inventory_directory, f"inventaire_trie_{inventory_date}.csv")
            with open(output_file, 'w', encoding='utf-8') as file:
                # Ajouter un en-tête au CSV
                file.write("Code;Quantité\n")
                for key, value in articles_dictionnary.items():
                    file.write(f"{key};{value}\n")

            # Création du dossier pour les familles dans le dossier temporaire
            families_directory = os.path.join(temp_inventory_directory, "familles")
            if not os.path.exists(families_directory):
                log_and_display(f"Création du dossier {families_directory}...", self.text_box, self.root, 0.5)
                os.makedirs(families_directory)

            # Création de chaque fichier d'inventaire par famille
            for family in families:
                log_and_display(f"Création du fichier d'inventaire pour la famille {family}...", self.text_box, self.root, 1)
                family_file = os.path.join(families_directory, f"{family}.csv")
                with open(family_file, 'w', encoding='utf-8') as file:
                    # Ajouter un en-tête au CSV
                    file.write("Code;Quantité\n")
                    for key in articles_dictionnary.keys():
                        family_code = get_family(self.connection, key)[0]
                        if family_code is not None :
                            family_code = family_code.replace(".", "")
                            if family_code == family:
                                file.write(f"{key};{articles_dictionnary[key]}\n")

            # Exécution de la fonction update_stock
            log_and_display("Lancement de la mise à jour des stocks", self.text_box, self.root, 3)
            self.update_stock(articles_dictionnary)

            # Remplacer l'ancien inventaire si nécessaire
            if inventory_exists:
                log_and_display("Tout s'est bien passé, remplacement de l'ancien inventaire...", self.text_box, self.root, 1)
                
                # Essayer de supprimer l'ancien dossier d'inventaire avec une gestion d'erreur
                files_locked = True
                retry_count = 0
                max_retries = 3
                
                while files_locked and retry_count < max_retries:
                    try:
                        # Essayer de supprimer l'ancien dossier d'inventaire
                        log_and_display(f"Tentative de suppression de l'ancien inventaire {this_inventory_directory}...", self.text_box, self.root, 0.5)
                        shutil.rmtree(this_inventory_directory)
                        files_locked = False  # La suppression a fonctionné
                        log_and_display(f"Ancien inventaire supprimé avec succès.", self.text_box, self.root, 0.5)
                        
                    except PermissionError:
                        retry_count += 1
                        log_and_display(f"ERREUR: Des fichiers sont ouverts dans le dossier d'inventaire.", self.text_box, self.root, 1)
                        
                        # Demander à l'utilisateur de fermer les fichiers
                        retry = messagebox.askretrycancel(
                            "Fichiers ouverts", 
                            f"Certains fichiers du dossier d'inventaire sont actuellement ouverts.\n\n"
                            f"Veuillez fermer tous les fichiers PDF ou HTML qui pourraient être ouverts "
                            f"dans le dossier '{this_inventory_directory}' et cliquer sur 'Recommencer'.\n\n"
                            f"Tentative {retry_count}/{max_retries}"
                        )
                        
                        if not retry:
                            log_and_display("Opération annulée par l'utilisateur.", self.text_box, self.root, 0.5)
                            # Nettoyer le dossier temporaire
                            shutil.rmtree(temp_inventory_directory)
                            return
                    
                    except Exception as e:
                        # En cas d'autre erreur
                        error_msg = str(e)
                        log_and_display(f"ERREUR lors de la suppression de l'ancien inventaire: {error_msg}", self.text_box, self.root, 0.5)
                        
                        # Proposer des alternatives à l'utilisateur
                        response = messagebox.askquestion(
                            "Erreur de suppression",
                            f"Une erreur est survenue lors de la suppression de l'ancien inventaire:\n{error_msg}\n\n"
                            f"Souhaitez-vous tout de même créer un nouveau dossier d'inventaire?"
                        )
                        
                        if response == "yes":
                            # Renommer le dossier temporaire avec un suffixe pour éviter les conflits
                            new_inventory_name = f"{this_inventory_directory}_new"
                            log_and_display(f"Création d'un nouveau dossier d'inventaire: {new_inventory_name}", self.text_box, self.root, 0.5)
                            # Renommer le temporaire en nouveau dossier final
                            shutil.move(temp_inventory_directory, new_inventory_name)
                            log_and_display(f"Nouvel inventaire créé dans {new_inventory_name}", self.text_box, self.root, 0.5)
                            messagebox.showinfo(
                                "Inventaire terminé", 
                                f"L'inventaire a été créé dans un nouveau dossier: {os.path.basename(new_inventory_name)}.\n\n"
                                f"L'ancien inventaire n'a pas été remplacé en raison d'une erreur."
                            )
                            return
                        else:
                            log_and_display("Opération annulée par l'utilisateur.", self.text_box, self.root, 0.5)
                            # Nettoyer le dossier temporaire
                            shutil.rmtree(temp_inventory_directory)
                            return
                
                # Si trop de tentatives ont échoué
                if files_locked:
                    log_and_display(f"Impossible de supprimer l'ancien inventaire après {max_retries} tentatives.", self.text_box, self.root, 0.5)
                    
                    # Demander à l'utilisateur ce qu'il souhaite faire
                    response = messagebox.askquestion(
                        "Maximum de tentatives atteint",
                        f"Après {max_retries} tentatives, impossible de supprimer l'ancien inventaire.\n\n"
                        f"Souhaitez-vous créer un nouveau dossier d'inventaire sans supprimer l'ancien?"
                    )
                    
                    if response == "yes":
                        # Créer un nouveau dossier avec un suffixe
                        new_inventory_name = f"{this_inventory_directory}_new"
                        log_and_display(f"Création d'un nouveau dossier d'inventaire: {new_inventory_name}", self.text_box, self.root, 0.5)
                        shutil.move(temp_inventory_directory, new_inventory_name)
                        log_and_display(f"Nouvel inventaire créé dans {new_inventory_name}", self.text_box, self.root, 0.5)
                        messagebox.showinfo(
                            "Inventaire terminé", 
                            f"L'inventaire a été créé dans un nouveau dossier: {os.path.basename(new_inventory_name)}.\n\n"
                            f"L'ancien inventaire n'a pas été remplacé car des fichiers sont toujours ouverts."
                        )
                        return
                    else:
                        log_and_display("Opération annulée par l'utilisateur.", self.text_box, self.root, 0.5)
                        # Nettoyer le dossier temporaire
                        shutil.rmtree(temp_inventory_directory)
                        return
            
            # Renommer le dossier temporaire en dossier final
            log_and_display(f"Finalisation de l'inventaire...", self.text_box, self.root, 0.5)
            shutil.move(temp_inventory_directory, this_inventory_directory)

            # Après avoir finalisé l'inventaire
            log_and_display("Inventaire terminé.", self.text_box, self.root, 1)

            # Génération des rapports HTML par famille
            log_and_display("Génération des rapports par famille...", self.text_box, self.root, 1)

            # Pour chaque famille, collecter les articles et générer un rapport
            for family in families:
                log_and_display(f"Génération du rapport pour la famille {family}...", self.text_box, self.root, 0.5)
                # Récupérer tous les articles de cette famille
                families_articles = {}
                for key, value in articles_dictionnary.items():
                    try:
                        article_family = get_family(self.connection, key)[0].replace(".", "")
                        if article_family == family:
                            # Récupérer les détails de l'article
                            article_data = get_article_stock(self.connection, key)
                            if article_data:
                                # Créer une entrée dans le dictionnaire
                                families_articles[key] = {
                                    "nom": article_data[7],       # Le nom/libellé de l'article
                                    "quantite": value,            # La quantité scannée
                                    "prix": article_data[5]       # Le prix moyen pondéré (PAMP)
                                }
                    except Exception as e:
                        write_log(f"[ERREUR] Impossible de récupérer les détails de l'article {key}: {str(e)}")
                
                # Générer le rapport pour cette famille si des articles sont présents
                if families_articles:
                    family_name = get_family_name(self.connection, family)
                    
                    # Générer le rapport HTML
                    family_report = generate_family_report(family, family_name, families_articles)
                    log_and_display(f"Rapport généré pour la famille {family}: {os.path.basename(family_report)}", self.text_box, self.root)

            # Génération du rapport d'exécution global
            log_and_display("Génération du rapport d'exécution...", self.text_box, self.root, 1)
            report = generate_report(self.report_data)

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
        
        finally:
            # Réactiver les boutons
            self.launch_inventory_button.config(state=tk.NORMAL)
            self.browse_button.config(state=tk.NORMAL)
            self.root.update()

    # But : permet de mettre à jour le stock en se basant sur un dictionnaire code => quantité
    def update_stock(self, correct_stock):
        # Récupérer tous les articles de la base de données
        all_articles = get_all_articles(self.connection)
        
        if all_articles is not None:
            transaction_success = True
            try:
                log_and_display("Début de la transaction de mise à jour du stock...", self.text_box, self.root, 0.5)
                
                for article in all_articles:
                    commercial_num = article[6]
                    if commercial_num in correct_stock:
                        stock = correct_stock[commercial_num]
                    else:
                        stock = 0
                    # Metrre à jour avec le stock correct
                    success = self.compare_and_update_article_stock(commercial_num, stock)
                    if not success:
                        # Si une mise à jour échoue, marquer la transaction comme échouée
                        transaction_success = False
                        log_and_display(f"Échec de la mise à jour à 0 pour l'article {commercial_num}", self.text_box, self.root, 0.5)
                        break
                
                # Valider ou annuler les modifications selon le résultat
                if transaction_success:
                    log_and_display("Validation de toutes les mises à jour en base de données...", self.text_box, self.root, 0.5)
                    self.connection.commit()
                    log_and_display("Mises à jour validées avec succès!", self.text_box, self.root, 0.5)
                else:
                    log_and_display("Annulation de toutes les mises à jour en raison d'erreurs...", self.text_box, self.root, 0.5)
                    self.connection.rollback()
                    log_and_display("Modifications annulées avec succès.", self.text_box, self.root, 0.5)
                    # Lever une exception pour informer l'appelant de l'échec
                    raise Exception("La mise à jour du stock a échoué pour certains articles.")
                    
            except Exception as e:
                # En cas d'erreur inattendue, annuler toutes les modifications
                log_and_display(f"ERREUR lors de la mise à jour du stock: {str(e)}", self.text_box, self.root, 0.5)
                log_and_display("Annulation de toutes les modifications...", self.text_box, self.root, 0.5)
                self.connection.rollback()
                log_and_display("Modifications annulées avec succès.", self.text_box, self.root, 0.5)
                # Relever l'exception pour qu'elle soit gérée par la méthode appelante
                raise

    # But : Comparer la quantité théorique et la réelle afin de réaliser un mouvement de stock
    def compare_and_update_article_stock(self, commercial_num, real_quantity):
        try:
            # Récupérer la quantité en stock théorique
            bd_article = get_article_stock(self.connection, commercial_num)
            if bd_article is None:
                log_and_display(f"L'article {commercial_num} n'existe pas dans la base de données", self.text_box, self.root, 0.05)
                return False

            code = bd_article[0].replace(".", "")
            pamp = bd_article[5]

            supply_qty = bd_article[2]
            consumption_qty = bd_article[3]
            stock_qty = supply_qty - consumption_qty

            # Créer un mouvement de stock pour corriger la différence
            diff = abs(stock_qty - real_quantity)
            if stock_qty > real_quantity:
                movement_type = 'S'
            elif stock_qty < real_quantity:
                movement_type = 'E'
            else:
                movement_type = None

            # Mettre à jour la somme de l'inventaire de la famille
            family_article = get_family(self.connection, commercial_num)

            # Vérifier si la famille existe
            if family_article is None:
                # Gérer le cas d'une famille inexistante
                log_and_display(f"L'article {code} n'a pas de famille valide associée. Il a été mis à jour mais ne figurera dans aucun inventaire.", self.text_box, self.root, 0.05)
                self.report_data["errors"][f"Famille inexistante pour l'article {code}"] = f"L'article {code} n'a pas de famille valide associée. Mis à jour, mais ne figurera dans aucun inventaire par famille, opération reprise."
            else:
                family_code = family_article[0]
                family_libelle = family_article[1]

                # Mettre à jour les valeurs dans le rapport
                if self.report_data["families_values"].get(family_code, None) is None:
                    self.report_data["families_values"][family_code] = {
                        "libelle": family_libelle,
                        "value": 0
                    }
                self.report_data["families_values"][family_code]["value"] += pamp * real_quantity

            # Gestion d'une potentielle erreur lors de la mise à jour
            if movement_type is not None:
                log_and_display(f"Mise à jour de l'article {code} à sa nouvelle quantité : {real_quantity}", self.text_box, self.root, 0.02)
                if not create_movement(self.connection, movement_type, bd_article, diff):
                    log_and_display(f"La mise à jour de l'article {code} a échoué", self.text_box, self.root, 0.02)
                    return False
                else:
                    write_log(f"Mise à jour de l'article {code} réussie")
                    return True
            else:
                log_and_display(f"Aucune mise à jour nécessaire pour l'article {code}", self.text_box, self.root, 0.02)
                return True
        
        except Exception as e:
            log_and_display(f"Erreur lors de la mise à jour de l'article {commercial_num}: {str(e)}", self.text_box, self.root, 0.05)
            return False

    # But : Formater le message d'erreur pour l'affichage
    def format_article_error_message(self, code, position, prev_code=None, next_code=None):
        if position == "first":
            article_suivant = get_article_name(self.connection, next_code)
            return f"Article {code} absent en base de données. Situé en première position, avant {next_code} ({article_suivant}). Ignoré, opération reprise"
        elif position == "last":
            article_precedent = get_article_name(self.connection, prev_code)
            return f"Article {code} absent en base de données. Situé en dernière position, après {prev_code} ({article_precedent}). Ignoré, opération reprise"
        else:
            article_precedent = get_article_name(self.connection, prev_code)
            article_suivant = get_article_name(self.connection, next_code)
            line_number = position + 1
            return f"Article {code} absent en base de données. Situé entre {prev_code} ({article_precedent}) et {next_code} ({article_suivant}) à la ligne {line_number}. Ignoré, opération reprise"