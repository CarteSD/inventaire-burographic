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
        self.InventoryFilePath = tk.StringVar()

        # Création de la fenêtre principale
        self.mainFrame = tk.Frame(self.root, padx=10, pady=10)
        self.mainFrame.pack(fill=tk.BOTH, expand=True, anchor="w")

        # Création du message d'arrivée sur l'application
        self.welcomeMessage = tk.Label(
            self.mainFrame, text="Bienvenue dans le module d'inventaire de BUROGRAPHIC",
           font=("Arial", 16, "bold")
        )
        self.description = tk.Label(
            self.mainFrame,
            text="Ce module vous permet de créer un inventaire pour votre entreprise",
            font=("Arial", 12)
        )
        self.welcomeMessage.pack(pady=20)
        self.description.pack()

        # Cadre pour la sélection de fichier
        self.fileSelectionFrame = tk.Frame(self.mainFrame)
        self.fileSelectionFrame.pack(fill=tk.X, pady=15)

        self.fileLabel = tk.Label(self.fileSelectionFrame, text="Fichier sélectionné:", font=("Arial", 10))
        self.fileLabel.pack(anchor=tk.W)

        # Création du bouton pour choisir le fichier d'inventaire
        self.filePathEntry = tk.Entry(self.fileSelectionFrame, textvariable=self.InventoryFilePath, width=50)
        self.filePathEntry.pack(side=tk.LEFT, fill=tk.X, expand=True, pady=5)

        self.browseButton = tk.Button(self.fileSelectionFrame, text="Parcourir...", command=self.select_file)
        self.browseButton.pack(side=tk.RIGHT, padx=5, pady=5)

        # Bouton pour lancer l'inventaire
        self.launchInventoryButton = tk.Button(
            self.mainFrame,
            text="Lancer l'inventaire",
            command=self.launch_inventory,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 12, "bold"),
            padx=10,
            pady=5
        )
        self.launchInventoryButton.pack(pady=20)

        # Cadre pour les résultats
        self.results_frame = tk.Frame(self.mainFrame)
        self.results_frame.pack(fill=tk.BOTH, expand=True)

        self.text_scrollbar = tk.Scrollbar(self.results_frame)
        self.text_box = tk.Text(self.results_frame, height=15, width=80, yscrollcommand=self.text_scrollbar.set,
                                state="normal")
        self.text_scrollbar.config(command=self.text_box.yview)

        # Placement de la boîte de texte et de la barre de défilement
        self.text_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.text_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Initialisation du tableau de données pour le rapport d'exécution
        self.reportDatas = {
            "errors": {},
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
            self.InventoryFilePath.set(filename)

    # But : Lancer la création de l'inventaire
    def launch_inventory(self):
        # Vider la zone d'informations
        self.text_box.delete(1.0, tk.END)

        # Affichage du message de récupération du fichier d'inventaire
        log_and_display("Récupération du fichier d'inventaire...", self.text_box, self.root)

        # Récupération du fichier sélectionné
        filePath = self.InventoryFilePath.get()
        if not filePath:
            messagebox.showerror("Erreur", "Veuillez sélectionner un fichier d'inventaire.")
            return
        if not os.path.exists(filePath):
            messagebox.showerror("Erreur", "Le fichier sélectionné n'existe pas.")
            return
        if not filePath.endswith(".txt"):
            messagebox.showerror("Erreur", "Le fichier sélectionné n'est pas un fichier texte.")
            return

        # Affichage du message de lecture du fichier d'inventaire
        log_and_display("Lecture du fichier d'inventaire...", self.text_box, self.root, 0.5)

        # Lecture du fichier d'inventaire
        try:
            inventairesDirectory = ".\\inventaires"
            if not os.path.exists(inventairesDirectory):
                log_and_display(f"Création du dossier {inventairesDirectory}...", self.text_box, self.root, 0.5)
                os.makedirs(inventairesDirectory)

            # Affichage du message de création du fichier d'inventaire
            log_and_display("Création du dossier d'inventaire à la date du jour", self.text_box, self.root, 0.5)

            currentDate = datetime.now().strftime("%Y-%m-%d")
            log_and_display(f"Date d'inventaire : {currentDate}", self.text_box, self.root)
            thisInventoryDirectory = os.path.join(inventairesDirectory, f"inventaire_{currentDate}")
            tempInventoryDirectory = os.path.join(inventairesDirectory, f"temp_inventaire_{currentDate}_{int(time.time())}")

            inventoryExists = os.path.exists(thisInventoryDirectory)
            
            # Création du dossier temporaire pour préparer le nouvel inventaire
            log_and_display(f"Création du dossier temporaire {tempInventoryDirectory}...", self.text_box, self.root, 0.5)
            os.makedirs(tempInventoryDirectory)

            if inventoryExists:
                log_and_display(f"Le dossier {thisInventoryDirectory} existe déjà", self.text_box, self.root, 0.5)
                time.sleep(2)
                overwrite = messagebox.askyesno("Dossier déjà existant", f"Le dossier {thisInventoryDirectory} existe déjà. Voulez-vous l'écraser ?")
                if not overwrite:
                    log_and_display("Annulation de l'opération.", self.text_box, self.root, 0.5)

                    # Nettoyage du dossier temporaire
                    shutil.rmtree(tempInventoryDirectory)
                    return
    
            with open(filePath, 'r') as file:
                rawDatas = file.readlines()

            # Affichage du message de récupération des articles
            log_and_display("Récupération des articles...", self.text_box, self.root, 1)

            # Ajout du nombre d'article total au rapport d'exécution
            self.reportDatas["stats"]["total_articles"] = len(rawDatas)

            # Création d'un dictionnaire pour transformer le fichier en code => quantité
            articlesDictionnary = {}
            undefinedArticles = []
            for code in rawDatas:
                code = code.replace("\n", "")
                # Vérification de l'existence de l'article dans la base de données
                if not article_exists(self.connection, code):
                    if code not in undefinedArticles:
                        log_and_display(f"L'article {code} n'existe pas dans la base de données", self.text_box, self.root)
                        errorName = f"Article {code} inexistant"
                        skip = messagebox.askyesno("Article inexistant", f"L'article {code} n'existe pas dans la base de données.\n\n Voulez-vous l'ignorer et continuer ?")
                        if skip:
                            undefinedArticles.append(code)
                            log_and_display(f"Article {code} ignoré.", self.text_box, self.root, 0.5)
                            indexation = list(articlesDictionnary.keys())
                            indexActuel = rawDatas.index(code + '\n')
                            if indexActuel == 0:
                                articleSuivant = rawDatas[indexActuel + 1]
                                self.reportDatas["errors"][errorName] = f"Article {code} absent en base de données. Situé en première position, avant {articleSuivant}. Ignoré, opération reprise"
                            elif indexActuel == len(rawDatas) - 1:
                                articlePrecedent = rawDatas[indexActuel - 1]
                                self.reportDatas["errors"][errorName] = f"Article {code} absent en base de données. Situé en dernière position, après {articlePrecedent}. Ignoré, opération reprise"
                            else:
                                articleSuivant = rawDatas[indexActuel + 1]
                                articlePrecedent = rawDatas[indexActuel - 1]
                                self.reportDatas["errors"][errorName] = f"Article {code} absent en base de données. Situé entre {articlePrecedent} et {articleSuivant}. Ignoré, opération reprise"
                            continue
                        else:
                            log_and_display("Annulation de l'opération.", self.text_box, self.root, 0.5)
                            self.reportDatas["errors"][errorName] = f"Interruption de l'opération suite à l'erreur d'article inexistant {code}"

                            # Nettoyage du dossier temporaire
                            shutil.rmtree(tempInventoryDirectory)
                            return
                else:
                    if code in articlesDictionnary:
                        articlesDictionnary[code] += 1
                    else:
                        articlesDictionnary[code] = 1

            # Création du fichier code;quantite dans le dossier temporaire
            outputFile = os.path.join(tempInventoryDirectory, f"inventaire_{currentDate}.txt")
            with open(outputFile, 'w', encoding='utf-8') as file:
                for key, value in articlesDictionnary.items():
                    file.write(f"{key};{value}\n")
                    self.reportDatas["stats"]["different_articles"] += 1

            # Création du tableau des familles scannées
            familles = []
            for key in articlesDictionnary.keys():
                # Récupération de la famille de l'article
                famille = get_famille(self.connection, key)
                if famille is None and key not in undefinedArticles:
                    log_and_display(f"La récupération de la famille pour l'article {key} a échoué", self.text_box, self.root)
                    errorName = f"Erreur de récupération de famille pour l'article {key}"
                    skip = messagebox.askyesno("Échec de récupération",
                                            f"La récupération de la famille pour l'article {key} a échoué.\n\n Voulez-vous l'ignorer et continuer ?")
                    if skip:
                        log_and_display(f"Article {key} ignoré.", self.text_box, self.root, 0.5)
                        self.reportDatas["errors"][errorName] = f"Article {key} ignoré, dû a une erreur de récupération de famille en base de données. Ignoré, opération reprise."
                        continue
                    else:
                        log_and_display("Annulation de l'opération.", self.text_box, self.root, 0.5)
                        self.reportDatas["errors"][errorName] = f"Interruption de l'opération suite à l'erreur de récupération de la famille de l'article {key}"

                        # Nettoyage du dossier temporaire
                        shutil.rmtree(tempInventoryDirectory)
                        return
                if famille not in familles:
                    if not famille_exists(self.connection, famille):
                        log_and_display(f"La famille {famille} n'existe pas dans la base de données", self.text_box, self.root)
                        errorName = f"Famille {famille} inexistante"
                        skip = messagebox.askyesno("Famille inexistante",
                                                f"La famille {famille} n'existe pas dans la base de données.\n\n Voulez-vous l'ignorer et continuer ?")
                        if skip:
                            log_and_display(f"Famille {famille} ignorée.", self.text_box, self.root, 0.5)
                            self.reportDatas["errors"][errorName] = f"Famille {famille} absente en base de données. Ignoré, opération reprise"
                            continue
                        else:
                            log_and_display("Annulation de l'opération.", self.text_box, self.root, 0.5)
                            self.reportDatas["errors"][errorName] = f"Interruption de l'opération suite à l'erreur de famille inexistante {famille}"

                            # Nettoyage du dossier temporaire
                            shutil.rmtree(tempInventoryDirectory)
                            return
                    familles.append(famille)

            # Ajout du nombre de familles au rapport d'exécution
            self.reportDatas["stats"]["familles_count"] = len(familles)

            # Création du dossier pour les familles dans le dossier temporaire
            famillesDirectory = os.path.join(tempInventoryDirectory, "familles")
            if not os.path.exists(famillesDirectory):
                log_and_display(f"Création du dossier {famillesDirectory}...", self.text_box, self.root, 0.5)
                os.makedirs(famillesDirectory)

            # Création de chaque fichier d'inventaire par famille
            for famille in familles:
                log_and_display(f"Création du fichier d'inventaire pour la famille {famille}...", self.text_box, self.root, 1)
                familleFile = os.path.join(famillesDirectory, f"{famille}.txt")
                with open(familleFile, 'w', encoding='utf-8') as file:
                    for key in articlesDictionnary.keys():
                        if get_famille(self.connection, key) == famille:
                            file.write(f"{key};{articlesDictionnary[key]}\n")

            # Exécution de la fonction update_stock
            log_and_display("Lancement de la mise à jour des stocks", self.text_box, self.root, 3)
            self.update_stock(articlesDictionnary)

            # Remplacer l'ancien inventaire si nécessaire
            if inventoryExists:
                log_and_display("Tout s'est bien passé, remplacement de l'ancien inventaire...", self.text_box, self.root, 1)
                
                # Supprimer l'ancien dossier d'inventaire
                log_and_display(f"Suppression de l'ancien inventaire {thisInventoryDirectory}...", self.text_box, self.root, 0.5)
                shutil.rmtree(thisInventoryDirectory)
            
            # Renommer le dossier temporaire en dossier final
            log_and_display(f"Finalisation de l'inventaire...", self.text_box, self.root, 0.5)
            shutil.move(tempInventoryDirectory, thisInventoryDirectory)

            log_and_display("Inventaire terminé.", self.text_box, self.root, 1)

            # Génération du rapport d'exécution
            log_and_display("Génération du rapport d'exécution...", self.text_box, self.root, 1)
            self.reportDatas["stats"]["errors_count"] = len(self.reportDatas["errors"])
            report = generate_report(self.reportDatas)

            log_and_display(f"Rapport d'exécution généré : {report}", self.text_box, self.root, 1)
            userWantOpen = messagebox.askyesno("Rapport généré", f"Le rapport d'exécution d'inventaire a été généré à l'emplacement {report}.\n\n Souhaitez-vous l'ouvrir ?")
            if userWantOpen:
                log_and_display(f"Ouverture du rapport d'exécution...", self.text_box, self.root)
                webbrowser.open(f"file:///{os.path.abspath(report)}")

        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors du traitement du fichier: {str(e)}")
            write_log(f"[ERREUR] {str(e)}")
            # En cas d'erreur, nettoyer le dossier temporaire s'il existe
            if os.path.exists(tempInventoryDirectory):
                try:
                    shutil.rmtree(tempInventoryDirectory)
                    log_and_display(f"Nettoyage du dossier temporaire suite à une erreur", self.text_box, self.root)
                except:
                    pass
            return

    # But : permet de mettre à jour le stock en se basant sur un dictionnaire code => quantité
    def update_stock(self, correctStock):
        print(correctStock)
        # Récupérer tous les articles de la base de données
        allArticles = get_all_articles(self.connection)

        if allArticles is not None:
            for article in allArticles:
                codeArticle = article[0]
                if codeArticle in correctStock:
                    # Mettre à jour l'article dans la base de données
                    self.compare_and_update_article_stock(codeArticle, correctStock[codeArticle])
                else:
                    # Mettre la valeur du stock à 0
                    self.compare_and_update_article_stock(codeArticle, 0)

    # But : Comparer la quantité théorique et la réelle afin de réaliser un mouvement de stock
    def compare_and_update_article_stock(self, code, realQuantity):
        # Récupérer la quantité en stock théorique
        bdArticle = get_article_stock(self.connection, code)
        qteAppro = bdArticle[2]
        qteConso = bdArticle[3]
        qteStock = qteAppro - qteConso

        # Créer un mouvement de stock pour corriger la différence
        log_and_display(f"Mise à jour de l'article {code} à sa nouvelle quantité : {realQuantity}", self.text_box, self.root)
        diff = abs(qteStock - realQuantity)
        if qteStock > realQuantity:
            typeMvt = 'S'
        elif qteStock < realQuantity:
            typeMvt = 'E'
        else:
            typeMvt = None

        # Gestion d'une potentielle erreur lors de la mise à jour
        if typeMvt is not None:
            if not create_mvt(self.connection, typeMvt, bdArticle, diff) :
                log_and_display(f"La mise à jour de l'article {code} a échoué")
            else:
                write_log(f"Mise à jour de l'article {code} réussie")
        else:
            log_and_display(f"Aucune mise à jour nécessaire pour l'article {code}", self.text_box, self.root)