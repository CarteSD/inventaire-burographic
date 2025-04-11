import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pyodbc
from datetime import datetime
from constantes import *
import shutil
import time

class Inventaire:
    def __init__(self, root):
        self.root = root;
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
        self.welcomeMessage = tk.Label(self.mainFrame, text="Bienvenue dans le module d'inventaire de BUROGRAPHIC", font=("Arial", 16, "bold"))
        self.description = tk.Label(self.mainFrame, text="Ce module vous permet de créer un inventaire semestriel pour votre entreprise", font=("Arial", 12))
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
        self.text_box = tk.Text(self.results_frame, height=15, width=80, yscrollcommand=self.text_scrollbar.set, state="normal")
        self.text_scrollbar.config(command=self.text_box.yview)

        # Placement de la boîte de texte et de la barre de défilement
        self.text_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.text_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def select_file(self):
        filename = filedialog.askopenfilename(
            title="Sélectionnez un fichier texte",
            filetypes=(("Fichiers texte", "*.txt"), ("Tous les fichiers", "*.*"))
        )
        if filename:
            self.InventoryFilePath.set(filename)

    def get_famille(self, item):
        try:
            if not self.connection:
                self.connection = database_connection()

            cursor = self.connection.cursor()
            query = "SELECT Famille FROM ElementDef WHERE Code = ?"
            cursor.execute(query, item)
            result = cursor.fetchone()
            if result:
                return result[0].rstrip('.') # Suppression du point final
            else:
                return None

        except pyodbc.Error as e:
            messagebox.showerror("Erreur", f"Erreur lors de la récupération de la famille : {str(e)}")
            self.write_log(f"[ERREUR] {str(e)}")
            return None

    def launch_inventory(self):
        # Vider la zone d'informations
        self.text_box.delete(1.0, tk.END)

        # Affichage du message de récupération du fichier d'inventaire
        self.log_and_display("Récupération du fichier d'inventaire...")

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
        self.log_and_display("Lecture du fichier d'inventaire...", 0.5)

        # Lecture du fichier d'inventaire
        try:
            with open(filePath, 'r') as file:
                rawDatas = file.readlines()

            # Affichage du message de récupération des articles
            self.log_and_display("Récupération des articles...", 1)

            # Création d'un dictionnaire pour transformer le fichier en code => quantité
            articlesDictionnary = {}
            for code in rawDatas:
                code = code.replace("\n", "")
                if code in articlesDictionnary:
                    articlesDictionnary[code] += 1
                else:
                    articlesDictionnary[code] = 1

            inventairesDirectory = ".\\inventaires"
            if not os.path.exists(inventairesDirectory):
                self.log_and_display(f"Création du dossier {inventairesDirectory}...", 0.5)
                os.makedirs(inventairesDirectory)

            # Affichage du message de création du fichier d'inventaire
            self.log_and_display("Création du dossier d'inventaire à la date du jour", 0.5)

            currentDate = datetime.now().strftime("%Y-%m-%d")
            self.log_and_display(f"Date d'inventaire : {currentDate}")
            thisInventoryDirectory = os.path.join(inventairesDirectory, f"inventaire_{currentDate}")

            if not os.path.exists(thisInventoryDirectory):
                self.log_and_display(f"Création du dossier {thisInventoryDirectory}...", 0.5)
                os.makedirs(thisInventoryDirectory)
            else:
                self.log_and_display(f"Le dossier {thisInventoryDirectory} existe déjà", 0.5)

                time.sleep(2)
                overwrite = messagebox.askyesno("Dossier déjà existant", f"Le dossier {thisInventoryDirectory} existe déjà. Voulez-vous l'écraser ?")

                if overwrite:
                    self.log_and_display(f"Suppression du dossier {thisInventoryDirectory}...", 0.5)
                    shutil.rmtree(thisInventoryDirectory)

                    self.log_and_display(f"Réécriture du dossier {thisInventoryDirectory}...", 0.5)
                    os.makedirs(thisInventoryDirectory)
                else:
                    self.log_and_display("Annulation de l'opération.", 0.5)
                    return

            # Création du fichier code;quantite
            outputFile = os.path.join(thisInventoryDirectory, f"inventaire_{currentDate}.txt")
            with open(outputFile, 'w') as file:
                for key, value in articlesDictionnary.items():
                    file.write(f"{key};{value}\n")

            # Création du tableau des familles scannées
            familles = []
            for key in articlesDictionnary.keys():
                famille = self.get_famille(key)
                if famille is None:
                    self.log_and_display(f"La récupération de la famille pour l'article {key} a échoué")
                    skip = messagebox.askyesno("Dossier déjà existant", f"La récupération de la famille pour l'article {key} a échoué.\nIl se peut que cet article ne soit pas enregsitré dans la base de données.\n\n Voulez-vous l'ignorer et continuer ?")
                    if skip:
                        self.log_and_display(f"Article {key} ignoré.", 0.5)
                        continue
                    else:
                        self.log_and_display("Annulation de l'opération.", 0.5)
                        return
                if famille not in familles:
                    familles.append(famille)

            # Création du dossier pour les familles
            famillesDirectory = os.path.join(thisInventoryDirectory, "familles")
            if not os.path.exists(famillesDirectory):
                self.log_and_display(f"Création du dossier {famillesDirectory}...", 0.5)
                os.makedirs(famillesDirectory)

            # Création de chaque fichier d'inventaire par famille
            for famille in familles:
                self.log_and_display(f"Création du fichier d'inventaire pour la famille {famille}...", 1)
                familleFile = os.path.join(famillesDirectory, f"{famille}.txt")
                with open(familleFile, 'w') as file:
                    for key in articlesDictionnary.keys():
                        if self.get_famille(key) == famille:
                            file.write(f"{key};{articlesDictionnary[key]}\n")

            # Exécution de la fonction create_inventories
            self.log_and_display("Lancement de la création des inventaires par famille...", 3)
            self.create_inventories(famillesDirectory)

        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors du traitement du fichier: {str(e)}")
            self.write_log(f"[ERREUR] {str(e)}")
            return

    def create_inventories(self, directory):
        if not os.path.exists(directory):
            messagebox.showerror("Erreur", f"Le dossier d'inventaire par famille {directory} n'existe pas.")
            self.write_log("[ERREUR] Le dossier d'inventaire n'existe pas.")
            return

        # Récupération de la liste des fichiers d'inventaire
        files = []
        for f in os.listdir(directory):
            file_path = os.path.join(directory, f)
            if os.path.isfile(file_path) and f.endswith(".txt"):
                files.append(f)

        if not self.connection:
            self.connection = database_connection()

        for file in files:
            file_path = os.path.join(directory, file)
            self.log_and_display(f"Traitement de la famille {file.replace('.txt', '')}...", 1)
            self.create_inventory_famille(file_path)

    def write_log(self, message):
        log_file = LOG_FILE
        with open(log_file, 'a') as f:
            f.write(f"{datetime.now()} - {message}\n")

    def log_and_display(self, message, delay = 0):
        if delay:
            time.sleep(delay)
        self.text_box.insert(tk.END, message + "\n")
        self.root.update()
        self.write_log(message)

    def create_inventory_famille(self, file):
        with open(file, 'r') as f:
            lines = f.readlines()
            for line in lines:
                self.insert_line_in_inventory(line)

    def insert_line_in_inventory(self, line):
        args = line.replace("\n", "").split(";")
        self.log_and_display(f"Insertion de l'article {args[0]} : quantité {args[1]}", 1)


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

root = tk.Tk()
app = Inventaire(root)
root.mainloop()