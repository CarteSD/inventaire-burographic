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
            print(f"Erreur lors de la récupération de la famille : {e}")
            self.write_log(f"[ERREUR] {e}")
            return None

    def launch_inventory(self):
        # Vider la zone d'informations
        self.text_box.delete(1.0, tk.END)

        # Affichage du message de récupération du fichier d'inventaire
        self.text_box.insert(tk.END, "Récupération du fichier d'inventaire...\n")
        self.root.update()
        self.write_log("Récupération du fichier d'inventaire...")

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

        time.sleep(0.5)
        # Affichage du message de lecture du fichier d'inventaire
        self.text_box.insert(tk.END, "Lecture du fichier d'inventaire...\n")
        self.root.update()
        self.write_log("Lecture du fichier d'inventaire...")

        time.sleep(1)
        # Lecture du fichier d'inventaire
        try:
            with open(filePath, 'r') as file:
                rawDatas = file.readlines()

            time.sleep(0.5)
            # Affichage du message de récupération des articles
            self.text_box.insert(tk.END, "Récupération des quantités de chaque article...\n")
            self.root.update()
            self.write_log("Récupération des quantités de chaque article...")

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
                time.sleep(0.5)
                self.text_box.insert(tk.END, f"Création du dossier {inventairesDirectory}...\n")
                self.root.update()
                self.write_log(f"Création du dossier {inventairesDirectory}...")
                os.makedirs(inventairesDirectory)

            time.sleep(0.5)
            # Affichage du message de création du fichier d'inventaire
            self.text_box.insert(tk.END, "Création du dossier d'inventaire à la date du jour...\n")
            self.root.update()
            self.write_log("Création du dossier d'inventaire à la date du jour...")

            currentDate = datetime.now().strftime("%Y-%m-%d")
            self.text_box.insert(tk.END, f"Date d'inventaire : {currentDate}\n")
            self.root.update()
            thisInventoryDirectory = os.path.join(inventairesDirectory, f"inventaire_{currentDate}")
            time.sleep(0.5)
            if not os.path.exists(thisInventoryDirectory):
                self.text_box.insert(tk.END, f"Création du dossier {thisInventoryDirectory}...\n")
                self.root.update()
                self.write_log(f"Création du dossier {thisInventoryDirectory}...")
                os.makedirs(thisInventoryDirectory)
            else:
                time.sleep(0.5)
                self.text_box.insert(tk.END, f"Le dossier {thisInventoryDirectory} existe déjà.\n")
                self.root.update()
                time.sleep(2)
                self.write_log(f"Le dossier {thisInventoryDirectory} existe déjà.")
                overwrite = messagebox.askyesno("Dossier existe déjà", f"Le dossier {thisInventoryDirectory} existe déjà. Voulez-vous l'écraser ?")
                if overwrite:
                    time.sleep(0.5)
                    self.text_box.insert(tk.END, f"Suppression du dossier {thisInventoryDirectory}...\n")
                    self.root.update()
                    self.write_log(f"Suppression du dossier {thisInventoryDirectory}...")
                    shutil.rmtree(thisInventoryDirectory)
                    time.sleep(0.5)
                    self.text_box.insert(tk.END, f"Création du dossier {thisInventoryDirectory}...\n")
                    self.root.update()
                    self.write_log(f"Création du dossier {thisInventoryDirectory}...")
                    os.makedirs(thisInventoryDirectory)
                else:
                    time.sleep(0.5)
                    self.text_box.insert(tk.END, "Annulation de l'opération.\n")
                    self.root.update()
                    self.write_log("Annulation de l'opération.")
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
                if famille is not None and famille not in familles:
                    familles.append(famille)

            # Création du dossier pour les familles
            famillesDirectory = os.path.join(thisInventoryDirectory, "familles")
            if not os.path.exists(famillesDirectory):
                time.sleep(0.5)
                self.text_box.insert(tk.END, f"Création du dossier {famillesDirectory}...\n")
                self.root.update()
                self.write_log(f"Création du dossier {famillesDirectory}...")
                os.makedirs(famillesDirectory)

            # Création de chaque fichier d'inventaire par famille
            for famille in familles:
                time.sleep(1)
                self.text_box.insert(tk.END, f"Création du fichier d'inventaire pour la famille {famille}...\n")
                self.root.update()
                self.write_log(f"Création du fichier d'inventaire pour la famille {famille}...")
                familleFile = os.path.join(famillesDirectory, f"{famille}.txt")
                with open(familleFile, 'w') as file:
                    for key in articlesDictionnary.keys():
                        if self.get_famille(key) == famille:
                            file.write(f"{key};{articlesDictionnary[key]}\n")

            # Exécution de la fonction create_inventories
            time.sleep(3)
            self.text_box.insert(tk.END, "Lancement de la création des inventaires...\n")
            self.root.update()
            self.write_log("Lancement de la création des inventaires...")
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
            time.sleep(1)
            self.text_box.insert(tk.END, f"Traitement du fichier {file}...\n")
            self.root.update()
            self.write_log(f"Traitement du fichier {file}...")
            with open(file_path, 'r') as f:
                lines = f.readlines()

    def write_log(self, message):
        log_file = LOG_FILE
        with open(log_file, 'a') as f:
            f.write(f"{datetime.now()} - {message}\n")


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