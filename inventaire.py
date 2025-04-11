import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pyodbc

class Inventaire:
    def __init__(self, root):
        self.root = root;
        self.root.title("BUROGRAPHIC - Inventaire")
        self.root.geometry("800x600")
        self.root.iconbitmap("icone.ico")

        # Variables pour le chemin du fichier d'inventaire
        self.InventoryfilePath = tk.StringVar()

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

        self.file_label = tk.Label(self.fileSelectionFrame, text="Fichier sélectionné:", font=("Arial", 10))
        self.file_label.pack(anchor=tk.W)

        # Création du bouton pour choisir le fichier d'inventaire
        self.filePathEntry = tk.Entry(self.fileSelectionFrame, textvariable=self.InventoryfilePath, width=50)
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

    def select_file(self):
        filename = filedialog.askopenfilename(
            title="Sélectionnez un fichier texte",
            filetypes=(("Fichiers texte", "*.txt"), ("Tous les fichiers", "*.*"))
        )
        if filename:
            self.InventoryfilePath.set(filename)

    def get_famille(self, item):
        try:
            connection = database_connection()
            if not connection:
                return "Erreur de connexion"

            cursor = connection.cursor()
            query = "SELECT Famille FROM ElementDef WHERE Code = ?"
            cursor.execute(query, item)
            result = cursor.fetchone()
            if result:
                return result[0]
            else:
                return None

        except pyodbc.Error as e:
            print(f"Erreur lors de la récupération de la famille : {e}")
            return None

    def launch_inventory(self):
        # Récupération du fichier sélectionné
        filePath = self.InventoryfilePath.get()
        if not filePath:
            messagebox.showerror("Erreur", "Veuillez sélectionner un fichier d'inventaire.")
            return
        if not os.path.exists(filePath):
            messagebox.showerror("Erreur", "Le fichier sélectionné n'existe pas.")
            return
        if not filePath.endswith(".txt"):
            messagebox.showerror("Erreur", "Le fichier sélectionné n'est pas un fichier texte.")
            return

        # Lecture du fichier d'inventaire
        try:
            with open(filePath, 'r') as file:
                rawDatas = file.readlines()

            # Création d'un dictionnaire pour transformer le fichier en code => quantité
            articlesDictionnary = {}
            for code in rawDatas:
                code = code.replace("\n", "")
                if code in articlesDictionnary:
                    articlesDictionnary[code] += 1
                else:
                    articlesDictionnary[code] = 1

            # Création du fichier code;quantite
            outputFile = "inventaire_quantite.txt"
            with open(outputFile, 'w') as file:
                for key, value in articlesDictionnary.items():
                    file.write(f"{key};{value}\n")

            # Création du tableau des familles scannées
            familles = []

            for key in articlesDictionnary.keys():
                famille = self.get_famille(key)
                if famille is not None and famille not in familles:
                    familles.append(famille)

        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors du traitement du fichier: {str(e)}")
            return

def database_connection():
    try:
        connection = pyodbc.connect(
            "Driver={SQL Server};"
            "Server=DESKTOP-D5H040D\\SAGEBAT;" 
            "Database=BTG_DOS_SOC01;"
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