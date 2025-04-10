import tkinter as tk
from tkinter import filedialog

class Inventaire:
    def __init__(self, root):
        self.root = root;
        self.root.title("BUROGRAPHIC - Inventaire")
        self.root.geometry("800x600")

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
        return

    def launch_inventory(self):
        return

root = tk.Tk()
app = Inventaire(root)
root.mainloop()