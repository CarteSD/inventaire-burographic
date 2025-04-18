# Module d'inventaire Python pour Batigest Connect
Ce module permet le traitement automatique d'un fichier texte extrait d'une douchette, en vue de l'importer dans Batigest Connect. Il est conçu pour fonctionner avec un fichier texte contenant l'ensemble des références scannées, séparées par un retour à la ligne.  

Celui-ci a été réalisé dans le cadre d'un stage en entreprise chez Burographic. En pleine transition de logiciel de gestion interne, il a fallu recréer leur module d'automisation d'inventaire annuel. Après avoir échangé avec les intéressés (secrétaires et directeur), j'ai pu me lancer dans la création de ce module du début à la fin, en prenant en main la base de données internes de Batigest Connect.

## Fonctionnement 
Avant d'effectuer le traitement même du fichier, le programme vérifie qu'un inventaire n'ait pas déjà été effectué à la date du jour. Si tel est le cas, il propose à l'utilisateur d'écraser l'ancien inventaire ou d'interrompre le procédé.

Le traitement du fichier se fait **en 5 étapes** majeures :
### 1. **Lecture du fichier texte** :  

Le fichier texte est ouvert et lu ligne par ligne et converti sous forme de fichier recensant `code:quantité`.  

> [!WARNING]
> Le fichier texte à fournir doit obligatoirement être une série de code barre séparé par des sauts de ligne.
> Dans le cas contraire, des erreurs peuvent survenir.

*Fichier importé :*
```
ABCDEFGHI
ABCDEFGHI
123456789
ABCDEFGHI
123456789
```

*Fichier traité :*
```
ABCDEFGHI;3
123456789;2
```

Cela permet de connaître la quantité scannée de chaque article.

### 2. **Récupération des familles d'articles** : 

Le fichier généré est parcouru, et la famille de chaque article est récupérée depuis la base de données de Batigest Connect. Si cette famille n'avait pas encore été récupérée, elle est ajoutée à un tableau recensant toutes les familles différentes.

### 3. **Création du fichier d'inventaire pour chaque famille** :

Chaque famille fera l'objet d'un fichier texte à son nom, et un fichier d'inventaire pour chaque famille sera créé dans le répertoire de destination. Le fichier d'inventaire contiendra les articles de la famille, ainsi que leur quantité scannée.  
*N.b. : Ces fichiers ne sont pas utiles pour la suite du traitement, mais peuvent l'être afin de réaliser différents traitements manuels annexes.*

### 4. **Mise à jour du stock en base de données** :

Le stock est géré par deux tables différentes, l'une recense l'ensemble des mouvements de stock (Entrées et Sorties), tandis que la seconde recense chaque article avec sa quantité approvisionnée et consommée.

Afin de corriger les différences, le programme calcule la valeur aboslue de la différence entre la quantité en stock (récupérer en soustrayant la quantité approvisonnée à la quantité consommée) et la quantité scanné par les secrétaires. Il créé par la suite le mouvement de stock adapté avec la différence calculée.

Une fois ce mouvement de stock inséré, la table recensant les articles avec leur apprivisonnement et leur consommation est à son tour modifiée. Selon le mouvement de stock effectué auparavant, le programme modifie automatiquement la quantité concernée (approvsionnée ou consommée).

### 5. **Génération du rapport de fin d'exécution** :

Une fois la réalisation complète de l'inventaire terminée, le programme écrit un rapport d'exécution à destination de l'utilisateur, afin de consulter les différentes statistiques telles que le nombre d'articles traités ou les erreurs survenues durant le déroulé.

## Fonctionnalités

Le programme répond à quelques questions de sûreté et de cohérence des données :
- Vérification de l'existence de chaque article avant de le traiter
- Vérification de l'existence de chaque famille d'article
- Création de l'inventaire dans un dossier temporaire, supprimé en cas d'interruption du programme

Il permet également à l'utilisateur de comprendre et de suivre le déroulé du traitement :
- Affichage de la succession des tâches réalisées
- Utilisation de boîtes de dialogue lors de choix à réaliser
- Édition d'un rapport d'exécution à la fin du traitement


## Déploiement

Afin de déployer ce module, assurez-vous d'avoir Python 3.7 ou supérieur installé sur votre machine. Vous pouvez le télécharger depuis le site officiel de Python : [python.org](https://www.python.org/downloads/).  
Il est également nécessaire d'avoir Git installé sur sa machine pour cloner ce dépôt GitHub. Vous pouvez le télécharger depuis le site officiel de Git : [git-scm.com](https://git-scm.com/downloads).

### 1. **Cloner le dépôt** :
Ouvrez un terminal et exécutez la commande suivante pour cloner le dépôt :
```bash
git clone https://github.com/CarteSD/inventaire-burographic.git
cd inventaire-burographic
```

### 2. **Modifier le fichier de constantes** :
Avant de lancer le module, il est nécessaire de modifier le fichier `constantes.example.py` pour y indiquer les chemins d'accès aux fichiers et répertoires nécessaires au bon fonctionnement du module. Une fois fait, renommer le en `constantes.py`.

### 3. **Installer les dépendances** :
Installez les dépendances nécessaires en exécutant la commande suivante :
```bash
pip install -r requirements.txt
```

### 4. **Lancer le module** :
Une fois les dépendances installées, vous pouvez lancer le module en exécutant la commande suivante :
```bash
python main.py
```

Vous pouvez également construire l'exécutable afin de pouvoir lancer le module depuis le chemin que vous souhaitez :
```bash
pyinstaller --onefile --windowed --icon=icone.ico --add-data "constantes.py;." --add-data "icone.ico;." --add-data "report_template.html;." --name "BUROGRAPHIC_Inventaire" main.py
```

Une fois la compilation terminée, l'exécutable se retrouve dans le dossier `/dist`. Vous pouvez le déplacer où vous le souhaitez.