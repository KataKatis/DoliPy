# DoliPy
[![Python 3.9.6](https://img.shields.io/badge/Python-3.9.6-blue.svg)](https://www.python.org/) [![Selenium 4.1.1](https://img.shields.io/badge/selenium-4.1.1-green.svg)](https://pypi.org/project/selenium/) [![Openpyxl 3.0.9](https://img.shields.io/badge/openpyxl-3.0.9-green.svg)](https://openpyxl.readthedocs.io/en/stable/)

DoliPy est un outil censé faciliter la saisie de notes de frais, de devis ou encore de transaction de comptabilité dans Dolibarr, grâce à des documents excel. <br>
:warning::rotating_light: **Avant de commencer, il faut savoir que DoliPy nécessite quelques installations indispensables à son exécution :**
- Python (la version 3.9.6 de préférence)
- Excel, OpenOffice ou LibreOffice afin de lire des fichiers excel
- Un driver afin d'utiliser google chrome avec python

L'installation de Python et du driver sont expliquées en détails ci-dessous. <br>
:warning::rotating_light:DoliPy n'est compatible qu'avec le navigateur Chrome.

## Installation

Avant de démarrer les installations suivantes, créez un nouveau dossier dans votre machine (sur le bureau par exemple) dans lequel seront stockés les fichiers python, le driver...

> Le nom de votre dossier n'influencera pas l'exécution du programme, vous êtes libres de choisir celui que vous voulez. <br>
> Si vous voulez le même dossier qui sera pris en exemple dans le tutoriel, créez-le sur le bureau et appelez-le "*dolipy*".

### Python

Pour installer **python** sur votre ordinateur, rendez-vous sur [le site officiel de python](https://www.python.org/), dans l'onglet *Download* puis téléchargez **python** en fonction de votre système d'exploitation (MAC OS / Windows / Linux).

Si vous êtes sur Windows, téléchargez un executable installer, ici entourés en rouge :

![python-installer](https://exposesnt.herokuapp.com/Images/python-installer.png)

Pour savoir si votre système Windows est en 32 et 64-bit, clickez sur le bouton **Démarrer > Paramètres > Système > À propos de**. Sur la droite, sous **Spécifications de l’appareil**, consultez **Type de système**

Exécutez ensuite le fichier ".exe" qui vient d'être téléchargé.
:warning::rotating_light: Assurez-vous lors de l'installation de sélectionner **Install launcher for all users** et **Add Python 3.7 to PATH** (cf image ci-dessous). Cette étape est indispensable.

![python-setup](https://phoenixnap.com/kb/wp-content/uploads/2021/04/python-setup.png)

Dans la boîte de dialogue suivante, selectionner **Disable path length limit**.

![python-limit](https://phoenixnap.com/kb/wp-content/uploads/2021/04/python-setup-completed.png)

Pour vous assurer que **python** est bien installer sur votre machine, ouvrez "**Invite de commandes**" dans le menu démarrer en tapant **cmd**. Dans le terminal exécutez la commande : `python --version`. Si **python** est bien installé, sa version s'affichera.


### Modules

Pour installer les modules complémentaires nécessaires à l'exécution de *script.py*, ouvrez d'abord un terminal. Dans le menu démarrer, tapez **cmd** et ouvrez "**Invite de commandes**". Il est recommandé de copier-coller les commandes suivantes afin d'éviter toute erreur de syntaxe.

Entrez cette première commande et appuyez sur votre touche "Entrer" : `pip install selenium`<br>
Entrez cette seconde commande et appuyez sur votre touche "Entrer" : `pip install openpyxl`

À présent, les modules complémentaires à **python** sont installés.

### Driver

Pour que **python** puisse utiliser votre navigateur, il est obligatoire d'installer un driver (ou "pilote" en français). Il s'agit d'un simple exécutable à placer dans le dossier que vous avez créer au début.

Le non-respect de cette étape vous donnera une erreur :
*selenium.common.exceptions.WebDriverException: Message: ‘geckodriver’ executable needs to be in PATH*

Notez dans un premier temps la version de votre navigateur chrome. Pour cela, clickez dans chrome sur les 3 petits points en haut à droite, puis **Aide > À propos de chrome** :

![python-limit](https://zupimages.net/up/21/22/i0an.jpg)

Ici, la version de mon navigateur est 91 (deux premiers chiffres). Rendez-vous ensuite sur ce [site](https://sites.google.com/chromium.org/driver/) et téléchargez le driver correspondant à votre version de chrome. Lorsque le pilote est téléchargé, déplacez-le dans le dossier que vous avez créé au début.

### Scripts python

Pour installer les scripts python, clickez en haut de la page sur **Code > Download ZIP**. Extrayez maintenant les fichiers contenus dans le zip, dans le dossier créé au début. Vous devez y trouvez plusieurs fichiers python en *.py* ainsi qu'un fichier excel au format *xlsx*.

### Mise à jour

Aucune notification ne vous sera envoyé en cas de mise à jour. Il faut retélécharger les scripts de temps à autre pour être sur qu'ils soient à jour.

## Utilisation

> Le fichier *template.xlsx* est composé de plusieurs feuilles (NDF[^NDF], Devis, Compta) que vous devez remplir au préalable (elles ne sont pas toutes à remplir), pour utiliser DoliPy.

Commencez par ouvrir un terminal : dans le menu démarrer tapez **cmd** et ouvrez "**Invite de commandes**". Utilisez la commande **cd** pour vous déplacer dans les répertoire/dossiers de votre PC. Si le dossier que vous avez créé au début s'appelle "*dolipy*" et qu'il se situe sur votre bureau, vous taperez la commande :
```shell
cd Desktop\dolipy
```

Pour vous assurez que vous vous situez dans le bon répertoire, tapez la commande[^dir] `dir`. S'afficheront ensuite les fichiers présents dans le répertoire, vous devriez voir : *script.py*, *note_de_frais.py*, *devis.py*, *chromedriver.exe*, *template.xlsx*...

La commande à utiliser pour lancer DoliPy est : 
```shell
python script.py --Site "lien" --Username "nom d'utilisateur" --Password "mdp" --Start ligne_de_début_(sans guillemets) --End ligne_de_fin(sans guillemets)
```

Chaque élément après *--xxx* est un paramètre et est obligatoire.

Pour le paramètre *--Site*, copiez-collez le lien Dolibarr avec lequel vous travaillez (si vous voulez remplir un devis par exemple, vous prenez le lien du devis).

À *--Username*, entrez votre nom d'utilisateur, et *--Password* votre mot de passe. Ces trois premiers paramètres sont obligatoirement entre guillemets.

Après *--Start* et *--End*, entrez les lignes de début et de fin de votre document excel. Si vous souhaitez sélectionné qu'une ligne, vous mettrez la même valeur à *--Start* et *--End*.

Une fois ces étapes terminé, tapez sur votre touche **Entrez**.

### Besoin d'aide sur les commandes ??

Tapez 
```shell
python script.py -h
```
OU
```shell
python script.py --help
```
et vous obtiendrez une page d'aide sur les commandes.

## Les erreurs à ne pas faire

- Le fichier *template* doit forcément être aux format *.xlsx* et non pas *.xls* ou autre
- Dans les feuilles NDF et Devis, toutes les colonnes devront être remplies (cela ne saura plus obligatoire par la suite).
- Ne pas oublier de paramètre lors de l'exécution de la commande.
- Le "Site", "Username" et "Password" que vous entrez lors de l'exécution de la commande sont entre guillemets.
- Le "Start" et "End" que vous entrez lors de l'exécution de la commande ne sont pas entre guillemets et sont des nombres entiers.
- Enregistrer *template.xlsx* après toute modifications, auquel cas python ne pourra pas les prendre en compte.
- Les paramètres de la commande doivent être dans le bon ordre (Site, Username, Password, Start, End)

[^NDF]: Note De Frais
[^dir]: en tout cas sur Windows