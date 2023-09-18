# Bons_de_vol
_**Création de bons authentifiés**_

Les bons sont générés au format PDF avec un QR-Code contenant les informations signées permettant de les authentifier.

## Pré-requis
* Environnement (au choix) :
  * Microsoft Windows 10
  * Linux Ubuntu 20
* Logiciels Windows :
  * Microsoft Office ou [LibreOffice](https://www.libreoffice.org/) (expérimental)
  * [Gpg4win](https://www.gpg4win.org/)
  * [Python](https://www.python.org/downloads/) (seulement pour Microsoft Office)
  * [Git](https://git-scm.com/download/win)
* Logiciels Linux :
  * LibreOffice
  * GnuPG
  * Python
  * Git
---
* Pour Windows :
  * Ajouter la variable d'environnement `PYTHONUTF8=1`
  * Si LibreOffice est utilisé (expérimental) :
    * Télécharger [get-pip.py](https://bootstrap.pypa.io/get-pip.py) et l'enregistrer dans `C:\Program Files\LibreOffice\program`
    * Ajouter à PATH les chemins suivants :
      * `C:\Program Files\LibreOffice\program`
      * `C:\Program Files\LibreOffice\program\python-core-x.y.zz\Scripts` (`x.y.zz` est la version de Python livrée avec l'installation de LibreOffice)
    * Exécuter les commandes suivantes dans une fenêtre _Invite de commandes_ en tant qu'administrateur :
      * `cd "C:\Program Files\LibreOffice\program"`
      * `python get-pip.py`
      * `move mailmerge.py mailmerge.py.back` (En conflit avec `docx-mailmerge`)
    * Modifier la ligne 43 du fichier `%APPDATA%\Python\Pythonxy\site-packages\unoserver\server.py` (`xy` est la version de Python livrée avec l'installation de LibreOffice)
      * `process = subprocess.Popen(cmd, shell=True)`
* Pour Linux :
  * Si nécessaire, ajouter les polices de caractères Microsoft :
     * Impact
     * Consolas
     * Times_New_Roman.
---
## Installation
* Télécharger les fichiers :
  * `git clone https://github.com/zepingouin/Bons_de_vol.git`
* Installer les modules Python nécessaires :
  * `pip install -r requirements.txt`
## Utilisation

Le fichier ```config.ini``` est à éditer selon la configuration, certains des paramètres
de ce fichier sont configurables depuis le logiciel dans le menu ```Fichier/Préférences```.

Exécuter le fichier ```Bons.py``` avec Python.
