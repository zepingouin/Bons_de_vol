# Bons_de_vol
_**Création de bons authentifiés par QR-Code**_

Les bons sont générés au format PDF avec un QR-Code permettant de les authentifier.

## Pré-requis
* Environnement (au choix) :
  * Microsoft Windows 10
  * Linux Ubuntu 20
* Logiciels Windows :
  * Microsoft Office
  * Gpg4win / GnuPG
  * Python
* Logiciels Linux :
  * LibreOffice
  * GnuPG
  * Python
```
pip install -r requirements.txt
```
Pour Windows, ajouter la variable d'environnement `PYTHONUTF8=1`.
## Utilisation

Le fichier ```config.ini``` est à éditer selon la configuration, certains des paramètres
de ce fichier sont configurables depuis le logiciel dans le menu ```Fichier/Préférences```.

Exécuter le fichier ```Bons.py``` avec Python.
