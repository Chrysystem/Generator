# Compilation de l'Application Generator Formation

Ce guide vous explique comment compiler l'application `excel_to_word_app.py` en fichier exécutable `.EXE` avec toutes ses dépendances.

## Prérequis

- Python 3.7 ou supérieur installé
- Windows 10/11
- Connexion Internet (pour télécharger les dépendances)

## Méthode 1 : Compilation Automatique (Recommandée)

1. **Double-cliquez sur le fichier `compiler.bat`**
   - Ce script installera automatiquement toutes les dépendances
   - Compilera l'application en .EXE
   - Nettoiera les fichiers temporaires

2. **Attendez la fin de la compilation**
   - Le processus peut prendre 2-5 minutes
   - Vous verrez des messages de progression

3. **Récupérez votre exécutable**
   - L'exécutable se trouve dans le dossier `dist/`
   - Nom du fichier : `Generator_Formation.exe`

## Méthode 2 : Compilation Manuelle

Si la méthode automatique ne fonctionne pas :

1. **Ouvrez PowerShell ou Command Prompt dans ce dossier**

2. **Installez les dépendances :**
   ```bash
   pip install -r requirements.txt
   ```

3. **Lancez la compilation :**
   ```bash
   python build_exe.py
   ```

## Structure des Fichiers Requis

Assurez-vous que ces fichiers sont présents dans le dossier :
- `excel_to_word_app.py` (application principale)
- `logo-Toyota-Solo.ico` (icône de l'application)
- `LogoTMH.png` (logo TMH)
- `Datas/` (dossier avec les documents)

## Résultat de la Compilation

Après compilation réussie :
- ✅ `dist/Generator_Formation.exe` - Votre application exécutable
- ✅ L'exécutable inclut toutes les dépendances Python
- ✅ L'application peut être exécutée sur n'importe quel PC Windows

## Distribution

Pour distribuer l'application :
1. Copiez le fichier `dist/Generator_Formation.exe`
2. Copiez le dossier `Datas/` (s'il n'est pas inclus dans l'exe)
3. L'application fonctionne sans installation de Python

## Dépannage

### Erreur "Module not found"
- Vérifiez que toutes les dépendances sont installées
- Relancez la compilation

### Erreur "File not found"
- Vérifiez que tous les fichiers requis sont présents
- Vérifiez les noms de fichiers (sensibles à la casse)

### L'exécutable ne démarre pas
- Essayez de l'exécuter depuis un terminal pour voir les erreurs
- Vérifiez que le dossier `Datas/` est présent

## Support

Si vous rencontrez des problèmes :
1. Vérifiez que Python est bien installé : `python --version`
2. Vérifiez que pip fonctionne : `pip --version`
3. Essayez de réinstaller les dépendances : `pip install --upgrade -r requirements.txt`








