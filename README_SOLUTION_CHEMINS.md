# Solution au problème des chemins dans l'exécutable

## Problème initial
La fonction `generate_filtered_mailmerge_without_tmhf` ne fonctionnait pas correctement après compilation en .exe car les chemins vers les fichiers dans le dossier `Datas` n'étaient pas résolus correctement.

## Cause du problème
Dans l'exécutable PyInstaller, la fonction `resource_path` utilisait `sys._MEIPASS` pour tous les fichiers, mais les fichiers du dossier `Datas` ne sont pas embarqués dans l'exécutable et restent dans le répertoire de l'exécutable.

## Solution implémentée

### 1. Modification de la fonction `resource_path`
La fonction a été modifiée pour distinguer les fichiers `Datas` des autres ressources :

```python
def resource_path(relative_path):
    """Obtenir le chemin absolu vers une ressource, compatible dev et .exe PyInstaller"""
    if hasattr(sys, '_MEIPASS'):
        # Pour les exécutables PyInstaller, utiliser le répertoire de l'exécutable
        # pour les fichiers de données qui ne sont pas embarqués
        if relative_path.startswith("Datas"):
            # Les fichiers de données restent dans le répertoire de l'exécutable
            exe_dir = os.path.dirname(sys.executable)
            full_path = os.path.join(exe_dir, relative_path)
            
            # Créer le dossier Datas s'il n'existe pas
            datas_dir = os.path.join(exe_dir, "Datas")
            if not os.path.exists(datas_dir):
                os.makedirs(datas_dir, exist_ok=True)
                # Créer aussi les sous-dossiers
                os.makedirs(os.path.join(datas_dir, "documents"), exist_ok=True)
                os.makedirs(os.path.join(datas_dir, "config"), exist_ok=True)
                os.makedirs(os.path.join(datas_dir, "Log"), exist_ok=True)
            
            return full_path
        else:
            # Les autres ressources peuvent être dans _MEIPASS
            return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)
```

### 2. Comportement de la fonction
- **En mode développement** : Utilise le répertoire courant
- **En mode exécutable** : 
  - Pour les fichiers `Datas/*` : Utilise le répertoire de l'exécutable
  - Pour les autres fichiers : Utilise `sys._MEIPASS`

### 3. Création automatique des dossiers
La fonction crée automatiquement les dossiers `Datas`, `Datas/documents`, `Datas/config`, et `Datas/Log` s'ils n'existent pas dans le répertoire de l'exécutable.

## Distribution de l'application

### Fichiers nécessaires
1. `Generator_Formation.exe` (l'exécutable)
2. Dossier `Datas/` avec tous ses sous-dossiers et fichiers

### Structure recommandée
```
Generator_Formation.exe
Datas/
├── documents/
│   ├── source_publipostage.xlsx
│   ├── source_publipostage_sans_TMHF.xlsx
│   ├── CONVENTION-Sxx 2025-BUSSY.docx
│   └── ... (autres fichiers)
├── config/
│   └── default_excel.txt
└── Log/
    └── Log_export.xlsx
```

## Test de la solution

### Script de test
Un script `test_resource_path.py` a été créé pour tester la fonction `resource_path` :

```bash
python test_resource_path.py
```

### Test de l'exécutable
1. Compiler l'application : `python build_exe.py`
2. Copier le dossier `Datas` dans `dist/` : `Copy-Item -Recurse -Force Datas dist/`
3. Tester l'exécutable : `dist/Generator_Formation.exe`

## Fonctions affectées
- `generate_filtered_mailmerge_without_tmhf()` : Génère le fichier filtré sans TMHF
- `send_billing_email()` : Utilise le fichier filtré pour créer un mail de facturation
- Toutes les fonctions qui ouvrent des fichiers dans le dossier `Datas`

## Notes importantes
- L'application fonctionne maintenant correctement en mode développement ET en mode exécutable
- Les fichiers de données restent accessibles et modifiables
- La structure des dossiers est créée automatiquement si nécessaire
