# Script de recompilation automatique pour Generator Formation
# Auteur: Assistant IA
# Date: $(Get-Date -Format "yyyy-MM-dd")

Write-Host "=== Recompilation de Generator Formation ===" -ForegroundColor Green
Write-Host ""

# Étape 1: Nettoyage des anciens fichiers
Write-Host "1. Nettoyage des anciens fichiers de compilation..." -ForegroundColor Yellow
try {
    if (Test-Path "build") {
        Remove-Item -Recurse -Force build -ErrorAction Stop
        Write-Host "   ✓ Dossier 'build' supprimé" -ForegroundColor Green
    }
    if (Test-Path "dist") {
        Remove-Item -Recurse -Force dist -ErrorAction Stop
        Write-Host "   ✓ Dossier 'dist' supprimé" -ForegroundColor Green
    }
} catch {
    Write-Host "   ⚠ Erreur lors du nettoyage: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "   Continuer quand même..." -ForegroundColor Yellow
}

Write-Host ""

# Étape 2: Compilation
Write-Host "2. Début de la compilation..." -ForegroundColor Yellow
try {
    python build_exe.py
    if ($LASTEXITCODE -eq 0) {
        Write-Host "   ✓ Compilation réussie!" -ForegroundColor Green
    } else {
        Write-Host "   ❌ Erreur lors de la compilation" -ForegroundColor Red
        exit 1
    }
} catch {
    Write-Host "   ❌ Erreur lors de la compilation: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""

# Étape 3: Copie du dossier Datas
Write-Host "3. Copie du dossier Datas..." -ForegroundColor Yellow
try {
    if (Test-Path "Datas") {
        Copy-Item -Recurse -Force Datas dist/ -ErrorAction Stop
        Write-Host "   ✓ Dossier Datas copié avec succès" -ForegroundColor Green
    } else {
        Write-Host "   ⚠ Dossier Datas non trouvé" -ForegroundColor Yellow
    }
} catch {
    Write-Host "   ❌ Erreur lors de la copie du dossier Datas: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""

# Étape 4: Vérification finale
Write-Host "4. Vérification de la compilation..." -ForegroundColor Yellow
if (Test-Path "dist/Generator_Formation.exe") {
    $exeSize = (Get-Item "dist/Generator_Formation.exe").Length / 1MB
    Write-Host "   ✓ Exécutable trouvé: Generator_Formation.exe ($([math]::Round($exeSize, 1)) MB)" -ForegroundColor Green
} else {
    Write-Host "   ❌ Exécutable non trouvé!" -ForegroundColor Red
}

if (Test-Path "dist/Datas") {
    $datasFiles = (Get-ChildItem "dist/Datas" -Recurse -File).Count
    Write-Host "   ✓ Dossier Datas trouvé avec $datasFiles fichiers" -ForegroundColor Green
} else {
    Write-Host "   ⚠ Dossier Datas non trouvé dans dist/" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=== Recompilation terminée ===" -ForegroundColor Green
Write-Host ""
Write-Host "L'exécutable se trouve dans: dist/Generator_Formation.exe" -ForegroundColor Cyan
Write-Host ""
Write-Host "Pour tester l'application:" -ForegroundColor White
Write-Host "1. Aller dans le dossier dist/" -ForegroundColor White
Write-Host "2. Double-cliquer sur Generator_Formation.exe" -ForegroundColor White
Write-Host ""
Write-Host "Pour distribuer l'application:" -ForegroundColor White
Write-Host "1. Copier le fichier dist/Generator_Formation.exe" -ForegroundColor White
Write-Host "2. Copier le dossier dist/Datas/ (avec tous ses sous-dossiers)" -ForegroundColor White
Write-Host "3. L'application peut être exécutée sur n'importe quel PC Windows" -ForegroundColor White
