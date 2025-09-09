# Script de recompilation automatique pour Generator Formation
# Auteur: Assistant IA
# Date: $(Get-Date -Format "yyyy-MM-dd")

Write-Host "=== Recompilation de Generator Formation ===" -ForegroundColor Green
Write-Host ""

# Etape 1: Nettoyage des anciens fichiers
Write-Host "1. Nettoyage des anciens fichiers de compilation..." -ForegroundColor Yellow
try {
    if (Test-Path "build") {
        Remove-Item -Recurse -Force build -ErrorAction Stop
        Write-Host "   [OK] Dossier 'build' supprime" -ForegroundColor Green
    }
    if (Test-Path "dist") {
        Remove-Item -Recurse -Force dist -ErrorAction Stop
        Write-Host "   [OK] Dossier 'dist' supprime" -ForegroundColor Green
    }
} catch {
    Write-Host "   [WARN] Erreur lors du nettoyage: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "   Continuer quand meme..." -ForegroundColor Yellow
}

Write-Host ""

# Etape 2: Compilation
Write-Host "2. Debut de la compilation..." -ForegroundColor Yellow
try {
    python build_exe.py
    if ($LASTEXITCODE -eq 0) {
        Write-Host "   [OK] Compilation reussie!" -ForegroundColor Green
    } else {
        Write-Host "   [ERR] Erreur lors de la compilation" -ForegroundColor Red
        exit 1
    }
} catch {
    Write-Host "   [ERR] Erreur lors de la compilation: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""

# Etape 3: Copie du dossier Datas
Write-Host "3. Copie du dossier Datas..." -ForegroundColor Yellow
try {
    if (Test-Path "Datas") {
        Copy-Item -Recurse -Force Datas dist/ -ErrorAction Stop
        Write-Host "   [OK] Dossier Datas copie avec succes" -ForegroundColor Green
    } else {
        Write-Host "   [WARN] Dossier Datas non trouve" -ForegroundColor Yellow
    }
} catch {
    Write-Host "   [ERR] Erreur lors de la copie du dossier Datas: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""

# Etape 4: Verification finale
Write-Host "4. Verification de la compilation..." -ForegroundColor Yellow
if (Test-Path "dist/Generator_Formation.exe") {
    $exeSize = (Get-Item "dist/Generator_Formation.exe").Length / 1MB
    $exeSizeRounded = [math]::Round($exeSize, 1)
    Write-Host ("   [OK] Exe trouve: Generator_Formation.exe ({0} MB)" -f $exeSizeRounded) -ForegroundColor Green
} else {
    Write-Host "   [ERR] Exe non trouve!" -ForegroundColor Red
}

if (Test-Path "dist/Datas") {
    $datasFiles = (Get-ChildItem "dist/Datas" -Recurse -File).Count
    Write-Host "   [OK] Dossier Datas trouve avec $datasFiles fichiers" -ForegroundColor Green
} else {
    Write-Host "   [WARN] Dossier Datas non trouve dans dist/" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=== Recompilation terminee ===" -ForegroundColor Green
Write-Host ""
Write-Host "L'executable se trouve dans: dist/Generator_Formation.exe" -ForegroundColor Cyan
Write-Host ""
Write-Host "Pour tester l'application:" -ForegroundColor White
Write-Host "1. Aller dans le dossier dist/" -ForegroundColor White
Write-Host "2. Double-cliquer sur Generator_Formation.exe" -ForegroundColor White
Write-Host ""
Write-Host "Pour distribuer l'application:" -ForegroundColor White
Write-Host "1. Copier le fichier dist/Generator_Formation.exe" -ForegroundColor White
Write-Host "2. Copier le dossier dist/Datas/ (avec tous ses sous-dossiers)" -ForegroundColor White
Write-Host "3. L'application peut etre executee sur n'importe quel PC Windows" -ForegroundColor White
