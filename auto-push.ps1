# Script per push automatico su GitHub
# Email Print Service - Auto Push Script

param(
    [string]$CommitMessage = "Auto-commit: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
    [switch]$Force = $false
)

Write-Host "🚀 Email Print Service - Auto Push Script" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green

# Controlla se ci sono modifiche
$status = git status --porcelain
if (-not $status) {
    Write-Host "✅ Nessuna modifica da committare" -ForegroundColor Yellow
    exit 0
}

Write-Host "📝 Modifiche trovate:" -ForegroundColor Cyan
git status --short

# Aggiungi tutti i file
Write-Host "📦 Aggiungendo file..." -ForegroundColor Cyan
git add .

# Commit
Write-Host "💾 Creando commit..." -ForegroundColor Cyan
git commit -m $CommitMessage

if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ Errore durante il commit" -ForegroundColor Red
    exit 1
}

# Push
Write-Host "🚀 Invio su GitHub..." -ForegroundColor Cyan
if ($Force) {
    git push origin main --force
} else {
    git push origin main
}

if ($LASTEXITCODE -eq 0) {
    Write-Host "✅ Push completato con successo!" -ForegroundColor Green
    Write-Host "🌐 Repository: https://github.com/mccoy88f/eps-windows" -ForegroundColor Blue
} else {
    Write-Host "❌ Errore durante il push" -ForegroundColor Red
    Write-Host "💡 Suggerimento: Verifica le credenziali GitHub" -ForegroundColor Yellow
    exit 1
}

Write-Host "🎉 Operazione completata!" -ForegroundColor Green 