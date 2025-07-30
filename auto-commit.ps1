# Script per commit automatico con monitoraggio
# Email Print Service - Auto Commit Script

param(
    [string]$WatchPath = ".",
    [int]$IntervalSeconds = 30,
    [switch]$Continuous = $false
)

Write-Host "Email Print Service - Auto Commit Monitor" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green
Write-Host "Monitorando: $WatchPath" -ForegroundColor Cyan
Write-Host "Intervallo: $IntervalSeconds secondi" -ForegroundColor Cyan
Write-Host "Continuo: $Continuous" -ForegroundColor Cyan
Write-Host ""

if ($Continuous) {
    Write-Host "Avvio monitoraggio continuo... (Ctrl+C per fermare)" -ForegroundColor Yellow
    Write-Host ""
}

do {
    # Controlla modifiche
    $status = git status --porcelain
    
    if ($status) {
        Write-Host "Modifiche rilevate: $(Get-Date -Format 'HH:mm:ss')" -ForegroundColor Yellow
        
        # Mostra file modificati
        git status --short
        
        # Commit automatico
        $commitMsg = "Auto-commit: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $(git status --short | Measure-Object | Select-Object -ExpandProperty Count) files"
        
        Write-Host "Commit automatico..." -ForegroundColor Cyan
        git add .
        git commit -m $commitMsg
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "Commit completato" -ForegroundColor Green
            
            # Push automatico
            Write-Host "Push automatico..." -ForegroundColor Cyan
            git push origin main
            
            if ($LASTEXITCODE -eq 0) {
                Write-Host "Push completato" -ForegroundColor Green
            } else {
                Write-Host "Errore push - prova manuale" -ForegroundColor Red
            }
        } else {
            Write-Host "Errore commit" -ForegroundColor Red
        }
        
        Write-Host ""
    } else {
        if ($Continuous) {
            Write-Host "Nessuna modifica... $(Get-Date -Format 'HH:mm:ss')" -ForegroundColor Gray
        }
    }
    
    if ($Continuous) {
        Start-Sleep -Seconds $IntervalSeconds
    }
} while ($Continuous)

if (-not $Continuous) {
    # Esegui una sola volta
    if ($status) {
        Write-Host "Modifiche trovate, eseguendo commit..." -ForegroundColor Yellow
        & .\auto-push.ps1
    } else {
        Write-Host "Nessuna modifica da committare" -ForegroundColor Green
    }
} 