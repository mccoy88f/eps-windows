# 🔄 Automazione Git per Email Print Service

## 📋 Script Disponibili

### 1. **auto-push.ps1** - Push Manuale
Script per commit e push manuale con messaggio personalizzato.

```powershell
# Push normale
.\auto-push.ps1

# Push con messaggio personalizzato
.\auto-push.ps1 -CommitMessage "Fix: UI improvements"

# Push forzato (solo se necessario)
.\auto-push.ps1 -Force
```

### 2. **auto-commit.ps1** - Commit Automatico
Script per commit automatico con monitoraggio opzionale.

```powershell
# Commit una volta sola
.\auto-commit.ps1

# Monitoraggio continuo (ogni 30 secondi)
.\auto-commit.ps1 -Continuous

# Monitoraggio personalizzato (ogni 60 secondi)
.\auto-commit.ps1 -Continuous -IntervalSeconds 60
```

## 🚀 Setup Rapido

### **Prima Esecuzione**
```powershell
# Abilita esecuzione script PowerShell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# Testa lo script di push
.\auto-push.ps1
```

### **Configurazione Credenziali**
Se richiesto durante il primo push:
- **Username**: `mccoy88f`
- **Password**: Usa il Personal Access Token (NON la password GitHub)

## 🔧 Configurazione GitHub

### **Genera Personal Access Token**
1. Vai su [GitHub Settings > Developer settings > Personal access tokens](https://github.com/settings/tokens)
2. Clicca "Generate new token (classic)"
3. Seleziona permessi:
   - `repo` (tutti i permessi del repository)
   - `workflow` (opzionale, per GitHub Actions)
4. Copia il token generato

### **Configura Credenziali Locali**
```powershell
# Configura credential manager
git config --global credential.helper manager-core

# Testa autenticazione
git push origin main
```

## 📝 Esempi di Utilizzo

### **Sviluppo Normale**
```powershell
# Dopo aver fatto modifiche
.\auto-push.ps1 -CommitMessage "Add new feature"
```

### **Sviluppo Continuo**
```powershell
# Avvia monitoraggio continuo
.\auto-commit.ps1 -Continuous

# In un altro terminale, lavora sui file
# Le modifiche verranno committate automaticamente
```

### **Debug e Testing**
```powershell
# Commit con messaggio dettagliato
.\auto-push.ps1 -CommitMessage "Debug: Fix UI blocking issue"

# Monitoraggio rapido (ogni 10 secondi)
.\auto-commit.ps1 -Continuous -IntervalSeconds 10
```

## 🛠️ Troubleshooting

### **Errore: Authentication failed**
```powershell
# Rimuovi credenziali salvate
git config --global --unset credential.helper
git config --system --unset credential.helper

# Reinstalla credential manager
git config --global credential.helper manager-core
```

### **Errore: Push rejected**
```powershell
# Pull ultime modifiche
git pull origin main

# Risolvi conflitti se necessario
git add .
git commit -m "Merge conflicts resolved"

# Riprova push
.\auto-push.ps1
```

### **Errore: Execution policy**
```powershell
# Abilita esecuzione script
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# Oppure esegui con bypass
powershell -ExecutionPolicy Bypass -File .\auto-push.ps1
```

## 📊 Monitoraggio

### **Status Repository**
```powershell
# Verifica stato
git status

# Log recenti
git log --oneline -10

# Remote info
git remote -v
```

### **Controllo Modifiche**
```powershell
# Modifiche non committate
git diff

# File modificati
git status --short

# Storia commit
git log --graph --oneline --all
```

## 🔒 Sicurezza

### **Best Practices**
- ✅ **Token Sicuro**: Non condividere mai il Personal Access Token
- ✅ **Scadenza**: Ricorda di rinnovare i token scaduti
- ✅ **Permessi Minimi**: Usa sempre i permessi minimi necessari
- ✅ **Backup**: Mantieni sempre una copia locale del codice

### **Token Management**
```powershell
# Verifica configurazione
git config --global --list | findstr credential

# Rimuovi credenziali (se necessario)
git config --global --unset credential.helper
```

## 🎯 Workflow Consigliato

### **Sviluppo Individuale**
1. **Lavora sui file** normalmente
2. **Testa le modifiche** localmente
3. **Commit manuale**: `.\auto-push.ps1 -CommitMessage "Descrizione"`
4. **Verifica push** su GitHub

### **Sviluppo Continuo**
1. **Avvia monitoraggio**: `.\auto-commit.ps1 -Continuous`
2. **Lavora sui file** - commit automatici ogni 30 secondi
3. **Ferma monitoraggio**: `Ctrl+C`
4. **Verifica risultati** su GitHub

### **Release Management**
1. **Prepara release**: `.\auto-push.ps1 -CommitMessage "Release v1.0.2"`
2. **Crea tag**: `git tag v1.0.2`
3. **Push tag**: `git push origin v1.0.2`
4. **Crea release** su GitHub

## 📈 Statistiche

### **Comandi Utili**
```powershell
# Conta commit
git rev-list --count HEAD

# Statistiche autore
git shortlog -sn

# File più modificati
git log --pretty=format: --name-only | sort | uniq -c | sort -rn | head -10
```

## 🆘 Supporto

### **Problemi Comuni**
- **Push rifiutato**: Fai `git pull` prima
- **Conflitti**: Risolvi manualmente, poi commit
- **Token scaduto**: Genera nuovo token su GitHub
- **Credenziali perse**: Riconfigura con `git config`

### **Contatti**
- **Repository**: https://github.com/mccoy88f/eps-windows
- **Issues**: https://github.com/mccoy88f/eps-windows/issues
- **Documentazione**: Vedi `github-setup.md`

---

**🎉 Automazione Git configurata con successo!** 