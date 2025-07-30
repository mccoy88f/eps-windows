# 🔐 Configurazione GitHub per Push Automatico

## 📋 Prerequisiti

1. **Account GitHub**: Deve essere configurato con accesso al repository `mccoy88f/eps-windows`
2. **Personal Access Token**: Necessario per l'autenticazione

## 🚀 Setup Automatico

### 1. **Genera Personal Access Token**

1. Vai su [GitHub Settings > Developer settings > Personal access tokens](https://github.com/settings/tokens)
2. Clicca "Generate new token (classic)"
3. Seleziona i seguenti permessi:
   - `repo` (tutti i permessi del repository)
   - `workflow` (se vuoi usare GitHub Actions)
4. Copia il token generato

### 2. **Configura Credenziali Locali**

```powershell
# Configura il credential manager per GitHub
git config --global credential.helper manager-core
```

### 3. **Testa l'Autenticazione**

```powershell
# Esegui il primo push per configurare le credenziali
git push origin main
```

Quando richiesto:
- **Username**: `mccoy88f`
- **Password**: Usa il Personal Access Token (NON la password GitHub)

### 4. **Usa lo Script di Push Automatico**

```powershell
# Push normale
.\auto-push.ps1

# Push con messaggio personalizzato
.\auto-push.ps1 -CommitMessage "Fix: UI improvements"

# Push forzato (solo se necessario)
.\auto-push.ps1 -Force
```

## 🔧 Configurazione Alternativa

### **Usando GitHub CLI**

```powershell
# Installa GitHub CLI
winget install GitHub.cli

# Login
gh auth login

# Push automatico
gh repo sync
```

### **Usando SSH Keys**

```powershell
# Genera chiave SSH
ssh-keygen -t ed25519 -C "antonellomigliorelli@gmail.com"

# Aggiungi la chiave a GitHub
# Copia il contenuto di ~/.ssh/id_ed25519.pub su GitHub Settings > SSH and GPG keys

# Cambia remote URL
git remote set-url origin git@github.com:mccoy88f/eps-windows.git
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

### **Errore: Permission denied**

1. Verifica che il token abbia i permessi corretti
2. Controlla che l'account abbia accesso al repository
3. Verifica che il repository esista: https://github.com/mccoy88f/eps-windows

### **Errore: Repository not found**

```powershell
# Verifica remote URL
git remote -v

# Se necessario, aggiorna l'URL
git remote set-url origin https://github.com/mccoy88f/eps-windows.git
```

## 📝 Note Importanti

- **Token Sicurezza**: Non condividere mai il Personal Access Token
- **Scadenza Token**: I token hanno una scadenza, ricorda di rinnovarli
- **Permessi Minimi**: Usa sempre i permessi minimi necessari
- **Backup**: Mantieni sempre una copia locale del codice

## 🎯 Comandi Utili

```powershell
# Status repository
git status

# Log commit recenti
git log --oneline -10

# Verifica remote
git remote -v

# Pull ultime modifiche
git pull origin main

# Push manuale
git push origin main
``` 