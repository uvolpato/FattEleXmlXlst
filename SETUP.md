# Setup — Convertitore XML → Excel (Fatturazione Elettronica)

Prerequisiti e istruzioni per installare e avviare l'applicazione.

## Prerequisiti

| Requisito | Versione | Note |
|---|---|---|
| Node.js | ≥ 20 (LTS) | runtime JavaScript |
| npm | ≥ 10 | incluso con Node.js |
| Git | qualunque recente | per clonare il repository |

- **Nessun database richiesto**: lo storage è su filesystem (cartelle su disco), nessun servizio esterno da installare.
- Sistema operativo: Windows, macOS o Linux.

## Installazione

### 1. Clona il repository

```bash
git clone <repo-url>
cd <cartella-progetto>
```

### 2. Installa le dipendenze

```bash
npm install
```

### 3. Configura l'ambiente

Copia il file di esempio e valorizza le variabili:

```bash
cp .env.example .env
```

Variabili principali:

- `SESSION_SECRET` — chiave per firmare i cookie di sessione (JWT)
- `DATA_DIR` — cartella dati (default `./data`)

### 4. Avvia in sviluppo

```bash
npm run dev
```

L'app è disponibile su `http://localhost:3000`.

### 5. Build e avvio in produzione

```bash
npm run build
npm start
```

## Struttura dati (filesystem)

```
data/
├── users.json          # utenti (username, passwordHash, role)
└── uploads/
    └── {username}/     # file XML caricati per utente
```

Il file `data/users.json` e la cartella `data/uploads/` vengono creati automaticamente al primo avvio.

## Primo accesso

- Il **super user `admin`** viene creato automaticamente al primo avvio (seed).
- Cambiare la password subito dopo il primo login.
- Il super user è non eliminabile e può creare/eliminare gli altri utenti.

## Stack tecnologico

- **Next.js** (App Router) + TypeScript
- **Bootstrap 5** + Bootstrap Icons (UI)
- **fast-xml-parser** — parsing FatturaPA
- **exceljs** — generazione Excel
- **fflate** + **tar** — estrazione archivi (.zip/.tar/.tar.gz/.tgz/.gz)
- **busboy** — upload multipart
- **bcryptjs** (hash password) + **jose** (sessioni JWT)

## Risoluzione problemi

- Porta 3000 occupata: `npm run dev -- -p 3001`
- Errori di permessi su `data/`: assicurarsi che la cartella sia scrivibile dall'utente che esegue Node.
