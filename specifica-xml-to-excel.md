# Specifica — Convertitore XML Fatturazione Elettronica → Excel

## 1. Panoramica

Applicazione web che consente di caricare fatture elettroniche XML (formato SDI italiano, schema FatturaPA 1.2.x), visualizzarle in un elenco gestibile, ed esportare i dati in un singolo file Excel (.xlsx) dove ogni fattura diventa una riga.

App full-stack **Next.js (TypeScript)** con storage su **filesystem** (nessun database): ogni utente ha una cartella dedicata in cui vengono salvati fisicamente i file XML caricati.

## 2. Obiettivi

- **Caricamento multiplo**: drag-and-drop o selezione file, senza limite superiore ragionevole
- **Elenco gestibile**: vedere tutti i file caricati, con info sintetiche (mittente, numero, data, totale)
- **Selezione e cancellazione**: cancellare uno, una selezione, o tutti i file
- **Esportazione Excel**: un click genera il file Excel e avvia il download

## 3. Utente target

- Commercialisti, consulenti fiscali, uffici amministrativi
- Gestiscono volumi di fatture elettroniche e devono analizzarli in Excel
- Livello tecnico: medio-basso (devono poter usare l'app senza istruzioni)

## 4. Architettura

### 4.1 Modello dati (fattura)

```
Fattura (parsed da XML)
├── id: string (UUID generato lato client)
├── fileName: string
├── fileSize: number
├── parsed: boolean
├── data: {
│   ├── tipoDocumento: string (TD01-TD28)
│   ├── numero: string
│   ├── data: string (YYYY-MM-DD)
│   ├── divisa: string (EUR)
│   ├── cedentePrestatore: {
│   │   ├── denominazione: string
│   │   ├── partitaIva: string
│   │   ├── indirizzo: string
│   │   ├── cap: string
│   │   ├── comune: string
│   │   └── provincia: string
│   │ }
│   ├── cessionarioCommittente: {
│   │   ├── denominazione: string
│   │   ├── partitaIva: string
│   │   ├── indirizzo: string
│   │   ├── cap: string
│   │   ├── comune: string
│   │   └── provincia: string
│   │ }
│   ├── beniServizi: [{
│   │   ├── lineaNumero: number
│   │   ├── descrizione: string
│   │   ├── quantita: number
│   │   ├── unitaMisura: string
│   │   ├── prezzoUnitario: number
│   │   ├── prezzoTotale: number
│   │   ├── aliquotaIVA: number
│   │   └── natura: string
│   │ }]
│   ├── datiGenerali: {
│   │   ├── importoTotaleDocumento: number
│   │   ├── importoPagamento: number
│   │   ├── modalitaPagamento: string
│   │   └── dataScadenzaPagamento: string
│   │ }
│   └── datiRitenuta: {
│       ├── tipoRitenuta: string
│       ├── importoRitenuta: number
│       └── aliquotaRitenuta: number
│     }
│ }
└── errori: string[]
```

### 4.2 Formati accettati ed estrazione contenuto

Formati file accettati: `.xml`, `.p7m`, `.p7s` e archivi compressi `.zip`, `.tar`, `.tar.gz`, `.tgz`, `.gz` (estensioni case-insensitive).

- **`.xml`**: letto direttamente come testo.
- **`.p7m` / `.p7s`**: fattura firmata digitalmente (busta PKCS#7 / CAdES). Il contenuto XML della fattura è incapsulato nella busta; il programma individua l'inizio del documento (`<?xml`, `<FatturaElettronica` o `<FileMetadati`) e taglia i byte della firma presenti dopo la chiusura del root (`</FatturaElettronica>` / `</FileMetadati>`), così il parser non si blocca sui dati binari finali. Se il contenuto XML non è estraibile in chiaro (busta interamente DER), il file viene marcato come "non leggibile".
- **Archivi compressi**: all'upload l'archivio viene aperto e ogni voce con estensione `.xml`/`.p7m`/`.p7s` è estratta e trattata come un singolo file; cartelle e altri formati vengono ignorati. In v1 non è prevista la ricorsione su archivi annidati.

### 4.3 Parsing XML

- Parsing XML lato server con `fast-xml-parser`
- Namespace FatturaPA 1.2: `http://ivaservizi.agenziaentrate.gov.it/docs/xsd/fatture/v1.2`
- Namespace FileMetadati SdI: `http://ivaservizi.agenziaentrate.gov.it/docs/xsd/fattura/messaggi/v1.0`
- Il parsing usa i localname degli elementi, indipendenti da namespace/prefix, così gestisce sia le fatture FatturaPA sia le notifiche SdI (FileMetadati)
- Gestione errori: file non valido, XML malformato, namespace sconosciuto
- Il parsing è interamente lato server (nessun database)

### 4.4 Classificazione dei file (per tipo)

Ogni file caricato viene classificato immediatamente in uno dei seguenti tipi, mostrato nell'elenco come badge di stato:

| Tipo | Condizione | In export? |
|------|-----------|-----------|
| **Fattura elettronica** (`fattura`) | XML valido con root `FatturaElettronica` + `FatturaElettronicaBody` | ✅ |
| **Notifica SdI** (`metadati`) | XML valido con root `FileMetadati` (notifica di trasmissione/ricezione, non è una fattura) | ❌ |
| **Non è una fattura** (`nonFattura`) | XML valido ma di altro tipo (nessun root FatturaPA né FileMetadati) | ❌ |
| **Fattura corrotta** (`corrotta`) | Il contenuto contiene i marcatori FatturaPA (`FatturaElettronica` / namespace `ivaservizi.agenziaentrate.gov.it`) ma il parse fallisce (XML malformato, root senza `FatturaElettronicaBody`, file troncato) | ❌ |
| **Non leggibile** (`nonLeggibile`) | File `.p7m`/`.p7s` senza contenuto XML estraibile in chiaro | ❌ |

Criterio di rilevamento "corrotta vs non-fattura": se il testo contiene i marcatori FatturaPA ma il parsing genera errore, il file è considerato una **fattura corrotta**; altrimenti è un file **non di fatturazione** (o non leggibile).

### 4.5 Struttura Excel di output

Ogni riga = 1 fattura. L'export allinea il modello utente `ReportFattureRicevute` con **24 colonne fisse** (intestazioni e ordine riportati sotto):

| # | Colonna | Path XML | Note |
|---|---------|----------|------|
| 1 | Numero | `FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/Numero` | |
| 2 | Nome file | — | Nome del file XML originale |
| 3 | ID SdI | `FatturaElettronicaHeader/DatiTrasmissione/IdTrasmittente/IdCodice` + `ProgressivoInvio` | Concatenati; da `FileMetadati/IdentificativoSdI` per notifiche |
| 4 | Data ricezione | `FileMetadati` (metadato SdI) | Vuota per fattura singola |
| 5 | Data documento | `.../DatiGeneraliDocumento/Data` | |
| 6 | Tipo documento | `.../TipoDocumento` | Formato "Fattura - TD01" (vedi §4.6) |
| 7 | Fornitore | `CedentePrestatore/DatiAnagrafici/Denominazione` (o Nome+Cognome) | |
| 8 | P.IVA | `CedentePrestatore/DatiAnagrafici/IdFiscaleIVA/IdCodice` | |
| 9 | Codice Fiscale | `CedentePrestatore/DatiAnagrafici/CodiceFiscale` | |
| 10 | Metodo di pagamento | `DatiPagamento/DettaglioPagamento/ModalitaPagamento` | Formato "MP05 - Bonifico" (vedi §4.7) |
| 11 | Totale imponibile | `DatiBeniServizi/DatiRiepilogo/ImponibileImporto` (somma) | Formato € |
| 12 | N1 - Escluso | `DatiRiepilogo` con `Natura=N1` | |
| 13 | N2 - Non soggetto | `DatiRiepilogo` con `Natura=N2/N2.1/N2.2` | |
| 14 | N3 - Non imponibile | `DatiRiepilogo` con `Natura=N3…N3.6` | |
| 15 | N4 - Esente | `DatiRiepilogo` con `Natura=N4` | |
| 16 | N5 - Margine RP | `DatiRiepilogo` con `Natura=N5/N5.1/N5.2` | |
| 17 | N6 - Inversione contabile | `DatiRiepilogo` con `Natura=N6…N6.5` | |
| 18 | N7 - IVA assolta UE | `DatiRiepilogo` con `Natura=N7` | |
| 19 | Totale IVA | `DatiRiepilogo/Imposta` (somma) | Formato € |
| 20 | Totale documento | `.../DatiGeneraliDocumento/ImportoTotaleDocumento` | Formato € |
| 21 | Netto a pagare | `DatiPagamento/DettaglioPagamento/ImportoPagamento` | Formato € |
| 22 | Pagamenti | — | Fisso "Non pagata" (come modello) |
| 23 | Data pagamento | — | Vuota (da valorizzare a pagamento avvenuto) |
| 24 | Stato | — | Fisso "Non letta" (come modello) |

Nota: i file `FileMetadati` SdI non sono fatture e non producono righe; servono solo per recuperare ID SdI / data ricezione associati alla fattura.

### 4.6 Mapping TipoDocumento

| Codice | Descrizione |
|--------|-------------|
| TD01 | Fattura |
| TD02 | Acconto/anticipo su fattura |
| TD03 | Acconto/anticipo su parcella |
| TD04 | Nota di credito |
| TD05 | Nota di debito |
| TD06 | Parcella |
| TD07 | Fattura RT |
| TD08 | Fattura per acquisto intracomunitario |
| TD09 | Autofattura |
| TD10 | Fattura di sola indicazione IVA |
| TD11 | Fattura accompagnatoria |
| TD12 | Acconto/anticipo su fattura accompagnatoria |
| TD13 | Acconto/anticipo su parcella accompagnatoria |
| TD14 | Fattura RT accompagnatoria |
| TD15 | Fattura per acquisto intracomunitario accompagnatoria |
| TD16 | Integrazione fattura reverse charge interno |
| TD17 | Integrazione/autofattura per acquisto servizi estero |
| TD18 | Integrazione per acquisto beni ex art.17 comma 2 DPR 633/72 |
| TD19 | Integrazione/autofattura per acquisto beni ex art.17 comma 6-bis DPR 633/72 |
| TD20 | Autofattura per regolarizzazione e integrazione delle fatture |
| TD21 | Autofattura per splafonamento |
| TD22 | Estrazione beni da Deposito IVA |
| TD23 | Estrazione beni da Deposito IVA con versamento IVA |
| TD24 | Fattura differita art.21, comma 4, lett. a) |
| TD25 | Fattura differita art.21, comma 4, terzo periodo lett. b) |
| TD26 | Cessione di beni ammortizzabili e per passaggi interni |
| TD27 | Fattura per autoconsumo o cessioni gratuite senza rivalsa |
| TD28 | Acquisti da San Marino con IVA |
| TD29 | Acquisti da San Marino senza IVA |

### 4.7 Mapping ModalitaPagamento

| Codice | Descrizione |
|--------|-------------|
| MP01 | Contanti |
| MP02 | Assegno |
| MP03 | Assegno circolare |
| MP04 | Contanti presso Tesoreria |
| MP05 | Bonifico |
| MP06 | Vaglia cambiario |
| MP07 | Bollettino bancario |
| MP08 | Carta di pagamento |
| MP09 | RID |
| MP10 | RID utenze |
| MP11 | RID veloce |
| MP12 | Riba |
| MP13 | MAV |
| MP14 | Quietanza erario |
| MP15 | Giroconto su conti di contabilità speciale |
| MP16 | Bonifico bancario postale |
| MP17 | Contanti con RIA (ricevuta fiscale) |
| MP18 | Contanti con RIA (scontrino) |
| MP19 | SEPA Direct Debit |
| MP20 | SEPA Direct Debit CORE |
| MP21 | SEPA Direct Debit B2B |
| MP22 | Accredito su conto corrente postale |
| MP23 | Accredito su conto di tesoreria |
| MP24 | Addebito diretto |
| MP25 | Contanti |
| MP26 | Assegno bancario e circolare |
| MP27 | Altri mezzi di pagamento |
| MP28 | Bollo |

## 5. Funzionalità UI

### 5.1 Area di caricamento (Dropzone)

- Zona drag-and-drop centrale con icona e testo invitante
- Click per aprire il selettore file
- Accetta `.xml`, `.p7m`, `.p7s` e archivi `.zip`, `.tar`, `.tar.gz`, `.tgz`, `.gz` (fatture firmate digitalmente incluse)
- Feedback visivo durante il drag (bordi evidenziati)
- Conteggio file caricati in tempo reale
- Gestione duplicati: avviso se un file con lo stesso nome è già caricato

### 5.2 Elenco file

- Tabella con colonne: Nome file, Dimensione, Tipo/Stato, Data caricamento
- Badge di classificazione per ogni file (vedi §4.4): "Fattura elettronica", "Notifica SdI", "Non è una fattura", "Fattura corrotta", "Non leggibile"
- Checkbox per selezione multipla
- Header con checkbox "Seleziona tutto"
- Pulsante "Elimina selezionati" (visibile solo con selezione attiva)
- Pulsante "Elimina tutti" (sempre visibile se ci sono file)
- Conteggio riepilogativo: totale file, numero fatture, numero notifiche, numero esclusi
- Per file con errore (corrotta/non leggibile): tooltip con messaggio di errore

### 5.3 Filtri e ricerca

- **Filtro per stato**: menu a tendina (`<select>`) per mostrare solo i file di un determinato tipo — "Tutti", "Fatture", "Notifiche SdI", "Non fatture", "Corrotte", "Non leggibili" (mappato sui tipi di §4.4)
- **Ricerca per nome**: campo di testo (`<input type="search">`) con filtro incrementale sul nome file (case-insensitive)
- Filtro e ricerca sono combinabili e si applicano lato client sull'elenco già caricato
- Il conteggio riepilogativo e i pulsanti di selezione operano sull'elenco filtrato

### 5.4 Barra azioni

- Pulsante primario "Esporta Excel" (disabilitato se 0 fatture elettroniche valide)
- Indicatore di stato: "X fatture, Y notifiche, Z esclusi"
- L'export considera **solo** le fatture elettroniche valide (esclude notifiche SdI, non-fatture, fatture corrotte e file non leggibili)

### 5.5 Feedback e stati

- Spinner durante il parsing XML
- Toast di successo dopo esportazione
- Toast di errore per file non validi
- Progress bar durante il caricamento multiplo

## 6. Design directions

### 6.1 UI framework

- **Bootstrap 5.3** + **Bootstrap Icons** — framework CSS/iconografico, compatibile con Next.js (App Router) via CDN o `react-bootstrap`.
- Nessun colore brand custom: si usa la palette Bootstrap di default.

### 6.2 Palette

- Sfondo pagina: `#f5f6f8` (grigio-azzurro chiaro)
- Superfici / card: `#ffffff` con `shadow-sm`
- Navbar: scura `#212529` con testo `rgba(255,255,255,.5)`
- Testo primario: `#212529`
- Testo secondario: `#6c757d`
- Azione primaria: `#0d6efd` (blu Bootstrap)
- Bordi: `#dee2e6`
- Status: `#198754` (success), `#ffc107` (warning), `#dc3545` (danger)

### 6.3 Tipografia

- Font di sistema (stack Bootstrap): `system-ui, -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif`
- Mono per numeri/codici: `ui-monospace, "Cascadia Code", Menlo, monospace`

### 6.4 Layout

- **Navbar scura fissa** in alto: brand a sinistra, menu a tendina e utente a destra
- **Contenuto** su sfondo `#f5f6f8` con card a superficie bianca
- **Tabelle**: header sticky (`position:sticky; top:0; background:#e9ecef`)
- **Login**: card centrata (max-width ~360px), password con toggle occhio, checkbox "Resta collegato", pulsante primario a tutta larghezza
- **Mobile**: layout verticale stack, dropzone in alto, elenco sotto

## 7. Stack tecnologico

- **Framework**: Next.js (App Router) + TypeScript — un'unica app full-stack: interfaccia e API (route handler / server actions) nello stesso progetto
- **Storage**: **filesystem** (nessun database). I file XML vengono salvati su disco in cartelle dedicate.
- **Parsing XML**: `fast-xml-parser` lato server, con supporto `.p7m`/`.p7s` (estrazione del contenuto XML dalla busta PKCS#7/CAdES, vedi §4.2)
- **Generazione Excel**: `exceljs` (in alternativa `xlsx`/SheetJS) — output a 24 colonne (vedi §4.5)
- **Upload**: route handler Next.js con `busboy` (parsing `multipart/form-data`)
- **Estrazione archivi**: `fflate` (`.zip`) e `tar` + `zlib` (`.tar`, `.tar.gz`, `.tgz`, `.gz`) — solo dipendenze pure JS o built-in Node
- **Runtime**: Node.js

## 8. Architettura filesystem

```
data/
└── uploads/
    └── {userId}/
        ├── 2026-09-01_fattura-28.xml
        ├── 2026-09-02_fattura-31.xml.p7m
        └── ...
```

- Ogni utente ha una cartella dedicata in `data/uploads/{userId}/`.
- I file caricati vengono scritti fisicamente su disco; nessun contenuto XML finisce in un database.
- Il file XML è la sorgente di verità: la classificazione (fattura / notifica SdI / non-fattura / corrotta, vedi §4.4) è ricalcolata a ogni lettura.
- L'elenco dei file è ottenuto scandendo la cartella dell'utente (nessun indice centralizzato).

### Route / API

| Metodo | Route | Descrizione |
|---|---|---|
| POST | `/api/files` | Upload multiplo (.xml/.p7m/.p7s e archivi compressi) → estrae, salva su filesystem + classifica |
| GET | `/api/files` | Elenco file dell'utente (nome, dimensione, data, tipo) |
| DELETE | `/api/files/:id` | Elimina un singolo file |
| DELETE | `/api/files` | Elimina una selezione o tutti (body con elenco id) |
| POST | `/api/export` | Genera il file .xlsx con solo le fatture elettroniche e avvia il download |
| POST | `/api/auth/login` | Verifica credenziali e imposta il cookie di sessione |
| POST | `/api/auth/logout` | Invalida la sessione |
| POST | `/api/auth/change-password` | Cambio password (verifica della vecchia password) |
| GET/POST/DELETE | `/api/users` | Gestione utenti (solo admin): crea / elimina / reset password |

### Autenticazione e utenti (senza database)

- **Store utenti**: file JSON su filesystem `data/users.json` — una voce per utente `{ username, passwordHash, role, createdAt }`. Nessuna tabella database.
- **Password**: hash con `bcrypt` (libreria `bcryptjs`, puro JavaScript, nessuna dipendenza nativa); le password in chiaro non vengono mai salvate.
- **Sessione**: cookie httpOnly firmato. `jose` (JWT) per firmare un token `{ username, role }`; in alternativa `iron-session` (cookie cifrato lato server).
- **Super user**: voce fissa non eliminabile e con password non modificabile dal pannello; può creare/eliminare gli altri utenti.
- **Ruoli**: `admin` (super user) e `user`.
- **Guard**: funzioni `requireAuth` e `requireAdmin` applicate a tutte le route protette.
- **Cartella per utente**: a ogni utente corrisponde `data/uploads/{username}/` (vedi §8).

## 9. Non incluso (v1)

- Database relazionale (uso esclusivo del filesystem, inclusi gli utenti in `data/users.json`)
- Archivi `.rar` e `.7z` (richiedono componenti extra: WASM o binario 7z)
- Personalizzazione colonne Excel
- Ricorsione su archivi compressi annidati
- Verifica crittografica della firma P7M (si estrae solo il contenuto, la firma non viene validata)

## 10. Modello di marchio

Nessun brand specifico. Direzione visiva: **Tech / Utility** — grigio neutro con blu professionale come accento. Font system per leggibilità. Interfaccia pulita, densa, funzionale.

## 11. Distribuzione e installazione

### 11.1 Distribuzione

- L'app viene distribuita tramite **GitHub** (repository versionato con Git).
- Ramo `main` come riferimento; rilasci tramite tag semantici (`v1.0.0`, …).
- Il repository contiene il codice Next.js e il file di setup **`SETUP.md`** con prerequisiti e istruzioni di installazione.

### 11.2 Prerequisiti

- **Node.js** ≥ 20 (LTS)
- **npm** ≥ 10 (incluso con Node.js)
- **Git** (per clonare il repository)
- **Nessun database**: lo storage è su filesystem, nessun servizio esterno da installare

### 11.3 Installazione

1. `git clone <repo-url>`
2. `npm install`
3. Configurare le variabili d'ambiente copiando `.env.example` in `.env` (es. `SESSION_SECRET`)
4. `npm run dev` (sviluppo) oppure `npm run build && npm start` (produzione)

### 11.4 Struttura dati (creata al primo avvio)

- `data/users.json` — utenti (super user `admin` seed)
- `data/uploads/{username}/` — file XML caricati per utente

### 11.5 Primo accesso

- Super user `admin` creato automaticamente al primo avvio; password da cambiare subito dopo il login.
