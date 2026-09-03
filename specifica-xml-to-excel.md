# Specifica — Convertitore XML Fatturazione Elettronica → Excel

## 1. Panoramica

Applicazione web che consente di caricare fatture elettroniche XML (formato SDI italiano, schema FatturaPA 1.2.x), visualizzarle in un elenco gestibile, ed esportare i dati in un singolo file Excel (.xlsx) dove ogni fattura diventa una riga.

App full-stack **Next.js (TypeScript)** con storage su **filesystem** (nessun database): ogni utente ha una cartella dedicata in cui vengono salvati fisicamente i file XML caricati.

## 2. Obiettivi

- **Caricamento multiplo**: drag-and-drop o selezione file, senza limite superiore ragionevole
- **Elenco gestibile**: vedere tutti i file caricati, con info sintetiche (mittente, numero, data, totale)
- **Selezione e cancellazione**: cancellare uno, una selezione, o tutti i file
- **Esportazione Excel**: un click genera il file Excel e avvia il download
- **Ispezione XML**: visualizzare il contenuto XML di un file, indentato e navigabile

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
- **`.p7m` / `.p7s`**: fattura firmata digitalmente (busta PKCS#7 / CAdES). Il contenuto XML della fattura è incapsulato nella busta. **L'estrazione deve essere fatta per via strutturale, non con ricerca di pattern, e deve valere per ogni file importato** (vedi sotto).

**Estrazione robusta del contenuto XML da una busta PKCS#7/CAdES (obbligatoria per ogni `.p7m`/`.p7s`):**

1. **Tokenizzazione ASN.1 (BER/DER)** — il file viene decodificato come lista di TLV (Tag-Length-Value), gestendo sia la codifica **DER** (lunghezze definite) sia la **BER** (lunghezze indefinite `0x80` con terminatore end-of-contents `00 00`). La tokenizzazione è ricorsiva e scala fino agli elementi annidati. Si tollerano byte non-ASN.1 fuori contesto (es. BOM o header estranei) senza fallire.
2. **Navigazione alla `SignedData`** — si individua la struttura `ContentInfo` (SEQ OID `1.2.840.113549.1.7.2` = `signedData`) e si entra nel campo `SignedData.encapContentInfo.eContent`, l'OCTET STRING che contiene il documento firmato (la fattura XML).
3. **Ricostruzione del contenuto** — `eContent` (e, per la BER frammentata, la concatenazione dei suoi blocchi OCTET STRING interni) viene letto come flusso di byte, **scartando i byte di intestazione/framing dei TLV** e un eventuale BOM (`EF BB BF`) iniziale. Questo gestisce automaticamente i casi in cui il contenuto è spezzato in più blocchi a lunghezza indefinita (frammentazione BER) con intestazioni residue incollate nel testo.
4. **Troncamento alla chiusura del root** — il documento viene tagliato alla prima occorrenza della chiusura del root (`</FatturaElettronica>` / `</FileMetadati>`), scartando i byte di firma (SignerInfo, certificati) che lo seguono nella busta.
5. **Validazione** — il risultato deve essere XML ben formato; in caso contrario il file viene marcato `nonLeggibile` (o `corrotta` se contiene i marcatori FatturaPA).

> Perché strutturale e non a pattern: una ricerca per `<?xml`/`<FatturaElettronica` fallisce quando la corruzione è **interna** al testo (es. intestazioni BER `04 82 …` incollate nel mezzo ai punti di giunzione dei blocchi, o BOM non all'inizio). L'approccio TLV/BER ricostruisce il documento originale indifferentemente da 1 o N blocchi, lunghezze variabili e certificati diversi, ed è quindi robusto per **tutti** i file importati.

- **Archivi compressi**: all'upload l'archivio viene aperto e ogni voce con estensione `.xml`/`.p7m`/`.p7s` è estratta e trattata come un singolo file (ogni `.p7m`/`.p7s` segue l'estrazione robusta sopra); cartelle e altri formati vengono ignorati. In v1 non è prevista la ricorsione su archivi annidati.

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
- **Gestione duplicati** (due livelli, vedi sotto)
- **Caricamento additivo**: la dropzone resta sempre disponibile (compatto quando l'elenco è popolato) e ogni nuova selezione/drag **aggiunge** i file all'elenco esistente senza sovrascriverlo — i file si possono caricare in più lotti successivi
- **Pulsante "Carica file"** nella toolbar, disponibile anche quando ci sono righe in tabella: apre lo stesso selettore file della dropzone

**Controllo anti-duplicato** (al caricamento):

- **Livello 1 — stesso nome file**: se un documento con lo stesso nome è già presente nell'elenco (per gli estratti da archivio vale il nome interno del documento, non il nome della .zip), il nuovo file viene **saltato automaticamente** con un avviso. Nessuna riga duplicata.
- **Livello 2 — stessa fattura, nome file diverso**: dopo il parsing, se un documento ha la **stessa identità di fattura** di uno già caricato (chiave = ID SdI `IdTrasmittente-ProgressivoInvio`; per le notifiche `IdentificativoSdI`; fallback `fornitorePiva + numero + data`) ma con **nome file diverso**, viene mostrato un **dialogo di conferma** ("Possibile duplicato con *fileX*… Aggiungere comunque?"). Se l'utente conferma, il documento entra comunque in elenco (e quindi anche nell'export).
- **Nell'export Excel** le righe non vengono rimosse: l'eventuale duplicato presente è quello che l'utente ha **confermato esplicitamente** al caricamento.

### 5.2 Elenco file

- Tabella con colonne: Nome file, Dimensione, Tipo/Stato, Data caricamento
- Badge di classificazione per ogni file (vedi §4.4): "Fattura elettronica", "Notifica SdI", "Non è una fattura", "Fattura corrotta", "Non leggibile"
- Checkbox per selezione multipla
- Header con checkbox "Seleziona tutto"
- Pulsante "Elimina selezionati" (visibile solo con selezione attiva)
- Pulsante "Elimina tutti" (sempre visibile se ci sono file)
- Conteggio riepilogativo: totale file, numero fatture, numero notifiche, numero esclusi
- Per file con errore (corrotta/non leggibile): tooltip con messaggio di errore
- Azioni per riga: pulsante **"Vedi XML"** (lente, presente solo quando è disponibile del contenuto XML) e pulsante di **eliminazione singola** (✕)
- **Non è presente** un pulsante "Recupera": se un XML risulta corrotto in modo non riparabile automaticamente all'importazione (§4.2), l'utente lo riscarica dal portale dell'Agenzia delle Entrate/SdI e lo ricarica nel sistema

### 5.2bis Persistenza dei file sul filesystem (flusso previsto — non nel prototipo)

Nell'implementazione Next.js (non nel prototipo client-side):

- A ogni upload, il file originale viene **salvato nella cartella dedicata dell'utente** sul filesystem (`data/uploads/{userId}/`), come da architettura §1/§3. Il file XML salvato è la sorgente di verità.
- All'accesso, la pagina **ricarica l'elenco dei file già presenti** nella cartella dell'utente: uscendo e rientrando, l'utente ritrova i file caricati in precedenza.
- L'utente può poi **aggiungerne di nuovi** (d&d o pulsante "Carica file") che si accumulano all'elenco già presente (§5.1 additivo).
- Eliminazione singola/multipla/tutti: rimuove anche il file fisico dalla cartella dell'utente.
- Il prototipo attuale mantiene i file solo in memoria (nella sessione di pagina, ricaricata dall'utente); la persistenza filesystem è demandata all'app Next.js.

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

### 5.6 Gestione utenti (pagina admin separata)

- La gestione degli utenti è una **pagina separata** (`/admin/users`), non più un pannello dentro la home.
- Raggiungibile dal link **"Gestione utenti"** nella navbar, visibile **solo** agli utenti con ruolo `admin`.
- **Guard lato client**: la pagina controlla la sessione all'avvio e mostra "Accesso negato" (con link di ritorno) se l'utente non è autenticato o non ha ruolo `admin`.
- Funzionalità: creazione utente (username, password, ruolo), reset password (tramite modal, non più `prompt`), eliminazione utente (con conferma). L'eliminazione rimuove anche la cartella file dell'utente.
- Il super user `admin` è sempre presente in lista, non eliminabile e con password non modificabile dal pannello.
- Prototipo: la pagina corrisponde a un file HTML separato (`admin-users.html`), collegato dalla home tramite link mostrato solo agli admin.

### 5.7 Visualizzazione XML del file (lente)

- Ogni riga dell'elenco può avere un pulsante con **lente** ("Vedi XML"), mostrato solo quando dal file è estraibile del contenuto XML (fatture, notifiche SdI e file corrotti con marcatori FatturaPA).
- Al click si apre una **modale ampia** che mostra il contenuto XML del file:
  - titolo = nome file;
  - XML **indentato** (pretty-print) in un blocco monospace **scrollabile** (verticale e orizzontale);
  - **evidenziati i nomi dei tag** (syntax highlight minimale) su sfondo scuro stile editor;
  - chiusura tramite pulsante "Chiudi" o click sullo sfondo.
- Il testo XML viene conservato in memoria durante la sessione ma **non è persistito** su filesystem/localStorage (evita di appesantire lo storage con file grandi); in implementazione Next.js il contenuto verrà riletto dal file su disco al momento dell'apertura.

### 5.8 Funzionalità del prototipo da implementare in produzione

Questa sezione elenca le **funzionalità** dimostrate dal prototipo che devono essere implementate lato server (Next.js). È il riferimento puntuale per replicare le funzioni, non i dettagli di demo del prototipo (toast, localStorage, login finto).

| # | Funzionalità | Specifiche di implementazione in produzione (Next.js) |
|---|---|---|
| R1 | **Caricamento additivo**: aggiungere file all'elenco esistente senza sovrascrivere. | `POST /api/files` salva il file in `data/uploads/{userId}/` **senza** sovrascrivere quelli esistenti; l'elenco (`GET /api/files`) riflette sempre tutti i file della cartella. |
| R2 | **Doppia via di caricamento**: drag&drop e pulsante "Carica file", entrambi disponibili anche con l'elenco già popolato. | Entrambe le vie pubblicano a `POST /api/files`; la dropzone resta visibile (in forma compatta) anche quando la tabella contiene righe. |
| R3 | **Gestione duplicati — stesso nome file**: non caricare due file con lo stesso nome di documento. | Il server rifiuta un file il cui **nome** esiste già nella cartella utente (per gli estratti da archivio vale il nome interno del documento, non quello della .zip): nessuna riga duplicata. |
| R4 | **Gestione duplicati — stessa fattura, nome file diverso**: riconoscere che due file con nome diverso sono la stessa fattura. | Il server confronta l'**identità fattura** (chiave = ID SdI `IdTrasmittente-ProgressivoInvio`; notifiche `IdentificativoSdI`; fallback `fornitorePiva + numero + data`) con i file già presenti. Se coincide ma il nome è diverso, l'API segnala un "probabile duplicato" e il file non viene salvato finché l'utente non conferma di volerlo aggiungere comunque. |
| R5 | **Riprendere un duplicato** (se l'utente conferma): il documento aggiunto resta nell'elenco e quindi anche nell'export Excel. | Nessuna rimozione forzata in export: un duplicato presente è quello **confermato** dall'utente al caricamento. |
| R6 | **Chiave di identità fattura** (`invoiceKey`). | Implementare una funzione server (es. `utils/invoiceKey.ts`) che calcola la medesima chiave usata nel controllo duplicati (R4). |
| R7 | **Export Excel**: solo fatture elettroniche valide, 24 colonne. | `GET /api/export` genera l'.xls (SheetJS) dalle fatture valide della cartella utente, allineate al modello ReportFattureRicevute (vedi §4.5/§4.6). |

> Nota: il prototipo è **client-only**; in produzione tutta la logica di business (salvataggio file, estrazione BER-DER, parsing, classificazione, gestione duplicati, export) è eseguita **lato server**. Queste righe definiscono le funzionalità da implementare; i dettagli di presentazione del prototipo (toast, modal di conferma, localStorage) non sono vincolanti e in producción vengono realizzati secondo il design UI dell'app Next.js.

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
- **Pagina admin utenti**: stessa struttura (navbar scura + contenuto su sfondo `#f5f6f8`), card unica con form di creazione e tabella utenti; schermata "Accesso negato" centrata quando l'utente non è admin
- **Mobile**: layout verticale stack, dropzone in alto, elenco sotto

## 7. Stack tecnologico

- **Framework**: Next.js (App Router) + TypeScript — un'unica app full-stack: interfaccia e API (route handler / server actions) nello stesso progetto
- **Storage**: **filesystem** (nessun database). I file XML vengono salvati su disco in cartelle dedicate.
- **Parsing XML**: `fast-xml-parser` lato server, con supporto `.p7m`/`.p7s` (estrazione strutturale del contenuto XML dalla busta PKCS#7/CAdES via ASN.1 BER/DER, vedi §4.2)
- **Decodifica ASN.1/PKCS#7**: libreria Node dedicata (es. `@peculiar/asn1-schema` + `@peculiar/asn1-cms`, oppure `asn1js`) per tokenizzare la lingua BER/DER in TLV e navigare la `SignedData` fino a `encapContentInfo.eContent` (vedi §4.2); alternativa leggera: decoder ASN.1 minimal custom in ~100 righe se si preferisce zero dipendenze
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

**Risposta di `POST /api/files` in caso di duplicati** (vedi §5.8 R3/R4):

- **Stesso nome file**: il server non salva il file e risponde con un esito per-row di tipo `duplicateName` (il client mostra "salta: file con lo stesso nome già caricato").
- **Stessa fattura, nome diverso**: il server rileva la potenziale identità ma **non** salva finché il client non conferma; l'API espone un esito per-row `possibleDuplicate` con il nome del file già presente, e il client mostra il dialogo "Possibile duplicato". Solo a conferma avvenuta l'upload del file viene effettivamente salvato (seconda chiamata esplicita di conferma o parametro `force`).
- La risposta complessiva include i conteggi `{ ok, esclusi, duplicati }` per il toast di riepilogo (§5.8 R5).

### Autenticazione e utenti (senza database)

- **Store utenti**: file JSON su filesystem `data/users.json` — una voce per utente `{ username, passwordHash, role, createdAt }`. Nessuna tabella database.
- **Password**: hash con `bcrypt` (libreria `bcryptjs`, puro JavaScript, nessuna dipendenza nativa); le password in chiaro non vengono mai salvate.
- **Sessione**: cookie httpOnly firmato. `jose` (JWT) per firmare un token `{ username, role }`; in alternativa `iron-session` (cookie cifrato lato server).
- **Super user**: voce fissa non eliminabile e con password non modificabile dal pannello; può creare/eliminare gli altri utenti.
- **Ruoli**: `admin` e `user`. Il super user `admin` è una voce fissa con ruolo `admin`; anche altri utenti possono essere creati con ruolo `admin`.
- **Gestione utenti su pagina separata**: l'interfaccia di gestione utenti è su una pagina dedicata (`/admin/users`), protetta dal guard `requireAdmin`; non è più un pannello nella home (vedi §5.6).
- **Guard**: funzioni `requireAuth` e `requireAdmin` applicate a tutte le route protette (incluse le pagine admin).
- **Cartella per utente**: a ogni utente corrisponde `data/uploads/{username}/` (vedi §8).

## 9. Non incluso (v1)

- Database relazionale (uso esclusivo del filesystem, inclusi gli utenti in `data/users.json`)
- Archivi `.rar` e `.7z` (richiedono componenti extra: WASM o binario 7z)
- Personalizzazione colonne Excel
- Ricorsione su archivi compressi annidati
- Verifica crittografica della firma P7M (si estrae solo il contenuto, la firma non viene validata)
- Ri-firma dell'XML: l'estrazione (§4.2) produce un XML non firmato; il programma non rigenera una busta PKCS#7 valida (modificando l'XML la firma originale è comunque invalidata)

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
