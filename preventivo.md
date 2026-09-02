# Preventivo — Convertitore XML Fatturazione Elettronica → Excel

**Redatto da:** Ugo Volpato — AI Consultant
**Data:** 03/09/2026
**Cliente:** ________________ (da indicare)
**Modalità di sviluppo:** assistita da AI (analisi, generazione e codifica tramite modelli AI)

---

## 1. Oggetto

Web app full-stack (**Next.js + TypeScript**, storage su **filesystem senza database**) per:

- caricare fatture elettroniche XML (FatturaPA 1.2) e file firmati `.p7m`/`.p7s`;
- riconoscere automaticamente fatture / notifiche SdI / file non-fattura / fatture corrotte;
- esportare un file **Excel a 24 colonne** (modello `ReportFattureRicevute`), una riga per fattura;
- gestire **login multi-utente** con super user non eliminabile e cartella dedicata per utente;
- importare archivi compressi (`.zip`, `.tar`, `.tar.gz`, `.tgz`, `.gz`);
- filtrare per stato e cercare per nome file.

Riferimento tecnico: file `specifica-xml-to-excel.md` (§1–§11) e `SETUP.md`.

## 2. Tariffe applicate (indicative, mercato Italia)

| Voce | Tariffa oraria |
|---|---|
| Analisi, consulenza e prototipazione | € 80,00 |
| Sviluppo assistito da AI | € 55,00 |
| Sviluppo tradizionale (riferimento) | € 70,00 |

## 3. Attività già svolte (questa sessione)

| # | Attività | Ore | Tariffa | Importo |
|---|---|---|---|---|
| 1 | Analisi file XML (FatturaPA 1.2, FileMetadati SdI, buste p7m) ed estrazione del modello Excel a 24 colonne | 2 | € 80,00 | € 160,00 |
| 2 | Redazione della specifica tecnica (`specifica-xml-to-excel.md`) | 2 | € 80,00 | € 160,00 |
| 3 | Prototipo HTML interattivo (upload, elenco, classificazione, export, archivi, filtro/ricerca) | 3 | € 80,00 | € 240,00 |
| 4 | Login client-side mock + gestione utenti (super user, ruoli, cartella per utente) | 1 | € 80,00 | € 80,00 |
| 5 | Restyling Bootstrap 5 + documentazione setup/deploy | 1 | € 80,00 | € 80,00 |
| 6 | Debug e rifiniture | 1 | € 80,00 | € 80,00 |
| **Subtotale già svolto** | | **10** | | **€ 800,00** |

## 4. Attività da svolgere (implementazione Next.js)

La colonna "Ore (AI)" è la stima con sviluppo assistito; "Ore (senza AI)" è il tempo che richiederebbe uno sviluppo tradizionale (riferimento).

| # | Attività | Ore (AI) | Ore (senza AI) | Importo (AI) |
|---|---|---|---|---|
| 1 | Setup progetto Next.js + repository GitHub + CI/CD | 2 | 4 | € 110,00 |
| 2 | Autenticazione server (`users.json`, `bcryptjs`, JWT `jose`, guard `requireAuth`/`requireAdmin`) | 3 | 8 | € 165,00 |
| 3 | API upload multipart (`busboy`) + estrazione archivi (`fflate` + `tar`) | 4 | 10 | € 220,00 |
| 4 | Parsing FatturaPA server-side (`fast-xml-parser`) + classificazione file | 3 | 10 | € 165,00 |
| 5 | Export Excel 24 colonne (`exceljs`) | 2 | 6 | € 110,00 |
| 6 | Frontend UI (Bootstrap 5: pagina principale + login) | 6 | 16 | € 330,00 |
| 7 | Filtro/ricerca + gestione elenco e cancellazione | 2 | 5 | € 110,00 |
| 8 | Test e rifinitura finale | 3 | 11 | € 165,00 |
| **Subtotale da svolgere** | | **25** | **70** | **€ 1.375,00** |

## 5. Riepilogo economico

| Voce | Ore | Importo |
|---|---|---|
| Attività già svolte | 10 | € 800,00 |
| Attività da svolgere | 25 | € 1.375,00 |
| **Totale complessivo** | **35** | **€ 2.175,00** |

**Riferimento senza AI** — lo sviluppo tradizionale richiederebbe **70 ore** (≈ 2,8× le ore assistite): 70 h × € 70,00 = **€ 4.900,00**, per un totale di **€ 5.700,00** (€ 800,00 già svolto + € 4.900,00 sviluppo).

L'approccio assistito da AI consente un **risparmio di € 3.525,00** sul solo sviluppo (€ 4.900,00 → € 1.375,00, **−72%**).

## 6. Note

- Le ore della sezione "già svolte" sono una **stima** del tempo impiegato in questa sessione (analisi, specifica, prototipo e iterazioni); da confermare.
- Il costo ridotto dello sviluppo riflette la **velocizzazione dell'AI**, non una riduzione di qualità o delle garanzie.
- Tariffe indicative e personalizzabili; **IVA esclusa**.
- Esclusioni v1 (vedi specifica §9): database relazionale, `.rar`/`.7z`, personalizzazione colonne Excel, verifica crittografica della firma p7m.
