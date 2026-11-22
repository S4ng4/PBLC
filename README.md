# PLCB Invoice Processor

Applicazione Python locale con interfaccia Tkinter per estrarre dati da fatture PLCB (Special Order, LSC AR, Wholesale Licensee), calcolare l’upcharge per bottiglia in stile MarginEdge e generare un file Excel a tre fogli (`GCA`, `MarginEdge`, `MarginEdgeFinal`).

## Requisiti

- Python 3.10+
- Dipendenze: `pip install -r requirements.txt`

`tkinter` è incluso nella installazione standard di Python su macOS/Windows. Il drag & drop dei PDF nella finestra richiede `tkinterdnd2`; se non è installato, resta disponibile la selezione file con clic.

## Avvio

```bash
cd plcb-invoice-processor
pip install -r requirements.txt
python3 main.py
```

### macOS: finestra nera con `python3` di sistema

Il Python dei **Command Line Tools** usa **Tk 8.5**: spesso i `Frame`/`Label` non si disegnano (schermo nero con solo i pulsanti).

- **Consigliato:** interfaccia nel browser (nessun Tkinter visibile):

  ```bash
  python3 main.py --web
  ```

  Poi segui la pagina che si apre; in **Download** trovi Excel, copia identica del PDF PLCB e un PDF `{Invoice}_fattura_GCA.pdf` che è l’**export del foglio `MarginEdgeFinal`** (richiede `reportlab`).

- In alternativa installa **Python da [python.org](https://www.python.org/downloads/macos/)** o **Homebrew** e usa quel `python3` con `main.py` senza `--web`.

Trascina uno o più PDF (su Windows/Linux, se `tkinterdnd2` è attivo) oppure usa **Sfoglia… PDF**. Dopo l’analisi, **Genera Excel** crea nella stessa cartella del PDF `{InvoiceNumber}_MarginEdge.xlsx` e `{InvoiceNumber}_fattura_GCA.pdf` (export del foglio `MarginEdgeFinal`).

## Verifica totali

Prima del salvataggio viene confrontato il totale ricostruito con le formule MarginEdge (somma delle righe P) con il **Total Net Price** estratto dal PDF. Se la differenza supera 1 centesimo, viene mostrato un avviso con la discrepanza; puoi comunque scegliere di salvare.

## Struttura

Vedi layout dei moduli in `parser/`, `calculator/`, `excel/` come da specifica del progetto.

## Foglio «MarginEdgeFinal» (grafica PLCB)

Il foglio include una **fascia blu** (stile istituzionale), **striscia oro** con testo PLCB, **logo** a sinistra e intestazioni tabella in blu con bordi.  
Il logo è preso da `assets/plcb_logo.png` se presente; altrimenti viene generato un wordmark «PLCB» con **Pillow** (già in `requirements.txt`).

**Stampa:** ogni foglio è impostato su carta **US Letter verticale** (8,5×11 pollici), margini ~0,7″ e **adatta tutte le colonne alla larghezza di una pagina** (`fitToWidth`), come in anteprima di stampa tipica su foglio Letter in verticale.

## Automazione Email (IMAP + SMTP)

Il worker `email_worker.py` monitora una **casella email predefinita** (o un alias come `tuamail+plcb@gmail.com`), elabora automaticamente le fatture PLCB allegate in PDF e invia i file a gruppi di destinatari predisposti:

- **`PDF_ONLY_RECIPIENTS`** (es. email di import **MarginEdge**): riceve **solo** `{Invoice}_fattura_GCA.pdf`.
- **`FULL_RECIPIENTS`** (es. amministrazione / contabilità): riceve `{Invoice}_fattura_GCA.pdf`, `{Invoice}_MarginEdge.xlsx` e (opzionalmente) il PDF originale PLCB.
- **`ALERT_RECIPIENTS`** (controllo/admin): se la verifica totali supera $0.01 di scarto o il PDF dà errore, **blocca l'invio automatico** ai destinatari sopra e invia un alert con i file per verifica manuale.

Le email processate con successo vengono spostate in `PLCB/Processed`; quelle con scarto/errore in `PLCB/Error`.

### 1. Setup Gmail (App Password)

1. Attiva la **Verifica in due passaggi** sull'account Google della casella dedicata.
2. Genera una **Password per le app** (16 caratteri) in *Account Google → Sicurezza → Password per le app*.
3. Assicurati che **IMAP** sia attivo nelle impostazioni di Gmail.

### 2. Inserire le email desiderate e avviare in locale (`--setup` / `--web` / `--watch`)

Puoi inserire la tua casella e i destinatari predisposti in **due modi semplici** (entrambi salvano in locale nel file `.env`):

- **Da browser:** avvia `python3 main.py --web` e compila la sezione **«Automazione Email — Destinatari predisposti»** (puoi anche lanciare subito *Prova (Dry-Run)* o *Controlla casella ora*).
- **Da terminale (procedura guidata):**

```bash
python3 email_worker.py --setup            # Procedura guidata interattiva per inserire le email
python3 email_worker.py --check-config     # Verifica configurazione
python3 email_worker.py --once --dry-run   # Prova senza inviare né spostare email
python3 email_worker.py --once             # Singolo controllo (ideale per cron)
python3 email_worker.py --watch            # Ascolto continuo (default ogni 60s)
```

### 3. Esecuzione automatica su GitHub Actions (a computer spento)

Il workflow [`.github/workflows/plcb-email-worker.yml`](.github/workflows/plcb-email-worker.yml) controlla la casella ogni **15 minuti** (e può essere lanciato a mano con opzione `dry_run`).

Nel repository GitHub vai su **Settings → Secrets and variables → Actions** e aggiungi i seguenti **Secrets**:

- `IMAP_USER` — indirizzo email della casella predefinita
- `IMAP_PASSWORD` — App Password di Gmail (16 caratteri)
- `PDF_ONLY_RECIPIENTS` — destinatari solo PDF (es. email MarginEdge), separati da virgola
- `FULL_RECIPIENTS` — destinatari PDF + Excel, separati da virgola
- `ALERT_RECIPIENTS` — email di controllo per errori/discrepanze

Opzionali (come **Variables** se vuoi cambiare i default): `IMAP_HOST`, `IMAP_PORT`, `IMAP_MAILBOX`, `IMAP_PROCESSED_FOLDER`, `IMAP_ERROR_FOLDER`, `IMAP_TO_FILTER`, `IMAP_ALLOWED_SENDERS`, `IMAP_SUBJECT_CONTAINS`, `SMTP_HOST`, `SMTP_PORT`, `SMTP_FROM`, `INCLUDE_ORIGINAL_PDF_IN_FULL`, `ALERT_ON_UNRECOGNIZED_PDF`.

