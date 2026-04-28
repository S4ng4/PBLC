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
