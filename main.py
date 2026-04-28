from __future__ import annotations

import os

os.environ.setdefault("TK_SILENCE_DEPRECATION", "1")

import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

from excel.builder import build_and_save_excel
from pdf_margin_edge_final import write_margin_edge_final_pdf
from processing import safe_invoice_filename, verify_invoice

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:  # pragma: no cover
    TkinterDnD = None
    DND_FILES = None

# Tk 8.5 (Python Apple): Frame/Label spesso invisibili → usiamo solo Text + Button (Text disegna sempre).
_BG = "#d4d4d4"
_TEXT_BG = "#ffffff"
_TEXT_FG = "#334155"
_ACCENT = "#0d9488"
_FONT = ("Arial", 12)
_FONT_H = ("Arial", 15, "bold")


def _root() -> tk.Tk:
    if sys.platform == "darwin" or TkinterDnD is None:
        return tk.Tk()
    return TkinterDnD.Tk()


def _tk_patchlevel() -> str:
    try:
        return tk.Tcl().eval("info patchlevel")
    except tk.TclError:
        return ""


class App:
    def __init__(self) -> None:
        self.root = _root()
        self.root.title("PLCB Invoice Processor — Gran Caffè L'Aquila")
        self.root.geometry("780x620")
        self.root.minsize(680, 520)
        self.root.configure(bg=_BG)

        self._pdf_path: str | None = None
        self._summary = None
        self._raw_text = ""
        self._sum_p = 0.0
        self._ok = False
        self._diff = 0.0
        self._last_out: Path | None = None

        self._build_ui()
        self._setup_drop()
        self.root.update_idletasks()

    def _help_text(self) -> str:
        return (
            "COME SI USA\n"
            "──────────\n"
            "1. Premi «Sfoglia… PDF» e scegli una fattura PLCB (Special Order, LSC AR o Wholesale).\n"
            "2. Qui sotto compare il riepilogo. «Verifica OK» = totale calcolato = Total Net del PDF (±1 cent).\n"
            "3. «Genera Excel» crea il file accanto al PDF: {numero}_MarginEdge.xlsx (fogli GCA, MarginEdge, MarginEdgeFinal).\n"
            "4. «Apri cartella» apre la cartella dell’ultimo Excel o del PDF.\n\n"
            "Se vedi solo pulsanti su sfondo nero: nel Terminale usa  python3 main.py --web\n"
        )

    def _build_ui(self) -> None:
        bar = tk.Frame(self.root, bg=_BG, highlightthickness=0)
        bar.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)

        def btn(text: str, cmd, *, primary: bool = False, state=tk.NORMAL) -> tk.Button:
            bg, fg, abg = (_ACCENT, "#ffffff", "#0f766e") if primary else ("#e2e8f0", _TEXT_FG, "#cbd5e1")
            return tk.Button(
                bar,
                text=text,
                command=cmd,
                state=state,
                bg=bg,
                fg=fg,
                activebackground=abg,
                activeforeground=fg if not primary else "#ffffff",
                font=("Arial", 11, "bold") if primary else _FONT,
                padx=14,
                pady=8,
                highlightthickness=0,
                relief=tk.RAISED,
                bd=1,
            )

        self.btn_browse = btn("Sfoglia… PDF", self._pick_file)
        self.btn_browse.pack(side=tk.LEFT, padx=4)
        self.btn_excel = btn("Genera Excel", self._on_generate, primary=True, state=tk.DISABLED)
        self.btn_excel.pack(side=tk.LEFT, padx=4)
        self.btn_open = btn("Apri cartella", self._on_open_dir, state=tk.DISABLED)
        self.btn_open.pack(side=tk.LEFT, padx=4)

        self.body = tk.Text(
            self.root,
            wrap=tk.WORD,
            bg=_TEXT_BG,
            fg=_TEXT_FG,
            font=_FONT,
            relief=tk.SUNKEN,
            bd=1,
            padx=12,
            pady=12,
            highlightthickness=0,
        )
        self.body.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 0))
        self.body.tag_configure("title", font=_FONT_H, foreground="#0f766e")
        self.body.tag_configure("warn", foreground="#b45309")
        self.body.tag_configure("ok", foreground="#047857")
        self._render_body()

    def _render_body(self) -> None:
        self.body.configure(state=tk.NORMAL)
        self.body.delete("1.0", tk.END)
        self.body.insert(tk.END, "PLCB Invoice Processor\n", "title")
        self.body.insert(tk.END, "Gran Caffè L'Aquila — MarginEdge\n\n")
        self.body.insert(tk.END, self._help_text())
        self.body.insert(tk.END, "──────────\nRIEPILOGO\n──────────\n")
        if self._summary is None:
            self.body.insert(tk.END, "(Nessun PDF caricato.)\n")
        else:
            s = self._summary
            self.body.insert(tk.END, f"Tipo: {s.invoice_type}\n")
            self.body.insert(tk.END, f"Invoice Number: {s.invoice_number}\n")
            self.body.insert(tk.END, f"Total Net Price: ${s.total_net_price:,.2f}\n")
            line = f"Totale calcolato (verifica): ${self._sum_p:,.2f}  —  "
            self.body.insert(tk.END, line)
            if self._ok:
                self.body.insert(tk.END, "Verifica OK\n", "ok")
            else:
                self.body.insert(
                    tk.END,
                    f"ATTENZIONE (delta ${self._diff:,.4f})\n",
                    "warn",
                )
            units = sum(i.unit_qty for i in s.items)
            self.body.insert(tk.END, f"Prodotti: {len(s.items)}  —  Bottiglie: {units}\n")
        self.body.configure(state=tk.DISABLED)

    def _setup_drop(self) -> None:
        if sys.platform == "darwin":
            return
        if TkinterDnD is None or not isinstance(self.root, TkinterDnD.Tk):
            return
        try:
            self.body.drop_target_register(DND_FILES)
            self.body.dnd_bind("<<Drop>>", self._on_drop)
        except tk.TclError:
            pass

    def _on_drop(self, event) -> None:
        paths = self.root.tk.splitlist(event.data)
        pdfs = [p.strip("{}") for p in paths if p.strip("{}").lower().endswith(".pdf")]
        if pdfs:
            self._load_pdf(pdfs[0])

    def _pick_file(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf"), ("Tutti i file", "*.*")])
        if path:
            self._load_pdf(path)

    def _load_pdf(self, path: str) -> None:
        self._pdf_path = path
        try:
            summary, raw, sp, ok, diff = verify_invoice(path)
        except Exception as exc:  # pragma: no cover
            messagebox.showerror("Errore", f"Impossibile leggere il PDF:\n{exc}")
            return
        self._summary = summary
        self._raw_text = raw
        self._sum_p = sp
        self._ok = ok
        self._diff = diff
        self._render_body()
        self.btn_excel.configure(state=tk.NORMAL)
        self.btn_open.configure(state=tk.NORMAL)

    def _on_generate(self) -> None:
        if not self._pdf_path or not self._summary:
            return
        if not self._ok:
            if not messagebox.askyesno(
                "Verifica totali",
                f"La somma delle righe P (${self._sum_p:,.2f}) non coincide con il Total Net "
                f"(${self._summary.total_net_price:,.2f}).\nDifferenza: ${self._diff:,.4f}.\n\nSalvare comunque?",
                icon="warning",
            ):
                return
        try:
            inv = safe_invoice_filename(self._summary.invoice_number)
            out = Path(self._pdf_path).resolve().parent / f"{inv}_MarginEdge.xlsx"
            self._last_out = build_and_save_excel(self._summary, self._pdf_path, self._raw_text, str(out))
            dest_gca = out.parent / f"{inv}_fattura_GCA.pdf"
            write_margin_edge_final_pdf(self._summary, dest_gca)
        except Exception as exc:  # pragma: no cover
            messagebox.showerror("Errore", f"Salvataggio non riuscito:\n{exc}")
            return
        messagebox.showinfo(
            "Completato",
            f"File creati:\n{self._last_out}\n\nPDF fattura GCA (export MarginEdgeFinal):\n{dest_gca}",
        )

    def _on_open_dir(self) -> None:
        if self._last_out is not None:
            p = self._last_out
        elif self._pdf_path:
            p = Path(self._pdf_path)
        else:
            return
        folder = str(p.parent if p.is_file() else p)
        try:
            if sys.platform == "darwin":
                os.system(f'open "{folder}"')
            elif sys.platform.startswith("win"):
                os.startfile(folder)  # type: ignore[attr-defined]
            else:
                os.system(f'xdg-open "{folder}"')
        except OSError as exc:  # pragma: no cover
            messagebox.showwarning("Apri cartella", str(exc))

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    App().run()


if __name__ == "__main__":
    if "--web" in sys.argv:
        from web_ui import run_web_ui

        run_web_ui()
    else:
        if sys.platform == "darwin" and _tk_patchlevel().startswith("8.5"):
            print(
                "\n*** Tk 8.5 rilevato (tipico del python3 Apple). Se la finestra è nera, usa:\n"
                "    python3 main.py --web\n",
                file=sys.stderr,
            )
        main()
