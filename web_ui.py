"""
Interfaccia web locale (nessun tkinter): consigliata su macOS con Tk 8.5.x
(Python dei Command Line Tools), dove la finestra Tk resta spesso nera.

Avvio: python3 main.py --web
Poi apri il browser all’indirizzo indicato (di solito http://127.0.0.1:8765).
"""
from __future__ import annotations

import base64
import json
import os
import secrets
import sys
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Dict
from urllib.parse import parse_qs, urlparse

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from excel.builder import build_and_save_excel
from models import verify_totals
from pdf_margin_edge_final import write_margin_edge_final_pdf
from processing import compute_sum_p, load_summary, safe_invoice_filename

_DOWNLOADS: Dict[str, Path] = {}


def _normalized_request_path(raw_path: str) -> str:
    """Path assoluto senza slash finale (eccetto la root `/`)."""
    p = (urlparse(raw_path).path or "/").rstrip("/")
    return p if p else "/"


def _download_content_type(p: Path) -> str:
    s = p.suffix.lower()
    if s == ".pdf":
        return "application/pdf"
    if s in (".xlsx", ".xlsm"):
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return "application/octet-stream"


# Favicon 1×1 trasparente (GIF): evita 404 su /favicon.ico nella console del browser
_FAVICON_GIF = base64.b64decode(
    "R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7"
)

HTML_PAGE = """<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<link rel="icon" href="/favicon.ico" type="image/gif"/>
<title>PLCB Invoice Processor</title>
<style>
  body { font-family: system-ui, sans-serif; max-width: 44rem; margin: 2rem auto; padding: 0 1rem;
         background: #f1f5f9; color: #1e293b; }
  h1 { color: #0f766e; font-size: 1.35rem; }
  .card { background: #fff; border: 1px solid #e2e8f0; border-radius: 10px; padding: 1.25rem; margin: 1rem 0; }
  label.btn, button.btn { display: inline-block; background: #0d9488; color: #fff; padding: 0.6rem 1.2rem;
               border-radius: 8px; cursor: pointer; font-weight: 600; border: none; font: inherit; }
  label.btn:hover, button.btn:hover { background: #0f766e; }
  input[type=file] { display: none; }
  #out { white-space: pre-wrap; font-size: 0.95rem; background: #f8fafc; padding: 1rem; border-radius: 8px; }
  #previewWrap { margin-top: 1rem; }
  #previewMeta { font-size: 0.9rem; color: #475569; margin: 0.25rem 0 0.75rem; }
  #pdfPreview { width: 100%; height: 480px; border: 1px solid #e2e8f0; border-radius: 10px; background: #0b1220; }
  #loading { margin-top: 1rem; }
  #loading[hidden] { display: none; }
  #bar { height: 10px; background: #e2e8f0; border-radius: 999px; overflow: hidden; }
  #bar > div { height: 100%; width: 35%; background: #0d9488; border-radius: 999px; animation: slide 1.0s infinite linear; }
  @keyframes slide { 0% { transform: translateX(-120%); } 100% { transform: translateX(320%); } }
  #loadingTxt { margin-top: 0.5rem; font-size: 0.9rem; color: #475569; }
  a.dl { color: #0d9488; font-weight: 600; }
  ul { margin: 0.5rem 0 0 1.2rem; }
  #afterOk[hidden] { display: none; }
  #afterOk .dl-row { display: flex; flex-wrap: wrap; gap: 0.5rem; align-items: center; margin-top: 0.75rem; }
  #afterOk .dl-note { font-size: 0.88rem; color: #475569; margin-top: 0.5rem; }
  a.btn.dl-btn { text-decoration: none; text-align: center; }
  a.btn.dl-btn.secondary { background: #64748b; }
  a.btn.dl-btn.secondary:hover { background: #475569; }
</style>
</head>
<body>
  <h1>PLCB Invoice Processor — Gran Caffè L'Aquila</h1>
  <div class="card">
    <p><strong>Come si usa</strong></p>
    <ul>
      <li>Scegli un PDF di fattura PLCB (Special Order, LSC AR o Wholesale).</li>
      <li>Controlla il riepilogo: se «Verifica» è OK, il totale calcolato coincide con il netto del PDF.</li>
      <li>Clicca «Analizza e genera Excel»: poi usa i pulsanti sotto per scaricare subito i file (il PDF GCA è l’<strong>export del foglio MarginEdgeFinal</strong>).</li>
    </ul>
    <form id="f" method="post" enctype="multipart/form-data" style="margin-top:1rem;">
      <input type="file" id="pdf" name="pdf" accept=".pdf,application/pdf"/>
      <label class="btn" for="pdf">1 — Scegli PDF</label>
      <button id="submitBtn" type="submit" class="btn" style="margin-left:0.5rem;" disabled>2 — Analizza e genera Excel</button>
    </form>
    <div id="previewWrap" hidden>
      <div id="previewMeta"></div>
      <p class="dl-note" style="margin:0.35rem 0 0.5rem;">Anteprima = PDF <strong>caricato</strong> (nessun logo). Il PDF con logo GCA esiste solo dopo l’analisi: usa «Scarica PDF fattura (layout PLCB + logo GCA)».</p>
      <iframe id="pdfPreview" title="Anteprima PDF caricato"></iframe>
    </div>
    <div id="loading" hidden>
      <div id="bar"><div></div></div>
      <div id="loadingTxt">Analisi in corso…</div>
    </div>
    <div class="card" id="afterOk" hidden style="margin-top:1rem;">
      <p><strong>Download</strong></p>
      <div class="dl-row">
        <a id="btnPdfGca" class="btn dl-btn" href="#" download>Scarica PDF fattura (layout PLCB + logo GCA)</a>
        <a id="btnExcel" class="btn dl-btn secondary" href="#" download>Scarica Excel</a>
        <a id="btnPdfPlcb" class="btn dl-btn secondary" href="#" download>Scarica copia PDF PLCB</a>
      </div>
      <p class="dl-note">Il PDF «GCA» è l’export del foglio <strong>MarginEdgeFinal</strong> (tabella pronta per import in MarginEdge).</p>
    </div>
  </div>
  <div class="card" id="out" hidden></div>
  <script>
  const pdfEl = document.getElementById('pdf');
  const submitBtn = document.getElementById('submitBtn');
  const previewWrap = document.getElementById('previewWrap');
  const previewMeta = document.getElementById('previewMeta');
  const pdfPreview = document.getElementById('pdfPreview');
  const loading = document.getElementById('loading');
  const loadingTxt = document.getElementById('loadingTxt');
  const out = document.getElementById('out');
  const afterOk = document.getElementById('afterOk');
  const btnPdfGca = document.getElementById('btnPdfGca');
  const btnExcel = document.getElementById('btnExcel');
  const btnPdfPlcb = document.getElementById('btnPdfPlcb');
  let blobUrl = null;

  function hideDownloads() {
    afterOk.hidden = true;
    btnPdfGca.href = '#';
    btnExcel.href = '#';
    btnPdfPlcb.href = '#';
  }

  pdfEl.addEventListener('change', () => {
    const f = pdfEl.files && pdfEl.files[0];
    out.hidden = true;
    hideDownloads();
    if (!f) {
      submitBtn.disabled = true;
      previewWrap.hidden = true;
      if (blobUrl) URL.revokeObjectURL(blobUrl);
      blobUrl = null;
      return;
    }
    submitBtn.disabled = false;
    previewWrap.hidden = false;
    previewMeta.textContent = `Selezionato: ${f.name} — ${(f.size/1024/1024).toFixed(2)} MB`;
    if (blobUrl) URL.revokeObjectURL(blobUrl);
    blobUrl = URL.createObjectURL(f);
    pdfPreview.src = blobUrl;
  });

  document.getElementById('f').onsubmit = async (e) => {
    e.preventDefault();
    if (!pdfEl.files || !pdfEl.files[0]) return;
    const fd = new FormData(e.target);
    submitBtn.disabled = true;
    loading.hidden = false;
    loadingTxt.textContent = 'Analisi in corso… (può richiedere 5–30s)';
    try {
      const r = await fetch('/', { method: 'POST', body: fd });
      const j = await r.json();
      out.hidden = false;
      if (!j.ok) {
        out.textContent = 'Errore: ' + (j.error || r.status);
        hideDownloads();
        return;
      }
      let t = 'Tipo: ' + j.invoice_type + '\\nInvoice: ' + j.invoice_number + '\\nTotal Net: $' + j.total_net.toFixed(2)
        + '\\nTotale calcolato: $' + j.sum_p.toFixed(2) + '\\nVerifica: ' + (j.verify_ok ? 'OK' : 'ATTENZIONE (delta $' + j.diff.toFixed(4) + ')')
        + '\\nProdotti: ' + j.n_items + ' — Bottiglie: ' + j.units + '\\n\\n';
      if (j.download_path) t += 'Excel:\\n' + j.download_path + '\\n';
      if (j.pdf_saved_path) t += '\\nPDF (copia PLCB):\\n' + j.pdf_saved_path + '\\n';
      if (j.pdf_branded_path) t += '\\nPDF con logo GCA:\\n' + j.pdf_branded_path + '\\n';
      if (j.download_url) t += '\\nScarica Excel (browser): ' + j.download_url;
      if (j.download_url_pdf_plcb) t += '\\nScarica PDF PLCB (browser): ' + j.download_url_pdf_plcb;
      if (j.download_url_pdf_gca) t += '\\nScarica PDF con logo GCA (browser): ' + j.download_url_pdf_gca;
      out.textContent = t;
      if (j.download_url_pdf_gca) {
        btnPdfGca.href = j.download_url_pdf_gca;
        afterOk.hidden = false;
      } else {
        hideDownloads();
      }
      btnExcel.href = j.download_url || '#';
      btnPdfPlcb.href = j.download_url_pdf_plcb || '#';
      btnExcel.style.display = j.download_url ? 'inline-block' : 'none';
      btnPdfPlcb.style.display = j.download_url_pdf_plcb ? 'inline-block' : 'none';
    } catch (err) {
      out.hidden = false;
      out.textContent = 'Errore: ' + (err && err.message ? err.message : String(err));
      hideDownloads();
    } finally {
      loading.hidden = true;
      submitBtn.disabled = false;
    }
  };
  </script>
</body>
</html>
"""


def _parse_multipart(body: bytes, ctype: str) -> tuple[bytes, str]:
    if "boundary=" not in ctype:
        raise ValueError("multipart senza boundary")
    boundary = ctype.split("boundary=", 1)[1].strip().strip('"').encode("ascii")
    delim = b"--" + boundary + b"\r\n"
    end = b"--" + boundary + b"--"
    if delim not in body:
        raise ValueError("corpo multipart non valido")
    chunk = body.split(delim, 1)[1]
    if b"\r\n\r\n" not in chunk:
        raise ValueError("intestazioni file mancanti")
    heads, rest = chunk.split(b"\r\n\r\n", 1)
    if end in rest:
        rest = rest.split(end)[0]
    if rest.endswith(b"\r\n"):
        rest = rest[:-2]
    filename = "upload.pdf"
    for line in heads.split(b"\r\n"):
        if line.lower().startswith(b"content-disposition:"):
            if b'filename="' in line:
                fn = line.split(b'filename="', 1)[1].split(b'"', 1)[0].decode("utf-8", "replace")
                filename = fn or filename
    return rest, filename


class Handler(BaseHTTPRequestHandler):
    server_version = "PLCBWeb/1.0"

    def log_message(self, fmt: str, *args) -> None:
        sys.stderr.write("%s - %s\n" % (self.address_string(), fmt % args))

    def do_GET(self) -> None:
        u = urlparse(self.path)
        path = _normalized_request_path(self.path)
        if path.lower() == "/favicon.ico":
            self.send_response(200)
            self.send_header("Content-Type", "image/gif")
            self.send_header("Content-Length", str(len(_FAVICON_GIF)))
            self.send_header("Cache-Control", "public, max-age=86400")
            self.end_headers()
            self.wfile.write(_FAVICON_GIF)
            return
        if path == "/download":
            qs = parse_qs(u.query or "")
            tok = (qs.get("t") or [""])[0]
            path = _DOWNLOADS.pop(tok, None)
            if not path or not path.is_file():
                self.send_error(404, "Link scaduto o non valido")
                return
            data = path.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", _download_content_type(path))
            self.send_header("Content-Disposition", f'attachment; filename="{path.name}"')
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            try:
                path.unlink(missing_ok=True)
            except OSError:
                pass
            return
        if path != "/":
            self.send_error(404)
            return
        b = HTML_PAGE.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(b)))
        self.end_headers()
        self.wfile.write(b)

    def do_POST(self) -> None:
        if self.path != "/":
            self.send_error(404)
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            length = 0
        if length <= 0:
            self._json(400, {"ok": False, "error": "Corpo vuoto"})
            return
        body = self.rfile.read(length)
        ctype = self.headers.get("Content-Type", "")
        try:
            raw_pdf, _fn = _parse_multipart(body, ctype)
        except ValueError as e:
            self._json(400, {"ok": False, "error": str(e)})
            return
        tmp = NamedTemporaryFile(suffix=".pdf", delete=False)
        try:
            tmp.write(raw_pdf)
            tmp.close()
            path = tmp.name
            summary, raw = load_summary(path)
            sp = compute_sum_p(summary)
            ok, diff = verify_totals(summary, sp)
            units = sum(i.unit_qty for i in summary.items)
            inv_stem = safe_invoice_filename(summary.invoice_number)
            out_name = f"{inv_stem}_MarginEdge.xlsx"
            out_dir = Path.home() / "Downloads"
            out_dir.mkdir(parents=True, exist_ok=True)
            dest = out_dir / out_name
            build_and_save_excel(summary, path, raw, str(dest))
            # Copia byte-identica PDF PLCB + PDF GCA (export del foglio MarginEdgeFinal)
            dest_pdf = out_dir / f"{inv_stem}_fattura_PLCB.pdf"
            dest_pdf.write_bytes(raw_pdf)
            dest_pdf_gca = out_dir / f"{inv_stem}_fattura_GCA.pdf"
            write_margin_edge_final_pdf(summary, dest_pdf_gca)
            tok = secrets.token_urlsafe(16)
            tok_pdf = secrets.token_urlsafe(16)
            tok_gca = secrets.token_urlsafe(16)
            _DOWNLOADS[tok] = dest
            _DOWNLOADS[tok_pdf] = dest_pdf
            _DOWNLOADS[tok_gca] = dest_pdf_gca
            host = self.headers.get("Host", "127.0.0.1:8765")
            base = f"http://{host}"
            self._json(
                200,
                {
                    "ok": True,
                    "invoice_type": summary.invoice_type,
                    "invoice_number": summary.invoice_number,
                    "total_net": summary.total_net_price,
                    "sum_p": sp,
                    "verify_ok": ok,
                    "diff": diff,
                    "n_items": len(summary.items),
                    "units": units,
                    "download_path": str(dest),
                    "pdf_saved_path": str(dest_pdf),
                    "pdf_branded_path": str(dest_pdf_gca),
                    "download_url": f"{base}/download?t={tok}",
                    "download_url_pdf_plcb": f"{base}/download?t={tok_pdf}",
                    "download_url_pdf_gca": f"{base}/download?t={tok_gca}",
                },
            )
        except Exception as e:
            self._json(500, {"ok": False, "error": str(e)})
        finally:
            try:
                os.unlink(tmp.name)
            except OSError:
                pass

    def _json(self, code: int, obj: dict) -> None:
        b = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(b)))
        self.end_headers()
        self.wfile.write(b)


def run_web_ui(host: str = "127.0.0.1", port: int = 8765) -> None:
    httpd = None
    chosen = None
    for p in range(int(port), int(port) + 25):
        try:
            httpd = ThreadingHTTPServer((host, p), Handler)
            chosen = p
            break
        except OSError as e:
            # macOS: Errno 48 = address already in use
            if getattr(e, "errno", None) == 48:
                continue
            raise
    if httpd is None or chosen is None:
        raise RuntimeError("Impossibile avviare il server web: nessuna porta libera (range 25).")

    url = f"http://{host}:{chosen}/"
    print(f"Interfaccia web: {url}", flush=True)
    print("Premi Ctrl+C per terminare.", flush=True)
    threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nChiusura.")
    finally:
        httpd.server_close()
