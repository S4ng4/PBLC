Logo opzionale per il foglio Excel «MarginEdgeFinal»:
- Salva qui un file PNG chiamato  plcb_logo.png  (consigliato ~200×60 px, sfondo trasparente o blu scuro).
- Se il file è assente, viene usato un wordmark generato automaticamente (richiede il pacchetto Pillow).

Logo per il PDF con marchio Gran Caffè L'Aquila (sovrapposto al PDF fattura):
- Predefinito:  gca_logo_builtin.jpg  (JPEG incluso nel repository).
- Opzionale:  gca_logo.png  (PNG con trasparenza consigliata, es. ~300×60 px) — se presente, ha priorità sul JPEG.
- Se mancano entrambi, viene generato un wordmark testuale (Pillow).
- Il PDF di output replica il file PLCB (PyMuPDF) con il logo in alto a destra su ogni pagina.
