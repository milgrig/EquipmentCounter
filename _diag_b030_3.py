"""B030 diagnostic part 3: enable DEBUG logging for match_symbols and run."""
from __future__ import annotations
import sys
import logging
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
logging.basicConfig(level=logging.DEBUG, format="%(name)s %(levelname)s %(message)s")
# Quiet pdfplumber/pdfminer
logging.getLogger("pdfplumber").setLevel(logging.WARNING)
logging.getLogger("pdfminer").setLevel(logging.WARNING)
logging.getLogger("PIL").setLevel(logging.WARNING)
logging.getLogger("fitz").setLevel(logging.WARNING)

PDF_PATH = Path(
    r"Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf"
)
from pdf_legend_parser import parse_legend
from pdf_count_visual import match_symbols

legend = parse_legend(str(PDF_PATH))
res = match_symbols(str(PDF_PATH), legend)
print("\n=== FINAL ===")
for k, v in sorted(res.counts.items()):
    print(f"  [{k}] = {v}")
