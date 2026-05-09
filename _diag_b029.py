"""B029 diagnostic: does the 005 floor-plan PDF contain L=Nm tokens?

Also dumps all tokens matching rough length-like patterns so we can
decide whether floor plans carry length tokens at all.
"""
from __future__ import annotations
import re
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

import pdfplumber

PDF_FLOOR = Path(r"Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf")
PDF_SCHEMA = Path(r"Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\003 - Схемы ЩО, ЩАО-ЩО-3.pdf")

LEN_STRICT = re.compile(r"^L=(\d+(?:[,\.]\d+)?)м$")
LEN_LOOSE = re.compile(r"L\s*=\s*\d+[,\.]?\d*\s*[мm]?", re.IGNORECASE)
# Possibly length-like: '=Nm', 'мN=L', 'Nм', 'N m'
LEN_ANY = re.compile(r"\d+(?:[,\.]\d+)?\s*м\b", re.IGNORECASE)
REV_LEN = re.compile(r"^м(\d+(?:[,\.]\d+)?)=L$")

def scan(pdf_path: Path, label: str) -> None:
    if not pdf_path.exists():
        print(f"[SKIP] {label}: {pdf_path.name} not found")
        return
    print("=" * 80)
    print(f"{label}: {pdf_path.name}")
    print("=" * 80)
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_i, page in enumerate(pdf.pages):
            words = page.extract_words() or []
            strict, loose_only, rev, any_m = [], [], [], []
            for w in words:
                t = w["text"]
                if LEN_STRICT.match(t):
                    strict.append(t)
                elif LEN_LOOSE.search(t):
                    loose_only.append(t)
                elif REV_LEN.match(t):
                    rev.append(t)
                elif LEN_ANY.search(t):
                    any_m.append(t)
            print(f"\n  Page {page_i}: {len(words)} words")
            print(f"    strict L=Nм matches ({len(strict)}): {strict[:20]}")
            print(f"    loose L=N... ({len(loose_only)}): {loose_only[:20]}")
            print(f"    reversed мN=L ({len(rev)}): {rev[:10]}")
            print(f"    other 'Nм' tokens ({len(any_m)}): {any_m[:20]}")

scan(PDF_FLOOR, "FLOOR PLAN")
scan(PDF_SCHEMA, "SCHEMA")
