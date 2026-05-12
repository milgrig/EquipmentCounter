"""Диагностика наличия текста в PDF (pdfplumber vs fitz)."""
import sys
import pdfplumber
import fitz

sys.stdout.reconfigure(encoding='utf-8')

PATHS = [
    r'Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\019-Лотки отм. 0.000.pdf',
    r'Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\020-Лотки отм. +4.200.pdf',
    r'Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf',
]

for p in PATHS:
    print('=' * 80)
    print(p)
    # pdfplumber
    with pdfplumber.open(p) as pdf:
        page = pdf.pages[0]
        words = page.extract_words()
        chars = page.chars
        print(f'  pdfplumber: words={len(words)}, chars={len(chars)}')
        # examples
        if chars:
            for c in chars[:5]:
                print(f'    char: {c.get("text", "?")!r} @ ({c.get("x0",0):.0f},{c.get("top",0):.0f}) size={c.get("size","?")}')
        if words:
            for w in words[:5]:
                print(f'    word: {w.get("text", "?")!r} @ ({w.get("x0",0):.0f},{w.get("top",0):.0f})')

    # fitz
    doc = fitz.open(p)
    page = doc[0]
    text = page.get_text("text")
    print(f'  fitz get_text len={len(text)}')
    if text:
        # ищем ключевые слова
        for line in text.splitlines()[:30]:
            line = line.strip()
            if line and ('отм' in line.lower() or 'гр.' in line.lower() or 'гр ' in line.lower()):
                print(f'    [match] {line[:100]}')
    # ищем через get_text("dict") — даже outlined могут иметь span
    pdict = page.get_text("dict")
    nblocks = len(pdict.get("blocks", []))
    nspans = sum(len(l.get("spans", [])) for b in pdict.get("blocks", []) for l in b.get("lines", []))
    print(f'  fitz dict: blocks={nblocks}, spans={nspans}')
    doc.close()
