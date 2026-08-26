# -*- coding: utf-8 -*-
"""
SAP metin ciktilarini (ALV list export) okuyup satirlara donusturur.

SAP ciktilari UTF-16LE + BOM olarak kaydedilir, sutunlar TAB ile ayrilir ve
veri satirlarinin arasina bos satir konur. Basliklar ilk ~8 satirdadir,
son satirda "TOPLAM SITE SAYISI :" bulunur.
"""

import re
import unicodedata

# --- Sayi bicimi ------------------------------------------------------------

# SAP Turkce sayilari nokta binlik ayraci ile yazar: 64.145 -> 64145
_NUM_RE = re.compile(r"^-?[\d.  ]+$")


class ParseError(Exception):
    """SAP dosyasi beklenen yapida degil."""


def parse_int(text):
    """'64.145' -> 64145. Bozuk deger icin ParseError."""
    t = (text or "").strip().replace(" ", "").replace(" ", "")
    if not t or not _NUM_RE.match(t):
        raise ParseError("Sayi okunamadi: %r" % text)
    t = t.replace(".", "")
    try:
        return int(t)
    except ValueError:
        raise ParseError("Sayi okunamadi: %r" % text)


# --- Dosya okuma ------------------------------------------------------------

def read_sap_text(path):
    """SAP dosyasini kodlamasini otomatik saptayarak okur, satir listesi doner."""
    with open(path, "rb") as fh:
        raw = fh.read()

    if not raw.strip():
        raise ParseError("Dosya bos: %s" % path)

    # BOM'a gore kodlama sec; SAP normalde UTF-16LE yazar ama
    # bazi sistemlerde UTF-8 veya Windows-1254 gelebilir.
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        candidates = ["utf-16"]
    elif raw[:3] == b"\xef\xbb\xbf":
        candidates = ["utf-8-sig"]
    else:
        candidates = ["utf-8", "cp1254", "latin-1"]

    text = None
    for enc in candidates:
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise ParseError("Dosya kodlamasi cozulemedi: %s" % path)

    # Bicim isaretlerini ve sayfa besleme karakterlerini temizle
    text = text.replace("\x0c", "\n").replace("­", "")
    return [ln.rstrip() for ln in text.splitlines()]


def _clean(cell):
    """Hucre metnini normalize eder (NBSP, cift bosluk, bas/son bosluk)."""
    s = (cell or "").replace(" ", " ")
    s = unicodedata.normalize("NFC", s)
    return re.sub(r"\s+", " ", s).strip()


def _split_cells(line):
    return [_clean(c) for c in line.split("\t")]


# --- Ortak yapi -------------------------------------------------------------

_TOTAL_RE = re.compile(r"TOPLAM\s+S[İI]TE\s+SAYISI", re.IGNORECASE)


def _find_total(lines):
    """Dosya sonundaki TOPLAM degerini bulur (dogrulama icin)."""
    for line in reversed(lines):
        if _TOTAL_RE.search(line):
            cells = [c for c in _split_cells(line) if c]
            for cell in reversed(cells):
                try:
                    return parse_int(cell)
                except ParseError:
                    continue
    return None


def extract_period(lines):
    """Basliktaki '01.01.2026-07.08.2026' donemini bulur."""
    pat = re.compile(r"(\d{2}\.\d{2}\.\d{4})\s*-\s*(\d{2}\.\d{2}\.\d{4})")
    for line in lines[:15]:
        m = pat.search(line.replace(" ", "") if line.count(" ") > 20 else line)
        if m:
            return m.group(1), m.group(2)
    # Bosluklu yazilmis olabilir
    joined = "".join(lines[:15]).replace(" ", "")
    m = pat.search(joined)
    if m:
        return m.group(1), m.group(2)
    return None


def _data_rows(lines):
    """Basliktan sonraki, TOPLAM'dan onceki dolu satirlari doner."""
    rows = []
    for line in lines:
        if _TOTAL_RE.search(line):
            break
        if "\t" not in line:
            continue
        cells = _split_cells(line)
        # Veri satiri ilk hucresi sira numarasi olan satirdir
        if cells and cells[0].isdigit():
            rows.append(cells)
    if not rows:
        raise ParseError("Dosyada veri satiri bulunamadi")
    return rows


# --- Tekil ayristiricilar ---------------------------------------------------

def parse_simple(path):
    """
    Tek bloklu raporlar (icerik / site durumu / sube).

    Doner: {'items': [(ad, sayi), ...], 'total': int|None, 'period': (bas, bit)|None}
    """
    lines = read_sap_text(path)
    items = []
    for cells in _data_rows(lines):
        # Beklenen: [sira, ad, sayi]  (bos hucreler olabilir)
        parts = [c for c in cells if c != ""]
        if len(parts) < 3:
            raise ParseError("Eksik sutun: %r" % (cells,))
        name, value = parts[1], parts[2]
        items.append((name, parse_int(value)))
    return {
        "items": items,
        "total": _find_total(lines),
        "period": extract_period(lines),
    }


def parse_countries(path):
    """
    Ulke raporu iki blok halinde yazilir (sol: 1..n, sag: n+1..son).
    Sira numarasina gore birlestirip tek listeye cevirir.

    Doner: {'items': [(ulke, sayi), ...], 'total': int|None, 'period': ...}
    """
    lines = read_sap_text(path)
    by_rank = {}
    for cells in _data_rows(lines):
        parts = [c for c in cells if c != ""]
        # Satirda 1 veya 2 blok olabilir: [s,ulke,sayi] veya [s,ulke,sayi,s,ulke,sayi]
        if len(parts) not in (3, 6):
            raise ParseError("Beklenmeyen sutun sayisi: %r" % (cells,))
        for i in range(0, len(parts), 3):
            rank_s, name, value = parts[i], parts[i + 1], parts[i + 2]
            if not rank_s.isdigit():
                raise ParseError("Sira numarasi okunamadi: %r" % rank_s)
            rank = int(rank_s)
            if rank in by_rank:
                raise ParseError("Sira numarasi tekrar ediyor: %d" % rank)
            by_rank[rank] = (name, parse_int(value))

    ranks = sorted(by_rank)
    if ranks != list(range(1, len(ranks) + 1)):
        missing = set(range(1, max(ranks) + 1)) - set(ranks)
        raise ParseError("Sira numaralarinda bosluk var: %s" % sorted(missing))

    return {
        "items": [by_rank[r] for r in ranks],
        "total": _find_total(lines),
        "period": extract_period(lines),
    }
