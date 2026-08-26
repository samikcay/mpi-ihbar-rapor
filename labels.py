# -*- coding: utf-8 -*-
"""
SAP etiketlerini Excel'de kullanilan yazima cevirir.

SAP ciktilarinda Turkce karakterler bazen kaybolur (BTK ENGELLI) ve
buyuk/kucuk harf kullanimi rapordan farklidir. Grafikler bu hucreleri
etiket olarak kullandigi icin rapordaki yazimi korumak isteriz.

Eslesme aksan ve buyuk/kucuk harf duyarsizdir; listede olmayan yeni bir
etiket gelirse SAP'tan geldigi gibi yazilir (veri kaybolmaz).
"""

# Excel'de gorulmesini istedigimiz yazimlar
TERCIH_EDILEN = [
    # Site icerigine gore
    "DİREK OYNATAN",
    "TANITIM, REKLAM VE YÖNLENDİRME",
    # Site durumuna gore
    "BTK ENGELLİ",
    "SİTE DEVAM EDİYOR",
    "REKLAM İLE YÖNLENDİRME",
    "DOMAİN",
]

# Turkce harfleri ASCII karsiligina indirger (I/İ/ı ayrimi dahil)
_ASCII = str.maketrans({
    "ı": "i", "İ": "i", "I": "i", "i": "i",
    "ş": "s", "Ş": "s",
    "ğ": "g", "Ğ": "g",
    "ü": "u", "Ü": "u",
    "ö": "o", "Ö": "o",
    "ç": "c", "Ç": "c",
    "â": "a", "Â": "a",
    "î": "i", "Î": "i",
    "û": "u", "Û": "u",
})


def anahtar(metin):
    """Karsilastirma anahtari: aksansiz, kucuk harf, tek bosluk."""
    return " ".join((metin or "").translate(_ASCII).lower().split())


_HARITA = {anahtar(x): x for x in TERCIH_EDILEN}


def duzelt(metin, ek_tercihler=None):
    """
    SAP etiketini rapordaki yazima cevirir.

    ek_tercihler: calisma kitabindan okunan mevcut etiketler
                  (sube adlari, ulke adlari gibi degisken listeler icin).
    """
    if not metin:
        return metin
    k = anahtar(metin)
    if ek_tercihler:
        for mevcut in ek_tercihler:
            if anahtar(mevcut) == k:
                return mevcut
    return _HARITA.get(k, metin)
