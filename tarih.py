# -*- coding: utf-8 -*-
"""
Rapor donemi tarihlerini bicimler ve dogrular.

Uc ayri bicim var:
  1. Sekme basliklari (A3/A4)  : "01.01.2026 - 31.07.2026"
  2. 'Yillar Bazli' A23        : "01.01.2026\n31.07.2026"  (alt alta)
  3. Dosya adi                 : "(01 Ocak-31 Temmuz 2026)"  (ay ADI ile)
"""

import re
import datetime

AYLAR = {
    1: "Ocak", 2: "Şubat", 3: "Mart", 4: "Nisan", 5: "Mayıs", 6: "Haziran",
    7: "Temmuz", 8: "Ağustos", 9: "Eylül", 10: "Ekim", 11: "Kasım", 12: "Aralık",
}

# Dosya adindaki ay adini sayiya cevirmek icin (aksan/buyuk-kucuk duyarsiz)
_AY_NO = {}
for _no, _ad in AYLAR.items():
    _AY_NO[_ad.lower()] = _no

GUN_DESENI = re.compile(r"^\s*(\d{1,2})[./](\d{1,2})[./](\d{4})\s*$")


class TarihHatasi(Exception):
    """Kullanicinin girdigi tarih gecersiz."""


def gun_coz(metin):
    """
    Kullanicidan gelen '07.08.2026' / '7.8.2026' / '07/08/2026' metnini
    date nesnesine cevirir. Gecersizse TarihHatasi.
    """
    if not metin or not metin.strip():
        raise TarihHatasi("Bitiş tarihi boş olamaz.")
    m = GUN_DESENI.match(metin)
    if not m:
        raise TarihHatasi(
            "Tarihi GG.AA.YYYY biçiminde yazın (örnek: 07.08.2026).\n"
            "Girilen: %r" % metin)
    g, a, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    try:
        return datetime.date(y, a, g)
    except ValueError:
        raise TarihHatasi("Böyle bir tarih yok: %s" % metin)


def nokta(d):
    """date -> '07.08.2026'"""
    return "%02d.%02d.%04d" % (d.day, d.month, d.year)


def ad_ile(d):
    """date -> '07 Ağustos 2026' (dosya adi icin)"""
    return "%02d %s %04d" % (d.day, AYLAR[d.month], d.year)


def donem_metni(bas, bit):
    """Sekme basligi bicimi: '01.01.2026 - 07.08.2026'"""
    return "%s - %s" % (nokta(bas), nokta(bit))


def donem_alt_alta(bas, bit):
    """'Yillar Bazli' bicimi: iki satir."""
    return "%s\n%s" % (nokta(bas), nokta(bit))


# --- Dosya adi --------------------------------------------------------------

# "(01 Ocak-31 Temmuz 2026)" veya "(01 Ocak - 31 Temmuz 2026)"
_AD_DESENI = re.compile(
    r"\(\s*(\d{1,2})\s+([^\s\-]+)\s*-\s*(\d{1,2})\s+([^\s\-]+)\s+(\d{4})\s*\)")


def dosya_adi_guncelle(ad, bas, bit):
    """
    'İhbar Site Rapor (01 Ocak-31 Temmuz 2026).xlsx'
      -> 'İhbar Site Rapor (01 Ocak-07 Ağustos 2026).xlsx'

    Desen bulunamazsa ad DEGISMEDEN doner (None ikinci deger).
    Doner: (yeni_ad, degisti_mi)
    """
    m = _AD_DESENI.search(ad)
    if not m:
        return ad, False

    yeni_parca = "(%s-%s %04d)" % (
        "%02d %s" % (bas.day, AYLAR[bas.month]),
        "%02d %s" % (bit.day, AYLAR[bit.month]),
        bit.year,
    )
    yeni = ad[:m.start()] + yeni_parca + ad[m.end():]
    return yeni, yeni != ad


def dosya_adindan_donem(ad):
    """Dosya adindaki donemi okur; bulunamazsa None."""
    m = _AD_DESENI.search(ad)
    if not m:
        return None
    g1, a1, g2, a2, yil = m.groups()
    n1 = _AY_NO.get(a1.lower())
    n2 = _AY_NO.get(a2.lower())
    if not n1 or not n2:
        return None
    try:
        return datetime.date(int(yil), n1, int(g1)), datetime.date(int(yil), n2, int(g2))
    except ValueError:
        return None
