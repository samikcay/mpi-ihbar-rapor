# -*- coding: utf-8 -*-
"""
Rapor sekmelerini Excel COM ile doldurur.

Neden COM? Calisma kitabindaki grafikler ve bicimler openpyxl ile
kaydedildiginde bozuluyor. Excel'in kendisini kullanarak grafikler,
kosullu bicimler ve formuller oldugu gibi kalir.

Temel yaklasim: veri blogunun satir sayisi degistiginde satirlari
blogun ICINE ekler / icinden siler. Boylece Excel formulleri
(=SUM(...)) ve grafik kaynak araliklarini kendisi genisletir.
"""

import os

import win32com.client as win32

import labels
import tarih

# Excel sabitleri (gencache olmadan da calissin diye elle taniml.)
XL_UP = -4162
XL_SHIFT_DOWN = -4121
XL_VALUES = -4163
XL_FORMULAS = -4123
XL_WHOLE = 1


class WriteError(Exception):
    """Yazma sirasinda beklenmeyen durum."""


# --- Sekme yerlesimleri -----------------------------------------------------
# ilk_satir : ilk veri satiri
# sutunlar  : (sira, ad, sayi) sutun harfleri
# Toplam satiri veri blogunun hemen altindadir ve formulle bulunur.

YERLESIM = {
    "icerik": {
        "sekme": "Site İçeriğine Göre ",
        "ilk_satir": 6,
        "sira": "A", "ad": "B", "sayi": "C", "oran": "D",
    },
    "durum": {
        "sekme": "Site Durumuna Göre",
        "ilk_satir": 6,
        "sira": "A", "ad": "B", "sayi": "C", "oran": "D",
    },
    "sube": {
        "sekme": "Şube Bazlı",
        "ilk_satir": 7,
        "sira": "A", "ad": "B", "sayi": "C", "oran": "D",
    },
}

ULKE_SEKME = "2026 Ülke Dağılımı"
ULKE_ILK_SATIR = 5

# 'Ülke Grafiği' sekmesi sabit yerlesimlidir:
#   7..21 -> ilk 15 ulke, 22 -> Diger, 23 -> TOPLAM
GRAFIK_SEKME = "2026 Ülke Grafiği"
GRAFIK_ILK_SATIR = 7
GRAFIK_ULKE_ADEDI = 15
GRAFIK_DIGER_ETIKETI = "Diğer"


class ExcelOturum(object):
    """
    Excel uygulamasini acar/kapatir; hata olsa da temizlik yapar.

    ONEMLI: DispatchEx kullanilir, Dispatch DEGIL.
    Dispatch zaten calisan bir Excel'e baglanir; bu iki soruna yol acar:
      1. Kullanicinin acik Excel'i varsa Quit() onun penceresini kapatir,
         kaydedilmemis calismasi gidebilir.
      2. O Excel'de acik bir uyari penceresi varsa program sonsuza kadar
         bekler (bizim DisplayAlerts=False ayarimiz o pencereyi kapatmaz).
    DispatchEx her zaman AYRI ve gizli bir Excel baslatir; kullanicinin
    penceresine dokunmayiz, kilitlenme riski kalkar.
    """

    def __init__(self, gorunur=False):
        self.gorunur = gorunur
        self.xl = None
        self.wb = None

    def __enter__(self):
        try:
            # AYRI bir Excel ornegi (kullanicininkine dokunma)
            self.xl = win32.DispatchEx("Excel.Application")
        except Exception:
            # Cok eski kurulumlarda DispatchEx olmayabilir
            self.xl = win32.Dispatch("Excel.Application")

        self.xl.Visible = self.gorunur
        # Uyari/onay pencereleri ACILMASIN: aksi halde kimse cevaplamaz
        # ve program asili kalir.
        for ozellik, deger in (("DisplayAlerts", False),
                               ("ScreenUpdating", False),
                               ("EnableEvents", False),
                               ("AskToUpdateLinks", False),
                               ("AlertBeforeOverwriting", False),
                               ("DisplayStatusBar", False)):
            try:
                setattr(self.xl, ozellik, deger)
            except Exception:
                pass  # bu Excel surumunde yoksa onemli degil
        return self

    def ac(self, yol):
        """
        Calisma kitabini acar.

        UpdateLinks=0  -> dis baglanti guncelleme sorusu sorulmaz
        ReadOnly=False -> yazmak icin aciyoruz
        CorruptLoad    -> bozuk dosyada Excel onarim penceresi acmasin
        """
        self.wb = self.xl.Workbooks.Open(
            os.path.abspath(yol), UpdateLinks=0, ReadOnly=False)
        return self.wb

    def __exit__(self, *exc):
        # Sirayla: kitabi kapat -> tum kitaplari kapat -> uygulamayi kapat.
        # Her adim ayri korunur ki biri takilirsa digerleri yine calissin.
        try:
            if self.wb is not None:
                self.wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if self.xl is not None:
                # Artakalan kitap varsa kapat (yoksa Quit takilabilir)
                try:
                    while self.xl.Workbooks.Count:
                        self.xl.Workbooks(1).Close(SaveChanges=False)
                except Exception:
                    pass
                self.xl.Quit()
        except Exception:
            pass
        self.wb = None
        self.xl = None
        return False


def _sekme_bul(wb, ad):
    """Sekmeyi adiyla bulur; bosluk/aksan farkini tolere eder."""
    hedef = labels.anahtar(ad)
    for i in range(1, wb.Worksheets.Count + 1):
        ws = wb.Worksheets(i)
        if labels.anahtar(ws.Name) == hedef:
            return ws
    mevcut = [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]
    raise WriteError("Sekme bulunamadi: %r\nMevcut sekmeler: %s" % (ad, mevcut))


def _toplam_satiri(ws, ilk_satir, sayi_sutunu):
    """
    Veri blogunun altindaki TOPLAM satirini bulur.
    =SUM(...) iceren ilk satiri arar; bulamazsa hata verir.
    """
    for r in range(ilk_satir, ilk_satir + 500):
        f = ws.Cells(r, _sutun_no(sayi_sutunu)).Formula or ""
        if f.upper().startswith("=SUM(") or f.upper().startswith("=TOPLA("):
            return r
    raise WriteError("TOPLAM satiri (=SUM) bulunamadi")


def _sutun_no(harf):
    n = 0
    for ch in harf.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _blok_boyutunu_ayarla(ws, ilk_satir, son_satir, hedef_adet, ornek_satir=None):
    """
    Veri blogunu hedef satir sayisina getirir.

    Satirlari blogun ICINE ekler/siler ki Excel formulleri ve grafik
    araliklarini kendisi guncellesin. Yeni satirlar ornek satirdan
    bicim ve formul kopyalar.

    Doner: yeni son_satir
    """
    mevcut = son_satir - ilk_satir + 1
    if hedef_adet < 1:
        raise WriteError("Veri bulunamadi (0 satir yazilamaz)")

    if hedef_adet > mevcut:
        ekle = hedef_adet - mevcut
        # Sondan bir onceki satirin altina ekle -> blogun icinde kalir,
        # boylece SUM ve grafik araliklari otomatik genisler.
        hedef = son_satir
        ws.Rows("%d:%d" % (hedef, hedef + ekle - 1)).Insert(Shift=XL_SHIFT_DOWN)
        # Bicim/formul icin ornek satiri kopyala
        kaynak = ornek_satir if ornek_satir else ilk_satir
        ws.Rows(kaynak).Copy()
        ws.Rows("%d:%d" % (hedef, hedef + ekle - 1)).PasteSpecial(Paste=-4104)  # xlPasteAll
        ws.Application.CutCopyMode = False
        son_satir += ekle

    elif hedef_adet < mevcut:
        sil = mevcut - hedef_adet
        # Blogun icinden sil (en alttaki veri satirlarindan)
        ws.Rows("%d:%d" % (son_satir - sil + 1, son_satir)).Delete()
        son_satir -= sil

    return son_satir


def _mevcut_etiketler(ws, sutun, ilk_satir, son_satir):
    """Sekmedeki mevcut ad yazimlarini toplar (yazim korumak icin)."""
    out = []
    col = _sutun_no(sutun)
    for r in range(ilk_satir, son_satir + 1):
        v = ws.Cells(r, col).Value
        if v:
            out.append(str(v))
    return out


def yaz_basit(wb, anahtar_ad, kayitlar, log=None):
    """
    Tek sutunlu sekmeleri doldurur (icerik / durum / sube).

    kayitlar: [(ad, sayi), ...]
    """
    y = YERLESIM[anahtar_ad]
    ws = _sekme_bul(wb, y["sekme"])
    ilk = y["ilk_satir"]

    toplam_satir = _toplam_satiri(ws, ilk, y["sayi"])
    son = toplam_satir - 1

    onceki_etiketler = _mevcut_etiketler(ws, y["ad"], ilk, son)

    son = _blok_boyutunu_ayarla(ws, ilk, son, len(kayitlar), ornek_satir=ilk)

    for i, (ad, sayi) in enumerate(kayitlar):
        r = ilk + i
        ws.Cells(r, _sutun_no(y["sira"])).Value = i + 1
        ws.Cells(r, _sutun_no(y["ad"])).Value = labels.duzelt(ad, onceki_etiketler)
        ws.Cells(r, _sutun_no(y["sayi"])).Value = sayi

    if log:
        log("  %-22s %2d satir yazildi" % (y["sekme"].strip(), len(kayitlar)))
    return len(kayitlar)


def yaz_ulkeler(wb, kayitlar, sekme_adi=ULKE_SEKME, log=None):
    """
    Ulke sekmesi iki sutun blogu halindedir:
      sol  (A,B,C) : 1 .. ceil(n/2)
      sag  (D,E,F) : ceil(n/2)+1 .. n
    Satir sayisi sol bloga gore belirlenir.
    """
    ws = _sekme_bul(wb, sekme_adi)
    ilk = ULKE_ILK_SATIR

    n = len(kayitlar)
    sol_adet = (n + 1) // 2  # tek sayida ise fazlalik sol tarafta

    # Mevcut blok sonunu bul: A sutununda sira numarasi olan son satir
    son = ilk
    r = ilk
    while r < ilk + 1000:
        v = ws.Cells(r, 1).Value
        if isinstance(v, (int, float)) and float(v).is_integer() and v > 0:
            son = r
            r += 1
        else:
            break

    onceki = _mevcut_etiketler(ws, "B", ilk, son) + _mevcut_etiketler(ws, "E", ilk, son)

    son = _blok_boyutunu_ayarla(ws, ilk, son, sol_adet, ornek_satir=ilk)

    # Once her iki blogu da temizle (eski uzun listeden kalinti kalmasin)
    ws.Range(ws.Cells(ilk, 1), ws.Cells(son, 6)).ClearContents()

    for i in range(sol_adet):
        r = ilk + i
        ad, sayi = kayitlar[i]
        ws.Cells(r, 1).Value = i + 1
        ws.Cells(r, 2).Value = labels.duzelt(ad, onceki)
        ws.Cells(r, 3).Value = sayi

        j = sol_adet + i
        if j < n:
            ad2, sayi2 = kayitlar[j]
            ws.Cells(r, 4).Value = j + 1
            ws.Cells(r, 5).Value = labels.duzelt(ad2, onceki)
            ws.Cells(r, 6).Value = sayi2

    # TOPLAM satiri: blogun hemen altinda, F sutununda
    toplam_satir = son + 1
    ws.Cells(toplam_satir, 6).Value = sum(v for _, v in kayitlar)

    if log:
        log("  %-22s %2d ulke (%d sol / %d sag)"
            % (sekme_adi, n, sol_adet, n - sol_adet))
    return n


def yaz_ulke_grafigi(wb, kayitlar, sekme_adi=GRAFIK_SEKME, log=None):
    """
    'Ülke Grafiği' sekmesini 'Ülke Dağılımı' verisinden uretir.

    Ilk 15 ulke aynen yazilir, 16. satira "Diğer" olarak
    (genel toplam - ilk 15 toplami) yazilir. Yerlesim sabittir
    (7..21 ulke, 22 Diger, 23 TOPLAM) - satir eklenmez/silinmez.
    """
    ws = _sekme_bul(wb, sekme_adi)
    ilk = GRAFIK_ILK_SATIR

    genel_toplam = sum(v for _, v in kayitlar)
    ilk_n = kayitlar[:GRAFIK_ULKE_ADEDI]
    diger = genel_toplam - sum(v for _, v in ilk_n)

    if diger < 0:
        raise WriteError("'Diğer' degeri negatif cikti (%d)" % diger)

    onceki = _mevcut_etiketler(ws, "B", ilk, ilk + GRAFIK_ULKE_ADEDI)

    for i in range(GRAFIK_ULKE_ADEDI):
        r = ilk + i
        if i < len(ilk_n):
            ad, sayi = ilk_n[i]
            ws.Cells(r, 1).Value = i + 1
            ws.Cells(r, 2).Value = labels.duzelt(ad, onceki)
            ws.Cells(r, 3).Value = sayi
        else:
            # 15'ten az ulke varsa kalan satirlari bosalt
            ws.Cells(r, 1).ClearContents()
            ws.Cells(r, 2).ClearContents()
            ws.Cells(r, 3).ClearContents()

    diger_satir = ilk + GRAFIK_ULKE_ADEDI
    ws.Cells(diger_satir, 1).Value = len(ilk_n) + 1
    ws.Cells(diger_satir, 2).Value = GRAFIK_DIGER_ETIKETI
    ws.Cells(diger_satir, 3).Value = diger

    if log:
        log("  %-22s ilk %d ulke + Diğer (%s)"
            % (sekme_adi, len(ilk_n), "{:,}".format(diger).replace(",", ".")))
    return diger


def donem_yaz(wb, bas, bit, log=None):
    """
    Rapor donemini ilgili sekmelerde gunceller.

    SADECE cari donem sekmelerine dokunur. Gecmis yil sekmeleri
    (2023/2024/2025) ve 'Suç Duyuruları' gibi tarihsel kayitlar
    ELLENMEZ - oradaki tarihler o donemlere aittir.
    """
    import re
    desen = re.compile(r"\d{2}\.\d{2}\.\d{4}\s*-\s*\d{2}\.\d{2}\.\d{4}")
    yeni = tarih.donem_metni(bas, bit)

    sayac = 0
    hedefler = [YERLESIM[k]["sekme"] for k in YERLESIM] + [ULKE_SEKME, GRAFIK_SEKME]
    for ad in hedefler:
        try:
            ws = _sekme_bul(wb, ad)
        except WriteError:
            continue
        for r in range(1, 7):
            v = ws.Cells(r, 1).Value
            if isinstance(v, str) and desen.search(v):
                ws.Cells(r, 1).Value = yeni
                sayac += 1
                break

    # 'Yıllar Bazlı' sekmesinde cari yil satiri alt alta iki tarih tasir.
    yil_sayac = _yillar_bazli_donem(wb, bas, bit)

    if log:
        log("  Donem guncellendi: %s (%d sekme%s)"
            % (yeni, sayac, " + Yıllar Bazlı" if yil_sayac else ""))
    return sayac


# Yillar Bazli sekmesinde cari yil satirinin SAP'tan gelen sutunlari.
# E-I sutunlari (erisime engellenen, EEK, mahkeme karari, pasif/domain,
# islemleri devam eden) SAP ciktilarinda YOKTUR - elle girilir, ELLENMEZ.
YILLAR_SAP_SUTUNLARI = {
    "B": "toplam",   # BTK'ya ihbar edilen toplam site sayisi
    "C": "direk",    # direk oynatan
    "D": "reklam",   # tanitim, reklam ve yonlendirme
}


def yillar_bazli_degerler(veri):
    """SAP verisinden B/C/D sutunlarinin degerlerini hesaplar."""
    icerik = dict(veri["icerik"]["items"])
    ulkeler = veri["ulke"]["items"]

    def sade(x):
        return (x or "").lower().replace("ı", "i").replace("İ", "i") \
            .replace("ş", "s").replace("ğ", "g").replace("ü", "u") \
            .replace("ö", "o").replace("ç", "c")

    def bul(*anahtarlar):
        for k, v in icerik.items():
            for a in anahtarlar:
                if a in sade(k):
                    return v
        return None

    return {
        "toplam": sum(v for _, v in ulkeler),
        "direk": bul("direk"),
        "reklam": bul("tanitim", "reklam"),
    }


def yaz_yillar_bazli(wb, veri, bas, log=None):
    """
    'Yillar Bazli' sekmesindeki CARI YIL satirinin SAP'tan gelen
    sutunlarini (B/C/D) gunceller.

    Bu satir kumulatif TOPLAM formullerini besler (B25 = toplam tespit,
    E25 = erisime engellenen); Word raporundaki "2006-2026 doneminde
    toplam ..." cumlesi de dolayli olarak buradan gelir. Bu yuzden bu
    satir guncel degilse kumulatif sayilar da eski kalir.

    E-I sutunlari SAP'ta bulunmadigi icin DEGISTIRILMEZ.
    Doner: (satir_no, {sutun: (eski, yeni)})
    """
    import re
    try:
        ws = _sekme_bul(wb, "Yıllar Bazlı")
    except WriteError:
        return None, {}

    degerler = yillar_bazli_degerler(veri)

    # Cari yil satiri: A sutununda '01.01.2026\n17.08.2026' bicimi
    desen = re.compile(r"\d{2}\.\d{2}\.(\d{4})\s*\n")
    hedef = None
    for r in range(1, 60):
        v = ws.Cells(r, 1).Value
        if isinstance(v, str):
            m = desen.search(v)
            if m and m.group(1) == str(bas.year):
                hedef = r
                break
    if hedef is None:
        return None, {}

    degisim = {}
    for harf, anahtar in sorted(YILLAR_SAP_SUTUNLARI.items()):
        yeni_deger = degerler.get(anahtar)
        if yeni_deger is None:
            continue
        col = _sutun_no(harf)
        eski_deger = ws.Cells(hedef, col).Value
        if eski_deger is not None and int(eski_deger) == int(yeni_deger):
            continue
        ws.Cells(hedef, col).Value = yeni_deger
        degisim[harf] = (eski_deger, yeni_deger)

    if log and degisim:
        log("  %-22s satir %d guncellendi (%s)"
            % ("Yıllar Bazlı", hedef,
               ", ".join("%s: %s -> %s" % (h, int(e or 0), int(y))
                         for h, (e, y) in sorted(degisim.items()))))
    return hedef, degisim


def _yillar_bazli_donem(wb, bas, bit):
    """
    'Yıllar Bazlı' sekmesindeki cari donem satirini gunceller.
    Hucre '01.01.2026\\n31.07.2026' bicimindedir; sadece ayni yila ait
    olan satir degistirilir (gecmis yil satirlari '2025' gibi metindir).
    """
    import re
    try:
        ws = _sekme_bul(wb, "Yıllar Bazlı")
    except WriteError:
        return 0

    desen = re.compile(r"(\d{2}\.\d{2}\.(\d{4}))\s*\n\s*(\d{2}\.\d{2}\.\d{4})")
    for r in range(1, 40):
        v = ws.Cells(r, 1).Value
        if isinstance(v, str):
            m = desen.search(v)
            if m and m.group(2) == str(bas.year):
                ws.Cells(r, 1).Value = tarih.donem_alt_alta(bas, bit)
                return 1
    return 0
