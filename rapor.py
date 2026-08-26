# -*- coding: utf-8 -*-
"""
Ihbar Site Raporu - SAP ciktilarini Excel'e otomatik yazar.

Kullanim:
    python rapor.py                 -> pencere acilir
    python rapor.py --konsol ...    -> komut satirindan calisir

Program calisma kitabini degistirmeden once yedek alir; hata olursa
orijinal dosya bozulmaz.
"""

import os
import sys
import shutil
import argparse
import datetime
import traceback

import sap_parser
import excel_writer
import tarih
import word_writer
import hatalar
from hatalar import Kritik, Iptal
from excel_writer import ExcelOturum, WriteError
from sap_parser import ParseError
from tarih import TarihHatasi


# Rapor anahtari -> Excel sekmesi (UI ve gunlukte gorunen ad)
RAPOR_ADLARI = {
    "icerik": "Site Icerigine Gore",
    "durum": "Site Durumuna Gore",
    "sube": "Sube Bazli",
    "ulke": "Ulke Dagilimi + Grafigi",
}

# SAP dosya adlarini tanimak icin ipuclari (kucuk harf, aksansiz)
DOSYA_IPUCLARI = {
    "icerik": ["icerik", "içerik", "iceri"],
    "durum": ["durum", "sitedurum"],
    "sube": ["sube", "şube"],
    "ulke": ["ulke", "ülke"],
}


def _ad_anahtari(yol):
    ad = os.path.basename(yol).lower()
    return (ad.replace("ı", "i").replace("ş", "s").replace("ü", "u")
              .replace("ö", "o").replace("ç", "c").replace("ğ", "g"))


def klasoru_tara(klasor):
    """
    Klasordeki .txt dosyalarini icerigine gore siniflandirir.
    Dosya adi ipucu vermezse basliktan anlar.
    Doner: {'icerik': yol, 'durum': yol, 'sube': yol, 'ulke': yol}
    """
    bulunan = {}
    if not os.path.isdir(klasor):
        raise ParseError("Klasor bulunamadi: %s" % klasor)

    dosyalar = [os.path.join(klasor, f) for f in os.listdir(klasor)
                if f.lower().endswith(".txt")]
    if not dosyalar:
        raise ParseError("Klasorde .txt dosyasi yok: %s" % klasor)

    for yol in dosyalar:
        ad = _ad_anahtari(yol)
        tur = None
        for anahtar, ipuclari in DOSYA_IPUCLARI.items():
            if any(ip in ad for ip in ipuclari):
                tur = anahtar
                break
        if tur is None:
            tur = _basliktan_tur_bul(yol)
        if tur and tur not in bulunan:
            bulunan[tur] = yol

    return bulunan


def _basliktan_tur_bul(yol):
    """Dosya adi ipucu vermezse SAP basligina bakar."""
    try:
        satirlar = sap_parser.read_sap_text(yol)
    except ParseError:
        return None
    bas = " ".join(satirlar[:8]).upper()
    bas = (bas.replace("İ", "I").replace("Ş", "S").replace("Ğ", "G")
              .replace("Ü", "U").replace("Ö", "O").replace("Ç", "C"))
    if "ICERIK" in bas:
        return "icerik"
    if "DURUM" in bas:
        return "durum"
    if "SUBE" in bas:
        return "sube"
    if "ULKE" in bas:
        return "ulke"
    return None


def yedek_al(xlsx_yolu):
    """Calisma kitabinin zaman damgali yedegini alir, yedek yolunu doner."""
    klasor = os.path.join(os.path.dirname(os.path.abspath(xlsx_yolu)), "yedek")
    if not os.path.isdir(klasor):
        os.makedirs(klasor)
    damga = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    ad = os.path.basename(xlsx_yolu)
    kok, uzanti = os.path.splitext(ad)
    hedef = os.path.join(klasor, "%s_%s%s" % (kok, damga, uzanti))
    shutil.copy2(xlsx_yolu, hedef)
    return hedef


def dosyalari_dogrula(dosyalar):
    """
    Secilen dosyalarin gercekten o rapora ait olup olmadigini icerigindeki
    BASLIKTAN kontrol eder.

    Dosya adi her seferinde degisebildigi icin ada guvenilmez; baslik okunur.
    Uyusmazlik engel degil UYARIDIR - SAP basligi degismis olabilir.

    Doner: uyari listesi
    """
    uyarilar = []
    for anahtar in ("icerik", "durum", "sube", "ulke"):
        yol = dosyalar.get(anahtar)
        if not yol:
            continue
        try:
            bulunan = _basliktan_tur_bul(yol)
        except Exception:
            bulunan = None
        if bulunan and bulunan != anahtar:
            uyarilar.append(
                "'%s' icin secilen dosya (%s) icerik olarak '%s' raporuna benziyor."
                % (RAPOR_ADLARI.get(anahtar, anahtar),
                   os.path.basename(yol),
                   RAPOR_ADLARI.get(bulunan, bulunan)))
    return uyarilar


GENISLIK = 74


def _baslik(metin):
    """'=== EXCEL RAPORU ===' seklinde bolum basligi."""
    return "%s %s" % ("=" * 3, metin) + " " + "=" * max(0, GENISLIK - len(metin) - 5)


def _ozet_yaz(log, xlsx_yolu, word_yolu, word_ozet, yb_satir, yb_degisim,
              bas, bit, uyarilar, basarili=None, basarisiz=None,
              atlanan=None, sorun=None):
    """
    Calisma sonunda tek bakista gorulen ozet:
    NE guncellendi, NE guncellenmedi, nereye bakilmali.
    """
    log("")
    log(_baslik("OZET"))

    # --- Guncellenenler ---
    log("")
    log("  GUNCELLENDI")
    log("  " + "-" * (GENISLIK - 2))
    log("    Excel: %s" % os.path.basename(xlsx_yolu))
    yazilanlar = basarili if basarili is not None else [
        "Site İçeriğine Göre", "Site Durumuna Göre", "Şube Bazlı",
        excel_writer.ULKE_SEKME, excel_writer.GRAFIK_SEKME]
    for ad in yazilanlar:
        log("      - %s" % ad)
    if yb_satir and yb_degisim:
        sutunlar = ", ".join(sorted(yb_degisim))
        log("      - Yıllar Bazlı (satır %d: %s sütunları)" % (yb_satir, sutunlar))
    elif yb_satir:
        # Degerler zaten SAP ile ayni -> "atlandi" gibi gorunmesin
        log("      - Yıllar Bazlı (satır %d: değerler zaten güncel)" % yb_satir)
    else:
        log("      - Yıllar Bazlı  ! cari yıl satırı BULUNAMADI")
    log("      - Tüm cari sekmelerdeki dönem tarihi (%s)"
        % tarih.donem_metni(bas, bit))

    if word_yolu and word_ozet:
        log("")
        log("    Word: %s" % os.path.basename(word_yolu))
        n = len(word_ozet.get("degisen") or [])
        if n:
            log("      - %d paragraf (tarihler, site sayısı, ülke yüzdeleri," % n)
            log("        Excel'den gelen kümülatif ve dava sayıları)")
        else:
            log("      - (değişiklik gerekmedi)")

    # --- Guncellenmeyenler ---
    log("")
    log("  GUNCELLENMEDI  -  elle kontrol edin")
    log("  " + "-" * (GENISLIK - 2))

    if basarisiz:
        log("    ! YAZILAMAYAN sekmeler (hata olustu):")
        for ad in basarisiz:
            log("        - %s" % ad)
    if atlanan:
        log("    ! ATLANAN sekmeler (kaynak veri yok):")
        for ad in atlanan:
            log("        - %s" % ad)

    log("    Excel:")
    log("      - Yıllar Bazlı > E23-I23 (erişime engellenen, EEK, mahkeme,")
    log("        pasif/domain, devam eden)  -> SAP çıktısında yok")
    log("      - Suç Duyuruları > yeni dönem satırı  -> elle eklenir")
    log("      - Mobil Uygulamalar / Sosyal Medya / 850'li Hatlar /")
    log("        Ödeme Kuruluşları > B-C sütunları  -> dava sonucu, elle")
    log("      - Faaliyet Cetveli  -> program bu sekmeye hiç dokunmaz")
    log("      - Geçmiş yıl sekmeleri (2023/2024/2025)  -> dokunulmaz")

    if word_yolu and word_ozet:
        elle = word_ozet.get("elle") or []
        eksik = word_ozet.get("excel_eksik") or []
        if elle or eksik:
            log("    Word:")
            for aciklama, nerede in elle:
                log("      - %s" % aciklama)
                log("          duzeltme yeri: %s" % nerede)
            for aciklama, sebep in eksik:
                log("      ! %s  -> ESKI DEGER KALDI" % aciklama)
                log("          sebep: %s" % sebep)

    tum_uyarilar = list(uyarilar or [])
    if sorun is not None:
        for onem, baslik, ayrinti, nerede in sorun.liste():
            if onem == hatalar.BILGI:
                continue
            tum_uyarilar.append(
                "%s%s" % (baslik, (" - %s" % ayrinti) if ayrinti else ""))

    if tum_uyarilar:
        log("")
        log("  UYARILAR")
        log("  " + "-" * (GENISLIK - 2))
        for u in tum_uyarilar:
            for i, satir in enumerate(str(u).splitlines()):
                log("    %s %s" % ("!" if i == 0 else " ", satir.strip()))

    log("")
    log("=" * GENISLIK)


def calistir(xlsx_yolu, sap_klasoru=None, bitis_tarihi=None, log=print,
             ulke_sekmesi=None, grafik_sekmesi=None, yeniden_adlandir=True,
             sap_dosyalari=None, word_yolu=None, sorucu=None):
    """
    Ana is akisi: SAP dosyalarini oku -> dogrula -> Excel'e yaz.

    sap_klasoru   : klasor yolu (dosyalar otomatik taninir)
    sap_dosyalari : {'icerik': yol, 'durum': yol, 'sube': yol, 'ulke': yol}
                    Verilirse klasor taramasi YAPILMAZ.
    bitis_tarihi  : date veya 'GG.AA.YYYY' metni. None ise SAP'taki tarih.
    yeniden_adlandir : True ise dosya yeni donem adiyla KOPYALANIR
                       (orijinal dosya yerinde kalir).

    Doner: ozet sozlugu
    """
    sorun = hatalar.Rapor(log=log, sorucu=sorucu)

    if not xlsx_yolu or not os.path.isfile(xlsx_yolu):
        raise ParseError(
            "Excel raporu bulunamadi: %s\n"
            "Dosya tasinmis, adi degismis veya silinmis olabilir." % xlsx_yolu)

    if sap_dosyalari:
        dosyalar = dict(sap_dosyalari)
        eksik = [k for k in ("icerik", "durum", "sube", "ulke")
                 if not dosyalar.get(k)]
        if eksik:
            raise ParseError(
                "Su raporlar icin dosya secilmedi: %s"
                % ", ".join(RAPOR_ADLARI.get(k, k) for k in eksik))
        for k in ("icerik", "durum", "sube", "ulke"):
            if not os.path.isfile(dosyalar[k]):
                raise ParseError("Dosya bulunamadi: %s" % dosyalar[k])
        log("Secilen SAP dosyalari:")
    else:
        if not sap_klasoru:
            raise ParseError("SAP klasoru veya dosyalari verilmedi.")
        log("SAP klasoru taraniyor: %s" % sap_klasoru)
        dosyalar = klasoru_tara(sap_klasoru)

        eksik = [k for k in ("icerik", "durum", "sube", "ulke") if k not in dosyalar]
        if eksik:
            raise ParseError(
                "Su raporlarin dosyasi bulunamadi: %s\n"
                "Klasorde her rapor icin bir .txt olmali."
                % ", ".join(RAPOR_ADLARI.get(k, k) for k in eksik))

    for k in ("icerik", "durum", "sube", "ulke"):
        log("  %-22s -> %s" % (RAPOR_ADLARI.get(k, k),
                               os.path.basename(dosyalar[k])))

    # Secilen dosya gercekten o rapor mu? (ada degil, icerige bakilir)
    eslesme_uyarilari = dosyalari_dogrula(dosyalar)

    # --- Oku ---
    # Her dosya BAGIMSIZ okunur: biri bozuksa hangisi oldugu soylenir,
    # digerleri okunmaya devam eder.
    log("\nSAP ciktilari okunuyor...")
    okuyucular = {
        "icerik": sap_parser.parse_simple,
        "durum": sap_parser.parse_simple,
        "sube": sap_parser.parse_simple,
        "ulke": sap_parser.parse_countries,
    }
    veri, okunamayan = {}, []
    for k in ("icerik", "durum", "sube", "ulke"):
        try:
            veri[k] = okuyucular[k](dosyalar[k])
        except Exception as e:
            okunamayan.append(k)
            sorun.uyar("'%s' dosyasi okunamadi" % RAPOR_ADLARI.get(k, k),
                       "%s - %s" % (os.path.basename(dosyalar[k]),
                                    hatalar._kisa_hata(e)),
                       nerede=RAPOR_ADLARI.get(k, k))

    if okunamayan:
        # Ulke dosyasi hem Dagilim hem Grafik sekmesini besler; onsuz
        # anlamli bir rapor cikmaz.
        if "ulke" in okunamayan:
            raise Kritik(
                "Ulke dosyasi okunamadi; Ulke Dagilimi ve Grafik sekmeleri\n"
                "yazilamaz. Dosyayi kontrol edip tekrar deneyin.")
        if not sorun.sor(
                "%d SAP dosyasi okunamadi" % len(okunamayan),
                "ilgili sekmeler ESKI degerleriyle kalacak",
                varsayilan=True):
            raise Iptal("Okunamayan dosyalar nedeniyle islem iptal edildi.")

    # --- Dogrula ---
    uyarilar = list(eslesme_uyarilari)
    toplamlar = {}
    for k, d in sorted(veri.items()):
        s = sum(v for _, v in d["items"])
        toplamlar[k] = s
        if d["total"] is not None and d["total"] != s:
            uyarilar.append(
                "%s: satir toplami %d, SAP TOPLAM satiri %d (fark %d)"
                % (k, s, d["total"], s - d["total"]))
        log("  %-22s %3d satir, toplam %s"
            % (RAPOR_ADLARI.get(k, k), len(d["items"]),
               "{:,}".format(s).replace(",", ".")))

    farkli = set(toplamlar.values())
    if len(farkli) > 1:
        uyarilar.append(
            "Raporlarin genel toplamlari birbirini tutmuyor: %s\n"
            "  (Ayni doneme ait olmayan dosyalar karismis olabilir.)"
            % ", ".join("%s=%d" % kv for kv in sorted(toplamlar.items())))

    donemler = set(d["period"] for d in veri.values() if d["period"])
    if len(donemler) > 1:
        uyarilar.append("Dosyalarin donemleri farkli: %s" % sorted(donemler))
    sap_donem = sorted(donemler)[0] if donemler else None

    # --- Donem tarihleri ---
    # Baslangic: SAP'taki baslangic (yoksa dosya adindan, o da yoksa 1 Ocak)
    ad_donem = tarih.dosya_adindan_donem(os.path.basename(xlsx_yolu))
    if sap_donem:
        bas = tarih.gun_coz(sap_donem[0])
    elif ad_donem:
        bas = ad_donem[0]
    else:
        raise ParseError("Rapor donemi belirlenemedi (SAP basliginda tarih yok).")

    # Bitis: kullanici verdiyse o, yoksa SAP'taki
    if bitis_tarihi is None:
        if not sap_donem:
            raise ParseError("Bitis tarihi verilmedi ve SAP basliginda tarih yok.")
        bit = tarih.gun_coz(sap_donem[1])
    elif isinstance(bitis_tarihi, str):
        bit = tarih.gun_coz(bitis_tarihi)
    else:
        bit = bitis_tarihi

    if bit < bas:
        raise TarihHatasi(
            "Bitiş tarihi (%s) başlangıçtan (%s) önce olamaz."
            % (tarih.nokta(bit), tarih.nokta(bas)))

    if sap_donem:
        sap_bit = tarih.gun_coz(sap_donem[1])
        if bit != sap_bit:
            uyarilar.append(
                "Girilen bitiş tarihi (%s), SAP çıktısındaki tarihten (%s) farklı."
                % (tarih.nokta(bit), tarih.nokta(sap_bit)))

    log("\nRapor donemi: %s" % tarih.donem_metni(bas, bit))

    if uyarilar:
        log("\nUYARI:")
        for u in uyarilar:
            log("  ! %s" % u)

    # --- Hedef dosya ---
    # Yeni donem adiyla KOPYA olusturulur; orijinal dosya yerinde kalir.
    hedef_yolu = xlsx_yolu
    yeni_ad_bilgisi = None
    if yeniden_adlandir:
        klasor = os.path.dirname(os.path.abspath(xlsx_yolu))
        eski_ad = os.path.basename(xlsx_yolu)
        yeni_ad, degisti = tarih.dosya_adi_guncelle(eski_ad, bas, bit)
        if degisti:
            hedef_yolu = os.path.join(klasor, yeni_ad)
            if os.path.abspath(hedef_yolu) != os.path.abspath(xlsx_yolu):
                shutil.copy2(xlsx_yolu, hedef_yolu)
                yeni_ad_bilgisi = yeni_ad
                log("\nYeni dosya olusturuldu: %s" % yeni_ad)
                log("  (orijinal dosya yerinde birakildi)")

    # --- Yedek ---
    # Uzerine yazilacak dosyanin yedegi alinir.
    yedek = yedek_al(hedef_yolu)
    log("\nYedek alindi: %s" % os.path.basename(yedek))

    # --- Yaz ---
    log("")
    log(_baslik("EXCEL RAPORU"))
    yb_satir, yb_degisim = None, {}
    basarili, basarisiz, atlanan = [], [], []

    with ExcelOturum() as oturum:
        try:
            wb = oturum.ac(hedef_yolu)
        except Exception as e:
            raise Kritik(
                "Excel dosyasi acilamadi: %s\n"
                "Dosya baska bir programda acik olabilir; kapatip tekrar deneyin."
                % hatalar._kisa_hata(e))

        # Her sekme BAGIMSIZ yazilir: biri hata verirse digerleri yazilmaya
        # devam eder, kullaniciya sorulur.
        # (anahtar, sekme_adi, islem) - anahtar okunamadiysa adim atlanir
        adimlar = [
            ("icerik", "Site İçeriğine Göre", lambda:
                excel_writer.yaz_basit(wb, "icerik", veri["icerik"]["items"], log=log)),
            ("durum", "Site Durumuna Göre", lambda:
                excel_writer.yaz_basit(wb, "durum", veri["durum"]["items"], log=log)),
            ("sube", "Şube Bazlı", lambda:
                excel_writer.yaz_basit(wb, "sube", veri["sube"]["items"], log=log)),
            ("ulke", ulke_sekmesi or excel_writer.ULKE_SEKME, lambda:
                excel_writer.yaz_ulkeler(
                    wb, veri["ulke"]["items"],
                    sekme_adi=ulke_sekmesi or excel_writer.ULKE_SEKME, log=log)),
            ("ulke", grafik_sekmesi or excel_writer.GRAFIK_SEKME, lambda:
                excel_writer.yaz_ulke_grafigi(
                    wb, veri["ulke"]["items"],
                    sekme_adi=grafik_sekmesi or excel_writer.GRAFIK_SEKME, log=log)),
        ]

        for anahtar, ad, islem in adimlar:
            if anahtar not in veri:
                atlanan.append(ad)
                log("  %-22s ATLANDI (SAP dosyasi okunamadi)" % ad)
                continue
            try:
                islem()
                basarili.append(ad)
            except Exception as e:
                basarisiz.append(ad)
                if not sorun.sor(
                        "'%s' sekmesi yazilamadi" % ad,
                        hatalar._kisa_hata(e),
                        nerede=ad,
                        varsayilan=True):
                    raise Iptal(
                        "'%s' sekmesindeki hata nedeniyle islem iptal edildi."
                        % ad)

        # Yillar Bazli (kumulatif zinciri besler) - kaynak veri sartli
        try:
            if "ulke" not in veri or "icerik" not in veri:
                raise Exception("kaynak SAP verisi eksik")
            yb_satir, yb_degisim = excel_writer.yaz_yillar_bazli(
                wb, veri, bas, log=log)
            if yb_satir is None:
                sorun.uyar("Yıllar Bazlı cari yil satiri bulunamadi",
                           "kumulatif toplamlar ESKI kalacak",
                           nerede="Yıllar Bazlı")
        except Exception as e:
            basarisiz.append("Yıllar Bazlı")
            if not sorun.sor("Yıllar Bazlı guncellenemedi",
                             hatalar._kisa_hata(e),
                             nerede="Yıllar Bazlı", varsayilan=True):
                raise Iptal("Yıllar Bazlı hatasi nedeniyle islem iptal edildi.")

        # Donem tarihleri
        try:
            excel_writer.donem_yaz(wb, bas, bit, log=log)
        except Exception as e:
            sorun.uyar("Donem tarihi guncellenemedi", hatalar._kisa_hata(e))

        if not basarili:
            raise Kritik(
                "Hicbir sekme yazilamadi; dosya DEGISTIRILMEDI.\n"
                "Calisma kitabi beklenen yapida olmayabilir.")

        try:
            wb.Application.CalculateFull()
        except Exception as e:
            sorun.uyar("Formuller yeniden hesaplanamadi", hatalar._kisa_hata(e))

        try:
            wb.Save()
        except Exception as e:
            raise Kritik(
                "Dosya KAYDEDILEMEDI: %s\n"
                "Degisiklikler yazilmadi; yedek dosyaniz duruyor:\n  %s"
                % (hatalar._kisa_hata(e), yedek))

        log("\n  Kaydedildi: %s" % os.path.basename(hedef_yolu))
        if basarisiz:
            log("  ! Yazilamayan sekmeler: %s" % ", ".join(basarisiz))

    # --- Word raporu (istege bagli) ---
    word_ozet = None
    if word_yolu and ("ulke" not in veri or "icerik" not in veri):
        sorun.uyar("Word raporu guncellenmedi",
                   "SAP verisi eksik oldugu icin atlandi",
                   nerede="Word")
    elif word_yolu:
        log("")
        log(_baslik("WORD RAPORU"))
        # Word hatasi EXCEL'i gecersiz kilmaz: Excel zaten kaydedildi.
        try:
            word_ozet = word_writer.guncelle(word_yolu, veri, bas, bit,
                                             xlsx_yolu=hedef_yolu, log=log)
            log("\n  Yedek: %s" % os.path.basename(word_ozet["yedek"]))
            log("  Kaydedildi: %s" % os.path.basename(word_yolu))
        except Exception as e:
            word_ozet = None
            sorun.uyar("Word raporu guncellenemedi",
                       hatalar._kisa_hata(e), nerede="Word")
            log("  ! Word dosyasi DEGISTIRILMEDI; Excel raporu hazir.")

    # --- Kapanis ozeti: NE GUNCELLENDI / NE GUNCELLENMEDI ---
    _ozet_yaz(log, hedef_yolu, word_yolu, word_ozet, yb_satir, yb_degisim,
              bas, bit, uyarilar, basarili, basarisiz, atlanan, sorun)

    return {
        "toplamlar": toplamlar,
        "uyarilar": uyarilar,
        "donem": (bas, bit),
        "yedek": yedek,
        "hedef": hedef_yolu,
        "yeni_ad": yeni_ad_bilgisi,
        "word": word_ozet,
        "sorunlar": sorun.liste() if sorun else [],
        "yazilan": basarili,
        "yazilamayan": basarisiz,
        "atlanan": atlanan,
        "satirlar": {k: len(d["items"]) for k, d in veri.items()},
    }


# --------------------------------------------------------------------------
# Konsol
# --------------------------------------------------------------------------

def konsol(argv=None):
    p = argparse.ArgumentParser(
        description="SAP ciktilarini Ihbar Site Raporu'na yazar.")
    p.add_argument("--rapor", required=True, help="Excel dosyasi (.xlsx)")
    p.add_argument("--sap", default=None,
                   help="SAP .txt dosyalarinin klasoru (dosyalar otomatik taninir)")
    # Dosya adlari degisken oldugunda her raporu tek tek vermek icin:
    p.add_argument("--icerik", default=None, help="Site Icerigine Gore .txt")
    p.add_argument("--durum", default=None, help="Site Durumuna Gore .txt")
    p.add_argument("--sube", default=None, help="Sube Bazli .txt")
    p.add_argument("--ulke", default=None, help="Ulke Dagilimi .txt")
    p.add_argument("--bitis", default=None,
                   help="Rapor bitis tarihi GG.AA.YYYY (varsayilan: SAP'taki tarih)")
    p.add_argument("--ulke-sekmesi", default=None,
                   help="Ulke sekmesi adi (varsayilan: %s)" % excel_writer.ULKE_SEKME)
    p.add_argument("--grafik-sekmesi", default=None,
                   help="Ulke grafigi sekmesi adi (varsayilan: %s)" % excel_writer.GRAFIK_SEKME)
    p.add_argument("--word", default=None,
                   help="Word raporu (.docx) - verilirse o da guncellenir")
    p.add_argument("--ad-degistirme", action="store_true",
                   help="Dosya adini degistirme, mevcut dosyaya yaz")
    p.add_argument("--sor", action="store_true",
                   help="Sorun cikarsa ekrandan sor (E/h)")
    p.add_argument("--devam", action="store_true",
                   help="Sorun cikarsa sormadan devam et (varsayilan: DURDUR)")
    p.add_argument("--hata-durdur", action="store_true",
                   help="(varsayilan) Ilk sorunda isi iptal et")
    a = p.parse_args(argv)

    tekil = {"icerik": a.icerik, "durum": a.durum,
             "sube": a.sube, "ulke": a.ulke}
    verilen = {k: v for k, v in tekil.items() if v}
    if verilen and len(verilen) != 4:
        eksik = [RAPOR_ADLARI[k] for k in tekil if not tekil[k]]
        print("HATA: Dosyalari tek tek verirken DORDUNU birden verin.\n"
              "Eksik: %s" % ", ".join(eksik), file=sys.stderr)
        return 1
    if not verilen and not a.sap:
        print("HATA: --sap klasoru veya dort dosyanin yolu gerekli.", file=sys.stderr)
        return 1

    # VARSAYILAN: ilk sorunda DURDUR. Islem yarim kalmasin, kullanici
    # neyin bozuk oldugunu gorup duzeltsin.
    if a.sor:
        sorucu = _konsol_sorucu
    elif a.devam:
        sorucu = lambda baslik, ayrinti="": True
    else:
        sorucu = lambda baslik, ayrinti="": False

    try:
        ozet = calistir(a.rapor, a.sap, bitis_tarihi=a.bitis, log=print,
                        ulke_sekmesi=a.ulke_sekmesi,
                        grafik_sekmesi=a.grafik_sekmesi,
                        yeniden_adlandir=not a.ad_degistirme,
                        sap_dosyalari=verilen or None, word_yolu=a.word,
                        sorucu=sorucu)
        if ozet.get("yazilamayan") or ozet.get("atlanan"):
            print("\nTAMAMLANDI - bazi bolumler yazilamadi, ozeti kontrol edin.")
            return 3
        if ozet.get("uyarilar") or ozet.get("sorunlar"):
            print("\nTAMAMLANDI - uyarilar var, ozeti kontrol edin.")
            return 0
        print("\nTAMAMLANDI.")
        return 0
    except Iptal as e:
        print("\nIPTAL EDILDI: %s" % e, file=sys.stderr)
        print("Dosyalariniz DEGISTIRILMEDI.", file=sys.stderr)
        return 4
    except (Kritik, ParseError, WriteError, TarihHatasi,
            word_writer.WordHatasi) as e:
        print("\nHATA: %s" % e, file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("\nKullanici tarafindan durduruldu. Dosyalar degistirilmedi.",
              file=sys.stderr)
        return 4
    except Exception as e:
        # Beklenmeyen hata: coksek bile ne oldugu ANLASILIR yazilsin
        print("\nBEKLENMEYEN HATA: %s" % hatalar._kisa_hata(e), file=sys.stderr)
        print("\nAyrinti (destek icin):", file=sys.stderr)
        traceback.print_exc()
        print("\nDosyalariniz yedeklendi; 'yedek' klasorune bakabilirsiniz.",
              file=sys.stderr)
        return 2


def _konsol_sorucu(baslik, ayrinti=""):
    """Konsolda evet/hayir sorar. Cevap alinamazsa DEVAM eder."""
    try:
        mesaj = "\n  SORUN: %s" % baslik
        if ayrinti:
            mesaj += "\n         %s" % ayrinti
        print(mesaj, file=sys.stderr)
        cevap = input("  Devam edilsin mi? [E/h]: ").strip().lower()
        return cevap in ("", "e", "evet", "y", "yes")
    except (EOFError, KeyboardInterrupt, OSError):
        # Zamanlanmis gorevde girdi yok -> devam
        print("  (cevap alinamadi, devam ediliyor)", file=sys.stderr)
        return True


def _exe_mi():
    """PyInstaller ile paketlenmis exe icinde miyiz?"""
    return getattr(sys, "frozen", False)


def _konsola_bagla():
    """
    Pencere modunda derlenmis exe'de konsol yoktur; komut satirindan
    calistirilirsa cikti gorunmez. Varsa CAGIRAN konsola baglaniriz.
    """
    if not _exe_mi() or os.name != "nt":
        return
    try:
        import ctypes
        # ATTACH_PARENT_PROCESS = -1
        if ctypes.windll.kernel32.AttachConsole(-1):
            sys.stdout = open("CONOUT$", "w", encoding="utf-8",
                              errors="replace", buffering=1)
            sys.stderr = open("CONOUT$", "w", encoding="utf-8",
                              errors="replace", buffering=1)
    except Exception:
        pass  # konsol yoksa sessizce gec


def _ana():
    """
    Giris noktasi.

    Argumansiz  -> pencere acilir
    Argumanli   -> komut satiri modunda calisir
    Hata olursa exe SESSIZCE kapanmaz; sebep gosterilir.
    """
    komut_modu = len(sys.argv) > 1 and sys.argv[1] != "--pencere"

    if komut_modu:
        _konsola_bagla()
        return konsol()

    try:
        import gui
        gui.main()
        return 0
    except Exception as e:
        # Pencere hic acilamadiysa kullanici bos ekranla kalmasin
        mesaj = ("Program penceresi acilamadi.\n\n%s\n\n"
                 "Python/Tk kurulumu eksik olabilir." % hatalar._kisa_hata(e))
        try:
            import tkinter as tk
            from tkinter import messagebox
            k = tk.Tk(); k.withdraw()
            messagebox.showerror("Ihbar Rapor - Baslatilamadi", mesaj)
            k.destroy()
        except Exception:
            _konsola_bagla()
            print(mesaj, file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(_ana())
