# -*- coding: utf-8 -*-
"""
Merkezi hata yonetimi.

Amac: hicbir hata SESSIZCE yutulmasin, program COKMESIN, is gereksiz
yere IPTAL olmasin. Bunun icin hatalar uc siniftan birine girer:

  KRITIK   : devam edilemez (dosya yok, hicbir veri okunamadi).
             Islem durur, sebep acikca yazilir.
  SORULUR  : devam edilebilir ama kullanici bilmeli (bir sekme bozuk,
             tarih tutmuyor). Kullaniciya sorulur; cevaba gore devam
             edilir veya iptal edilir.
  BILGI    : kendi basina cozulur, sadece raporlanir (ornek satiri
             bulunamadi, kozmetik bir ayar uygulanamadi).

Kullanici arayuzu 'sorucu' geri cagrimini verir; konsolda sorulamiyorsa
varsayilan davranis KRITIK olmayan hatalarda DEVAM etmektir - boylece
zamanlanmis gorevde program takilip kalmaz, ama gunluge her sey yazilir.
"""


class Kritik(Exception):
    """Devam edilemeyecek durum; islem durur."""


class Iptal(Exception):
    """Kullanici devam etmek istemedi."""


# Onem dereceleri
KRITIK = "kritik"
SORULUR = "sorulur"
BILGI = "bilgi"


class Rapor(object):
    """
    Calisma boyunca olusan tum sorunlari toplar.

    Her kayit: (onem, baslik, ayrinti, nerede)
    """

    def __init__(self, log=None, sorucu=None):
        self.kayitlar = []
        self.log = log or (lambda s: None)
        # sorucu(baslik, ayrinti) -> True (devam) / False (iptal)
        self.sorucu = sorucu

    # -- kayit --

    def bilgi(self, baslik, ayrinti="", nerede=""):
        self.kayitlar.append((BILGI, baslik, ayrinti, nerede))
        self.log("    - %s%s" % (baslik, (" (%s)" % ayrinti) if ayrinti else ""))

    def uyar(self, baslik, ayrinti="", nerede=""):
        """Sorulmadan kaydedilen uyari (islem devam eder)."""
        self.kayitlar.append((SORULUR, baslik, ayrinti, nerede))
        self.log("    ! %s%s" % (baslik, (" - %s" % ayrinti) if ayrinti else ""))

    def sor(self, baslik, ayrinti="", nerede="", varsayilan=True):
        """
        Kullaniciya sorar: bu sorunla devam edilsin mi?

        Doner: True (devam) / False (iptal)
        Sorucu yoksa varsayilan uygulanir ve gunluge yazilir.
        """
        self.kayitlar.append((SORULUR, baslik, ayrinti, nerede))
        self.log("    ! %s%s" % (baslik, (" - %s" % ayrinti) if ayrinti else ""))

        if self.sorucu is None:
            self.log("      -> %s (soru sorulamadi, varsayilan)"
                     % ("devam ediliyor" if varsayilan else "iptal"))
            return varsayilan

        cevap = self.sorucu(baslik, ayrinti)
        self.log("      -> %s" % ("devam ediliyor" if cevap else "IPTAL"))
        return bool(cevap)

    # -- sorgu --

    def var_mi(self, onem=None):
        if onem is None:
            return bool(self.kayitlar)
        return any(k[0] == onem for k in self.kayitlar)

    def liste(self, onem=None):
        if onem is None:
            return list(self.kayitlar)
        return [k for k in self.kayitlar if k[0] == onem]


def guvenli(rapor, baslik, islem, *args, **kw):
    """
    Bir islemi calistirir; hata olursa PROGRAMI COKERTMEZ.

    Hatayi rapora yazar ve None doner. Cagiran taraf None'i gorup
    kendi kararini verir (atla / sor / dur).

    varsayilan_deger : hata halinde donecek deger
    """
    varsayilan = kw.pop("varsayilan_deger", None)
    try:
        return islem(*args, **kw)
    except Kritik:
        raise
    except Iptal:
        raise
    except Exception as e:
        rapor.uyar(baslik, _kisa_hata(e))
        return varsayilan


# Excel'in ise yaramaz COM ayrintilari (yardim dosyasi adi vb.)
_ANLAMSIZ = ("xlmain11.chm", "0x800a03ec", "None", "")


def _kisa_hata(e):
    """
    COM/Windows hatalarini okunabilir hale getirir.

    Excel COM hatalari cogu zaman yardim dosyasi adi ('xlmain11.chm')
    gibi kullaniciya hicbir sey anlatmayan metinler dondurur; bunlari
    ayiklayip anlamli bir aciklama birakiriz.
    """
    parcalar = []

    # pywin32: (hresult, kaynak, excepinfo, ...) seklinde gelir
    bilgi = getattr(e, "excepinfo", None)
    if bilgi:
        try:
            for x in bilgi:
                if isinstance(x, str) and x.strip() and \
                        x.strip().lower() not in _ANLAMSIZ:
                    parcalar.append(x.strip())
        except Exception:
            pass

    metin = str(e).strip()
    if metin and metin.lower() not in _ANLAMSIZ:
        parcalar.append(metin)

    # Anlamsiz parcalari at
    temiz = []
    for x in parcalar:
        x = " ".join(x.split())
        if x.lower() in _ANLAMSIZ:
            continue
        # 'xlmain11.chm' gibi tek kelimelik dosya adlarini at
        if x.lower().endswith((".chm", ".hlp")) and " " not in x:
            continue
        if x not in temiz:
            temiz.append(x)

    m = " | ".join(temiz) if temiz else e.__class__.__name__
    if len(m) > 200:
        m = m[:197] + "..."
    return m
