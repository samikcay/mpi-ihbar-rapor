# -*- coding: utf-8 -*-
"""
Basit pencere arayuzu: dosyalari sec, Calistir'a bas.
"""

import os
import sys
import json
import queue
import threading
import traceback

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import rapor
import tarih
import sap_parser
from sap_parser import ParseError
from excel_writer import WriteError
from tarih import TarihHatasi
from word_writer import WordHatasi
from hatalar import Kritik, Iptal

# Takvim bilesenini kullan; kurulu degilse elle yazmaya geri don.
try:
    from tkcalendar import DateEntry
    TAKVIM_VAR = True
except ImportError:
    TAKVIM_VAR = False

AYAR_DOSYASI = os.path.join(os.path.expanduser("~"), ".ihbar_rapor.json")

# Hangi SAP dosyasi hangi sekmeye yazilir (UI sirasi)
SEKME_SIRASI = ("icerik", "durum", "sube", "ulke")
SEKME_ETIKETLERI = {
    "icerik": "Site İçeriğine Göre",
    "durum": "Site Durumuna Göre",
    "sube": "Şube Bazlı",
    "ulke": "Ülke Dağılımı + Grafiği",
}


def ayar_yukle():
    try:
        with open(AYAR_DOSYASI, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def ayar_kaydet(d):
    try:
        with open(AYAR_DOSYASI, "w", encoding="utf-8") as fh:
            json.dump(d, fh, ensure_ascii=False, indent=1)
    except Exception:
        pass


class Uygulama(object):

    def __init__(self, kok):
        self.kok = kok
        kok.title("İhbar Site Raporu - SAP Aktarım")
        kok.geometry("860x600")
        kok.minsize(720, 500)

        ayar = ayar_yukle()
        self.rapor_yolu = tk.StringVar(value=ayar.get("rapor", ""))
        self.bitis = tk.StringVar()
        self.word_yolu = tk.StringVar(value=ayar.get("word", ""))
        self.word_acik = tk.BooleanVar(value=ayar.get("word_acik", False))
        self.kuyruk = queue.Queue()
        self.calisiyor = False

        # Her sekme icin ayri dosya yolu
        kayitli = ayar.get("dosyalar", {})
        self.dosya_yollari = {}
        for anahtar in SEKME_SIRASI:
            self.dosya_yollari[anahtar] = tk.StringVar(
                value=kayitli.get(anahtar, ""))

        # Seç pencerelerinin acilacagi klasor (son kullanilan)
        self.son_klasor = ayar.get("son_klasor", "")

        # Dosyalar degisince bitis tarihini otomatik doldur
        for sv in self.dosya_yollari.values():
            sv.trace_add("write", lambda *a: self.tarihi_oner())

        ana = ttk.Frame(kok, padding=12)
        ana.pack(fill="both", expand=True)
        ana.columnconfigure(1, weight=1)

        ttk.Label(ana, text="Excel raporu:").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(ana, textvariable=self.rapor_yolu).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(ana, text="Seç...", command=self.rapor_sec).grid(row=0, column=2)

        # --- Her sekme icin SAP dosyasi ---
        kutu = ttk.LabelFrame(
            ana, text="SAP çıktıları  —  hangi dosya hangi sekmeye yazılacak",
            padding=8)
        kutu.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(8, 2))
        kutu.columnconfigure(1, weight=1)

        self.dosya_notlari = {}
        for i, anahtar in enumerate(SEKME_SIRASI):
            ttk.Label(kutu, text=SEKME_ETIKETLERI[anahtar] + ":", width=22)\
                .grid(row=i, column=0, sticky="w", pady=3)
            ttk.Entry(kutu, textvariable=self.dosya_yollari[anahtar])\
                .grid(row=i, column=1, sticky="ew", padx=6)
            # Durum isareti butonun solunda -> pencere darsa da gorunur
            not_etiketi = ttk.Label(kutu, text="", width=2,
                                    font=("Segoe UI", 11, "bold"),
                                    anchor="center")
            not_etiketi.grid(row=i, column=2, padx=(0, 4))
            self.dosya_notlari[anahtar] = not_etiketi
            ttk.Button(kutu, text="Seç...", width=8,
                       command=lambda a=anahtar: self.dosya_sec(a))\
                .grid(row=i, column=3)

        ttk.Button(kutu, text="Klasörden otomatik doldur...",
                   command=self.klasor_sec)\
            .grid(row=len(SEKME_SIRASI), column=1, columnspan=3, sticky="w",
                  padx=6, pady=(6, 0))

        # --- Word raporu (istege bagli) ---
        word_kutu = ttk.Frame(ana)
        word_kutu.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        word_kutu.columnconfigure(1, weight=1)

        ttk.Checkbutton(word_kutu, text="Word raporunu da güncelle",
                        variable=self.word_acik,
                        command=self._word_durumu)\
            .grid(row=0, column=0, columnspan=4, sticky="w")

        ttk.Label(word_kutu, text="Word dosyası:", width=22)\
            .grid(row=1, column=0, sticky="w", pady=(2, 0))
        self.word_kutusu = ttk.Entry(word_kutu, textvariable=self.word_yolu)
        self.word_kutusu.grid(row=1, column=1, sticky="ew", padx=6, pady=(2, 0))
        self.word_dugmesi = ttk.Button(word_kutu, text="Seç...", width=8,
                                       command=self.word_sec)
        self.word_dugmesi.grid(row=1, column=3, pady=(2, 0))

        # --- Rapor bitis tarihi (takvim) ---
        tarih_cerceve = ttk.Frame(ana)
        tarih_cerceve.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        ttk.Label(tarih_cerceve, text="Rapor bitiş tarihi:").pack(side="left")

        if TAKVIM_VAR:
            # Ok tusuna basinca takvim acilir; kutuya elle de yazilabilir.
            self.takvim = DateEntry(
                tarih_cerceve,
                width=12,
                locale="tr_TR",
                date_pattern="dd.mm.yyyy",
                font=("Consolas", 10),
                justify="center",
                borderwidth=2,
                showweeknumbers=False,
                firstweekday="monday",
            )
            self.takvim.pack(side="left", padx=6)
            ttk.Button(tarih_cerceve, text="Takvim",
                       command=self.takvimi_ac, width=8).pack(side="left")
        else:
            self.takvim = None
            ttk.Entry(tarih_cerceve, textvariable=self.bitis, width=14,
                      font=("Consolas", 10)).pack(side="left", padx=6)
            ttk.Label(tarih_cerceve, text="(GG.AA.YYYY)",
                      foreground="#777").pack(side="left")

        self.tarih_notu = ttk.Label(tarih_cerceve, text="", foreground="#777")
        self.tarih_notu.pack(side="left", padx=(12, 0))

        ttk.Label(
            ana,
            text="Takvimden raporun son gününü seçin. Başlangıç 01.01. olarak "
                 "kalır; tablolardaki ve dosya adındaki tarih bu güne göre "
                 "güncellenir.",
            foreground="#777", wraplength=700, justify="left"
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(2, 0))

        dugmeler = ttk.Frame(ana)
        dugmeler.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(10, 6))
        self.calistir_dugmesi = ttk.Button(
            dugmeler, text="Aktar", command=self.calistir_bas)
        self.calistir_dugmesi.pack(side="left")
        ttk.Button(dugmeler, text="Kapat", command=kok.destroy).pack(side="right")

        self.durum = ttk.Label(ana, text="Hazır.", foreground="#444")
        self.durum.grid(row=6, column=0, columnspan=3, sticky="w")

        cerceve = ttk.LabelFrame(ana, text="İşlem günlüğü", padding=6)
        cerceve.grid(row=7, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
        ana.rowconfigure(7, weight=1)
        cerceve.columnconfigure(0, weight=1)
        cerceve.rowconfigure(0, weight=1)

        self.metin = tk.Text(cerceve, wrap="word", height=16,
                             font=("Consolas", 9), state="disabled")
        self.metin.grid(row=0, column=0, sticky="nsew")
        kaydirma = ttk.Scrollbar(cerceve, command=self.metin.yview)
        kaydirma.grid(row=0, column=1, sticky="ns")
        self.metin.configure(yscrollcommand=kaydirma.set)

        self.metin.tag_configure("hata", foreground="#b00020")
        self.metin.tag_configure("uyari", foreground="#a06000")
        self.metin.tag_configure("basari", foreground="#0a6b0a")
        self.metin.tag_configure("baslik", foreground="#333",
                                 font=("Consolas", 9, "bold"))
        self.metin.tag_configure("baslik_basari", foreground="#0a6b0a",
                                 font=("Consolas", 9, "bold"))
        self.metin.tag_configure("baslik_uyari", foreground="#a06000",
                                 font=("Consolas", 9, "bold"))

        self._word_durumu()
        kok.after(100, self.kuyrugu_isle)

    # -- dosya secimi --

    def rapor_sec(self):
        yol = filedialog.askopenfilename(
            title="Excel raporunu seçin",
            filetypes=[("Excel dosyaları", "*.xlsx *.xlsm"), ("Tümü", "*.*")])
        if yol:
            self.rapor_yolu.set(yol)

    def dosya_sec(self, anahtar):
        """Tek bir sekme icin SAP .txt dosyasi sectirir."""
        baslangic = self.son_klasor or os.path.expanduser("~")
        mevcut = self.dosya_yollari[anahtar].get().strip()
        if mevcut and os.path.isfile(mevcut):
            baslangic = os.path.dirname(mevcut)
        yol = filedialog.askopenfilename(
            title="%s  —  SAP çıktısını seçin" % SEKME_ETIKETLERI[anahtar],
            initialdir=baslangic,
            filetypes=[("SAP metin çıktıları", "*.txt *.TXT"), ("Tümü", "*.*")])
        if yol:
            self.dosya_yollari[anahtar].set(yol)
            self.son_klasor = os.path.dirname(yol)
            self.dosyalari_kontrol_et()

    def klasor_sec(self):
        """Klasor secip dort dosyayi otomatik doldurur (kolaylik icin)."""
        baslangic = self.son_klasor or os.path.expanduser("~")
        klasor = filedialog.askdirectory(
            title="SAP çıktılarının klasörünü seçin", initialdir=baslangic)
        if not klasor:
            return
        self.son_klasor = klasor
        try:
            bulunan = rapor.klasoru_tara(klasor)
        except Exception as e:
            messagebox.showwarning("Klasör okunamadı", str(e))
            return
        for anahtar in SEKME_SIRASI:
            if bulunan.get(anahtar):
                self.dosya_yollari[anahtar].set(bulunan[anahtar])
        eksik = [SEKME_ETIKETLERI[a] for a in SEKME_SIRASI
                 if not bulunan.get(a)]
        if eksik:
            messagebox.showinfo(
                "Bazı dosyalar bulunamadı",
                "Şu raporlar klasörde tanınamadı, lütfen elle seçin:\n\n- %s"
                % "\n- ".join(eksik))
        self.dosyalari_kontrol_et()

    def dosyalari_kontrol_et(self):
        """
        Secilen her dosyanin icerigine bakip dogru sekmeye mi gidiyor
        isaretler. Yesil onay = baslik uyusuyor, turuncu = suphe.
        """
        secili = {a: self.dosya_yollari[a].get().strip()
                  for a in SEKME_SIRASI}
        for anahtar, yol in secili.items():
            etiket = self.dosya_notlari[anahtar]
            if not yol or not os.path.isfile(yol):
                etiket.configure(text="")
                continue
            try:
                tur = rapor._basliktan_tur_bul(yol)
            except Exception:
                tur = None
            if tur is None:
                etiket.configure(text="?", foreground="#a06000")
            elif tur == anahtar:
                etiket.configure(text="\u2713", foreground="#0a6b0a")
            else:
                etiket.configure(text="!", foreground="#b00020")

    def _word_durumu(self):
        """Word kutusu isaretli degilse dosya alanini pasiflestirir."""
        dur = "normal" if self.word_acik.get() else "disabled"
        self.word_kutusu.configure(state=dur)
        self.word_dugmesi.configure(state=dur)

    def word_sec(self):
        baslangic = os.path.dirname(self.rapor_yolu.get().strip() or "") \
            or os.path.expanduser("~")
        mevcut = self.word_yolu.get().strip()
        if mevcut and os.path.isfile(mevcut):
            baslangic = os.path.dirname(mevcut)
        yol = filedialog.askopenfilename(
            title="Word raporunu seçin", initialdir=baslangic,
            filetypes=[("Word belgeleri", "*.docx"), ("Tümü", "*.*")])
        if yol:
            self.word_yolu.set(yol)

    def takvimi_ac(self):
        """Takvim düğmesi: açılır takvimi gösterir."""
        if self.takvim is not None:
            self.takvim.drop_down()

    def tarih_oku(self):
        """
        Secili tarihi date olarak doner.
        Takvim varsa ondan, yoksa metin kutusundan okur.
        """
        if self.takvim is not None:
            try:
                return self.takvim.get_date()
            except Exception:
                # Kutuya elle gecersiz bir sey yazilmis olabilir
                raise TarihHatasi(
                    "Tarih okunamadı. Takvimden bir gün seçin.")
        return tarih.gun_coz(self.bitis.get().strip())

    def tarih_yaz(self, d):
        """Onerilen tarihi bilesene yerlestirir."""
        if self.takvim is not None:
            self.takvim.set_date(d)
        else:
            self.bitis.set(tarih.nokta(d))

    def tarihi_oner(self):
        """
        Secilen SAP dosyalarindan bitis tarihini okuyup takvime yazar.
        Kullanici isterse takvimden baska bir gun secebilir.
        """
        try:
            for anahtar in SEKME_SIRASI:
                yol = self.dosya_yollari[anahtar].get().strip()
                if not yol or not os.path.isfile(yol):
                    continue
                d = sap_parser.extract_period(sap_parser.read_sap_text(yol))
                if d:
                    self.tarih_yaz(tarih.gun_coz(d[1]))
                    self.tarih_notu.configure(
                        text="← SAP çıktısındaki tarih", foreground="#0a6b0a")
                    return
        except Exception:
            pass
        self.tarih_notu.configure(text="")

    # -- gunluk --

    # Ozet bolumlerini renklendirmek icin anahtar kelimeler
    _RENK_KURALLARI = (
        ("GUNCELLENMEDI", "baslik_uyari"),
        ("UYARILAR", "baslik_uyari"),
        ("GUNCELLENDI", "baslik_basari"),
        ("=== ", "baslik"),
    )

    def yaz(self, satir, etiket=None):
        """
        Gunluge bir satir yazar. Etiket verilmezse satirin icerigine gore
        otomatik renklendirir (ozet basliklari, uyarilar).
        """
        if etiket is None:
            s = satir.lstrip()
            if s.startswith("!") or " ! " in satir[:8]:
                etiket = "uyari"
            else:
                for anahtar, tag in self._RENK_KURALLARI:
                    if s.startswith(anahtar):
                        etiket = tag
                        break
        self.kuyruk.put(("log", satir, etiket))

    def kuyrugu_isle(self):
        try:
            while True:
                tur, veri, etiket = self.kuyruk.get_nowait()
                if tur == "log":
                    self.metin.configure(state="normal")
                    self.metin.insert("end", veri + "\n", etiket or ())
                    self.metin.see("end")
                    self.metin.configure(state="disabled")
                elif tur == "durum":
                    self.durum.configure(text=veri)
                elif tur == "soru":
                    mesaj, kutu, olay = veri
                    try:
                        kutu["cevap"] = messagebox.askyesno(
                            "Sorun oluştu", mesaj, icon="warning",
                            default="yes")
                    except Exception:
                        kutu["cevap"] = True
                    finally:
                        olay.set()
                elif tur == "bitti":
                    self.calisiyor = False
                    self.calistir_dugmesi.configure(state="normal")
        except queue.Empty:
            pass
        self.kok.after(100, self.kuyrugu_isle)

    # -- calistirma --

    def calistir_bas(self):
        if self.calisiyor:
            return
        xlsx = self.rapor_yolu.get().strip()

        if not xlsx or not os.path.isfile(xlsx):
            messagebox.showerror("Eksik bilgi", "Geçerli bir Excel raporu seçin.")
            return

        # Dort dosya da secilmis mi?
        secilen = {}
        eksik = []
        for anahtar in SEKME_SIRASI:
            yol = self.dosya_yollari[anahtar].get().strip()
            if not yol or not os.path.isfile(yol):
                eksik.append(SEKME_ETIKETLERI[anahtar])
            else:
                secilen[anahtar] = yol
        if eksik:
            messagebox.showerror(
                "Eksik dosya",
                "Şu sekmeler için SAP dosyası seçilmedi:\n\n- %s"
                % "\n- ".join(eksik))
            return

        # Ayni dosya birden fazla sekmeye verilmis mi?
        ters = {}
        for anahtar, yol in secilen.items():
            ters.setdefault(os.path.abspath(yol).lower(), []).append(anahtar)
        cift = [v for v in ters.values() if len(v) > 1]
        if cift:
            adlar = ["  ve  ".join(SEKME_ETIKETLERI[a] for a in grup)
                     for grup in cift]
            if not messagebox.askokcancel(
                    "Aynı dosya birden fazla sekmede",
                    "Şu sekmeler için AYNI dosya seçilmiş:\n\n%s\n\n"
                    "Devam edilsin mi?" % "\n".join(adlar),
                    icon="warning", default="cancel"):
                return

        # Icerik-sekme uyusmazligi var mi? (ada degil basliga bakilir)
        eslesme = rapor.dosyalari_dogrula(secilen)
        if eslesme:
            if not messagebox.askokcancel(
                    "Dosya-sekme uyuşmazlığı",
                    "%s\n\nYine de devam edilsin mi?" % "\n\n".join(eslesme),
                    icon="warning", default="cancel"):
                return

        # Tarihi calistirmadan ONCE dogrula
        try:
            bitis = self.tarih_oku()
        except TarihHatasi as e:
            messagebox.showerror("Tarih hatası", str(e))
            return

        # Onizleme icin baslangici mevcut dosya adindan al (yoksa 1 Ocak)
        import datetime
        mevcut = tarih.dosya_adindan_donem(os.path.basename(xlsx))
        baslangic = mevcut[0] if mevcut else datetime.date(bitis.year, 1, 1)
        yeni_ad, degisti = tarih.dosya_adi_guncelle(
            os.path.basename(xlsx), baslangic, bitis)
        # Word secildiyse gecerli mi?
        word = self.word_yolu.get().strip() if self.word_acik.get() else ""
        if self.word_acik.get():
            if not word or not os.path.isfile(word):
                messagebox.showerror(
                    "Eksik bilgi",
                    "Word güncellemesi işaretli ama geçerli bir .docx seçilmedi.")
                return

        onay = (
            "Rapor bitiş tarihi: %s\n\n"
            "Tablolardaki tarihler bu güne göre güncellenecek." % tarih.nokta(bitis))
        if word:
            onay += ("\n\nWord raporu da güncellenecek:\n%s\n"
                     "(yedeği alınacak, SAP'tan türetilemeyen sayılara "
                     "dokunulmayacak)" % os.path.basename(word))
        if degisti:
            onay += "\n\nYeni dosya oluşturulacak:\n%s\n\n(mevcut dosya yerinde kalacak)" % yeni_ad
        if not messagebox.askokcancel("Onay", onay, icon="question", default="cancel"):
            return

        ayar_kaydet({"rapor": xlsx,
                     "dosyalar": secilen,
                     "son_klasor": self.son_klasor,
                     "word": self.word_yolu.get().strip(),
                     "word_acik": bool(self.word_acik.get())})

        self.metin.configure(state="normal")
        self.metin.delete("1.0", "end")
        self.metin.configure(state="disabled")

        self.calisiyor = True
        self.calistir_dugmesi.configure(state="disabled")
        self.kuyruk.put(("durum", "Çalışıyor...", None))

        t = threading.Thread(target=self._is,
                             args=(xlsx, secilen, bitis, word),
                             daemon=True)
        t.start()

    def _sorucu(self, baslik, ayrinti=""):
        """
        Is parcacigindan cagrilir; soruyu ANA is parcaciginda sordurur
        ve cevabi bekler. (Tkinter sadece ana is parcacigindan kullanilir.)
        """
        import threading as _th
        olay = _th.Event()
        kutu = {}

        mesaj = baslik
        if ayrinti:
            mesaj += "\n\n%s" % ayrinti
        mesaj += ("\n\nBu bölüm atlanarak devam edilsin mi?\n"
                  "(Hayır derseniz işlem iptal edilir, dosyalarınız "
                  "değişmez.)")

        # Soru ANA is parcaciginda sorulmali; kuyruga birak, kuyrugu_isle
        # cagirsin. (Tkinter'in after() metodu bile is parcacigindan
        # guvenli degildir.)
        self.kuyruk.put(("soru", (mesaj, kutu, olay), None))

        # Cevap gelene kadar bekle. Arayuz donmaz: burasi ayri parcacik.
        if not olay.wait(timeout=600):
            # Kimse cevaplamadiysa (pencere kapandi vb.) devam et
            return True
        return kutu.get("cevap", True)

    def _is(self, xlsx, secilen, bitis, word=None):
        # COM bu is parcaciginda kullanilacagi icin baslatilmali
        import pythoncom
        pythoncom.CoInitialize()
        try:
            ozet = rapor.calistir(xlsx, bitis_tarihi=bitis, log=self.yaz,
                                  sap_dosyalari=secilen,
                                  word_yolu=word or None,
                                  sorucu=self._sorucu)
            if ozet.get("yazilamayan") or ozet.get("atlanan"):
                self.kuyruk.put(("durum", "Tamamlandı - bazı bölümler yazılamadı.", None))
                self.yaz("\nTAMAMLANDI - bazı bölümler yazılamadı, özeti kontrol edin.",
                         "uyari")
            elif ozet.get("uyarilar") or ozet.get("sorunlar"):
                self.kuyruk.put(("durum", "Tamamlandı (uyarılarla).", None))
                self.yaz("\nTAMAMLANDI - lütfen yukarıdaki uyarıları kontrol edin.", "uyari")
            else:
                self.kuyruk.put(("durum", "Tamamlandı.", None))
                self.yaz("\nTAMAMLANDI.", "basari")
        except Iptal as e:
            self.yaz("\nIPTAL EDILDI: %s" % e, "uyari")
            self.yaz("Dosyalarınız değiştirilmedi.", "uyari")
            self.kuyruk.put(("durum", "İptal edildi.", None))
        except (Kritik, ParseError, WriteError, TarihHatasi, WordHatasi) as e:
            self.yaz("\nHATA: %s" % e, "hata")
            self.yaz("Excel dosyası değiştirilmedi.", "hata")
            self.kuyruk.put(("durum", "Hata.", None))
        except Exception as e:
            import hatalar as _h
            self.yaz("\nBEKLENMEYEN HATA: %s" % _h._kisa_hata(e), "hata")
            self.yaz("\nDosyalarınızın yedeği 'yedek' klasöründe duruyor.", "uyari")
            self.yaz("\nAyrıntı (destek için):", "hata")
            self.yaz(traceback.format_exc(), "hata")
            self.kuyruk.put(("durum", "Beklenmeyen hata.", None))
        finally:
            pythoncom.CoUninitialize()
            self.kuyruk.put(("bitti", None, None))


def main():
    kok = tk.Tk()
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    Uygulama(kok)
    kok.mainloop()


if __name__ == "__main__":
    main()
