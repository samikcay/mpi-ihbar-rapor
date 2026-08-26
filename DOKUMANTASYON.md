# İhbar Site Raporu — SAP Aktarım Programı

## Teknik Dokümantasyon

Sürüm: 25.08.2026
Hedef platform: Windows + Microsoft Excel

---

# 1. AMAÇ

SAP'tan alınan metin çıktılarını okur, `İhbar Site Rapor` çalışma kitabındaki
altı sekmeye yazar, isteğe bağlı olarak Word raporundaki sayıları günceller.

Elle yapıldığında hataya açık ve tekrarlayan bir işi otomatikleştirir.

---

# 2. VERİ AKIŞI

```
SAP .txt (4 dosya, UTF-16LE)
    |
    v
sap_parser  ->  {"items": [(ad, sayi)...], "total": int, "period": (bas, bit)}
    |
    v
rapor.calistir  --  doğrulama, dönem hesabı, yedekleme
    |
    +--> excel_writer (COM)  ->  6 sekme + dönem tarihleri
    |         |
    |         v
    |    Yıllar Bazlı B/C/D  ->  B25/E25 kümülatif formülleri
    |                                    |
    +--> word_writer  <------------------+  (Excel'den okunan sayılar)
              |
              v
         Word .docx paragrafları
```

Kritik bağımlılık: Word'deki kümülatif cümle, Excel'in `Yıllar Bazlı`
sekmesindeki formüllerden beslenir. Bu yüzden Word güncellemesi Excel
kaydedildikten **sonra** çalışır ve kaydedilmiş dosyayı okur.

---

# 3. MODÜLLER

| Modül | Sorumluluk |
|---|---|
| `rapor.py` | Ana akış, komut satırı, özet raporu, giriş noktası |
| `gui.py` | Pencere arayüzü |
| `sap_parser.py` | SAP metin çıktılarını ayrıştırma |
| `excel_writer.py` | Excel'e yazma (COM) |
| `word_writer.py` | Word'e yazma (python-docx) |
| `tarih.py` | Tarih biçimleri, dosya adı |
| `labels.py` | Etiket yazımı normalizasyonu |
| `hatalar.py` | Hata sınıflandırma ve kurtarma |

---

# 4. FONKSİYON REFERANSI

## 4.1 `sap_parser.py`

SAP çıktıları UTF-16LE + BOM, sekme ayraçlı, veri satırları arasında boş
satır içerir.

### `read_sap_text(path) -> list[str]`
Dosyayı kodlamasını saptayarak okur. BOM'a bakar; UTF-16, UTF-8-BOM,
UTF-8, CP1254, Latin-1 sırasıyla dener. Sayfa besleme karakterlerini
temizler. Satır listesi döner.

Hata: `ParseError` — dosya boş veya kodlama çözülemedi.

### `parse_int(text) -> int`
`'64.145'` → `64145`. Nokta binlik ayracını atar, NBSP temizler.

Hata: `ParseError` — sayı değil.

### `parse_simple(path) -> dict`
Tek bloklu raporlar (içerik / durum / şube).

Döner:
```python
{"items": [(ad, sayi), ...], "total": int|None, "period": (bas, bit)|None}
```

### `parse_countries(path) -> dict`
Ülke raporu iki blok halinde yazılır (sol `1..n`, sağ `n+1..son`).
Sıra numarasına göre birleştirir, tek listeye çevirir.

Hata: `ParseError` — sıra numarası tekrar ediyor veya boşluk var.

### `extract_period(lines) -> (str, str) | None`
Başlıktaki `01.01.2026-07.08.2026` dönemini bulur.

### `_find_total(lines) -> int | None`
Dosya sonundaki `TOPLAM SİTE SAYISI` değerini bulur. Doğrulama için
kullanılır: satır toplamı bu değerle karşılaştırılır.

---

## 4.2 `tarih.py`

Üç ayrı tarih biçimi vardır. Her birinin ayrı fonksiyonu bulunur.

### `gun_coz(metin) -> datetime.date`
`'07.08.2026'`, `'7.8.2026'`, `'07/08/2026'` kabul eder.

Hata: `TarihHatasi` — boş, biçimsiz veya olmayan tarih (`31.02.2026`).

### `nokta(d) -> str`
`date` → `'07.08.2026'`

### `ad_ile(d) -> str`
`date` → `'07 Ağustos 2026'` (dosya adı biçimi)

### `donem_metni(bas, bit) -> str`
Sekme başlığı biçimi: `'01.01.2026 - 07.08.2026'`

### `donem_alt_alta(bas, bit) -> str`
`Yıllar Bazlı` biçimi: `'01.01.2026\n07.08.2026'` (hücre içinde iki satır)

### `dosya_adi_guncelle(ad, bas, bit) -> (str, bool)`
`'... (01 Ocak-31 Temmuz 2026).xlsx'` → `'... (01 Ocak-07 Ağustos 2026).xlsx'`

Desen bulunamazsa ad değişmeden döner, ikinci değer `False`.

### `dosya_adindan_donem(ad) -> (date, date) | None`
Dosya adındaki dönemi okur. Ay adlarını sayıya çevirir.

---

## 4.3 `labels.py`

SAP çıktılarında Türkçe karakterler kaybolur (`BTK ENGELLI`), rapor doğru
yazımı kullanır (`BTK ENGELLİ`). Grafikler bu hücreleri etiket olarak
kullandığı için rapordaki yazım korunur.

### `anahtar(metin) -> str`
Karşılaştırma anahtarı: aksansız, küçük harf, tek boşluk.
`I/İ/ı/i` ayrımını ortadan kaldırır.

### `duzelt(metin, ek_tercihler=None) -> str`
SAP etiketini rapordaki yazıma çevirir. `ek_tercihler` çalışma kitabından
okunan mevcut etiketlerdir (şube/ülke adları gibi değişken listeler için).

Listede olmayan yeni etiket SAP'taki hâliyle döner. Veri kaybolmaz.

---

## 4.4 `excel_writer.py`

openpyxl grafikleri bozduğu için Excel COM kullanılır.

### `class ExcelOturum`
Bağlam yöneticisi (`with`). Excel'i açar, hata olsa da kapatır.

`DispatchEx` kullanır, `Dispatch` **değil**. Gerekçe:
- `Dispatch` çalışan Excel'e bağlanır; `Quit()` kullanıcının penceresini
  kapatabilir, kaydedilmemiş çalışması gidebilir.
- O Excel'de açık bir uyarı penceresi varsa program sonsuza kadar bekler.

Açılışta kapatılan ayarlar: `DisplayAlerts`, `ScreenUpdating`,
`EnableEvents`, `AskToUpdateLinks`, `AlertBeforeOverwriting`,
`DisplayStatusBar`. Amaç: kimsenin cevaplayamayacağı pencere açılmasın.

- `.ac(yol)` — `UpdateLinks=0` ile açar (bağlantı sorusu sorulmaz)
- `.__exit__` — kitabı kapat → artakalan kitapları kapat → `Quit()`

### `yaz_basit(wb, anahtar_ad, kayitlar, log=None) -> int`
Tek sütunlu sekmeleri doldurur. `anahtar_ad`: `"icerik" | "durum" | "sube"`.

Yerleşim `YERLESIM` sabitinden okunur. TOPLAM satırı `=SUM(` formülü
aranarak bulunur; bu yüzden satır sayısı değişse de doğru yeri bulur.

### `yaz_ulkeler(wb, kayitlar, sekme_adi, log=None) -> int`
İki sütunlu düzen. Sol blok `ceil(n/2)` satır alır, kalan sağa gider.
Tek sayıda ülke varsa fazlalık sol sütundadır.

### `yaz_ulke_grafigi(wb, kayitlar, sekme_adi, log=None) -> int`
İlk 15 ülkeyi aynen yazar. 16. satıra `Diğer = genel toplam − ilk 15`.
Yerleşim sabittir (7–21 ülke, 22 Diğer, 23 TOPLAM); satır eklenmez.

15'ten az ülke gelirse kalan satırlar boşaltılır, `Diğer` sıfır olur.

Hata: `WriteError` — `Diğer` negatif çıktı.

### `yaz_yillar_bazli(wb, veri, bas, log=None) -> (int|None, dict)`
`Yıllar Bazlı` cari yıl satırının SAP'tan gelen sütunlarını günceller.

| Sütun | İçerik | Kaynak |
|---|---|---|
| B | BTK'ya ihbar edilen toplam | SAP (ülke toplamı) |
| C | Direk oynatan | SAP (içerik) |
| D | Reklam, tanıtım, yönlendirme | SAP (içerik) |
| E–I | Engellenen, EEK, mahkeme, pasif, devam eden | **elle** |

Bu satır `B25` ve `E25` kümülatif formüllerini besler. Word'deki
`2006-2026 döneminde toplam ...` cümlesi dolaylı olarak buradan gelir.

Döner: `(satir_no, {sutun: (eski, yeni)})`. Değişiklik yoksa sözlük boş.

### `donem_yaz(wb, bas, bit, log=None) -> int`
Dönem tarihini cari sekmelerde günceller. Geçmiş yıl sekmeleri
(2023/2024/2025) ve `Suç Duyuruları` **ellenmez**.

`Yıllar Bazlı` satırı ayrı biçim taşıdığı için `_yillar_bazli_donem`
tarafından ayrıca işlenir.

### `_blok_boyutunu_ayarla(ws, ilk, son, hedef, ornek_satir) -> int`
Veri bloğunu hedef satır sayısına getirir.

Kilit nokta: satırlar bloğun **içine** eklenir/silinir. Böylece Excel
`=SUM()` aralıklarını ve grafik kaynak aralıklarını kendisi kaydırır.
Bloğun dışına eklenirse grafik aralığı genişlemez.

### `_sekme_bul(wb, ad)`
Sekmeyi adıyla bulur; boşluk ve aksan farkını tolere eder.

Hata: `WriteError` — mesajda mevcut sekmeler listelenir.

---

## 4.5 `word_writer.py`

Word bir cümleyi düzeltme geçmişi nedeniyle parçalara (run) böler.
`123.456` sayısı belgede `'1' + '23' + '.' + '456'` şeklinde durur.
Bu yüzden düz bul-değiştir **hiçbir şey bulamaz**.

Çözüm: paragrafın tam metnini al, regex ile değiştir, metni ilk run'a
yaz, kalanları boşalt. Hedef paragrafların tüm run'ları aynı biçimde
olduğu için (Times New Roman 11.5) biçim kaybı olmaz.

### `guncelle(docx_yolu, veri, bas, bit, xlsx_yolu=None, log=None) -> dict`
Ana giriş noktası.

Döner:
```python
{
  "degisen": [{"paragraf": int, "onceki": str, "sonraki": str,
               "farklar": [(eski, yeni), ...]}, ...],
  "elle": [(aciklama, nerede), ...],      # güncellenmeyenler
  "excel_eksik": [(aciklama, sebep), ...], # Excel'den okunamayanlar
  "yedek": str,
}
```

Hata: `WordHatasi` — dosya açılamadı veya kaydedilemedi.

### `excel_degerleri_oku(xlsx_yolu) -> (dict, list)`
Çalışma kitabından Word'e taşınacak sayıları okur (`data_only=True`,
formüllerin hesaplanmış değeri).

`EXCEL_KAYNAKLARI` haritası:

| Anahtar | Sekme | Hücre |
|---|---|---|
| `kumulatif_tespit` | Yıllar Bazlı | B25 |
| `kumulatif_engel` | Yıllar Bazlı | E25 |
| `mobil` | Mobil Uygulamalar | B4 |
| `sosyal_basvuru` | Sosyal Medya Hesapları | B4 |
| `sosyal_engel` | Sosyal Medya Hesapları | C4 |
| `turetilmis_site` | Suç Duyuruları | C26 |
| `kok_site` | Suç Duyuruları | D26 |
| `odeme_kurulusu` | Ödeme Kuruluşları | B6 |
| `hat_850` | 850'li Hatlar | B4 |
| `iban` | Faaliyet Cetveli | H9 |

Okunamayan değer **sessizce geçilmez**; ikinci dönüş değerinde sebebiyle
bildirilir. Böylece kullanıcı Word'de eski değerin kaldığını bilir.

### `_ulke_yuzdeleri(ulkeler, adet=10) -> (list, int)`
İlk N ülkenin yuvarlanmış yüzdesi + kalan dilim.

Her yüzde ayrı yuvarlanırsa toplam 100 olmayabilir
(`34+30+5+5+5+3+3+3+3+1+9 = 101`). Belgedeki kullanım toplamın 100
olmasıdır; bu yüzden `kalan = 100 − ilk N'in toplamı` şeklinde hesaplanır.
Yuvarlama farkını kalan dilim emer. Excel'deki `Diğer` ile aynı mantık.

### `_ulke_cumlesi_guncelle(metin, d) -> str`
Ülke **adları değişmez**; yalnızca addan sonra gelen `%XX` güncellenir.
Cümlenin dili ve sırası bozulmaz.

### `_sayilari_guncelle(metin, d, degisen) -> str`
Cümleyi tanıdıktan **sonra** sayıyı değiştirir. Aynı sayı başka bağlamda
geçiyorsa yanlışlıkla değişmez.

### `_elle_kontrol_listesi(belge, bas) -> list`
Programın güncellemediği sayıları, **nerede düzeltileceğiyle** listeler.

---

## 4.6 `hatalar.py`

Hiçbir hata sessizce yutulmaz, program çökmez, iş gereksiz yere iptal
olmaz. Üç sınıf:

| Sabit | Anlam |
|---|---|
| `KRITIK` | Devam edilemez. İşlem durur. |
| `SORULUR` | Devam edilebilir; kullanıcıya sorulur. |
| `BILGI` | Kendiliğinden çözülür; raporlanır. |

### `class Kritik(Exception)`
Devam edilemeyecek durum.

### `class Iptal(Exception)`
Kullanıcı devam etmek istemedi.

### `class Rapor`
Çalışma boyunca oluşan sorunları toplar.

- `.__init__(log=None, sorucu=None)` — `sorucu(baslik, ayrinti) -> bool`
- `.bilgi(baslik, ayrinti, nerede)` — kaydet, sorma
- `.uyar(baslik, ayrinti, nerede)` — uyarı kaydet, devam et
- `.sor(baslik, ayrinti, nerede, varsayilan=True) -> bool` — kullanıcıya
  sor. `sorucu` yoksa `varsayilan` uygulanır ve günlüğe yazılır.
- `.liste(onem=None) -> list[(onem, baslik, ayrinti, nerede)]`

### `_kisa_hata(e) -> str`
COM/Windows hatalarını okunabilir hale getirir. Excel COM hataları çoğu
zaman yardım dosyası adı (`xlmain11.chm`) gibi anlamsız metinler döndürür;
bunlar ayıklanır, Excel'in gerçek açıklaması bırakılır.

---

## 4.7 `rapor.py`

### `calistir(...) -> dict`

```python
calistir(xlsx_yolu,
         sap_klasoru=None,      # klasör yolu (otomatik tanıma)
         bitis_tarihi=None,     # date veya 'GG.AA.YYYY'; None -> SAP'taki
         log=print,
         ulke_sekmesi=None,
         grafik_sekmesi=None,
         yeniden_adlandir=True, # yeni dönem adıyla KOPYALA
         sap_dosyalari=None,    # {'icerik':yol, 'durum':yol, ...}
         sorucu=None)           # (baslik, ayrinti) -> bool
```

`sap_dosyalari` verilirse klasör taraması yapılmaz.

Döner:
```python
{
  "toplamlar": {anahtar: int},
  "uyarilar": [str],
  "donem": (bas, bit),
  "yedek": str,
  "hedef": str,          # yazılan dosya
  "yeni_ad": str|None,
  "word": dict|None,
  "sorunlar": [(onem, baslik, ayrinti, nerede)],
  "yazilan": [sekme],
  "yazilamayan": [sekme],
  "atlanan": [sekme],
  "satirlar": {anahtar: int},
}
```

İşlem sırası:
1. Excel dosyası doğrula (yoksa `ParseError`)
2. SAP dosyalarını **tek tek** oku — biri bozuksa diğerleri okunur
3. Toplamları karşılaştır, dönem hesapla
4. Yeni adla kopyala (istenirse), yedek al
5. Her sekmeyi **bağımsız** yaz — biri hata verirse sorulur
6. `Yıllar Bazlı` B/C/D, sonra dönem tarihleri
7. `CalculateFull()` → `Save()`
8. Word (isteğe bağlı) — hatası Excel'i geçersiz kılmaz
9. Özet raporu

Ülke dosyası okunamazsa **kritiktir**: hem Dağılım hem Grafik sekmesini
besler, onsuz anlamlı rapor çıkmaz.

### `klasoru_tara(klasor) -> dict`
Klasördeki `.txt` dosyalarını sınıflandırır. Önce ada bakar
(`DOSYA_IPUCLARI`), ad ipucu vermezse `_basliktan_tur_bul` ile içeriğe
bakar.

### `dosyalari_dogrula(dosyalar) -> list[str]`
Seçilen dosyaların gerçekten o rapora ait olup olmadığını **başlıktan**
kontrol eder. Dosya adı her seferinde değişebildiği için ada güvenilmez.

Uyuşmazlık engel değil **uyarıdır** — SAP başlığı değişmiş olabilir.

### `yedek_al(xlsx_yolu) -> str`
Dosyanın yanındaki `yedek/` klasörüne zaman damgalı kopya alır.

### `_ozet_yaz(...)`
Kapanış özeti: GÜNCELLENDİ / GÜNCELLENMEDİ / UYARILAR.

### `konsol(argv=None) -> int`
Komut satırı. Çıkış kodunu döner.

### `_ana() -> int`
Giriş noktası. Argümansız → pencere, argümanlı → komut satırı.
Pencere açılamazsa hata kutusu gösterir; exe sessizce kapanmaz.

---

## 4.8 `gui.py`

### `class Uygulama`

| Metot | İşlev |
|---|---|
| `.dosya_sec(anahtar)` | Tek sekme için SAP dosyası seçtirir |
| `.klasor_sec()` | Klasörden dördünü otomatik doldurur |
| `.dosyalari_kontrol_et()` | İçeriğe bakıp ✓ / ! / ? işaretler |
| `.word_sec()` | Word dosyası seçtirir |
| `.tarih_oku() -> date` | Takvimden veya kutudan tarihi okur |
| `.tarihi_oner()` | SAP dosyasından bitiş tarihini doldurur |
| `.yaz(satir, etiket=None)` | Günlüğe yazar; etiket yoksa içeriğe göre renklendirir |
| `.calistir_bas()` | Doğrulama + onay + iş parçacığı başlatma |
| `._sorucu(baslik, ayrinti) -> bool` | İş parçacığından soru sordurur |
| `._is(xlsx, secilen, bitis, word)` | Arka planda `rapor.calistir` |

`._sorucu` mekanizması: Tkinter yalnızca ana iş parçacığından çağrılabilir.
`after()` bile iş parçacığından güvenli değildir. Soru kuyruğa bırakılır,
`kuyrugu_isle` ana döngüde sorar, cevap `threading.Event` ile beklenir.

---

# 5. KULLANIM

## 5.1 Pencere

`IhbarRapor.exe` çift tıkla.

1. **Excel raporu** — `.xlsx` seç
2. **SAP çıktıları** — dört satırın her biri için `.txt` seç
   (veya `Klasörden otomatik doldur`)
3. **Rapor bitiş tarihi** — takvimden gün seç (SAP'taki tarihle dolu gelir)
4. **Word raporunu da güncelle** — isteğe bağlı
5. **Aktar**

Dosya–sekme kontrolü işaretleri:

| İşaret | Anlam |
|---|---|
| ✓ | Başlık bu sekmeye uyuyor |
| ! | Dosya başka bir rapora benziyor |
| ? | Başlık tanınamadı |

## 5.2 Komut satırı

```bash
IhbarRapor.exe --rapor "rapor.xlsx" --sap "SAP klasoru" --bitis 17.08.2026
```

Dosya adları değişkense (dördü birlikte zorunlu):

```bash
IhbarRapor.exe --rapor "rapor.xlsx" --icerik "a.txt" --durum "b.txt" --sube "c.txt" --ulke "d.txt" --bitis 17.08.2026
```

| Seçenek | İşlev |
|---|---|
| `--rapor` | Excel dosyası (zorunlu) |
| `--sap` | SAP klasörü |
| `--icerik --durum --sube --ulke` | Dosyaları tek tek ver |
| `--bitis GG.AA.YYYY` | Rapor bitiş tarihi |
| `--word` | Word raporunu da güncelle |
| `--ad-degistirme` | Dosya adını değiştirme |
| `--sor` | Sorun çıkarsa ekrandan sor |
| `--devam` | Sorun çıkarsa sormadan devam et |
| `--hata-durdur` | (varsayılan) İlk sorunda iptal |

## 5.3 Çıkış kodları

| Kod | Anlam |
|---|---|
| 0 | Başarılı |
| 1 | Kritik hata — dosyalar değişmedi |
| 2 | Beklenmeyen hata |
| 3 | Kısmen başarılı — bazı bölümler yazılamadı |
| 4 | İptal edildi |

---

# 6. GÜNCELLENMEYEN ALANLAR

Program bunlara dokunmaz. Elle girilir.

| Konum | İçerik |
|---|---|
| `Yıllar Bazlı` E23–I23 | Erişime engellenen, EEK, mahkeme kararı, pasif/domain, devam eden |
| `Suç Duyuruları` yeni dönem satırı | Dava dosyası bazlı |
| `Mobil Uygulamalar` B4/C4 | Dava sonucu |
| `Sosyal Medya Hesapları` B4/C4 | Dava sonucu |
| `850'li Hatlar` B4/C4 | Dava sonucu |
| `Ödeme Kuruluşları` B4/B5 | Dava sonucu |
| `Faaliyet Cetveli` | Program bu sekmeye hiç dokunmaz |
| Geçmiş yıl sekmeleri (2023/2024/2025) | Tarihsel veri |

Bu sekmelerin TOPLAM satırları formüllüdür. Üstteki hücre girilince
toplam kendiliğinden düzelir ve Word'e yansır.

---

# 7. YIL DEĞİŞTİĞİNDE

`excel_writer.py` içinde iki sabit güncellenir:

```python
ULKE_SEKME = "2026 Ülke Dağılımı"
GRAFIK_SEKME = "2026 Ülke Grafiği"
```

Diğer sekmelerin adı yıldan bağımsızdır.

---

# 8. DERLEME

```bash
pip install pyinstaller pywin32 tkcalendar python-docx openpyxl
pyinstaller --clean --noconfirm ihbar_rapor.spec
```

Sonuç: `dist\IhbarRapor.exe` — tek dosya, 33 MB, yan dosya gerekmez.

`ihbar_rapor.spec` içinde `babel` yerel verileri elle toplanır; aksi
halde exe açılırken `unknown locale: tr_TR` hatası verir.

---

# 9. BİLİNEN KISITLAR

- Microsoft Excel kurulu olmalıdır. COM olmadan grafikler korunamaz.
- Word `.doc` (eski biçim) desteklenmez. `.docx` gerekir.
- `Yıllar Bazlı` E–I sütunları SAP'ta bulunmadığı için `E25` (erişime
  engellenen kümülatifi) elle güncellenene kadar eski kalır.
- Word p113'teki kök site sayısı ile `Suç Duyuruları` D26 formülü
  arasında bir fark tespit edilmiştir.
  Program Excel'i esas alır.
- Makine yoğunken (tam ekran uygulama vb.) art arda Excel COM çağrıları
  yavaşlayabilir.
