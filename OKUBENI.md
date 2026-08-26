# İhbar Site Raporu — SAP Aktarım Programı

SAP'tan alınan metin çıktılarını okuyup **İhbar Site Rapor** çalışma kitabındaki
dört sekmeye otomatik yazar:

| Sekme | Kaynak |
|---|---|
| Site İçeriğine Göre | SAP — içerik bazında ihbar edilen site sayısı |
| Site Durumuna Göre | SAP — site durum sayısı |
| Şube Bazlı | SAP — şube bazında ihbar edilen site sayısı |
| 2026 Ülke Dağılımı | SAP — ihbar edilen sitelerin ülke bazında dağılımı |
| **2026 Ülke Grafiği** | **Ülke Dağılımı sekmesinden hesaplanır** (ilk 15 ülke + Diğer) |

Ayrıca rapor dönemi tarihini bütün cari sekmelerde ve **dosya adında** günceller.
İsteğe bağlı olarak **Word raporundaki** (Sanal Kumar ve Yasa Dışı Bahisle
Mücadele Çalışmaları) SAP kaynaklı sayıları da günceller.

## Kurulum

### A) EXE ile (önerilen - Python gerekmez)

`dist\IhbarRapor.exe` **tek dosyadır**; yanında hiçbir DLL, kütüphane veya
klasör gerekmez. Kopyalayıp çift tıklamanız yeterlidir.

Tek şart: bilgisayarda **Microsoft Excel** kurulu olmalıdır.

> Kurulum sihirbazına gerek yoktur — program tek dosya olarak dağıtılır.
> İsterseniz masaüstüne kısayol oluşturabilirsiniz.

Komut satırından da kullanılabilir:

```bash
IhbarRapor.exe --rapor "rapor.xlsx" --sap "SAP klasoru" --bitis 17.08.2026
```

### B) Python ile (geliştirme)

1. Bilgisayarda **Microsoft Excel** kurulu olmalıdır.
2. Python 3 kurulu değilse python.org'dan kurun; kurulumda
   **"Add Python to PATH"** kutusunu işaretleyin.

Gerekli bileşenler (`pywin32`, `tkcalendar`, `python-docx`) ilk çalıştırmada otomatik
kurulur. Elle kurmak isterseniz:

```bash
pip install pywin32 tkcalendar python-docx
```

> `tkcalendar` kurulamazsa program yine çalışır; takvim yerine tarihi
> elle yazacağınız bir kutu görünür.

## Kullanım

`Rapor Aktar.bat` dosyasına çift tıklayın.

1. **Excel raporu**: `İhbar Site Rapor (...).xlsx` dosyasını seçin.
2. **SAP çıktıları**: Her sekmenin yanındaki **Seç...** düğmesiyle o sekmeye
   yazılacak `.txt` dosyasını gösterin. Dört satır vardır:

   | Satır | Yazılacağı sekme |
   |---|---|
   | Site İçeriğine Göre | Site İçeriğine Göre |
   | Site Durumuna Göre | Site Durumuna Göre |
   | Şube Bazlı | Şube Bazlı |
   | Ülke Dağılımı + Grafiği | 2026 Ülke Dağılımı **ve** 2026 Ülke Grafiği |

   Dosyalar hep aynı klasördeyse **Klasörden otomatik doldur...** ile dördünü
   birden doldurup gerekirse tek tek düzeltebilirsiniz.
3. **Rapor bitiş tarihi**: Takvim, SAP çıktısındaki tarihle kendiliğinden
   dolar. Değiştirmek için kutunun sağındaki oka (veya **Takvim** düğmesine)
   basıp raporun son gününü seçin. Takvim Türkçedir, hafta pazartesi başlar;
   üstteki oklarla ay ve yıl değiştirilir.
4. **Aktar** düğmesine basın; onay penceresi ne yapılacağını özetler.

Seçtiğiniz yollar hatırlanır; sonraki kullanımda tekrar seçmeniz gerekmez.

### Dosya–sekme kontrolü

SAP dosyalarının adı her seferinde değişebildiği için program **ada değil,
dosyanın içindeki başlığa** bakar. Her satırın sağında bir işaret çıkar:

| İşaret | Anlamı |
|---|---|
| ✓ (yeşil) | Dosyanın başlığı bu sekmeye uyuyor |
| **!** (kırmızı) | Dosya başka bir rapora benziyor — muhtemelen karışmış |
| **?** (turuncu) | Başlık tanınamadı (SAP başlığı değişmiş olabilir) |

`!` veya aynı dosyanın iki sekmeye verilmesi durumunda program **Aktar'a
basınca sorar**; onaylarsanız yine de devam eder. Yani kontrol engel değil,
uyarıdır — SAP başlığı değişirse programa takılıp kalmazsınız.

## Word raporu (isteğe bağlı)

**"Word raporunu da güncelle"** kutusunu işaretleyip `.docx` dosyasını
seçerseniz, Excel bittikten sonra Word raporu da güncellenir. Dosyanın
**üzerine yazılır**, öncesinde `yedek/` klasörüne kopyası alınır.

### Güncellenen yerler

| Yer | Ne değişir |
|---|---|
| Başlık | `(01 Ocak-17 Ağustos 2026)` |
| Metindeki dönemler | `01 Ocak-17 Ağustos`, `01.01.2026-17.08.2026` |
| Tespit edilen site sayısı | `...döneminde ise N web sitesi...` |
| Ülke dağılımı yüzdeleri | `<Ülke>'nin %XX, <Ülke>'nin %XX, ...` |

Ülke yüzdeleri tek tek yuvarlandığında toplam 100'ü aşabilir; bu fark
**"kalan %..."** dilimine yansıtılır, böylece toplam her zaman %100 olur
(Excel'deki `Diğer` hücresiyle aynı mantık).

### Excel'den gelen sayılar

Bazı sayılar SAP çıktısında yoktur ama **çalışma kitabında** vardır
(personelin elle doldurduğu kaynak sekmeler). Program bunları Excel'den
okuyup Word'e taşır:

| Word cümlesi | Excel kaynağı |
|---|---|
| `2006-2026 döneminde toplam ...` | `Yıllar Bazlı` **B25** |
| `... erişime engellenmesi sağlanmıştır` | `Yıllar Bazlı` **E25** |
| `... adet mobil uygulama` | `Mobil Uygulamalar` **B4** |
| `... kök siteden türetilmiş ...` | `Suç Duyuruları` **D26 / C26** |
| `... hattın (850'li numaralar)` | `850'li Hatlar` **B4** |
| `... banka/ödeme/elektronik para hesabı` | `Faaliyet Cetveli` **H9** |

### Kümülatif zincir

`Yıllar Bazlı` sekmesindeki **cari yıl satırı** kümülatif toplamları besler.
Program bu satırın SAP'tan gelen sütunlarını otomatik günceller:

| Sütun | İçerik | Kaynak |
|---|---|---|
| B | BTK'ya ihbar edilen toplam | **SAP** (otomatik) |
| C | Direk oynatan | **SAP** (otomatik) |
| D | Reklam, tanıtım ve yönlendirme | **SAP** (otomatik) |
| E–I | Erişime engellenen, EEK, mahkeme kararı, pasif/domain, devam eden | **elle** |

B sütunu güncellenince `B25` (kümülatif tespit) kendiliğinden yeniden
hesaplanır ve Word'deki cümleye yansır. **E–I sütunları SAP çıktılarında
bulunmadığı için elle girilmelidir**; `E25` (erişime engellenen) bu yüzden
siz o sütunları güncelleyene kadar eski kalır.

### Elle kalan sayılar

Program şunlara **dokunmaz**, işlem günlüğünde hatırlatır:

- `Yıllar Bazlı` E–I sütunlarına bağlı "erişime engellenen" sayısı
- Reklam Kurulu hesap sayıları
- Geçmiş yıllara (2024/2025) ait suç duyurusu ve sosyal medya sayıları

## Tarih güncelleme

Başlangıç **01.01.** olarak kalır, yalnızca bitiş günü değişir.
`01.01.2026 - 31.07.2026` → `01.01.2026 - 07.08.2026`

Güncellenen yerler:

- Dört SAP sekmesi + `2026 Ülke Grafiği` başlıkları
- `Yıllar Bazlı` cari yıl satırının B/C/D sütunları (kümülatif toplamları besler)
- `Yıllar Bazlı` sekmesindeki cari yıl satırı (orada tarih alt alta yazılıdır)
- **Dosya adı** — ay adıyla: `(01 Ocak-07 Ağustos 2026)`

**Dokunulmayan yerler:** geçmiş yıl sekmeleri (2023 / 2024 / 2025) ve
`Suç Duyuruları` sekmesi. Oradaki tarihler kendi dönemlerine aittir.

### Dosya adı

Dosya, yeni dönem adıyla **kopyalanır**; eski dosya yerinde kalır.
Örnek: `...(01 Ocak-31 Temmuz 2026).xlsx` dosyasından
`...(01 Ocak-07 Ağustos 2026).xlsx` oluşur.

Adın değişmesini istemiyorsanız komut satırında `--ad-degistirme` kullanın.

## Ülke Grafiği sekmesi

`2026 Ülke Dağılımı` sekmesindeki **ilk 15 ülke** aynen yazılır; 16. satıra
`Diğer` olarak *(genel toplam − ilk 15'in toplamı)* yazılır. Bu sekmenin
yerleşimi sabittir (satır eklenmez), pasta grafikleri olduğu gibi kalır.

15'ten az ülke gelirse kalan satırlar boşaltılır, `Diğer` sıfır olur.

### SAP dosyaları

Dosyaları SAP'tan olduğu gibi kaydetmeniz yeterlidir; adının ne olduğu
önemli değildir, düzenlemeye de gerek yoktur. Hangi dosyanın hangi sekmeye
yazılacağını siz seçersiniz; program dosyanın içindeki başlığa bakıp
seçiminizi kontrol eder (bkz. *Dosya–sekme kontrolü*).

## Program ne yapar

- Satır sayısı değişse bile sekmeyi uyarlar: **şube veya ülke sayısı artarsa
  satır ekler, azalırsa siler.** `TOPLAM` satırı, `=SUM()` formülleri, yüzde
  oranları ve **grafikler** otomatik olarak yeni aralığa uyar.
- Ülke sekmesindeki iki sütunlu düzeni korur (sol sütun `1..n/2`,
  sağ sütun kalanlar) — tek sayıda ülke varsa fazlalık sol sütuna gider.
- Başlıktaki dönem bilgisini (`01.01.2026 - 31.07.2026`) SAP çıktısındaki
  tarihe göre günceller.
- SAP'ta Türkçe karakterler kaybolduğunda (`BTK ENGELLI`) rapordaki doğru
  yazımı (`BTK ENGELLİ`) kullanır. Listede olmayan **yeni** bir şube, ülke
  veya durum gelirse SAP'taki hâliyle yazılır — veri kaybolmaz.

## İşlem günlüğü

Her çalıştırmanın sonunda bir **ÖZET** bölümü yazılır:

- **GÜNCELLENDİ** — hangi dosyanın hangi sekmesi/paragrafı değişti
- **GÜNCELLENMEDİ** — elle girilmesi gereken yerler, *nerede düzeltileceğiyle*
  birlikte (örn. `Yıllar Bazlı > E23-I23`)
- **UYARILAR** — tarih uyuşmazlığı, okunamayan sekme gibi durumlar

Bir sekme bulunamazsa program durmaz; ilgili sayının **eski değeriyle kaldığını**
`!` işaretiyle bildirir. Pencerede bu satırlar renkli gösterilir.

## Hata yönetimi

Program **çökmez** ve bir sorun yüzünden **tüm iş iptal olmaz**. Her sorun
üç sınıftan birine girer:

| Sınıf | Ne olur |
|---|---|
| **Kritik** | Devam edilemez (dosya yok/bozuk, ülke verisi okunamadı). İşlem durur, sebep yazılır, **dosyalar değişmez**. |
| **Sorulur** | Devam edilebilir (bir sekme yazılamadı, bir SAP dosyası bozuk). Pencerede sorulur: *"Bu bölüm atlanarak devam edilsin mi?"* |
| **Bilgi** | Kendiliğinden çözülür, sadece özete yazılır. |

Her sekme **bağımsız** yazılır: biri hata verse bile diğerleri yazılmaya
devam eder. Örneğin `Şube Bazlı` sekmesinin TOPLAM formülü silinmişse,
o sekme atlanır, kalan beş sekme güncellenir ve özette
`! YAZILAMAYAN sekmeler: Şube Bazlı` diye bildirilir.

Word hatası Excel'i geçersiz kılmaz — Excel zaten kaydedilmiştir.

### Çıkış kodları (zamanlanmış görev için)

| Kod | Anlamı |
|---|---|
| 0 | Başarılı (uyarı olabilir) |
| 1 | Kritik hata — dosyalar değişmedi |
| 2 | Beklenmeyen hata — ayrıntı ekrana yazılır |
| 3 | Kısmen başarılı — bazı bölümler yazılamadı |
| 4 | İptal edildi |

### Komut satırı seçenekleri

| Seçenek | İşlevi |
|---|---|
| *(varsayılan)* | Sorun çıkarsa **sormadan devam eder**, özete yazar |
| `--sor` | Her sorunda ekrandan sorar (E/h) |
| `--hata-durdur` | İlk sorunda işi iptal eder |

> Zamanlanmış görevde soru sorulamayacağı için varsayılan **devam
> etmektir**; böylece görev takılıp kalmaz, ama her şey günlüğe yazılır.

## Güvenlik ve kontroller

- **Her çalıştırmadan önce yedek alınır.** Yedekler Excel dosyasının yanındaki
  `yedek` klasörüne tarih-saat damgasıyla kaydedilir.
- Her dosyanın satır toplamı, SAP'ın kendi `TOPLAM` satırıyla karşılaştırılır.
- Dört raporun genel toplamı birbirini tutmazsa **uyarı verilir** — bu
  genellikle dosyalardan birinin farklı bir döneme ait olduğunu gösterir.
  Uyarılar günlükte `!` ile işaretlenir; lütfen sonucu kontrol edin.
- Seçilen dosyanın gerçekten o rapora ait olup olmadığı içeriğinden
  kontrol edilir; karışıklık varsa Aktar'a basınca sorulur.
- Okuma sırasında hata olursa Excel dosyasına **hiç dokunulmaz**.

## Komut satırından kullanım

Zamanlanmış görev için:

```bash
python rapor.py --rapor "C:\yol\İhbar Site Rapor.xlsx" --sap "C:\yol\SAP text outputs" --bitis 07.08.2026
```

Dosya adları değişkense her raporu tek tek verebilirsiniz (dördü de zorunlu):

```bash
python rapor.py --rapor "rapor.xlsx" --icerik "a.txt" --durum "b.txt" --sube "c.txt" --ulke "d.txt" --bitis 07.08.2026
```

`--bitis` verilmezse SAP çıktısındaki tarih kullanılır.
Başarılıysa `0`, veri/dosya/tarih hatasında `1`, beklenmeyen hatada `2` döner.

Diğer seçenekler:

| Seçenek | İşlevi |
|---|---|
| `--bitis GG.AA.YYYY` | Rapor bitiş tarihi |
| `--icerik/--durum/--sube/--ulke` | Her raporun .txt yolu (dördü birlikte) |
| `--word yol.docx` | Word raporunu da güncelle |
| `--ad-degistirme` | Dosya adını değiştirme, mevcut dosyaya yaz |
| `--ulke-sekmesi` | Ülke dağılımı sekmesinin adı |
| `--grafik-sekmesi` | Ülke grafiği sekmesinin adı |

## Yıl değiştiğinde

Program varsayılan olarak **`2026 Ülke Dağılımı`** ve **`2026 Ülke Grafiği`**
sekmelerine yazar. 2027 raporuna geçildiğinde çalışma kitabında yeni sekmeler
açıldıktan sonra `excel_writer.py` içindeki şu iki satırı güncelleyin:

```python
ULKE_SEKME = "2026 Ülke Dağılımı"
GRAFIK_SEKME = "2026 Ülke Grafiği"
```

Diğer üç sekmenin adı yıldan bağımsızdır, değişiklik gerekmez.

## Dosyalar

| Dosya | İşlevi |
|---|---|
| `Rapor Aktar.bat` | Programı başlatır |
| `rapor.py` | Ana akış, komut satırı |
| `gui.py` | Pencere arayüzü |
| `sap_parser.py` | SAP metin çıktılarını okur |
| `excel_writer.py` | Excel'e yazar (COM) |
| `labels.py` | Etiket yazımlarını düzeltir |
| `tarih.py` | Tarih biçimleri ve dosya adı |
| `word_writer.py` | Word raporunu günceller |
| `hatalar.py` | Hata sınıflandırma ve kurtarma |

## Sık karşılaşılan durumlar

**"Sekme bulunamadi"** — Çalışma kitabında sekme adı değişmiş olabilir.
Hata mesajı mevcut sekmeleri listeler.

**"TOPLAM satiri (=SUM) bulunamadi"** — Sekmedeki toplam satırının formülü
silinmiş demektir. Yedekten geri alın veya `=SUM()` formülünü geri koyun.

**Program takılı kalırsa** — Açık kalmış bir Excel penceresi engelliyor
olabilir. Excel'i kapatıp tekrar deneyin.
