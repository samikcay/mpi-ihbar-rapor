# İhbar Site Raporu — SAP Aktarım Programı

SAP'tan alınan metin çıktılarını okuyup **İhbar Site Rapor** çalışma
kitabındaki altı sekmeye yazar. İstenirse Word raporundaki sayıları da
günceller.

> **Bu depoda veri yoktur.** Yalnızca program kodu ve dökümantasyon
> bulunur. Excel raporları, Word belgesi, SAP çıktıları ve yedekler
> `.gitignore` ile dışarıda tutulur.

## Kurulum

Bilgisayarda **Microsoft Excel** kurulu olmalıdır. Program Excel'i COM
üzerinden kullanır; grafiklerin bozulmaması bu yüzden mümkün olur.

```bash
pip install -r requirements.txt
python rapor.py
```

Derlenmiş sürüm için [Releases](../../releases) bölümüne bakın.
`IhbarRapor.exe` tek dosyadır, yanında başka bir şey gerekmez.

## Kullanım

Pencereden:

```bash
python rapor.py
```

Komut satırından:

```bash
python rapor.py --rapor "rapor.xlsx" --sap "SAP klasoru" --bitis 17.08.2026
```

SAP dosyalarının adı her seferinde değişiyorsa dördünü tek tek verin:

```bash
python rapor.py --rapor "rapor.xlsx" \
  --icerik "a.txt" --durum "b.txt" --sube "c.txt" --ulke "d.txt" \
  --bitis 17.08.2026 --word "rapor.docx"
```

## Ne yapar

| Sekme | Kaynak |
|---|---|
| Site İçeriğine Göre | SAP |
| Site Durumuna Göre | SAP |
| Şube Bazlı | SAP |
| 2026 Ülke Dağılımı | SAP |
| 2026 Ülke Grafiği | Ülke Dağılımı'ndan hesaplanır (ilk 15 + Diğer) |
| Yıllar Bazlı | Cari yıl satırının B/C/D sütunları |

Ayrıca rapor dönemi tarihini bütün cari sekmelerde ve dosya adında
günceller.

Satır sayısı değişse bile sekmeyi uyarlar. Şube veya ülke sayısı artarsa
satır ekler, azalırsa siler. `TOPLAM` satırı, `=SUM()` formülleri, yüzde
oranları ve **grafikler** yeni aralığa kendiliğinden uyar.

## Güvenlik

- Her çalıştırmadan önce yedek alınır (`yedek/` klasörüne, tarih damgalı).
- Okuma sırasında hata olursa Excel dosyasına dokunulmaz.
- Varsayılan davranış: ilk sorunda **durur**. `--devam` ile sürdürülebilir.
- Seçilen dosyanın doğru rapora ait olup olmadığı içeriğinden kontrol edilir.

## Modüller

| Dosya | İşlevi |
|---|---|
| `rapor.py` | Ana akış, komut satırı, özet raporu |
| `gui.py` | Pencere arayüzü |
| `sap_parser.py` | SAP metin çıktılarını okur |
| `excel_writer.py` | Excel'e yazar (COM) |
| `word_writer.py` | Word raporunu günceller |
| `tarih.py` | Tarih biçimleri ve dosya adı |
| `labels.py` | Etiket yazımını düzeltir |
| `hatalar.py` | Hata sınıflandırma ve kurtarma |

Ayrıntılı teknik dökümantasyon: [`DOKUMANTASYON.md`](DOKUMANTASYON.md)
Kullanım kılavuzu: [`OKUBENI.md`](OKUBENI.md)

## Derleme

```bash
pip install pyinstaller
pyinstaller --clean --noconfirm ihbar_rapor.spec
```

Sonuç `dist/IhbarRapor.exe` olur. Tek dosya, yaklaşık 34 MB.
