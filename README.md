<div align="center">

![Excel Editor — Şablonunu seç. Ürünlerini ekle. Excel’in hazır.](docs/banner.svg)

# Excel Editor

**Ürün Excel Otomasyonu**

Mevcut ürününü şablon al, yeni ürünlerini ekle, içe aktarma dosyanı oluştur.

![Python 3.14+](https://img.shields.io/badge/Python-3.14%2B-3776AB?style=flat-square&logo=python&logoColor=white)
![Arayüz Tkinter](https://img.shields.io/badge/Aray%C3%BCz-Tkinter-14B8A6?style=flat-square)
![Dosya biçimi XLSX](https://img.shields.io/badge/Excel-XLSX-217346?style=flat-square)
[![Lisans MIT](https://img.shields.io/badge/Lisans-MIT-A78BFA?style=flat-square)](LICENSE)

[Hızlı başlangıç](#hızlı-başlangıç) · [Kullanım](#kullanım) · [Komut satırı](#komut-satırı) · [Testler](#testler)

</div>

---

## Ürün eklerken aynı bilgileri tekrar yazma

Yeni ürünün açıklaması, kategorisi, stok ayarları ve diğer alanları mevcut bir ürüne benziyorsa, o ürünü şablon olarak kullanabilirsin. Excel Editor seçtiğin satırı biçimleriyle birlikte kopyalar, girdiğin bilgileri uygular ve sonucu ayrı bir Excel dosyasına kaydeder.

**Masaüstü arayüzünden çalış veya aynı akışı komut satırından yürüt.**

```mermaid
flowchart LR
    A["📂 Kaynak Excel"] --> B["📋 Şablon ürün"]
    B --> C["✏️ Yeni ürünler"]
    C --> D["✅ Yeni Excel dosyası"]
    style A fill:#102c29,stroke:#34d399,color:#ecfdf5
    style B fill:#102c29,stroke:#34d399,color:#ecfdf5
    style C fill:#102c29,stroke:#34d399,color:#ecfdf5
    style D fill:#065f46,stroke:#6ee7b7,color:#ecfdf5
```

## Neler sunuyor?

| | Özellik | Ne işe yarar? |
| :---: | --- | --- |
| 🖥️ | **Masaüstü arayüzü** | Dosyanı seç, ürünlerini forma gir ve listeye ekle. |
| 📋 | **Şablondan çoğaltma** | Açıklama, kategori, stok ve diğer alanları mevcut üründen devral. |
| 🎨 | **Biçimleri koruma** | Şablon satırının hücre biçimlerini yeni ürüne taşı. |
| 🔢 | **Otomatik ID ve model** | Yeni ürün ID’leri ve benzersiz `Kxxxxxx` model kodları oluştur. |
| 🧩 | **Toplu ürün ekleme** | Birden fazla ürünü aynı çıktı dosyasına ekle. |
| 🛡️ | **Giriş kontrolleri** | Tekrar eden modelleri, eksik şablonları ve geçersiz fiyat/stok girişlerini yakala. |
| 🇹🇷 | **Türkçe fiyat desteği** | Hem `1.250,50` hem `1250.50` biçimini kullan. |
| ⌨️ | **Komut satırı** | Excel giriş şablonu oluştur ve doldurulmuş dosyayı işle. |

## Hızlı başlangıç

**Gerekenler:** Python **3.14+** ve Tkinter. Windows’ta Python kurulurken **Tcl/Tk** bileşeni seçili olmalıdır.

Depoyu indir veya klonla; ardından proje klasöründe çalıştır:

```powershell
uv sync --locked
uv run python examples/create_sample.py
uv run python main.py
```

Bu komutlar bağımlılıkları kurar, yapay ürün verisi içeren `examples/products.xlsx` dosyasını oluşturur ve uygulamayı açar. Örnek dosya zaten varsa üretici üzerine yazmaz.

<details>
<summary><strong>uv kullanmıyorum — Python ve pip ile kurulum</strong></summary>

Windows’ta proje klasöründe:

```powershell
py -3.14 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -e .
.\.venv\Scripts\python.exe examples/create_sample.py
.\.venv\Scripts\python.exe main.py
```

Aşağıdaki tüm örneklerde `uv run python` yerine `.\.venv\Scripts\python.exe` kullanabilirsin.

</details>

> 💡 **Windows’ta tek tıkla aç:** Kurulumdan sonra `urun_ekle.bat` dosyasına çift tıklaman yeterli.

## Kullanım

### 1 · Kaynağını seç

**Gözat** ile ürün Excel’ini seç ve **Şablon modellerini yükle** düğmesine bas. İlk deneme için `examples/products.xlsx` dosyasını kullanabilirsin.

### 2 · Yeni ürününü tanımla

Örnek veriyi kullanıyorsan şu bilgilerle başlayabilirsin:

| Alan | Örnek değer |
| --- | --- |
| Ürün adı | `Yeni Defter` |
| Fiyat | `150` |
| Şablon ürün modeli | `K000001` |
| Yeni model kodu | Otomatik üretim için boş bırak |
| Stok, kategori, açıklama, görsel yolu | Şablondan devralmak için boş bırak |

### 3 · Listeye ekle ve kaydet

**Ürünü listeye ekle** düğmesine bas. Başka ürünler için işlemi tekrarla; ardından **Yüklemeye hazır Excel oluştur** ile sonuç dosyanı kaydet.

Kaynak ve giriş dosyalarının üzerine yazılması engellenir. **Var olan başka bir sonuç dosyasını seçersen o dosyanın üzerine yazılır.**

<details>
<summary><strong>Başlangıçta hangi dosya seçilir?</strong></summary>

Proje klasöründe `A12.xlsx` varsa uygulama başlangıçta onu seçer. Aksi halde `examples/products.xlsx` yolunu gösterir. Kaynağı her zaman **Gözat** ile değiştirebilirsin.

</details>

## Excel dosyan nasıl olmalı?

Kaynak dosyada **`Products`** adlı bir sayfa ve aşağıdaki sütunlar bulunmalı:

| Zorunlu sütun | İçerik |
| --- | --- |
| `Product ID` | Mevcut ürünün sayısal, tam sayı ID’si |
| `Name` | Ürün adı |
| `Model` | Şablon olarak seçilecek model kodu |
| `Price` | Ürün fiyatı |

En az bir kullanılabilir ürün ID’si ve şablon modeli gerekir. `Quantity`, `Categories`, `Description`, `Image` gibi isteğe bağlı sütunlar varsa kullanılır; eksik sütunlar otomatik eklenmez. Diğer sayfalar çıktı dosyasında korunur.

<details>
<summary><strong>Alan kuralları ve şablondan devralma</strong></summary>

- **Zorunlu girişler:** Ürün adı, fiyat ve şablon ürün modeli.
- **Model:** Boşsa benzersiz bir `Kxxxxxx` kodu oluşturulur; mevcut bir kod tekrar kullanılamaz.
- **Fiyat:** Sıfır veya pozitif, sonlu bir sayı olmalıdır.
- **Stok:** Sıfır veya pozitif bir tam sayı olmalıdır; boşsa şablondan alınır.
- **Yayın tarihi:** `YYYY-AA-GG` biçiminde girilir; boşsa bugünün tarihi kullanılır.
- **Kategori, açıklama ve görsel yolu:** Boş bırakılırsa şablondan devralınır.
- **Kategori ve diğer ID’ler:** Tahmin edilmez; kaynak dosyadaki veya senin girdiğin değerler kullanılır.

</details>

> **Uyumluluk:** Uygulama burada belirtilen ürün şemasını bekler. Örnek dosya deneme içindir; gerçek ürün paneline yüklemeden önce çıktının panelinin içe aktarma şemasına uyduğunu kontrol et.

## Komut satırı

**Giriş şablonunu oluştur:**

```powershell
uv run python main.py sablon examples/products.xlsx urun_giris.xlsx
```

`urun_giris.xlsx` içindeki **Yeni Ürünler** sayfasını doldur, kaydet ve dosyayı kapat.

**Sonuç Excel’ini üret:**

```powershell
uv run python main.py olustur examples/products.xlsx urun_giris.xlsx yuklemeye_hazir.xlsx
```

**Komutları görüntüle:**

```powershell
uv run python main.py --help
```

## Testler

```powershell
uv run python -m unittest discover -s tests -v
```

Testler geçici klasörlerde yapay verilerle çalışır. Ürün ekleme, kaynak dosyanın korunması, şablondan değer ve biçim kopyalama, benzersiz ID/model üretimi ve hatalı girişler denetlenir. GitHub Actions iş akışı bu testleri Windows üzerinde çalıştıracak şekilde yapılandırılmıştır.

## Proje yapısı

```text
excel-editor/
├── main.py                   # Masaüstü arayüzü ve Excel işlemleri
├── urun_ekle.bat              # Windows başlatıcısı
├── examples/
│   └── create_sample.py       # Yapay örnek veri üreticisi
├── tests/
│   └── test_workbooks.py      # Excel işlemlerinin testleri
├── docs/
│   └── banner.svg             # README kapak görseli
├── .github/workflows/
│   └── tests.yml              # Otomatik test iş akışı
├── pyproject.toml             # Proje ve bağımlılık tanımları
├── uv.lock                    # Kilitlenmiş bağımlılıklar
└── LICENSE                    # MIT lisansı
```

### Gerçek veriler yerelde kalır

Excel dosyaları, yerel veri klasörleri, `row_info.json`, `row_header.json`, geçici çalışma dosyaları ve `.venv`, `.gitignore` ile Git kapsamı dışında tutulur. Örnek dosyayı depodaki üreticiyle kendin oluşturabilirsin.

## Katkıda bulunma

Bir hata bulduysan kullandığın komutu veya işlem adımlarını, beklenen sonucu ve hata mesajını bir issue’da paylaşabilirsin. Örnek dosya gerekiyorsa gerçek ürün bilgileri yerine yapay veri kullan. Kod değişikliklerinde ilgili testleri çalıştırıp bir pull request açabilirsin.

## Lisans

[MIT](LICENSE) · Kullanabilir, değiştirebilir ve lisans koşullarıyla paylaşabilirsin.
