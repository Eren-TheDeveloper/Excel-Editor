# Ürün Excel Otomasyonu

Mevcut Excel ürün listelerine, seçilen bir ürünü şablon alarak yeni ürünler ekleyen Python uygulaması. Tkinter masaüstü arayüzü ve komut satırı desteği sunar. Sonuç ayrı bir `.xlsx` dosyasına yazılır; kaynak ve giriş dosyalarının üzerine yazılması engellenir.

## Özellikler

- Şablon üründen açıklama, kategori, stok, görsel yolu ve diğer alanları devralma.
- Birden fazla ürünü tek işlemde ekleme ve hücre biçimlerini kopyalama.
- Otomatik ürün ID'si ve benzersiz `Kxxxxxx` model kodu oluşturma.
- Tekrar eden modelleri, bulunamayan şablonları ve geçersiz fiyat/stok girişlerini reddetme.
- Türkçe fiyat biçimi (`1.250,50`) ve ondalık nokta (`1250.50`) desteği.
- Ürün giriş Excel'i oluşturma ve doldurulmuş giriş dosyasından çıktı üretme.

## Gereksinimler ve kurulum

Python **3.14 veya üzeri** gerekir. Masaüstü arayüzü Tkinter kullanır; Windows Python kurulumunda Tcl/Tk bileşeni seçili olmalıdır. Aşağıdaki komutları indirdiğiniz deponun klasöründe çalıştırın.

### uv ile

`uv` kuruluysa:

```powershell
uv sync --locked
uv run python main.py
```

### Python ve pip ile (Windows)

```powershell
py -3.14 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -e .
.\.venv\Scripts\python.exe main.py
```

Kurulumdan sonra Windows'ta `urun_ekle.bat` dosyasına çift tıklayarak da başlatabilirsiniz. Aşağıdaki örneklerde `uv run python` yerine `.\.venv\Scripts\python.exe` kullanılabilir.

## Örnek veriyle deneme

Depoya gerçek ürün verileri dahil edilmez. Yapay veriler içeren bir kaynak oluşturmak için:

```powershell
uv run python examples/create_sample.py
uv run python main.py
```

İlk komut `examples/products.xlsx` dosyasını oluşturur; mevcutsa üzerine yazmaz. Arayüzden bu dosyayı seçip **Şablon modellerini yükle** düğmesine basın. Yeni ürün için örneğin `Yeni Defter`, `150` fiyatı ve `K000001` şablon modelini seçin. **Ürünü listeye ekle**, ardından **Yüklemeye hazır Excel oluştur** düğmesini kullanın.

Bilgisayarda eski `A12.xlsx` dosyası varsa uygulama başlangıçta onu seçer; aksi halde örnek dosyanın yolunu gösterir. İstediğiniz kaynak dosyayı **Gözat** ile seçebilirsiniz.

## Kendi Excel dosyanızla kullanım

1. Kaynak Excel'i seçip şablon modellerini yükleyin.
2. Ürün adı, fiyat ve şablon ürün modelini girin. Model kodu boşsa otomatik oluşturulur.
3. İsteğe bağlı alanları doldurun. Boş bırakılan stok, kategori, açıklama ve görsel yolu şablondan devralınır. Boş yayın tarihi bugünün tarihi olur.
4. Ürünleri listeye ekleyin ve sonuç dosyasını kaydedin.

Kaynak Excel'de **`Products`** sayfası ve **`Product ID`, `Name`, `Model`, `Price`** sütunları zorunludur. En az bir sayısal, tam sayı ürün ID'si ve kullanılabilir şablon modeli bulunmalıdır. `Quantity`, `Categories`, `Description`, `Image` gibi isteğe bağlı sütunlar mevcutsa doldurulur; eksik sütunlar eklenmez. Diğer sayfalar çıktı dosyasında korunur.

Fiyat sıfır veya pozitif, sonlu bir sayı; stok sıfır veya pozitif bir tam sayı olmalıdır. Tarih biçimi `YYYY-AA-GG` şeklindedir. Kategori ve diğer ID'ler tahmin edilmez; kaynağınızdaki değerler kullanılır.

Bu araç genel amaçlı bir Excel düzenleyicisi değildir; belirtilen ürün şemasını bekler. Örnek dosya uygulamayı denemek içindir. Gerçek ürün panelinize yüklemeden önce çıktıyı panelinizin içe aktarma şemasına göre kontrol edin. Var olan bir sonuç dosyasını seçerseniz üzerine yazılır.

## Komut satırı

Ürün giriş şablonu oluşturun:

```powershell
uv run python main.py sablon examples/products.xlsx urun_giris.xlsx
```

`urun_giris.xlsx` içindeki **Yeni Ürünler** sayfasını doldurup dosyayı kapatın. Ardından:

```powershell
uv run python main.py olustur examples/products.xlsx urun_giris.xlsx yuklemeye_hazir.xlsx
```

Yardım: `uv run python main.py --help`.

## Testler

```powershell
uv run python -m unittest discover -s tests -v
```

Testler geçici klasörlerde yapay Excel dosyaları kullanır. Ürün ekleme, kaynak dosyanın korunması, şablondan biçim ve değer kopyalama, benzersiz ID/model üretimi ve hatalı girişleri kontrol eder. GitHub Actions aynı testleri Windows üzerinde çalıştırır.

## Paylaşılan dosyalar

Excel dosyaları, yerel veri klasörleri, `row_info.json`, `row_header.json`, geçici çalışma dosyaları ve `.venv` Git tarafından yok sayılır. Bu dosyalar bilgisayarınızda kalır. Örnek veri, depodaki `examples/create_sample.py` koduyla yeniden oluşturulur.

Ana dosyalar: `main.py` (uygulama), `urun_ekle.bat` (Windows başlatıcısı), `tests/` (testler), `pyproject.toml` ve `uv.lock` (bağımlılıklar).

## Lisans

Proje [MIT lisansı](LICENSE) ile sunulur.
