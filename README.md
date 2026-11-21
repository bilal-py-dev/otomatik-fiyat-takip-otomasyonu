# 🛒 Otomatik Fiyat Takip & Excel Analiz Sistemi (Python)

E-ticaret sitelerindeki ürünlerin fiyatlarını anlık olarak takip eden, hedef fiyata ulaşıldığında Windows masaüstü bildirimi gönderen ve verileri biçimlendirilmiş Excel (.xlsx) raporu olarak dışa aktaran Python otomasyon aracı.

## 🚀 Öne Çıkan Özellikler
- **Web Scraping Motoru:** `requests` ve `BeautifulSoup4` ile hızlı ürün adı ve fiyat çekme
- **Kalıcı Veritabanı:** `JSON` tabanlı yerel veri saklama ve tekrarlı bildirim koruması
- **Anlık Masaüstü Bildirimi:** Hedef fiyat yakalandığında Windows uyarı sistemi (`plyer`)
- **Otomatik Excel Raporlama:** `openpyxl` ile renkli ve tarih damgalı `.xlsx` tablo çıktısı

## 🛠️ Kurulum ve Çalıştırma
py -m pip install requests beautifulsoup4 plyer openpyxl
py fiyat_bot.py
