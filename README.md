# Website Email Finder

Excel dosyasındaki firma web sitelerini tarayarak e-posta adreslerini otomatik olarak bulur ve sonuçları Excel dosyasına aktarır.

Girdi dosyası: `firmalar.xlsx`
Çıktı dosyası: `firmalar_mailli.xlsx`

## Kurulum

Gerekli Python paketlerini yüklemek için:

```bash
pip install requests pandas beautifulsoup4 openpyxl urllib3 pytesseract Pillow
```

OCR özelliği için Mac'te ayrıca Tesseract kurulmalıdır:

```bash
brew install tesseract
```
