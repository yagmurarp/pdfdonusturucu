# Dönüştürücü Web Uygulaması (Flask)

**Yapabildikleri:**
- PDF → Word (DOCX)
- Word (DOCX) → PDF
- PDF → Excel (XLSX)  *(tablo içeren PDF'lerde)*
- Excel (XLSX) → PDF
- Word (DOCX) → Excel (XLSX)  *(paragrafları ve tabloları sayfalara döker)*
- Excel (XLSX) → Word (DOCX)  *(sayfaları tabloya çevirir)*

> Not: Tarayıcıdan yüklediğin dosyalara göre çıktı dosyasını indirmen için bağlantı verir.

---

## 0) Sistem Gereksinimleri (tek seferlik kurulum)
### macOS
1. **Python 3** yüklü (Terminal: `python3 --version`)
2. **LibreOffice** (Word/Excel → PDF için)
   - https://www.libreoffice.org/download/ adresinden .dmg indir, uygulamayı **Applications** klasörüne sürükle.
3. **Ghostscript** (PDF → Excel için)
   - Homebrew varsa: `brew install ghostscript`
   - Yoksa: https://ghostscript.com/releases/ adresinden macOS paketini indirip kur.

> Gerekirse `SOFFICE_PATH` ortam değişkeniyle LibreOffice yolunu verebilirsin:  
> `/Applications/LibreOffice.app/Contents/MacOS/soffice`

### Windows
1. **Python 3** yüklü (PowerShell: `py --version`)
2. **LibreOffice**: https://www.libreoffice.org/download/ üzerinden .msi indirip kur.
3. **Ghostscript**: https://ghostscript.com/releases/ üzerinden Windows installer indirip kur.

> Gerekirse `SOFFICE_PATH` ortam değişkenine örnek yol ver:  
> `C:\Program Files\LibreOffice\program\soffice.exe`

---

## 1) Projeyi kur
Konsolu **proje klasöründe** aç (bu dosyanın yanındayken).

### macOS (Terminal)
```bash
python3 -m venv venv
source venv/bin/activate
python -m pip install -U pip
pip install -r requirements.txt
```

### Windows (PowerShell)
```powershell
py -m venv venv
.env\Scripts\Activate.ps1
python -m pip install -U pip
pip install -r requirements.txt
```
> PowerShell "scripts disabled" derse:  
> `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` → Y ile onayla → tekrar `Activate.ps1`.

---

## 2) Çalıştır
Sanal ortam **aktifken**:
```bash
python app.py
```
Tarayıcı: `http://127.0.0.1:5000`

---

## 3) Kullanım notları
- **PDF → Excel**: PDF'in içinde gerçek metin tablosu olmalı. Tarama (görüntü) ise tablo çıkmayabilir. OCR gerekiyorsa haber ver.
- **Word ↔ Excel**: DOCX'teki tablolar ayrı Excel sayfalarına; paragraflar "Paragraflar" sayfasına yazılır. Excel → Word'de her sayfa bir tabloya dönüşür.
- **Word/Excel → PDF**: LibreOffice başlıksız (headless) çalıştırılır. Uygulama LibreOffice yolunu otomatik bulamazsa `SOFFICE_PATH` ver.

---

## 4) Sorun Giderme
- **LibreOffice bulunamadı**: Uygulama ana sayfada yolu gösterir; boşsa kur veya `SOFFICE_PATH` ayarla.
- **PDF → Excel boş geldi**: PDF gerçekten tablo içeriyor mu? Tarama ise OCR gerekebilir.
- **İzin / Güvenlik**: macOS Gatekeeper bazen ilk çalıştırmada uyarabilir; Sistem Ayarları → Güvenlik'ten izin ver.
- **Boyut limiti**: Varsayılan 200MB. `app.config['MAX_CONTENT_LENGTH']` ile değiştirilebilir.