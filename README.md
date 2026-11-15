# NÖBETRA

![NÖBETRA logosu](Nobetra.png)

**Geliştirici · Salih Balseven**

Windows için Türkçe bir masaüstü uygulaması. Devriye, Yatakhane Kat ve Yatakhane Giriş nöbetlerini hazırlamak, sonuçlarını kaydetmek ve geçmiş çizelgeleri incelemek için geliştirildi. Python ve Tkinter arayüzü kullanır. Nöbet kayıtları yerel SQLite veritabanında tutulur.

## Kurulum ve demo

Windows’ta Python 3.10 veya üstü ve Tkinter gerekir. VS Code’da depo klasörünü aç, **Terminal → New Terminal** seç ve sırayla çalıştır.

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe demo_verisi.py
.\.venv\Scripts\python.exe app.py
```

Demo, 16 sınıfa dağıtılmış 256 kurgusal kayıt, bir örnek ceza ve üç nöbet türünün başlangıç sırasını oluşturur. Öğrenci numaraları `DEMO-001` biçimindedir. Demo komutu mevcut bir veritabanı varsa durur. Boş kullanım için demo komutunu atla ve doğrudan `app.py` dosyasını çalıştır. Uygulama yerel `veri/nobet.db` dosyasını oluşturur. Öğrencilerini ekledikten sonra **Sıra başlangıçları** ekranında otomatik planlamanın başlayacağı öğrencileri seç. `baslat.py` mevcut kurulumun üst klasördeki veritabanını açan başlatıcıdır.

## Uygulama ekranları

| Ekran | Yapabileceğin işlemler |
| --- | --- |
| Öğrenciler | Sınıfa göre filtrele, tekil kayıt ve profil aç veya boş Excel şablonuyla toplu aktarım yap. |
| Nöbet Hazırla | Tarih ve nöbet türünü seç. Otomatik veya manuel çizelge oluştur, taslağı kontrol et, görevleri değiştir ve listeyi onayla. |
| Nöbet Geçmişi | Tarihe ve nöbet türüne göre kayıtları; asıl, yedek ve görev sonucu bilgilerini incele. |
| İstatistikler | Dönem, görev türü ve saat aralığına göre toplamları, tabloları ve grafikleri aç. |
| Cezalılar | Okul numarası ya da ad soyad ile ceza kayıtlarını getir, nöbet adedi ve görevlendirme durumunu izle. |
| Muafiyetler | Öğrenciler için tarih aralığına bağlı muafiyetleri yönet. |
| Sınıf Yasakları | Yatakhane Giriş nöbetinde günlere göre sınıf kısıtlarını düzenle. |
| Yedekleme | SQLite veritabanını ayrı bir dosyaya yedekle. |
| Asistan | Program içi kullanım ipuçlarını aç. Codex ile soru-cevap için ayrıca kurulu ve oturumu açık bir Codex gerekir. Yerel güncelleme paketleme özelliği ayrıca PyInstaller yapılandırması ister. |

## Nöbet hazırlama ve onaylama

1. **Nöbet Hazırla** ekranında tarihi ve nöbet türünü seç.
2. **Otomatik hazırla** ile sıra ve geçmişe göre önizleme oluştur. Varsa mazeretli öğrencileri çıkar, tekrar değerlendirilecek asılları ve uygun ceza görevlerini seç. Ceza vardiyalarını önizlemede değiştirebilirsin.
3. **Manuel hazırla** ekranında öğrencileri tek tek veya toplu seç. Excel satırlarını yapıştırabilir, öğrencileri boş sıralara sırayla doldurabilir veya belirli bir saati ve görevi seçerek yerleştirebilirsin. Hazırlanmış taslakta satırdaki **×** mevcut saat, yer ve rol için öğrenci değiştirmeni sağlar.
4. Devriye ve Yatakhane Kat aynı gece aynı sınıfa yazılamaz. Bu kontrol ceza görevlendirmesinde de geçerlidir. Sınıf devam sırası ekranında yaklaşan çakışmaları kontrol edip bir sonraki sınıf geçişini planla.
5. Gece vardiyaları 23.00–01.00, 01.00–03.00, 03.00–05.00 ve 05.00–07.00 olarak gösterilir. İlk vardiya çizelge tarihinde, gece yarısından sonra başlayan vardiyalar ertesi takvim günündedir.
6. Görev türüne göre asıl ve yedek satırlarını kontrol et.
7. **Yazdır / Excel / PDF** menüsünden Excel dosyası al veya yazdırma ekranında PDF kaydet. Sonuçları **Tuttu** veya **Tutmadı** olarak işle. Liste eksikse onay tamamlanana kadar kapalıdır.

Normal nöbetlerin saat dağılımı geçmiş normal görevleri dikkate alır. Manuel seçim ekranındaki `3/1` gösterimi, o saatte tutulan üç nöbetten birinin ceza nöbeti olduğunu belirtir. Bir ceza görevinin kalan adedi, görev **Asıl**, **Tuttu** ve çizelge **Onaylandı** olduğunda azalır.

## Eski çizelgeyi içe aktarma

**Nöbet Hazırla → Geçmiş nöbet listesi yükle** yolundan uygun türü seç ve Excel ya da fotoğrafı aç. Fotoğrafta tablonun dört dış köşesini işaretle. Tarihi, saatleri, öğrenci eşleşmelerini ve görev sonuçlarını önizlemede kontrol et. Fotoğraf okuma Windows OCR bileşenlerine bağlıdır. Tarih ve eşleşmeleri kaydetmeden önce ekrandan düzeltebilirsin.

## Kaynak kodu

| Dosya | İçerik |
| --- | --- |
| `app.py` | Uygulama başlangıcı, SQLite şeması ve ana ekran akışı |
| `nobet_planlayici.py` | Otomatik seçim, sınıf sırası, çakışma ve vardiya dağılımı |
| `nobet_manuel.py` | Manuel çizelge düzenleme |
| `ceza_nobet.py` | Ceza kayıtları, vardiya önerileri ve kalan görev hesapları |
| `nobet_arsiv.py`, `nobet_arsiv_ui.py`, `nobet_fotograf.py` | Geçmiş çizelge aktarımı, önizleme ve fotoğraf okuma |
| `nobet_yazdir.py`, `excel_*.py` | Çizelge şablonları ve çıktı işlemleri |
| `ogrenci_profili.py`, `ogrenci_excel.py` | Öğrenci profili ve Excel ile toplu aktarım |
| `sistem_islemleri.py`, `pencere_yonetimi.py` | Yedekleme ve pencere yönetimi |

## Veri gizliliği

Demo komutu kurgusal kayıtları bilgisayarında üretir. Öğrenci veritabanı, gerçek çizelgeler, öğrenci fotoğrafları, yedekler ve yerel paketleme ayarları depoya dahil edilmez. Tanıtım ekran görüntülerini yalnızca demo veritabanıyla al. Asistanın soru alanında kişisel bilgi veya canlı nöbet çizelgesi kullanma.
