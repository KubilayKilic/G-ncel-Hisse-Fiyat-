# README

Bu kod, belirli hisse senedi kodlarını bir Excel dosyasından alır, Yahoo Finance API'sini kullanarak her hisse senedi için güncel verileri alır ve son olarak bu verileri aynı Excel dosyasına kaydeder.

## Kullanım

1. **Python Kurulumu**: Bu kod Python programlama dili kullanılarak yazılmıştır. Python'un bilgisayarınızda yüklü olması gerekmektedir. Python'un resmi web sitesinden indirip yükleyebilirsiniz: [python.org](https://www.python.org/).

2. **Gerekli Kütüphaneler**: Kodun çalışması için `yfinance` ve `pandas` kütüphanelerinin yüklü olması gerekmektedir. Bu kütüphaneleri pip kullanarak yükleyebilirsiniz:

   ```bash
   pip install yfinance pandas openpyxl
   ```

3. **Excel Dosyası Hazırlığı**:
   - `hisse_takip.xlsx` adında bir Excel dosyası oluşturun. Bu dosya, hisse kodlarını içerecek ve aynı zamanda güncel hisse bilgilerini depolayacak.
   - Excel dosyasının ilk sütununda "Hisse Kodu" başlığı altında hisse kodlarını ekleyin.

4. **Kodun Çalıştırılması**:
   - `hisse_takip.xlsx` dosyasını hazırladıktan sonra Python betiğini çalıştırın. Kod, hisse kodlarını alacak, güncel verileri çekecek ve aynı Excel dosyasına kaydedecektir.
   - Kodu çalıştırmak için terminal veya komut istemcisinde şu komutu çalıştırabilirsiniz:

   ```bash
   python kod_adi.py
   ```

5. **Sonuçlar**:
   - Kodun çalışması tamamlandığında, güncellenmiş hisse bilgileri `hisse_takip.xlsx` dosyasının aynı sayfasına (`hisse_verileri`) kaydedilecektir.

## Notlar

- Kodun çalışması sırasında, herhangi bir hata veya uyarı durumunda ekrana ilgili mesajlar yazdırılacaktır.
- `warnings` modülü kullanılarak alınan uyarılar geçici olarak bastırılmıştır.
- Eğer hisse kodu için güncel veri alınamazsa, ilgili hisse için bir hata mesajı görüntülenecektir ve Excel dosyasına bilgi eklenmeyecektir.
