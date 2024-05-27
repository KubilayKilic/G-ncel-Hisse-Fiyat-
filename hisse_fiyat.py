import warnings
import yfinance as yf
import pandas as pd

def hisse_senedi_bilgisi(hisse_kodu):
    hisse = yf.Ticker(hisse_kodu)
    hisse_info = hisse.history(period="1d")
    
    if not hisse_info.empty:
        son_fiyat = hisse_info['Close'].iloc[0]
        gunluk_getiri = ((hisse_info['Close'].iloc[0] - hisse_info['Open'].iloc[0]) / hisse_info['Open'].iloc[0]) * 100

        hisse_bilgileri = {
            'Hisse Kodu': hisse_kodu,
            'Son Fiyat (TL)': son_fiyat,
            'Günlük Getiri (%)': gunluk_getiri,
        }
        
        print(f"{hisse_kodu} verisi başarıyla alındı.")
        return hisse_bilgileri
    else:
        print(f"Hata: {hisse_kodu} için veri alınamadı.")
        return None

# Giriş ve çıkış Excel dosyaları
input_excel = 'hisse_takip.xlsx'  # Giriş Excel dosyanız
output_excel = 'hisse_takip.xlsx'  # Çıkış Excel dosyanız
sheet_name = 'hisse_verileri'  # İşlenecek sayfa adı

# Giriş Excel dosyasını oku
df_hisse_kodlari = pd.read_excel(input_excel, sheet_name=sheet_name)

# Hisse bilgilerini saklamak için boş bir liste oluştur
hisse_bilgileri_listesi = []

# Hisse kodlarını al ve her birini işle
for index, row in df_hisse_kodlari.iterrows():
    hisse_kodu = row['Hisse Kodu']  # A sütununda hisse kodları olduğunu varsayıyoruz
    hisse_bilgileri = hisse_senedi_bilgisi(hisse_kodu)
    if hisse_bilgileri:
        hisse_bilgileri_listesi.append(hisse_bilgileri)

# Hisse bilgilerini DataFrame'e dönüştür
df_hisse_bilgileri = pd.DataFrame(hisse_bilgileri_listesi)

# Uyarıları geçici olarak bastır
with warnings.catch_warnings():
    warnings.simplefilter("ignore")
    # Mevcut Excel dosyasını oku ve üzerine yaz
    with pd.ExcelWriter(output_excel, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df_hisse_bilgileri.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Hisse bilgileri {output_excel} dosyasına kaydedildi.")
