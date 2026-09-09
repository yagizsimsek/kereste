import os
import json
import re
import requests
from bs4 import BeautifulSoup
from datetime import datetime, timezone, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import urllib3
from urllib.parse import urljoin

urllib3.disable_warnings()

# GitHub'a şifreleri açık açık yazmamak için ortam değişkenlerinden alıyoruz
GMAIL_USER = os.environ.get("GMAIL_USER")
GMAIL_PASSWORD = os.environ.get("GMAIL_PASSWORD")
RECEIVER_EMAIL = os.environ.get("RECEIVER_EMAIL")
GCP_CREDS_JSON = os.environ.get("GCP_CREDENTIALS")

def main():
    print("🚀 OGM Sabah Radarı Başlıyor...")
    bugun = datetime.now().strftime("%d.%m.%Y")
    
    # 1. Google Sheets Bağlantısı
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds_dict = json.loads(GCP_CREDS_JSON)
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    
    try:
        takip_sheet = client.open("Kereste_İhale_Sistemi").worksheet("Takip_Listesi")
        mevcut_liste = [x.strip().upper() for x in takip_sheet.col_values(1)[1:] if x.strip()]
    except Exception as e:
        print("❌ Google Sheets okunamadı:", e)
        return
        
    if not mevcut_liste:
        print("💤 Takip edilen bölge yok, çıkılıyor.")
        return
        
    print(f"👀 Takip edilen bölgeler: {', '.join(mevcut_liste)}")
    
    # 2. OGM Tarama (Sadece 1. Sayfa)
    bulunan_ihaleler = []
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml'
    }
    
    # Senin talebin üzerine SADECE 1. sayfayı tarıyoruz
    url = "https://esatis.ogm.gov.tr/ihaleler/odun?sayfa=1"
    print(f"🌐 Taranıyor: {url}")
    
    try:
        res = requests.get(url, headers=headers, verify=False, timeout=15)
        soup = BeautifulSoup(res.text, 'html.parser')
        
        tum_satirlar = soup.find_all('tr')
        for tr in tum_satirlar:
            satir_metni = tr.get_text(separator=' ', strip=True).upper()
            if not satir_metni: continue
            
            eslesen_bolge = None
            for bolge in mevcut_liste:
                if bolge in satir_metni:
                    eslesen_bolge = bolge
                    break
                    
            if eslesen_bolge:
                tr_tz = timezone(timedelta(hours=3)) # Türkiye Saati
                baslama_saati = "Belirtilmedi"
                bitis_saati = "Belirtilmedi"
                
                td_baslama = tr.find('td', class_='baslama')
                if td_baslama and td_baslama.get('data-millis'):
                    dt = datetime.fromtimestamp(int(td_baslama.get('data-millis')) / 1000, tz=tr_tz)
                    baslama_saati = dt.strftime("%H:%M")
                    
                td_bitis = tr.find('td', class_='bitis')
                if td_bitis and td_bitis.get('data-millis'):
                    dt = datetime.fromtimestamp(int(td_bitis.get('data-millis')) / 1000, tz=tr_tz)
                    bitis_saati = dt.strftime("%H:%M")
                    
                ilan_linki = res.url
                a_tag = tr.find('a', href=True)
                if a_tag:
                    ilan_linki = urljoin(res.url, a_tag['href'])
                    
                kayit = {
                    "Bölge": eslesen_bolge,
                    "Başlama": baslama_saati,
                    "Bitiş": bitis_saati,
                    "Link": ilan_linki
                }
                
                if not any(x['Bölge'] == kayit['Bölge'] and x['Başlama'] == kayit['Başlama'] for x in bulunan_ihaleler):
                    bulunan_ihaleler.append(kayit)
    except Exception as e:
        print("❌ OGM taranırken hata:", e)
        
    # 3. Mail Gönderimi
    if bulunan_ihaleler:
        print(f"🎯 {len(bulunan_ihaleler)} ihale bulundu! Mail atılıyor...")
        mail_icerik = f"""
        <html>
        <body style="font-family: Arial, sans-serif;">
            <h2 style="color: #2E7D32;">🌲 OGM Sabah Radarı</h2>
            <p>Başgan, takip ettiğin bölgelerde <b>BUGÜN ({bugun})</b> yapılacak ihaleler aşağıdadır:</p>
            <table border="1" cellpadding="10" cellspacing="0" style="border-collapse: collapse; width: 100%; text-align: left;">
                <tr style="background-color: #f2f2f2;">
                    <th>İşletme</th>
                    <th>Başlama</th>
                    <th>Bitiş</th>
                    <th>İlan Linki</th>
                </tr>
        """
        for ihale in bulunan_ihaleler:
            mail_icerik += f"""
                <tr>
                    <td><b>{ihale['Bölge']}</b></td>
                    <td style="color: #D32F2F;"><b>{ihale['Başlama']}</b></td>
                    <td>{ihale['Bitiş']}</td>
                    <td><a href="{ihale['Link']}">Detay / İlan</a></td>
                </tr>
            """
        mail_icerik += "</table></body></html>"
        
        msg = MIMEMultipart()
        msg['From'] = GMAIL_USER
        msg['To'] = RECEIVER_EMAIL
        msg['Subject'] = f"🚨 {bugun} OGM İhale Alarmı ({len(bulunan_ihaleler)} Adet)"
        msg.attach(MIMEText(mail_icerik, 'html'))
        
        try:
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(GMAIL_USER, GMAIL_PASSWORD)
            server.send_message(msg)
            server.quit()
            print("✅ Mail başarıyla gönderildi!")
        except Exception as e:
            print("❌ Mail SMTP hatası:", e)
    else:
        print("💤 Bugün için takip edilen bölgelerde ihale yok.")

if __name__ == "__main__":
    main()