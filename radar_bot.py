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
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

urllib3.disable_warnings()

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
    
    bulunan_ihaleler = []
    
    # 2. Session ve Genişletilmiş Retry Ayarları
    session = requests.Session()
    retries = Retry(
        total=5,
        backoff_factor=3,
        status_forcelist=[500, 502, 503, 504],
        raise_on_status=False
    )
    adapter = HTTPAdapter(max_retries=retries)
    session.mount('https://', adapter)
    session.mount('http://', adapter)
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'tr-TR,tr;q=0.9,en-US;q=0.8',
        'Connection': 'keep-alive'
    }
    
    base_url = "https://esatis.ogm.gov.tr/ihaleler"
    params = {
        "tarihsecim": "5",
        "baslangic": bugun,
        "bitis": bugun,
        "_odunTuru": "1",
        "_agacTuru": "1",
        "sayfa": 1
    }
    
    print(f"🌐 İstek gönderiliyor: {base_url} (Connect: 30s, Read: 90s)")
    try:
        res = session.get(
            base_url, 
            params=params, 
            headers=headers, 
            verify=False, 
            timeout=(30, 90)
        )
        print(f"📥 Yanıt Kodu: {res.status_code}, Boyut: {len(res.text)}")
        
        soup = BeautifulSoup(res.text, 'html.parser')
        tum_satirlar = soup.find_all('tr')
        print(f"📄 Bulunan satır sayısı: {len(tum_satirlar)}")
        
        tr_tz = timezone(timedelta(hours=3))

        for tr in tum_satirlar:
            satir_metni = tr.get_text(separator=' ', strip=True).upper()
            if not satir_metni:
                continue
                
            eslesen_bolge = None
            for bolge in mevcut_liste:
                if bolge in satir_metni:
                    eslesen_bolge = bolge
                    break
                    
            if eslesen_bolge:
                def millis_to_saat(td_class):
                    td = tr.find('td', class_=td_class)
                    if td and td.get('data-millis'):
                        try:
                            dt = datetime.fromtimestamp(int(td.get('data-millis')) / 1000, tz=tr_tz)
                            return dt.strftime("%H:%M")
                        except Exception:
                            pass
                    return "Belirtilmedi"
                
                baslama = millis_to_saat('baslama')
                bitis = millis_to_saat('bitis')
                
                ilan_linki = res.url
                a_tag = tr.find('a', href=True)
                if a_tag:
                    ilan_linki = urljoin(res.url, a_tag['href'])
                    
                kayit = {
                    "Bölge": eslesen_bolge,
                    "Başlama": baslama,
                    "Bitiş": bitis,
                    "Link": ilan_linki
                }
                
                if not any(x['Bölge'] == kayit['Bölge'] and x['Başlama'] == kayit['Başlama'] for x in bulunan_ihaleler):
                    bulunan_ihaleler.append(kayit)
                    print(f"✅ Yakalandı: {eslesen_bolge} ({baslama} - {bitis})")

    except Exception as e:
        print("❌ OGM taranırken hata oluştu:", e)
        
    # 3. Mail Gönderimi
    if bulunan_ihaleler:
        print(f"🎯 {len(bulunan_ihaleler)} adet ihale yakalandı! Mail gönderiliyor...")
        mail_icerik = f"""
        <html>
        <body style="font-family: Arial, sans-serif;">
            <h2 style="color: #2E7D32;">🌲 OGM Sabah Radarı Alarmı</h2>
            <p>Takip edilen bölgelerde <b>BUGÜN ({bugun})</b> yapılacak ihaleler:</p>
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
                    <td><a href="{ihale['Link']}">İlana Git</a></td>
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
