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

GMAIL_USER = os.environ.get("GMAIL_USER")
GMAIL_PASSWORD = os.environ.get("GMAIL_PASSWORD")
RECEIVER_EMAIL = os.environ.get("RECEIVER_EMAIL")
GCP_CREDS_JSON = os.environ.get("GCP_CREDENTIALS")

def log(msg):
    print(msg, flush=True)

def main():
    log("🚀 OGM Sabah Radarı Başlıyor...")
    bugun = datetime.now().strftime("%d.%m.%Y")
    
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        creds_dict = json.loads(GCP_CREDS_JSON)
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        takip_sheet = client.open("Kereste_İhale_Sistemi").worksheet("Takip_Listesi")
        mevcut_liste = [x.strip().upper() for x in takip_sheet.col_values(1)[1:] if x.strip()]
        log(f"👀 Takip edilen bölgeler: {', '.join(mevcut_liste)}")
    except Exception as e:
        log(f"❌ Google Sheets hatası: {e}")
        return
        
    if not mevcut_liste:
        log("💤 Takip edilen bölge yok, çıkılıyor.")
        return

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'tr-TR,tr;q=0.9,en-US;q=0.8'
    }

    url = "https://esatis.ogm.gov.tr/ihaleler/odun?sayfa=1"
    log(f"🌐 İstek atılıyor: {url}")
    
    bulunan_ihaleler = []
    try:
        res = requests.get(url, headers=headers, verify=False, timeout=30)
        log(f"📥 Yanıt Kodu: {res.status_code}, Boyut: {len(res.text)} karakter")
        
        soup = BeautifulSoup(res.text, 'html.parser')
        tum_satirlar = soup.find_all('tr')
        log(f"📄 Bulunan satır sayısı: {len(tum_satirlar)}")
        
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

                # Eğer millis yoksa düz metinden saat yakalamayı dene
                if baslama == "Belirtilmedi":
                    saat_match = re.search(r'(\d{2}:\d{2})', satir_metni)
                    if saat_match:
                        baslama = saat_match.group(1)

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
                    log(f"✅ Yakalandı: {eslesen_bolge} (Saat: {baslama})")

    except Exception as e:
        log(f"❌ OGM taranırken hata: {e}")

    if bulunan_ihaleler:
        log(f"🎯 {len(bulunan_ihaleler)} adet ihale bulundu, mail hazırlanıyor...")
        mail_icerik = f"""
        <html>
        <body style="font-family: Arial, sans-serif;">
            <h2 style="color: #2E7D32;">🌲 OGM Sabah Radarı</h2>
            <p>Takip ettiğin bölgelerde <b>BUGÜN ({bugun})</b> yapılacak ihaleler:</p>
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
            log("✅ Mail başarıyla gönderildi!")
        except Exception as e:
            log(f"❌ Mail SMTP hatası: {e}")
    else:
        log("💤 Takip edilen bölgeler için eşleşen ihale bulunamadı.")

if __name__ == "__main__":
    main()
