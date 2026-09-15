import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
import urllib3
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from collections import Counter
import hashlib
import uuid

# SSL Uyarılarını Kapat
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def tr_upper(text):
    """Python'un .upper() metodu Türkçe kurallarını bilmez: küçük 'i' harfini düz 'I' yapar,
    oysa Türkçede karşılığı noktalı 'İ' olmalı (Örn: 'Necati'.upper() -> 'NECATI', 'NECATİ' değil).
    Bu yüzden 'NECATİ KELEŞ' gibi aramalar sessizce kaçırılabiliyordu. Bu fonksiyon önce
    küçük i'leri İ'ye çevirip sonra büyük harfe çeviriyor, karşılaştırmalar güvenli oluyor."""
    if not text:
        return ""
    return str(text).replace('i', 'İ').upper()

def isletme_kisalt(text):
    """OİM/OBM/Müdürlüğü eklerini atıp sadece yer adını döndürür (Örn: 'MENGEN OİM' -> 'MENGEN')."""
    if not text:
        return ""
    t = tr_upper(str(text).strip())
    m = re.search(r'([A-ZÇĞİÖŞÜ]+)\s+(?:OİM|OBM)\b', t)
    if m:
        return m.group(1).strip()
    t = t.replace("OİM", "").replace("OBM", "").replace("MÜDÜRLÜĞÜ", "").strip()
    parcalar = t.split()
    return parcalar[-1] if parcalar else t

def sayi_parse(v):
    """'42,707' ya da '1.234,56' gibi Türkçe ondalıklı bir metni float'a çevirir, boş/bozuksa 0.0 döner."""
    s = str(v).strip()
    if not s:
        return 0.0
    if ',' in s and '.' in s:
        s = s.replace('.', '').replace(',', '.')
    else:
        s = s.replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return 0.0

_AGAC_TURU_ANAHTARLARI = [
    ("ÇAM", "Çam"), ("KÖKNAR", "Köknar"), ("GÖKNAR", "Köknar"), ("LADİN", "Ladin"),
    ("KAYIN", "Kayın"), ("MEŞE", "Meşe"), ("KAVAK", "Kavak"), ("GÜRGEN", "Gürgen"),
    ("DİŞBUDAK", "Dişbudak"), ("SEDİR", "Sedir"), ("KESTANE", "Kestane"), ("CEVİZ", "Ceviz"),
]

def agac_turu_cikar(cinsi_metni):
    """'3.Sn.Nb.Kl. Sarıçam Tomruk' gibi bir metinden sınıf/kalite kodlarını görmezden gelip
    sadece ana ağaç türünü çıkarır (Sarıçam/Karaçam/Kızılçam gibi alt türler hepsi 'Çam'
    altında toplanır — parti bazında sınıf ayrımı bu özet için önemli değil)."""
    t = tr_upper(cinsi_metni)
    for anahtar, etiket in _AGAC_TURU_ANAHTARLARI:
        if anahtar in t:
            return etiket
    temiz = str(cinsi_metni).strip()
    return temiz if temiz else "Bilinmeyen"

def tl_formatla(deger):
    s = f"{deger:,.0f}"
    s = s.replace(",", "§").replace(".", ",").replace("§", ".")
    return f"{s} ₺"

def m3_formatla(deger):
    s = f"{deger:,.1f}"
    s = s.replace(",", "§").replace(".", ",").replace("§", ".")
    return f"{s} m³"

def nakliye_drive_senkronize(spreadsheet, df):
    """Nakliyesi tamamlanmış partileri ana Drive dosyasında 'Nakliye_Tümü' ve her
    İşletme + İhale Tarihi kombinasyonu için ayrı bir sekmede günceller — Excel indirmeye
    gerek kalmadan Drive dosyası her zaman güncel dursun diye. Veri son senkronizasyondan
    beri değişmediyse (session_state'teki hash aynıysa) tekrar Sheets'e yazmaz; aksi halde
    her sayfa yenilemesinde (Streamlit her etkileşimde tüm sekmeleri yeniden çalıştırdığı
    için) gereksiz API isteği atılıp kotaya takılma riski olurdu."""
    if df.empty:
        return False

    yazilacak_df = df.copy()
    if "Fatura" in yazilacak_df.columns:
        yazilacak_df["Fatura"] = yazilacak_df["Fatura"].map(lambda v: "EVET" if v else "")

    veri_hash = hashlib.md5(yazilacak_df.to_csv(index=False).encode("utf-8")).hexdigest()
    if st.session_state.get("_nakliye_drive_hash") == veri_hash:
        return False

    def _sekmeye_yaz(baslik, grup_df):
        satirlar = [grup_df.columns.tolist()] + grup_df.astype(str).values.tolist()
        try:
            ws = spreadsheet.worksheet(baslik)
            ws.clear()
        except gspread.exceptions.WorksheetNotFound:
            ws = spreadsheet.add_worksheet(title=baslik, rows=str(len(satirlar) + 5), cols=str(len(grup_df.columns) + 2))
        ws.update(satirlar, value_input_option="USER_ENTERED")

    _sekmeye_yaz("Nakliye_Tümü", yazilacak_df)

    if "İşletme" in yazilacak_df.columns and "İhale Tarihi" in yazilacak_df.columns:
        for (isletme_adi, tarih), grup in yazilacak_df.groupby(["İşletme", "İhale Tarihi"]):
            sekme_adi = f"Nak_{isletme_adi}_{tarih}".strip() or "Bilinmeyen"
            for ch in ['\\', '/', '*', '[', ']', ':', '?']:
                sekme_adi = sekme_adi.replace(ch, '-')
            sekme_adi = sekme_adi[:40]
            _sekmeye_yaz(sekme_adi, grup)

    st.session_state["_nakliye_drive_hash"] = veri_hash
    return True

st.set_page_config(page_title="Kereste İhale & Maliyet Sistemi", layout="wide")

# --- OPENROUTESERVICE (OTOMATİK KM HESAPLAMA) ---
try:
    ORS_API_KEY = st.secrets["ORS_API_KEY"]
except Exception:
    ORS_API_KEY = None

ORS_BASLANGIC_NOKTASI = "Sakarya, Türkiye"

def ors_km_hesapla(hedef_yer):
    """OpenRouteService ile Sakarya'dan hedef yere sürüş mesafesini (km) hesaplar.
    API anahtarı yoksa veya herhangi bir adımda hata olursa sessizce None döner
    (çağıran taraf bu durumda manuel girişe / mesafe tablosuna düşer)."""
    if not ORS_API_KEY or not hedef_yer or hedef_yer == "Bilinmeyen İşletme":
        return None
    try:
        def geokodla(yer_adi):
            r = requests.get(
                "https://api.openrouteservice.org/geocode/search",
                params={"api_key": ORS_API_KEY, "text": yer_adi, "size": 1, "boundary.country": "TR"},
                timeout=10,
            )
            r.raise_for_status()
            data = r.json()
            return data["features"][0]["geometry"]["coordinates"]  # [lon, lat]

        baslangic = geokodla(ORS_BASLANGIC_NOKTASI)
        hedef = geokodla(f"{hedef_yer}, Türkiye")

        r = requests.get(
            "https://api.openrouteservice.org/v2/directions/driving-car",
            params={
                "api_key": ORS_API_KEY,
                "start": f"{baslangic[0]},{baslangic[1]}",
                "end": f"{hedef[0]},{hedef[1]}",
            },
            timeout=15,
        )
        r.raise_for_status()
        data = r.json()
        mesafe_metre = data["features"][0]["properties"]["segments"][0]["distance"]
        return round(mesafe_metre / 1000.0, 1)
    except Exception:
        return None

# --- GOOGLE SHEETS BAĞLANTISI ---
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
try:
    if os.path.exists("credentials.json"):
        creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    else:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)

    client = gspread.authorize(creds)
    spreadsheet = client.open("Kereste_İhale_Sistemi")
    sheet = spreadsheet.sheet1

    try:
        takip_sheet = spreadsheet.worksheet("Takip_Listesi")
    except:
        takip_sheet = spreadsheet.add_worksheet(title="Takip_Listesi", rows="100", cols="2")
        takip_sheet.append_row(["İşletme Adı"])

    try:
        kasa_sheet = spreadsheet.worksheet("Kasa_Takip")
    except:
        kasa_sheet = spreadsheet.add_worksheet(title="Kasa_Takip", rows="100", cols="17")
        kasa_sheet.append_row(["İşletme", "İhale Tarihi", "Parti No", "Cinsi", "Boy", "Miktar", "Birim Fiyat", "Taksitli Tutar", "Nakit Tutar", "Son Ödeme Tarihi", "Durum", "Not", "Nakliye Durumu", "Nakliye Notu", "Alan Firma", "Çekilen Miktar", "Fatura"])

    try:
        mesafe_sheet = spreadsheet.worksheet("Mesafe_Tablosu")
    except:
        mesafe_sheet = spreadsheet.add_worksheet(title="Mesafe_Tablosu", rows="100", cols="2")
        mesafe_sheet.append_row(["Yer", "KM"])

    sheets_baglantisi = True
except Exception as e:
    sheets_baglantisi = False
    hata_mesaji = e

st.title("🌲 Kereste İhale & Maliyet Takip Sistemi")

if not sheets_baglantisi:
    st.error(f"Google Sheets'e bağlanılamadı! Hata: {hata_mesaji}")

tab_islem, tab_gecmis, tab_odeme, tab_nakliye, tab_radar = st.tabs([
    "📥 Yeni İhale Çek",
    "📊 Geçmiş Alımlar",
    "💳 Kasa & Ödeme Takibi",
    "🚚 Nakliye Takibi",
    "🔔 İhale Radarı"
])

# --- İHALE ÇEKME SEKMESİ ---
with tab_islem:
    st.subheader("📥 Yeni İhale Ekle / Çek")
    islem_turu = st.radio("İşlem Türü Seçin:", ["🔗 OGM Sonuç Linkinden Toplu Çek (Bot)", "📄 İhale Öncesi PDF'den Hesapla"])

    if islem_turu == "🔗 OGM Sonuç Linkinden Toplu Çek (Bot)":
        ihale_linki = st.text_input("OGM İhale Sonuç Linki", placeholder="Örn: https://esatis.ogm.gov.tr/ihale/207249/sonuc")

        # --- LİNK YAPIŞTIRILINCA YER + KAYITLI MESAFE ÖNİZLEMESİ ---
        if ihale_linki:
            if st.session_state.get("_onizleme_link") != ihale_linki:
                onizleme_yer = None
                onizleme_hata = None
                try:
                    _r = requests.get(ihale_linki, headers={'User-Agent': 'Mozilla/5.0'}, verify=False, timeout=15)
                    _s = BeautifulSoup(_r.text, 'html.parser')
                    _m = re.search(r'([A-ZÇĞİÖŞÜ\s]+(?:OİM|OBM))', _s.text)
                    if _m:
                        onizleme_yer = isletme_kisalt(_m.group(1))
                except requests.exceptions.Timeout:
                    onizleme_hata = "timeout"
                except Exception:
                    onizleme_hata = "diger"
                st.session_state["_onizleme_link"] = ihale_linki
                st.session_state["_onizleme_yer"] = onizleme_yer
                st.session_state["_onizleme_hata"] = onizleme_hata
            else:
                onizleme_yer = st.session_state.get("_onizleme_yer")
                onizleme_hata = st.session_state.get("_onizleme_hata")

            if onizleme_yer:
                _onizleme_mesafe_verileri = mesafe_sheet.get_all_values()
                _onizleme_mesafe_sozlugu = {}
                if len(_onizleme_mesafe_verileri) > 1:
                    for mv in _onizleme_mesafe_verileri[1:]:
                        if len(mv) > 1 and str(mv[0]).strip():
                            try:
                                _onizleme_mesafe_sozlugu[isletme_kisalt(mv[0])] = float(str(mv[1]).replace(',', '.'))
                            except (ValueError, TypeError):
                                pass
                if onizleme_yer in _onizleme_mesafe_sozlugu:
                    st.caption(f"📍 Tespit edilen yer: **{onizleme_yer}** — mesafe tablosunda kayıtlı: **{_onizleme_mesafe_sozlugu[onizleme_yer]:g} km** (otomatik kullanılacak, bir şey girmene gerek yok).")
                elif ORS_API_KEY:
                    st.caption(f"📍 Tespit edilen yer: **{onizleme_yer}** — mesafe tablosunda kayıtlı değil, OpenRouteService ile otomatik hesaplanacak.")
                else:
                    st.caption(f"📍 Tespit edilen yer: **{onizleme_yer}** — mesafe tablosunda kayıtlı değil. Aşağıya KM gir (bir dahakine sorulmaz).")
            elif onizleme_hata == "timeout":
                st.caption("⏱️ OGM sunucusu 15 saniye içinde cevap vermedi (site yavaş olabilir). Aşağıdaki 'Kazandıklarımızı Çek ve Kaydet' butonu yine de dene, o 20 saniye bekliyor.")
            else:
                st.caption("⚠️ Linkten yer bilgisi tespit edilemedi (bağlantı hatalı olabilir ya da sayfa henüz açılmadı).")
        # -------------------------------------------------------------

        km_mesafe = st.number_input("Bu İhalenin Depoya Mesafesi (KM) — mesafe otomatik hesaplanamazsa bu alan kullanılır", min_value=0.0, step=1.0)
        mesafe_hatirla = st.checkbox("📌 Girdiğim bu KM değerini bu yer için hatırla (bir daha sorulmasın)", value=True)

        with st.expander("📍 Mesafe Tablosunu Görüntüle / Elle Ekle"):
            mesafe_goster = mesafe_sheet.get_all_values()
            if len(mesafe_goster) > 1:
                st.dataframe(pd.DataFrame(mesafe_goster[1:], columns=mesafe_goster[0]), use_container_width=True)
            else:
                st.caption("Henüz kayıtlı mesafe yok.")
            col_yer, col_km, col_ekle = st.columns([2, 1, 1])
            with col_yer:
                yeni_yer = st.text_input("Yer", placeholder="Örn: MENGEN", key="mesafe_yeni_yer")
            with col_km:
                yeni_km = st.number_input("KM", min_value=0.0, step=1.0, key="mesafe_yeni_km")
            with col_ekle:
                st.write("")
                st.write("")
                if st.button("Kaydet", key="mesafe_kaydet_btn"):
                    if yeni_yer:
                        mesafe_sheet.append_row([isletme_kisalt(yeni_yer), yeni_km])
                        st.success("Kaydedildi!")
                        st.rerun()

        if st.button("Kazandıklarımızı Çek ve Kaydet", type="primary", use_container_width=True):
            if not ihale_linki:
                st.warning("Lütfen OGM sonuç linkini yapıştırın.")
            else:
                with st.spinner("Taktik devrede, OGM taranıyor... Lütfen bekleyin..."):
                    try:
                        # --- ANA VERİTABANI MÜKERRER KONTROLÜ (İŞLETME + PARTİ) ---
                        mevcut_gecmis = sheet.get_all_values()
                        mevcut_gecmis_set = set()
                        if len(mevcut_gecmis) > 1:
                            for r in mevcut_gecmis[1:]:
                                if len(r) > 3:
                                    m_isl = isletme_kisalt(r[1])
                                    m_tarih = str(r[0]).strip()
                                    m_prt = str(r[3]).strip()
                                    # Tarih dahil: aynı işletmede farklı tarihli iki ayrı ihalede
                                    # Parti No tekrar edebilir (Örn. iki farklı ihalede de "Parti 1"
                                    # olabilir) — tarih olmadan bu ikinci gerçek alım sessizce
                                    # "zaten kayıtlı" sayılıp atlanırdı.
                                    mevcut_gecmis_set.add(f"{m_isl}_{m_tarih}_{m_prt}")

                        # --- MESAFE TABLOSU (yer -> km, otomatik hatırlama) ---
                        mesafe_verileri = mesafe_sheet.get_all_values()
                        mesafe_sozlugu = {}
                        if len(mesafe_verileri) > 1:
                            for mv in mesafe_verileri[1:]:
                                if len(mv) > 1 and str(mv[0]).strip():
                                    try:
                                        mesafe_sozlugu[isletme_kisalt(mv[0])] = float(str(mv[1]).replace(',', '.'))
                                    except (ValueError, TypeError):
                                        pass
                        # --------------------------------------------------------

                        headers = {'User-Agent': 'Mozilla/5.0'}
                        res = requests.get(ihale_linki, headers=headers, verify=False, timeout=20)
                        soup = BeautifulSoup(res.text, 'html.parser')

                        isletme_text = "Bilinmeyen İşletme"
                        isletme_match = re.search(r'([A-ZÇĞİÖŞÜ\s]+(?:OİM|OBM))', soup.text)
                        if isletme_match:
                            isletme_text = isletme_kisalt(isletme_match.group(1))

                        if isletme_text in mesafe_sozlugu:
                            kullanilacak_km = mesafe_sozlugu[isletme_text]
                            mesafe_kaynagi = "tablo"
                        else:
                            ors_sonuc = ors_km_hesapla(isletme_text)
                            if ors_sonuc is not None:
                                kullanilacak_km = ors_sonuc
                                mesafe_kaynagi = "ors"
                            elif km_mesafe > 0:
                                kullanilacak_km = km_mesafe
                                mesafe_kaynagi = "manuel"
                            else:
                                kullanilacak_km = 0
                                mesafe_kaynagi = "eksik"

                        # --- CLAUDE TAKTİĞİ (DATA-MILLIS OKUMA) KESİN ÇÖZÜMÜ ---
                        genel_ihale_tarihi = "Tarih Bulunamadı"
                        
                        tarih_elementleri = soup.find_all(attrs={"data-millis": True})
                        for el in tarih_elementleri:
                            try:
                                millis = int(el.get('data-millis'))
                                if millis > 1000000000000: 
                                    genel_ihale_tarihi = datetime.fromtimestamp(millis / 1000.0).strftime('%d.%m.%Y')
                                    break 
                            except:
                                continue

                        if genel_ihale_tarihi == "Tarih Bulunamadı":
                            bugun_str = datetime.now().strftime("%d.%m.%Y")
                            tarih_match = re.search(r'(\d{2}\.\d{2}\.\d{4})\s*[Tt]arihli', soup.text)
                            if tarih_match:
                                genel_ihale_tarihi = tarih_match.group(1)
                            else:
                                tum_tarihler = re.findall(r'\b\d{2}\.\d{2}\.\d{4}\b', soup.get_text(separator=' '))
                                gecerli_tarihler = [t for t in tum_tarihler if t != bugun_str]
                                if gecerli_tarihler:
                                    genel_ihale_tarihi = Counter(gecerli_tarihler).most_common(1)[0][0]
                        # --------------------------------------------------------

                        eklenecek_satirlar = []
                        atlanan_adet = 0
                        supheli_partiler = []
                        dogru_tablo = None
                        for tablo in soup.find_all('table'):
                            if 'Müşteri' in tablo.text and 'Parti No' in tablo.text:
                                dogru_tablo = tablo
                                break

                        if dogru_tablo:
                            rows = dogru_tablo.find_all('tr')
                            for row in rows:
                                row_text = tr_upper(row.get_text(separator=' ', strip=True))
                                if "KELEŞ AHŞAP" in row_text or "NECATİ KELEŞ" in row_text:
                                    cols = row.find_all('td')
                                    if len(cols) > 5:
                                        parti_no = cols[0].get_text(strip=True)

                                        kayit_id_bot = f"{isletme_text}_{genel_ihale_tarihi}_{str(parti_no).strip()}"
                                        if kayit_id_bot in mevcut_gecmis_set:
                                            atlanan_adet += 1
                                            continue
                                            
                                        cins = cols[1].get_text(strip=True)
                                        hesaplanan_boy = cols[2].get_text(strip=True)
                                        hesaplanan_kutur = 0.0
                                        miktar_float = 0.0
                                        satir_ihale_tarihi = genel_ihale_tarihi 

                                        hucre_metinleri = [td.get_text(strip=True) for td in cols if td.get_text(strip=True) != '']
                                        fiyat_int = 0

                                        fiyat_hucreleri = [td for td in cols if 'currency-format' in (td.get('class') or [])]
                                        if fiyat_hucreleri:
                                            son_fiyat_hucre = fiyat_hucreleri[-1] 
                                            data_amount = son_fiyat_hucre.get('data-amount')
                                            if data_amount:
                                                try:
                                                    fiyat_int = int(round(float(data_amount)))
                                                except (ValueError, TypeError):
                                                    fiyat_int = 0
                                            if fiyat_int == 0:
                                                metin = son_fiyat_hucre.get_text(strip=True)
                                                kurussuz = metin.split(',')[0]
                                                temiz_sayi = re.sub(r'\D', '', kurussuz)
                                                if temiz_sayi:
                                                    fiyat_int = int(temiz_sayi)

                                        if fiyat_int == 0 or fiyat_int > 999999999:
                                            fiyat_int = 0
                                            for metin in reversed(hucre_metinleri):
                                                if '₺' in metin or 'TL' in metin.upper():
                                                    kurussuz = metin.split(',')[0]
                                                    temiz_sayi = re.sub(r'\D', '', kurussuz)
                                                    if temiz_sayi and len(temiz_sayi) <= 9:
                                                        fiyat_int = int(temiz_sayi)
                                                        break

                                        alan_firma = "Necati Keleş" if "NECATİ KELEŞ" in row_text else "Keleş Ahşap"

                                        detay_a = row.find('a', href=True)
                                        if detay_a:
                                            detay_linki = urljoin(ihale_linki, detay_a['href'])
                                            pdf_isim = None
                                            try:
                                                d_res = requests.get(detay_linki, headers=headers, verify=False, timeout=15)
                                                d_soup = BeautifulSoup(d_res.text, 'html.parser')

                                                pdf_link = None
                                                for a_tag in d_soup.find_all('a', href=True):
                                                    if 'pdf' in a_tag['href'].lower() or 'ebat' in a_tag.text.lower():
                                                        pdf_link = urljoin(detay_linki, a_tag['href'])
                                                        break

                                                if pdf_link:
                                                    p_res = requests.get(pdf_link, headers=headers, verify=False, timeout=20)
                                                    # Benzersiz dosya adı: aynı anda iki kullanıcı/sekme bot
                                                    # çalıştırırsa aynı parti_no ile çakışıp birbirinin PDF'ini
                                                    # bozmasın diye (parti_no tek başına eşsiz olmayabiliyor).
                                                    pdf_isim = f"temp_bot_{uuid.uuid4().hex}.pdf"
                                                    with open(pdf_isim, "wb") as f:
                                                        f.write(p_res.content)

                                                    with pdfplumber.open(pdf_isim) as pdf:
                                                        pdf_text = pdf.pages[0].extract_text()

                                                        m3_match = re.search(r'Miktar\s*\(m3\)\s*[:\-]?\s*([\d\.,]+)', pdf_text, re.IGNORECASE)
                                                        if m3_match:
                                                            temiz_m3 = m3_match.group(1).replace('.', '').replace(',', '.')
                                                            try: miktar_float = float(temiz_m3)
                                                            except: pass

                                                        if miktar_float == 0.0:
                                                            m3_matches = re.findall(r'([\d\.,]+)\s*(?:m3|M3|m³)', pdf_text)
                                                            if m3_matches:
                                                                temiz_m3 = m3_matches[-1].replace('.', '').replace(',', '.')
                                                                try: miktar_float = float(temiz_m3)
                                                                except: pass

                                                        table_pdf = pdf.pages[0].extract_table()
                                                        caplar = []
                                                        boy_adet = {}
                                                        t_adet = 0
                                                        if table_pdf and len(table_pdf) > 1:
                                                            baslik = [str(h).strip().lower() if h else '' for h in table_pdf[0]]

                                                            def _kolon_bul(anahtar_kelimeler):
                                                                for i, h in enumerate(baslik):
                                                                    if any(k in h for k in anahtar_kelimeler):
                                                                        return i
                                                                return None

                                                            cap_idx = _kolon_bul(['çap'])
                                                            boy_idx = _kolon_bul(['boy'])
                                                            adet_idx = _kolon_bul(['adet'])

                                                            if cap_idx is None: cap_idx = 1
                                                            if boy_idx is None: boy_idx = 2
                                                            if adet_idx is None: adet_idx = 3

                                                            for p_row in table_pdf[1:]:
                                                                if len(p_row) <= max(cap_idx, boy_idx, adet_idx):
                                                                    continue
                                                                # PDF'deki "Toplam" özet satırını atla — Çap/Boy hücreleri boş
                                                                # (None) olur ama Adet hücresinde genel toplam adet yazar;
                                                                # bu satır dahil edilirse t_adet şişip %80 kuralı hiç
                                                                # tutturulamıyordu (gerçek oranlar yanlışlıkla sulanıyordu).
                                                                if p_row[boy_idx] is None or str(p_row[boy_idx]).strip() == '' or 'toplam' in str(p_row[0]).strip().lower():
                                                                    continue

                                                                try:
                                                                    raw_adet = str(p_row[adet_idx]).strip()
                                                                    adet_match = re.search(r'\d+', raw_adet.replace('.', ''))
                                                                    if not adet_match: continue
                                                                    adet = int(adet_match.group())
                                                                except:
                                                                    continue
                                                                    
                                                                t_adet += adet
                                                                
                                                                try:
                                                                    raw_b = str(p_row[boy_idx]).strip().replace(',', '.')
                                                                    b_match = re.search(r'[\d\.]+', raw_b)
                                                                    if b_match:
                                                                        f_b = float(b_match.group())
                                                                        s_b = str(f_b)
                                                                        t_boy = s_b[:-2] if s_b.endswith('.0') else s_b
                                                                        boy_adet[t_boy] = boy_adet.get(t_boy, 0) + adet
                                                                except:
                                                                    pass
                                                                    
                                                                try:
                                                                    raw_cap = str(p_row[cap_idx]).strip().replace(',', '.')
                                                                    c_match = re.search(r'[\d\.]+', raw_cap)
                                                                    if c_match:
                                                                        cap = float(c_match.group())
                                                                        if cap > 0:
                                                                            caplar.extend([cap] * adet)
                                                                except:
                                                                    pass

                                                        # %80 KURALI
                                                        if t_adet > 0:
                                                            for b_deg, b_ad in boy_adet.items():
                                                                if (b_ad / t_adet) >= 0.80:
                                                                    hesaplanan_boy = b_deg
                                                                    break

                                                        if caplar:
                                                            hesaplanan_kutur = round(sum(caplar) / len(caplar), 2)
                                                        else:
                                                            if t_adet > 0 and miktar_float > 0.0:
                                                                try:
                                                                    h_boy = float(hesaplanan_boy)
                                                                    if h_boy > 0:
                                                                        import math
                                                                        hesaplanan_kutur = round(math.sqrt((miktar_float * 40000) / (math.pi * h_boy * t_adet)), 2)
                                                                except:
                                                                    pass

                                            except Exception as e:
                                                pass
                                            finally:
                                                # PDF işlenirken ortada bir hata çıksa bile (bozuk PDF,
                                                # ayrıştırma hatası vb.) geçici dosya diskte unutulmasın.
                                                if pdf_isim and os.path.exists(pdf_isim):
                                                    os.remove(pdf_isim)

                                        yeni_satir = [satir_ihale_tarihi, isletme_text, alan_firma, str(parti_no), cins, str(hesaplanan_boy), float(round(miktar_float, 3)), float(round(hesaplanan_kutur, 2)), int(kullanilacak_km), int(fiyat_int)]
                                        eklenecek_satirlar.append(yeni_satir)
                                        mevcut_gecmis_set.add(kayit_id_bot)
                                        if fiyat_int == 0 or miktar_float == 0.0:
                                            supheli_partiler.append(f"Parti {parti_no} ({cins}, {hesaplanan_boy}) — fiyat veya miktar 0 okundu")

                        if len(eklenecek_satirlar) > 0:
                            sheet.append_rows(eklenecek_satirlar, value_input_option='USER_ENTERED')
                            st.success(f"🎉 Helal olsun! {len(eklenecek_satirlar)} adet yeni ihale işlendi! (Zaten kayıtlı olan {atlanan_adet} parti atlandı).")

                            if mesafe_kaynagi == "tablo":
                                st.info(f"📍 '{isletme_text}' için mesafe tablodan otomatik alındı: {kullanilacak_km:g} km.")
                            elif mesafe_kaynagi == "ors":
                                mesafe_sheet.append_row([isletme_text, kullanilacak_km])
                                st.info(f"🌍 '{isletme_text}' için sürüş mesafesi otomatik hesaplandı: {kullanilacak_km:g} km (OpenRouteService). Mesafe tablosuna kaydedildi, bir daha sorgulanmayacak.")
                            elif mesafe_kaynagi == "manuel":
                                if mesafe_hatirla:
                                    mesafe_sheet.append_row([isletme_text, kullanilacak_km])
                                    st.info(f"📍 '{isletme_text}' → {kullanilacak_km:g} km olarak mesafe tablosuna kaydedildi, bir daha sorulmayacak.")
                            elif mesafe_kaynagi == "eksik":
                                st.warning(f"⚠️ '{isletme_text}' için otomatik mesafe hesaplanamadı ve KM girilmedi, mesafe 0 olarak kaydedildi. Yukarıdaki 'Mesafe Tablosunu Görüntüle / Elle Ekle' kısmından bu yer için KM ekleyebilirsin.")

                            if supheli_partiler:
                                st.warning("⚠️ Şu partilerde fiyat veya miktar 0 olarak kaydedildi, sayfa/PDF ayrıştırması başarısız olmuş olabilir — lütfen Google Sheets'ten elle kontrol edin:\n\n" + "\n".join(f"- {p}" for p in supheli_partiler))
                        elif atlanan_adet > 0:
                            st.warning(f"Bu sayfadaki kazandığımız {atlanan_adet} partinin tümü zaten veritabanında var, o yüzden yeniden eklenmedi (Mükerrer koruması devrede).")
                        else:
                            st.error("Sayfa tarandı ancak firmalarımızın kazandığı herhangi bir parti bulunamadı.")
                    except requests.exceptions.Timeout:
                        st.error("⏱️ OGM sunucusu 20 saniye içinde cevap vermedi. Bu genelde OGM'nin sitesi yavaş çalıştığında ya da Streamlit Cloud'un sunucu adresini geçici olarak yavaşlattığında olur — bir kaç dakika sonra tekrar dene. Sürekli oluyorsa bana söyle, başka bir çözüm bulalım.")
                    except requests.exceptions.RequestException as e:
                        st.error(f"🌐 OGM sunucusuna bağlanılamadı: {e}")
                    except Exception as e:
                        st.error(f"Bot çalışırken hata oluştu: {e}")

    else:
        st.info("PDF hesaplama modülü aktif.")


# --- GEÇMİŞ ALIMLAR ---
with tab_gecmis:
    st.subheader("📊 Buluttaki Geçmiş Alımlar")
    
    if sheets_baglantisi:
        with st.spinner("Buluttaki ihale geçmişi çekiliyor..."):
            try:
                raw_data = sheet.get_all_values()
                
                if len(raw_data) > 1:
                    headers = raw_data[0]
                    df = pd.DataFrame(raw_data[1:], columns=headers)
                    
                    for col in df.columns:
                        def turkce_sayiyi_duzelt(val):
                            if not isinstance(val, str): return val
                            v = val.strip()
                            if not v: return None
                            
                            if re.match(r'^-?[\d\.,]+$', v):
                                if ',' in v and '.' in v:
                                    v = v.replace('.', '').replace(',', '.')
                                elif ',' in v:
                                    v = v.replace(',', '.')
                                elif '.' in v:
                                    if len(v.split('.')[-1]) == 3:
                                        v = v.replace('.', '')
                                try:
                                    f = float(v)
                                    return int(f) if f.is_integer() else f
                                except:
                                    return val
                            return val

                        df[col] = df[col].apply(turkce_sayiyi_duzelt)

                    df_filtered = df.copy()
                    
                    st.markdown("##### 🔍 Tabloyu Filtrele")
                    with st.expander("Filtreleri Göster / Gizle", expanded=False):
                        num_columns = 3
                        filter_cols = st.columns(num_columns)
                        
                        for i, col_name in enumerate(df.columns):
                            unique_values = df[col_name].dropna().unique().tolist()
                            try: unique_values.sort()
                            except TypeError: unique_values.sort(key=lambda x: str(x))
                                
                            selected_values = filter_cols[i % num_columns].multiselect(
                                label=f"{col_name}",
                                options=unique_values,
                                default=[]
                            )
                            
                            if selected_values:
                                df_filtered = df_filtered[df_filtered[col_name].isin(selected_values)]

                    st.dataframe(df_filtered, use_container_width=True)
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.caption(f"Filtrelenmiş Sonuç: **{len(df_filtered)}** / Toplam: **{len(df)}** adet ihale gösteriliyor.")
                    with col2:
                        csv = df_filtered.to_csv(index=False).encode('utf-8')
                        st.download_button(
                            label="📥 Süzülmüş Tabloyu İndir (CSV)",
                            data=csv,
                            file_name='filtrelenmis_ihale_gecmisi.csv',
                            mime='text/csv',
                        )
                else:
                    st.info("Henüz kaydedilmiş geçmiş bir ihale bulunmuyor.")
            except Exception as e:
                st.error(f"Veriler çekilirken bir hata oluştu: {e}")


# --- KASA VE ÖDEME TAKİP SEKMESİ ---
with tab_odeme:
    st.subheader("💳 Kasa ve Son Ödeme Tarihi Takibi")

    if sheets_baglantisi:

        # --- TABLO SÜTUN ONARICI ---
        kasa_data = kasa_sheet.get_all_values()
        headers = kasa_data[0] if len(kasa_data) > 0 else []
        ideal_headers = ["İşletme", "İhale Tarihi", "Parti No", "Cinsi", "Boy", "Miktar", "Birim Fiyat", "Taksitli Tutar", "Nakit Tutar", "Son Ödeme Tarihi", "Durum", "Not", "Nakliye Durumu", "Nakliye Notu", "Alan Firma", "Çekilen Miktar", "Fatura"]

        if kasa_sheet.col_count < len(ideal_headers):
            kasa_sheet.add_cols(len(ideal_headers) - kasa_sheet.col_count)

        eksik_var_mi = False
        for i, h in enumerate(ideal_headers):
            if i >= len(headers) or headers[i] != h:
                kasa_sheet.update_cell(1, i+1, h)
                eksik_var_mi = True

        if eksik_var_mi:
            kasa_data = kasa_sheet.get_all_values()
            headers = kasa_data[0]

        if len(kasa_data) > 1:
            df_kasa = pd.DataFrame(kasa_data[1:], columns=headers)
            df_kasa['SheetRow'] = df_kasa.index + 2
            df_kasa['_DurumTemiz'] = df_kasa["Durum"].astype(str).str.strip().apply(tr_upper)

            df_bekleyen = df_kasa[df_kasa['_DurumTemiz'] != "ÖDENDİ"].copy()

            st.markdown("### 📊 Bekleyen Satışların Özeti")
            _odeme_tutar = 0.0
            for _c in ["Taksitli Tutar", "Nakit Tutar"]:
                if _c in df_bekleyen.columns:
                    _odeme_tutar += df_bekleyen[_c].apply(sayi_parse).sum()

            col_ozet1, col_ozet2 = st.columns([1, 2])
            with col_ozet1:
                st.metric("💰 Bekleyen Toplam Tutar", tl_formatla(_odeme_tutar))
                st.metric("📦 Bekleyen Parti Sayısı", len(df_bekleyen))
            with col_ozet2:
                if not df_bekleyen.empty and "Cinsi" in df_bekleyen.columns and "Miktar" in df_bekleyen.columns:
                    _tur_ozet = (
                        df_bekleyen.assign(
                            _AgacTuru=df_bekleyen["Cinsi"].apply(agac_turu_cikar),
                            _MiktarSayi=df_bekleyen["Miktar"].apply(sayi_parse),
                        )
                        .groupby("_AgacTuru")["_MiktarSayi"].sum()
                        .sort_values(ascending=False)
                    )
                    if not _tur_ozet.empty:
                        st.markdown("**Ağaç Türüne Göre Bekleyen Miktar**")
                        _tur_cols = st.columns(min(len(_tur_ozet), 4))
                        for i, (tur, miktar) in enumerate(_tur_ozet.items()):
                            _tur_cols[i % len(_tur_cols)].metric(f"🌲 {tur}", m3_formatla(miktar))
                else:
                    st.caption("Ağaç türü kırılımı için henüz veri yok.")

            st.markdown("---")
            st.markdown("### ⏳ Son Ödeme Tarihi Yaklaşanlar (Tarih Sıralı)")

            if not df_bekleyen.empty:
                df_bekleyen['Tarih_Formatli'] = pd.to_datetime(df_bekleyen['Son Ödeme Tarihi'], format='%d.%m.%Y', errors='coerce')
                df_bekleyen = df_bekleyen.sort_values(by='Tarih_Formatli', ascending=True).drop(columns=['Tarih_Formatli'])

                gorsel_kolonlar_kasa = [c for c in ["İşletme", "Alan Firma", "İhale Tarihi", "Parti No", "Cinsi", "Boy", "Miktar", "Birim Fiyat", "Taksitli Tutar", "Nakit Tutar", "Son Ödeme Tarihi", "Durum"] if c in df_bekleyen.columns]
                st.dataframe(df_bekleyen[gorsel_kolonlar_kasa], use_container_width=True)

                st.markdown("### ✅ Ödemeyi Gerçekleştir ve Listeden Sil")

                col_secim, col_not, col_btn = st.columns([2, 2, 1])

                with col_secim:
                    secenekler = []
                    for idx, row in df_bekleyen.iterrows():
                        firma_etiket = f" [{row['Alan Firma']}]" if row.get('Alan Firma') else ""
                        secenekler.append(f"Satır {row['SheetRow']} | {row['İşletme']}{firma_etiket} - Parti No: {row['Parti No']} - Taksitli: {row['Taksitli Tutar']} ₺ - Nakit: {row['Nakit Tutar']} ₺")

                    secilen_islem = st.selectbox("Ödemesi Yapılan Partiyi Seç", secenekler)

                with col_not:
                    islem_notu = st.text_input("Satış / Ödeme Notu Ekle", placeholder="Örn: Ziraat Kartından Nakit İndirimli Çekildi")

                with col_btn:
                    st.write("")
                    st.write("")
                    if st.button("💳 Ödendi Olarak İşaretle", type="primary", use_container_width=True):
                        gercek_satir_no = int(secilen_islem.split("|")[0].replace("Satır", "").strip())
                        durum_col_num = headers.index("Durum") + 1
                        not_col_num = headers.index("Not") + 1

                        with st.spinner("Ödeme Google Sheets'e işleniyor..."):
                            # Tek seferde toplu yazma: iki ayrı update_cell çağrısı arasında bağlantı
                            # koparsa satır "Durum" güncellenip "Not" güncellenmemiş yarım kalabilirdi.
                            kasa_sheet.update_cells([
                                gspread.Cell(gercek_satir_no, durum_col_num, "ÖDENDİ"),
                                gspread.Cell(gercek_satir_no, not_col_num, islem_notu),
                            ], value_input_option='USER_ENTERED')

                            st.success("✅ Ödeme başarıyla işlendi ve arşive aktarıldı!")
                            st.rerun()
            else:
                st.success("🎉 Mükemmel! Şu an ödeme bekleyen hiçbir parti bulunmuyor. Kasa tertemiz!")

            st.markdown("---")
            st.markdown("### ✅ Ödemesi Gerçekleşen (Arşiv) Partiler")

            df_odenen = df_kasa[df_kasa['_DurumTemiz'] == "ÖDENDİ"].copy()
            if not df_odenen.empty:
                arsiv_kolonlar_kasa = [c for c in ["İşletme", "Alan Firma", "İhale Tarihi", "Parti No", "Cinsi", "Boy", "Miktar", "Birim Fiyat", "Taksitli Tutar", "Nakit Tutar", "Son Ödeme Tarihi", "Durum", "Not"] if c in df_odenen.columns]
                st.dataframe(df_odenen[arsiv_kolonlar_kasa], use_container_width=True)
            else:
                st.caption("Henüz ödemesi yapılıp arşivlenen bir parti bulunmuyor.")
        else:
            st.info("Kasa şu an boş. Lütfen aşağıdaki alana OGM ödeme tablosunu yapıştırarak sistemi başlatın.")

        st.markdown("---")
        st.markdown("### 📋 OGM Verilerini Kasaya Ekle")
        st.info("OGM 'Parti Satış' ekranındaki tabloyu seçip yapıştırın. Yapışık ya da boşluklu olması fark etmez, bot içinden verileri çeker.")

        pasted_data = st.text_area("OGM Parti Satış Tablosunu Buraya Yapıştırın (CTRL+V)", height=150)

        firma_secimi = st.radio("Bu yapıştırdığın partiler hangi firmaya ait? (OGM ekranında firma bilgisi olmadığı için elle seç)", ["Keleş Ahşap", "Necati Keleş"], horizontal=True, key="kasa_firma_secimi")

        if st.button("🔄 OGM Verilerini Kasaya Kaydet", type="primary"):
            if pasted_data:
                with st.spinner("Terminatör bot metni parçalıyor..."):

                    # Tek yapıştırma = tek firma (OGM'nin Parti Satış ekranı tek bir alıcı hesabına aittir).
                    # Ekrandaki metinde firma bilgisi olmadığı için yukarıdaki seçimi doğrudan kullanıyoruz.
                    yapistirma_firmasi = firma_secimi

                    # --- GÜVENLİ BOY ÇEKME VE EŞLEŞTİRME ---
                    gecmis_raw = sheet.get_all_values()
                    parti_boy_sozlugu = {}
                    if len(gecmis_raw) > 1:
                        g_headers = [str(h).strip().lower() for h in gecmis_raw[0]]

                        # Kolay çökmemesi için esnek indeks bulucu
                        p_idx, b_idx, i_idx = 3, 5, 1 # Varsayılan sütun sıralarımız
                        for idx, h in enumerate(g_headers):
                            if "parti" in h: p_idx = idx
                            elif "boy" in h: b_idx = idx
                            elif "işletme" in h or "birim" in h: i_idx = idx

                        for g_row in gecmis_raw[1:]:
                            if len(g_row) > max(p_idx, b_idx, i_idx):
                                p_val = str(g_row[p_idx]).strip()
                                b_val = str(g_row[b_idx]).strip()
                                i_val_kisa = isletme_kisalt(g_row[i_idx])

                                if p_val and b_val:
                                    parti_boy_sozlugu[f"{i_val_kisa}_{p_val}"] = b_val
                    # ----------------------------------------

                    mevcut_kayitlar = set()
                    if len(kasa_data) > 1:
                        for row in kasa_data[1:]:
                            if len(row) > 2:
                                m_isletme = isletme_kisalt(row[0])
                                m_tarih = str(row[1]).strip() if len(row) > 1 else ""
                                m_parti = str(row[2]).strip()
                                # Tarih dahil: aynı işletmenin farklı tarihli iki ihalesinde
                                # Parti No tekrarlanabilir, tarihsiz kontrol gerçek bir yeni
                                # alımı "zaten kayıtlı" sanıp sessizce atlayabilirdi.
                                mevcut_kayitlar.add(f"{m_isletme}_{m_tarih}_{m_parti}")

                    yeni_kayitlar = []
                    eklenen_adet = 0

                    # NOT: re.DOTALL eklendi. OGM sayfasından kopyalanan tabloda hücreler arasına bazen
                    # görünmeyen bir satır sonu (newline) karakteri giriyor; "." varsayılan olarak newline'ı
                    # eşleştirmediği için tek bir gizli satır sonu bile tüm deseni kırıp hiçbir eşleşme
                    # bulunamamasına sebep oluyordu (ve bu durum yanlışlıkla "zaten kasada mevcut" diye
                    # gösteriliyordu). re.DOTALL ile "." artık newline dahil her karakteri eşleştiriyor.
                    pattern = r'([A-ZÇĞİÖŞÜ\s]+(?:OİM|OBM))\s*(\d{2}\.\d{2}\.\d{4}).*?(\d+)\s*No.*?Parti\s*(.*?)\s*([\d\.,]+)\s*m³.*?([\d\.,]+)\s*₺.*?([\d\.,]+)\s*₺.*?(\d{2}\.\d{2}\.\d{4})'
                    matches = re.finditer(pattern, pasted_data, re.IGNORECASE | re.DOTALL)

                    found_count = 0
                    islenemedi_count = 0
                    son_hata = None
                    for match in matches:
                        found_count += 1
                        try:
                            isletme_ham = match.group(1).strip()
                            if "OİM" in tr_upper(isletme_ham) or "OBM" in tr_upper(isletme_ham):
                                parcalar = isletme_ham.split()
                                isletme_ham = " ".join([w for w in parcalar if w not in ["Son", "Satış", "Tarihi:"]][-2:])
                            isletme_ham = re.sub(r'^(son\s*satış\s*tarihi|seçiniz|evet|hayır|müşteri)\s*', '', isletme_ham, flags=re.IGNORECASE).strip()

                            # Sadece yer adını al (Örn: "ADAPAZARI OİM" -> "ADAPAZARI"), OİM/OBM karışıklığını önler
                            isletme = isletme_kisalt(isletme_ham)
                            isletme_kisa = isletme

                            ihale_tarihi = match.group(2)
                            parti_no = match.group(3).strip()
                            cinsi = match.group(4).strip()

                            # Boy'u sözlükten çek — SADECE İşletme+Parti eşleşmesiyle. Sadece Parti
                            # No'ya bakan bir yedek arama riskliydi: iki farklı işletmede aynı parti
                            # numarası oluşabilir ve o zaman başka bir işletmenin boy değeri buraya
                            # yanlışlıkla yazılırdı (görünüşte doğru ama gerçekte yanlış bir sayı).
                            # Eşleşme yoksa "-" (açıkça eksik) bırakmak, sessizce yanlış veriden iyidir.
                            bulunan_boy = parti_boy_sozlugu.get(f"{isletme_kisa}_{parti_no}", "-")

                            # Firma: yukarıda elle seçilen (tek yapıştırma = tek firma)
                            bulunan_firma = yapistirma_firmasi

                            miktar_raw = match.group(5)
                            if ',' in miktar_raw and '.' in miktar_raw:
                                miktar_raw = miktar_raw.replace('.', '').replace(',', '.')
                            else:
                                miktar_raw = miktar_raw.replace(',', '.')
                            try: miktar = float(miktar_raw)
                            except: miktar = 0.0
                            
                            birim_raw = match.group(6)
                            try: birim_fiyat = float(birim_raw.replace('.', '').replace(',', '.'))
                            except: birim_fiyat = 0.0
                            
                            tutar_raw = match.group(7)
                            try: taksitli_tutar = float(tutar_raw.replace('.', '').replace(',', '.'))
                            except: taksitli_tutar = 0.0
                            
                            nakit_tutar = round(taksitli_tutar * 0.25, 2)
                            
                            if birim_fiyat == 0.0 and miktar > 0:
                                birim_fiyat = round(taksitli_tutar / miktar, 2)
                            
                            son_tarih = match.group(8)
                            
                            kayit_id = f"{isletme}_{ihale_tarihi}_{parti_no}"
                            
                            if kayit_id not in mevcut_kayitlar:
                                yeni_kayitlar.append([isletme, ihale_tarihi, parti_no, cinsi, bulunan_boy, miktar, birim_fiyat, taksitli_tutar, nakit_tutar, son_tarih, "BEKLİYOR", "", "", "", bulunan_firma, "", ""])
                                mevcut_kayitlar.add(kayit_id)
                                eklenen_adet += 1
                        except Exception as e:
                            islenemedi_count += 1
                            son_hata = f"{type(e).__name__}: {e}"
                            continue
                            
                    if yeni_kayitlar:
                        kasa_sheet.append_rows(yeni_kayitlar, value_input_option='USER_ENTERED')
                        st.success(f"🎉 Harika! {eklenen_adet} adet parti ({firma_secimi}, İşletme+Parti No kontrolünden geçerek) Kasaya eklendi.")
                        st.rerun()
                    elif found_count == 0:
                        st.error("❌ Yapıştırılan metinde tanınabilir hiçbir parti satırı bulunamadı. OGM 'Parti Satış' ekranındaki tabloyu (başlıklar dahil) tam olarak kopyaladığınızdan emin olun.")
                    elif islenemedi_count == found_count:
                        st.error(f"❌ {found_count} satır regex ile yakalandı ama hiçbiri işlenemedi (kasaya yazılmadı). Son hata: `{son_hata}`. Bu, 'zaten kayıtlı' değil gerçek bir işleme hatası — lütfen bu hata mesajını ilet.")
                    else:
                        st.warning("⚠️ Yeni parti bulunamadı. Kopyaladığınız verideki ihaleler zaten kasada mevcut.")
            else:
                st.warning("Lütfen boş kutuya tabloyu yapıştırın.")


# --- NAKLİYE TAKİP SEKMESİ ---
with tab_nakliye:
    st.subheader("🚚 Nakliye Takibi")
    st.info("Ödemesi tamamlanmış partiler burada listelenir. Tamamı bir seferde çekildiyse toplu işaretle; sadece bir kısmı çekildiyse (örn. 80 m³'lük partiden 40 m³) kısmi çekim bölümünü kullan — kalan miktar otomatik takip edilir.")

    if sheets_baglantisi:
        nakliye_kasa_data = kasa_sheet.get_all_values()
        nakliye_headers = nakliye_kasa_data[0] if len(nakliye_kasa_data) > 0 else []

        def _m3_parse(v):
            """'42,707' gibi Türkçe ondalıklı bir metni float'a çevirir, boş/bozuksa 0.0 döner."""
            s = str(v).strip()
            if not s:
                return 0.0
            if ',' in s and '.' in s:
                s = s.replace('.', '').replace(',', '.')
            else:
                s = s.replace(',', '.')
            try:
                return float(s)
            except ValueError:
                return 0.0

        if len(nakliye_kasa_data) > 1 and "Durum" in nakliye_headers:
            df_nakliye = pd.DataFrame(nakliye_kasa_data[1:], columns=nakliye_headers)
            df_nakliye['SheetRow'] = df_nakliye.index + 2
            df_nakliye['_DurumTemiz'] = df_nakliye["Durum"].astype(str).str.strip().apply(tr_upper)

            if "Nakliye Durumu" in df_nakliye.columns:
                df_nakliye['_NakliyeTemiz'] = df_nakliye["Nakliye Durumu"].astype(str).str.strip().apply(tr_upper)
            else:
                df_nakliye['_NakliyeTemiz'] = ""

            df_nakliye['_ToplamM3'] = df_nakliye["Miktar"].apply(_m3_parse) if "Miktar" in df_nakliye.columns else 0.0
            if "Çekilen Miktar" in df_nakliye.columns:
                df_nakliye['_CekilenM3'] = df_nakliye["Çekilen Miktar"].apply(_m3_parse)
            else:
                df_nakliye['_CekilenM3'] = 0.0
            df_nakliye['Kalan Miktar'] = (df_nakliye['_ToplamM3'] - df_nakliye['_CekilenM3']).clip(lower=0).round(3)

            df_odemesi_biten = df_nakliye[df_nakliye['_DurumTemiz'] == "ÖDENDİ"].copy()
            df_bekleyen_nakliye = df_odemesi_biten[df_odemesi_biten['_NakliyeTemiz'] != "NAKLİYE YAPILDI"].copy()

            st.markdown("### 📊 Nakliyesi Bekleyen Partilerin Özeti")
            _nakliye_tutar = 0.0
            for _c in ["Taksitli Tutar", "Nakit Tutar"]:
                if _c in df_bekleyen_nakliye.columns:
                    _nakliye_tutar += df_bekleyen_nakliye[_c].apply(sayi_parse).sum()

            col_nak1, col_nak2 = st.columns([1, 2])
            with col_nak1:
                st.metric("💰 Depodaki Malın Değeri", tl_formatla(_nakliye_tutar))
                st.metric("📦 Bekleyen Parti Sayısı", len(df_bekleyen_nakliye))
            with col_nak2:
                if not df_bekleyen_nakliye.empty and "Cinsi" in df_bekleyen_nakliye.columns:
                    _tur_ozet_nak = (
                        df_bekleyen_nakliye.assign(_AgacTuru=df_bekleyen_nakliye["Cinsi"].apply(agac_turu_cikar))
                        .groupby("_AgacTuru")["Kalan Miktar"].sum()
                        .sort_values(ascending=False)
                    )
                    if not _tur_ozet_nak.empty:
                        st.markdown("**Ağaç Türüne Göre Depoda Kalan Miktar**")
                        _tur_cols_nak = st.columns(min(len(_tur_ozet_nak), 4))
                        for i, (tur, miktar) in enumerate(_tur_ozet_nak.items()):
                            _tur_cols_nak[i % len(_tur_cols_nak)].metric(f"🌲 {tur}", m3_formatla(miktar))
                else:
                    st.caption("Ağaç türü kırılımı için henüz veri yok.")

            st.markdown("---")
            st.markdown("### 📦 Depodan Çekilmeyi Bekleyen Partiler (Ödemesi Yapılmış)")

            if not df_bekleyen_nakliye.empty:
                df_bekleyen_nakliye['Tarih_Formatli'] = pd.to_datetime(df_bekleyen_nakliye['Son Ödeme Tarihi'], format='%d.%m.%Y', errors='coerce')
                df_bekleyen_nakliye = df_bekleyen_nakliye.sort_values(by='Tarih_Formatli', ascending=True).drop(columns=['Tarih_Formatli'])

                gorsel_kolonlar = [c for c in ["İşletme", "Alan Firma", "İhale Tarihi", "Parti No", "Cinsi", "Boy", "Miktar", "Kalan Miktar", "Nakliye Durumu", "Birim Fiyat", "Taksitli Tutar", "Nakit Tutar", "Son Ödeme Tarihi"] if c in df_bekleyen_nakliye.columns]
                st.dataframe(df_bekleyen_nakliye[gorsel_kolonlar], use_container_width=True)

                secenekler_nakliye = []
                for idx, row in df_bekleyen_nakliye.iterrows():
                    firma_etiket = f" [{row['Alan Firma']}]" if row.get('Alan Firma') else ""
                    secenekler_nakliye.append(f"Satır {row['SheetRow']} | {row['İşletme']}{firma_etiket} - Parti No: {row['Parti No']} - {row.get('Cinsi', '')} - Kalan: {row['Kalan Miktar']:g} m³ / Toplam: {row['_ToplamM3']:g} m³")

                st.markdown("### 🚚 Bu Seferki Çekimi Kaydet")
                st.caption("Bu seferki taşımada hangi partilerden çekildiğini seç (birden fazla olabilir). Bir parti tamamen, diğeri yarım kalmış olabilir — seçtiğin her parti için ayrı ayrı miktar gireceksin, varsayılan olarak kalanın tamamı gelir.")

                secilenler_nakliye = st.multiselect("Bu Seferki Taşımada Hangi Partilerden Çekildi?", secenekler_nakliye, key="nakliye_multiselect")

                if secilenler_nakliye:
                    girilen_miktarlar = {}
                    for secim in secilenler_nakliye:
                        satir_no = int(secim.split("|")[0].replace("Satır", "").strip())
                        satir_bilgi = df_bekleyen_nakliye[df_bekleyen_nakliye['SheetRow'] == satir_no].iloc[0]
                        kalan = float(satir_bilgi['Kalan Miktar'])
                        girilen_miktarlar[satir_no] = st.number_input(
                            f"{satir_bilgi['İşletme']} - Parti {satir_bilgi['Parti No']} — bu seferki çekilen (m³, kalan: {kalan:g})",
                            min_value=0.0, max_value=max(kalan, 0.01), value=kalan, step=1.0,
                            key=f"nakliye_miktar_{satir_no}",
                        )

                    nakliye_notu = st.text_input("Nakliye Notu (Kim getirdi / hangi araç)", placeholder="Örn: Mehmet'in kamyonuyla çekildi", key="nakliye_notu_input")

                    if st.button("🚚 Bu Seferki Çekimi Kaydet", type="primary", use_container_width=True):
                        if "Nakliye Durumu" not in nakliye_headers or "Nakliye Notu" not in nakliye_headers or "Çekilen Miktar" not in nakliye_headers:
                            st.error("Kasa_Takip sayfasında gerekli sütunlar bulunamadı. Kasa & Ödeme sekmesini bir kez açıp tekrar dene.")
                        elif all(m <= 0 for m in girilen_miktarlar.values()):
                            st.warning("En az bir parti için 0'dan büyük miktar gir.")
                        else:
                            with st.spinner("Nakliye bilgisi Google Sheets'e işleniyor..."):
                                nakliye_durum_col = nakliye_headers.index("Nakliye Durumu") + 1
                                nakliye_not_col = nakliye_headers.index("Nakliye Notu") + 1
                                cekilen_col = nakliye_headers.index("Çekilen Miktar") + 1
                                bugun_str = datetime.now().strftime("%d.%m.%Y")
                                ozet = []
                                yazilacak_hucreler = []

                                for secim in secilenler_nakliye:
                                    satir_no = int(secim.split("|")[0].replace("Satır", "").strip())
                                    girilen = girilen_miktarlar[satir_no]
                                    if girilen <= 0:
                                        continue
                                    satir_bilgi = df_bekleyen_nakliye[df_bekleyen_nakliye['SheetRow'] == satir_no].iloc[0]
                                    toplam = float(satir_bilgi['_ToplamM3'])
                                    eski_cekilen = float(satir_bilgi['_CekilenM3'])
                                    yeni_cekilen = min(eski_cekilen + girilen, toplam)
                                    yeni_kalan = round(toplam - yeni_cekilen, 3)
                                    yeni_durum = "NAKLİYE YAPILDI" if yeni_kalan <= 0.01 else "KISMİ ÇEKİLDİ"

                                    eski_not = str(satir_bilgi.get("Nakliye Notu", "") or "").strip()
                                    yeni_not_parcasi = f"{bugun_str}: {girilen:g} m³ çekildi" + (f" ({nakliye_notu})" if nakliye_notu else "")
                                    guncel_not = f"{eski_not} | {yeni_not_parcasi}" if eski_not else yeni_not_parcasi

                                    yazilacak_hucreler.append(gspread.Cell(satir_no, nakliye_durum_col, yeni_durum))
                                    yazilacak_hucreler.append(gspread.Cell(satir_no, nakliye_not_col, guncel_not))
                                    yazilacak_hucreler.append(gspread.Cell(satir_no, cekilen_col, yeni_cekilen))

                                    durum_metni = "tamamen çekildi" if yeni_durum == "NAKLİYE YAPILDI" else f"{yeni_kalan:g} m³ kaldı"
                                    ozet.append(f"{satir_bilgi['İşletme']} Parti {satir_bilgi['Parti No']}: {durum_metni}")

                                # Tüm partiler için tüm hücreler TEK API çağrısıyla yazılıyor — birden
                                # fazla parti seçiliyken yarıda bağlantı kopması bazı partileri
                                # güncellenmiş bazılarını güncellenmemiş bırakmasın diye.
                                if yazilacak_hucreler:
                                    kasa_sheet.update_cells(yazilacak_hucreler, value_input_option='USER_ENTERED')

                                st.success("✅ Kaydedildi:\n\n" + "\n".join(f"- {o}" for o in ozet))
                                st.rerun()
            else:
                st.success("🎉 Depoda bekleyen (ödemesi yapılmış ama henüz çekilmemiş) parti yok!")

            st.markdown("---")
            st.markdown("### 🚛 Nakliyesi Tamamlanmış (Arşiv) Partiler")
            df_nakliye_tamam = df_odemesi_biten[df_odemesi_biten['_NakliyeTemiz'] == "NAKLİYE YAPILDI"].copy()
            if not df_nakliye_tamam.empty:
                df_nakliye_tamam = df_nakliye_tamam.reset_index(drop=True)
                if "Fatura" in df_nakliye_tamam.columns:
                    df_nakliye_tamam['Fatura'] = df_nakliye_tamam["Fatura"].astype(str).str.strip().apply(tr_upper) == "EVET"
                else:
                    df_nakliye_tamam['Fatura'] = False

                arsiv_kolonlar = [c for c in ["İşletme", "Alan Firma", "İhale Tarihi", "Parti No", "Cinsi", "Boy", "Miktar", "Birim Fiyat", "Taksitli Tutar", "Nakit Tutar", "Son Ödeme Tarihi", "Nakliye Notu", "Fatura"] if c in df_nakliye_tamam.columns]
                gorsel_arsiv = df_nakliye_tamam[arsiv_kolonlar].copy()

                duzenlenen_arsiv = st.data_editor(
                    gorsel_arsiv,
                    column_config={
                        "Fatura": st.column_config.CheckboxColumn("📄 Fatura Geldi mi?", help="Muhasebeci fatura geldiğinde burayı tikleyip geçecek."),
                    },
                    disabled=[c for c in arsiv_kolonlar if c != "Fatura"],
                    hide_index=True,
                    use_container_width=True,
                    key="fatura_editor",
                )

                if "Fatura" in nakliye_headers:
                    fatura_col = nakliye_headers.index("Fatura") + 1
                    degisiklik_oldu = False
                    for i in range(len(gorsel_arsiv)):
                        eski_deger = gorsel_arsiv.iloc[i]['Fatura']
                        yeni_deger = duzenlenen_arsiv.iloc[i]['Fatura']
                        if bool(eski_deger) != bool(yeni_deger):
                            satir_no = int(df_nakliye_tamam.iloc[i]['SheetRow'])
                            kasa_sheet.update_cell(satir_no, fatura_col, "EVET" if yeni_deger else "")
                            degisiklik_oldu = True
                    if degisiklik_oldu:
                        st.rerun()
                else:
                    st.caption("⚠️ 'Fatura' sütunu henüz sayfada yok — Kasa & Ödeme sekmesini bir kez açıp tekrar dene, otomatik eklenecek.")

                st.markdown("---")
                st.markdown("#### ☁️ Ana Drive Dosyasına Otomatik Aktarım")
                st.caption("Bu tablo, 'Kereste_İhale_Sistemi' dosyasında 'Nakliye_Tümü' sekmesine ve her ihale (İşletme + İhale Tarihi) için kendi ayrı sekmesine otomatik olarak işleniyor — indirmene gerek yok, Drive'da hep güncel duruyor.")
                with st.spinner("Drive'daki sekmeler kontrol ediliyor..."):
                    _senkron_oldu = nakliye_drive_senkronize(spreadsheet, gorsel_arsiv)
                if _senkron_oldu:
                    st.success("✅ Drive'daki 'Nakliye_Tümü' ve ihale bazlı sekmeler güncellendi.")
                else:
                    st.caption("☁️ Drive sekmeleri zaten güncel.")

                # --- İŞLETME + İHALE TARİHİ KOMBİNASYONUNA GÖRE AYRI SEKMELİ EXCEL İNDİRME (opsiyonel, ekstra) ---
                # Aynı yer (Örn. ALADAĞ) farklı tarihlerde birden fazla ihale olabilir; bunları
                # tek sekmede birleştirmiyoruz, her ihale (yer + tarih) kendi sekmesinde ayrı duruyor.
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    duzenlenen_arsiv.to_excel(writer, sheet_name='Tümü', index=False)
                    if "İşletme" in duzenlenen_arsiv.columns and "İhale Tarihi" in duzenlenen_arsiv.columns:
                        for (isletme_adi, tarih), grup in duzenlenen_arsiv.groupby(['İşletme', 'İhale Tarihi']):
                            sheet_adi = f"{isletme_adi}_{tarih}".strip() or 'Bilinmeyen'
                            for ch in ['\\', '/', '*', '[', ']', ':', '?']:
                                sheet_adi = sheet_adi.replace(ch, '-')
                            sheet_adi = sheet_adi[:31]
                            grup.to_excel(writer, sheet_name=sheet_adi, index=False)
                    elif "İşletme" in duzenlenen_arsiv.columns:
                        for isletme_adi, grup in duzenlenen_arsiv.groupby('İşletme'):
                            sheet_adi = str(isletme_adi).strip() or 'Bilinmeyen'
                            for ch in ['\\', '/', '*', '[', ']', ':', '?']:
                                sheet_adi = sheet_adi.replace(ch, '-')
                            sheet_adi = sheet_adi[:31]
                            grup.to_excel(writer, sheet_name=sheet_adi, index=False)

                st.download_button(
                    "📥 Nakliye Arşivini Excel Olarak İndir (Her İhale Ayrı Sekmede)",
                    data=excel_buffer.getvalue(),
                    file_name=f"nakliye_arsiv_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.caption("Henüz nakliyesi tamamlanmış bir parti bulunmuyor.")
        else:
            st.info("Kasa henüz boş ya da veri okunamadı. Önce '💳 Kasa & Ödeme Takibi' sekmesinden veri ekle.")


# --- RADAR VE BİLDİRİM BÖLÜMÜ ---
with tab_radar:
    st.subheader("🎯 Bölge Radarı & Mail Testi (PDF + RÖNTGEN Modu)")

    if sheets_baglantisi:
        mevcut_liste = takip_sheet.col_values(1)[1:]

        st.markdown("### 1️⃣ Takip Edilecek İşletmeleri Gir")
        col_ekle, col_bos = st.columns([2, 1])
        with col_ekle:
            yeni_isletme = st.text_input("Bölge Ekle", placeholder="Örn: ZONGULDAK VEYA MENGEN")
            if st.button("Listeye Ekle"):
                if yeni_isletme and yeni_isletme.upper() not in [x.upper() for x in mevcut_liste]:
                    takip_sheet.append_row([yeni_isletme.upper()])
                    st.success("Eklendi!")
                    st.rerun()

        if mevcut_liste:
            st.write("Şu an takip edilenler: ", ", ".join(mevcut_liste))

        st.markdown("---")

        with st.expander("⚙️ Gelişmiş Filtre Ayarları (gerekirse düzelt)"):
            parametresiz_dene = st.checkbox("Hiç filtre parametresi gönderme (sade /ihaleler dene)", value=False)
            odun_turu = st.text_input("_odunTuru değeri", value="1")
            agac_turu = st.text_input("_agacTuru değeri", value="1")
            tarihsecim_degeri = st.text_input("tarihsecim değeri", value="5")

        st.markdown("### 2️⃣ Radarı Şimdi Test Et (Röntgen Raporlu)")

        with st.form("mail_test_form"):
            gonderici_mail = st.text_input("Gönderici Gmail Adresin", placeholder="ornek@gmail.com")
            uygulama_sifresi = st.text_input("16 Haneli Uygulama Şifren", type="password")
            alici_mail = st.text_input("Kime Gidecek?", placeholder="sirket_veya_kisisel_mailin@gmail.com")

            test_baslat = st.form_submit_button("🚀 Radarı Çalıştır (Listeyi Tara ve Raporla)")

            if test_baslat:
                if not gonderici_mail or not uygulama_sifresi or not alici_mail:
                    st.error("Lütfen mail bilgilerini eksiksiz gir!")
                elif not mevcut_liste:
                    st.warning("Önce yukarıdan takip edilecek işletme eklemelisin!")
                else:
                    st.markdown("### 🛠️ BOT RÖNTGEN RAPORU")
                    bugun = datetime.now().strftime("%d.%m.%Y")
                    st.info(f"📅 BOT'UN BİLDİĞİ 'BUGÜN' TARİHİ: {bugun}")
                    bulunan_ihaleler = []

                    headers = {
                        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8'
                    }

                    base_url = "https://esatis.ogm.gov.tr/ihaleler"

                    for sayfa in range(1, 4):  # ilk 3 sayfa
                        if parametresiz_dene:
                            params = {"sayfa": sayfa}
                        else:
                            params = {
                                "tarihsecim": tarihsecim_degeri,
                                "baslangic": bugun,
                                "bitis": bugun,
                                "_odunTuru": odun_turu,
                                "_agacTuru": agac_turu,
                                "sayfa": sayfa,
                            }

                        st.write(f"🌐 **Sayfa {sayfa} taranıyor...**")
                        try:
                            res = requests.get(base_url, params=params, headers=headers, verify=False, timeout=15)
                            st.caption(f"İstek atılan gerçek URL: {res.url}")
                            st.caption(f"Yanıt uzunluğu: {len(res.text)} karakter")

                            soup = BeautifulSoup(res.text, 'html.parser')

                            if len(res.text) < 2000:
                                st.error("⚠️ Yanıt çok kısa geldi. Site bu isteği engelliyor ya da parametreler yanlış olabilir.")

                            tum_satirlar = soup.find_all('tr')
                            st.caption(f"Sayfada bulunan toplam <tr> satır sayısı: {len(tum_satirlar)}")

                            if len(tum_satirlar) == 0:
                                st.error("❌ Hiç tablo satırı bulunamadı. Tablo muhtemelen JavaScript ile sonradan yükleniyor.")

                            for tr in tum_satirlar:
                                satir_metni = tr_upper(tr.get_text(separator=' ', strip=True))
                                if not satir_metni:
                                    continue

                                eslesen_bolge = None
                                for bolge in mevcut_liste:
                                    if tr_upper(bolge) in satir_metni:
                                        eslesen_bolge = bolge
                                        break

                                if eslesen_bolge:
                                    st.warning(f"👀 **{eslesen_bolge}** BULUNDU! Satır: {satir_metni[:120]}")
                                    with st.expander(f"🔬 '{eslesen_bolge}' satırının ham HTML'i"):
                                        st.code(str(tr), language="html")

                                    from datetime import timezone, timedelta as _timedelta
                                    tr_tz = timezone(_timedelta(hours=3))

                                    def _millis_to_saat(td_class):
                                        td = tr.find('td', class_=td_class)
                                        if not td:
                                            return "Belirtilmedi"
                                        millis = td.get('data-millis')
                                        if not millis:
                                            return "Belirtilmedi"
                                        try:
                                            dt = datetime.fromtimestamp(int(millis) / 1000, tz=tr_tz)
                                            return dt.strftime("%H:%M")
                                        except Exception:
                                            return "Belirtilmedi"

                                    baslama_saati = _millis_to_saat('baslama')
                                    bitis_saati = _millis_to_saat('bitis')

                                    ilan_linki = None
                                    a_tag = tr.find('a', href=True)
                                    if a_tag:
                                        ilan_linki = urljoin(res.url, a_tag['href'])

                                    kayit = {
                                        "Bölge": eslesen_bolge.upper(),
                                        "Başlama": baslama_saati,
                                        "Bitiş": bitis_saati,
                                        "Link": ilan_linki or res.url,
                                    }
                                    if not any(x['Bölge'] == kayit['Bölge'] and x['Başlama'] == kayit['Başlama'] for x in bulunan_ihaleler):
                                        bulunan_ihaleler.append(kayit)

                        except Exception as e:
                            st.error(f"🌐 Sayfa {sayfa} bağlantı hatası: {e}")

                    st.markdown("---")

                    if bulunan_ihaleler:
                        st.success(f"🎯 Rapor Bitti! {len(bulunan_ihaleler)} adet ihale torbaya atıldı! Mail gönderiliyor...")
                        mail_icerik = f"""
                        <html>
                        <body style="font-family: Arial, sans-serif;">
                            <h2 style="color: #2E7D32;">🌲 OGM Radar Alarmı</h2>
                            <p>Günaydın, takip ettiğin bölgelerde <b>BUGÜN ({bugun})</b> yapılacak ihaleler aşağıdadır:</p>
                            <table border="1" cellpadding="10" cellspacing="0" style="border-collapse: collapse; width: 100%; text-align: left;">
                                <tr style="background-color: #f2f2f2;">
                                    <th>İşletme</th>
                                    <th>Başlama</th>
                                    <th>Bitiş</th>
                                    <th>İlan / Detay</th>
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
                        msg['From'] = gonderici_mail
                        msg['To'] = alici_mail
                        msg['Subject'] = f"🚨 {bugun} OGM İhale Alarmı ({len(bulunan_ihaleler)} Adet)"
                        msg.attach(MIMEText(mail_icerik, 'html'))

                        try:
                            server = smtplib.SMTP('smtp.gmail.com', 587)
                            server.starttls()
                            server.login(gonderici_mail, uygulama_sifresi)
                            server.send_message(msg)
                            server.quit()
                            st.balloons()
                            st.success("✅ Mail başarıyla gönderildi!")
                        except Exception as e:
                            st.error(f"❌ Mail gönderilirken SMTP hatası oluştu: {e}")
                    else:
                        st.info("Röntgen tamamlandı, listeye eklenecek ihale bulunamadı.")
