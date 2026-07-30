import streamlit as st
import pdfplumber
import pandas as pd
import re
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
import os

st.set_page_config(page_title="Kereste İhale & Maliyet Sistemi", layout="wide")

# --- GOOGLE SHEETS BAĞLANTISI (AKILLI KONTROL) ---
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
try:
    if os.path.exists("credentials.json"):
        creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    else:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        
    client = gspread.authorize(creds)
    sheet = client.open("Kereste_İhale_Sistemi").sheet1
    sheets_baglantisi = True
except Exception as e:
    sheets_baglantisi = False
    hata_mesaji = e

st.title("🌲 Kereste İhale & Maliyet Takip Sistemi")
st.write("OGM E-Satış ebat listesini yükleyin, veriler anında Google Sheets tablonuza işlensin.")

if not sheets_baglantisi:
    st.error(f"Google Sheets'e bağlanılamadı! Lütfen bağlantı ayarlarını kontrol et. Hata: {hata_mesaji}")

col_sol, col_sag = st.columns([1, 2])

with col_sol:
    st.subheader("📥 Yeni İhale Ekle")
    uploaded_file = st.file_uploader("OGM İstif PDF Dosyası", type=["pdf"])
    
    m3_teklif = st.number_input("m³ Teklifimiz (TL)", min_value=0.0, step=50.0)
    km_mesafe = st.number_input("Depoya Mesafe (KM)", min_value=0.0, step=1.0)
    
    if st.button("Hesapla ve Sheets'e Kaydet", type="primary", use_container_width=True):
        if not sheets_baglantisi:
            st.error("Google bağlantısı olmadığı için kayıt yapılamıyor!")
        elif uploaded_file is not None and m3_teklif > 0:
            with st.spinner("PDF okunuyor ve buluta kaydediliyor..."):
                try:
                    with open("temp.pdf", "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    
                    caplar = []
                    cinsi_nevi = "Belirtilmemiş"
                    boy_bilgisi = "Belirtilmemiş"
                    isletme_mudurlugu = "Belirtilmemiş"
                    parti_no = "Belirtilmemiş"
                    toplam_m3 = 0.0

                    with pdfplumber.open("temp.pdf") as pdf:
                        text = pdf.pages[0].extract_text()
                        
                        isletme_match = re.search(r'İşletme\s*Müdürlüğü\s*[:\-]?\s*(.*)', text, re.IGNORECASE)
                        if isletme_match:
                            ham_isletme = isletme_match.group(1).strip()
                            
                            parti_match = re.search(r'Parti\s*(?:No)?\s*[:\-]?\s*(\d+)', text, re.IGNORECASE)
                            if parti_match:
                                parti_no = parti_match.group(1).strip()
                            
                            isletme_mudurlugu = re.split(r'\s*Parti', ham_isletme, flags=re.IGNORECASE)[0].strip()
                        
                        cins_match = re.search(r'Cinsi\s*[-–]\s*Nevi\s*[:\-]?\s*(.*)', text, re.IGNORECASE)
                        if not cins_match:
                            cins_match = re.search(r'Cinsi\s*Nevi\s*[:\-]?\s*(.*)', text, re.IGNORECASE)
                        
                        if cins_match:
                            tam_cins = cins_match.group(1).strip()
                            
                            boy_match = re.search(r'Boy\s*[:\-]?\s*([\d\.,\-]+)', tam_cins, re.IGNORECASE)
                            if boy_match:
                                boy_bilgisi = boy_match.group(1).strip()
                                cinsi_nevi = re.sub(r'(?i)\s*Boy\s*[:\-]?\s*[\d\.,\-]+', '', tam_cins).strip()
                            else:
                                cinsi_nevi = tam_cins
                        
                        m3_match = re.search(r'Miktar\s*\(m3\)\s*[:\-]?\s*([\d\.,]+)', text, re.IGNORECASE)
                        if m3_match:
                            temiz_m3 = m3_match.group(1).replace('.', '').replace(',', '.')
                            try:
                                toplam_m3 = float(temiz_m3)
                            except:
                                toplam_m3 = 0.0
                        
                        if toplam_m3 == 0.0:
                            m3_matches = re.findall(r'([\d\.,]+)\s*(?:m3|M3|m³)', text)
                            if m3_matches:
                                temiz_m3 = m3_matches[-1].replace('.', '').replace(',', '.')
                                try:
                                    toplam_m3 = float(temiz_m3)
                                except:
                                    toplam_m3 = 0.0

                        table = pdf.pages[0].extract_table()
                        if table:
                            for row in table:
                                try:
                                    cap = float(str(row[1]).replace(',', '.'))
                                    adet = int(str(row[3]).replace(',', ''))
                                    caplar.extend([cap] * adet)
                                except:
                                    continue
                        
                        ortalama_kutur = round(sum(caplar) / len(caplar), 2) if caplar else 0.0

                    if toplam_m3 == 0.0:
                        toplam_m3 = 0.01 
                    
                    bugun_tarih = datetime.now().strftime("%d.%m.%Y")
                    
                    yeni_satir = [
                        bugun_tarih, 
                        isletme_mudurlugu, 
                        parti_no,
                        cinsi_nevi, 
                        boy_bilgisi,
                        str(toplam_m3).replace('.', ','), 
                        str(ortalama_kutur).replace('.', ','), 
                        str(km_mesafe).replace('.', ','), 
                        str(m3_teklif).replace('.', ',')
                    ]
                    
                    sheet.append_row(yeni_satir)
                    st.success("İhale başarıyla buluta (Google Sheets) kaydedildi!")
                    
                except Exception as e:
                    st.error(f"Veri işlenirken hata oluştu: {e}")
        else:
            st.warning("Lütfen PDF yükleyin ve teklif fiyatını girin.")

with col_sag:
    st.subheader("📊 Buluttaki Geçmiş Alımlar")
    
    if sheets_baglantisi:
        try:
            tum_veriler = sheet.get_all_records()
            if tum_veriler:
                df = pd.DataFrame(tum_veriler)
                istenilen_sutunlar = ["Tarih", "İşletme", "Parti No", "Cins / Nevi", "Boy", "Toplam m³", "Ort. Kutur (cm)", "Mesafe (KM)", "m³ Teklifimiz (TL)"]
                
                if set(istenilen_sutunlar).issubset(df.columns):
                    df = df[istenilen_sutunlar]
                    
                    st.markdown("### 🔍 Tabloyu Filtrele")
                    filtre_col1, filtre_col2, filtre_col3 = st.columns(3)
                    
                    benzersiz_isletmeler = sorted([str(x) for x in df["İşletme"].unique() if str(x) != "nan" and str(x).strip() != ""])
                    benzersiz_cinsler = sorted([str(x) for x in df["Cins / Nevi"].unique() if str(x) != "nan" and str(x).strip() != ""])
                    # Boy verilerindeki .0 gibi ondalıkları temizleyip listeye alıyoruz
                    benzersiz_boylar = sorted(list(set([str(x).replace('.0', '') for x in df["Boy"].unique() if str(x) != "nan" and str(x).strip() != ""])))
                    
                    with filtre_col1:
                        secilen_isletmeler = st.multiselect("📍 İşletme", benzersiz_isletmeler, placeholder="İşletme seç...")
                    with filtre_col2:
                        secilen_cinsler = st.multiselect("🌳 Cins / Nevi", benzersiz_cinsler, placeholder="Cins seç...")
                    with filtre_col3:
                        secilen_boylar = st.multiselect("📏 Boy", benzersiz_boylar, placeholder="Boy seç...")
                    
                    df_filtrelenmis = df.copy()
                    
                    # Filtreleme yaparken DataFrame sütunlarını astype(str) ile geçici olarak yazıya çevirip eşleştiriyoruz (.0'ları temizleyerek)
                    if secilen_isletmeler:
                        df_filtrelenmis = df_filtrelenmis[df_filtrelenmis["İşletme"].astype(str).isin(secilen_isletmeler)]
                        
                    if secilen_cinsler:
                        df_filtrelenmis = df_filtrelenmis[df_filtrelenmis["Cins / Nevi"].astype(str).isin(secilen_cinsler)]
                        
                    if secilen_boylar:
                        df_filtrelenmis = df_filtrelenmis[df_filtrelenmis["Boy"].astype(str).str.replace('.0', '', regex=False).isin(secilen_boylar)]
                    
                    st.dataframe(df_filtrelenmis.iloc[::-1], use_container_width=True, hide_index=True)
                else:
                    st.warning("Google Sheets başlıklarında eksik var! 'Boy' sütununu eklediğinden emin ol.")
                    st.dataframe(df.iloc[::-1], use_container_width=True)
                    
            else:
                st.info("Henüz kayıt bulunmuyor.")
        except Exception as e:
            st.warning(f"Tablo verileri çekilirken hata: {e}")
    else:
        st.info("Bağlantı kurulamadığı için veriler gösterilemiyor.")
