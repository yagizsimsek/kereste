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
import sys
import traceback
from datetime import timezone, timedelta

# İhale Radarı geçici olarak kapalı (sekme bilgi mesajı gösteriyor). Tekrar açmak için True yap.
# Sabah mail botu radar_bot.py içindeki aynı isimli ayarla ayrıca kapatılıyor.
RADAR_AKTIF = False

# SSL Uyarılarını Kapat
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Streamlit Cloud sunucusu UTC saatinde çalışıyor — datetime.now() Türkiye'de gece
# 00:00-03:00 arası bir önceki günü veriyordu (gecikti işareti, not tarihleri kayıyordu).
TR_TZ = timezone(timedelta(hours=3))

def simdi():
    return datetime.now(TR_TZ).replace(tzinfo=None)

def bildir(mesaj, tur="success"):
    """st.rerun() öncesi gösterilen st.success mesajı, sayfa yenilendiği için hiç görünmüyordu.
    Mesajı session_state'e koyup sayfa yenilendikten sonra en üstte gösteriyoruz."""
    st.session_state.setdefault("_bildirimler", []).append((tur, mesaj))

_BILDIRIM_IKONLARI = {"success": "✅", "warning": "⚠️", "error": "❌", "info": "ℹ️"}

def bildirimleri_goster():
    # st.toast: köşede çıkan bildirim. Sayfanın üstüne kutu eklemek sekmelerin yerini
    # kaydırıp kullanıcıyı ilk sekmeye geri atıyordu; toast sayfa düzenini değiştirmiyor.
    for tur, mesaj in st.session_state.pop("_bildirimler", []):
        st.toast(mesaj.lstrip("✅🎉🔄⚠️ "), icon=_BILDIRIM_IKONLARI.get(tur, "ℹ️"))

def kalici_bildir(sekme, mesaj, tur="warning"):
    """Sayfa yenilendikten sonra ilgili sekmenin EN ÜSTÜNDE kutu olarak gösterilecek uyarı.
    Köşedeki kısa bildirim (toast) birkaç saniyede kayboluyor; kontrol edilmesi gereken
    sorunlar (eksik boy, okunamayan tutar vb.) gözden kaçmasın diye burada kalıyor."""
    st.session_state.setdefault("_sekme_uyarilari", {}).setdefault(sekme, []).append((tur, mesaj))

def kalici_bildirimleri_goster(sekme):
    for tur, mesaj in st.session_state.get("_sekme_uyarilari", {}).pop(sekme, []):
        getattr(st, tur)(mesaj)

class guvenli_bolum:
    """Bir sekmede beklenmeyen bir hata olursa (Google bağlantısı kopması, kota, bozuk veri vb.)
    kırmızı teknik hata ekranı yerine anlaşılır bir mesaj gösterir; diğer sekmeler çalışmaya
    devam eder. Teknik detay Streamlit Cloud loglarına (Manage app) yazılır.
    NOT: st.rerun()/st.stop() BaseException olduğu için burada yakalanmaz, normal çalışır."""
    def __init__(self, ad):
        self.ad = ad

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        if exc_type is None or not issubclass(exc_type, Exception):
            return False
        print(f"[HATA] {self.ad}:", file=sys.stderr)
        traceback.print_exception(exc_type, exc, tb, file=sys.stderr)
        metin = str(exc)
        if "429" in metin or "Quota exceeded" in metin or "RATE_LIMIT" in metin:
            st.warning("⏳ Google Sheets şu an çok yoğun (kısa süreli istek sınırı). 1 dakika bekleyip sayfayı yenileyin — girdiğiniz işlem kaydedilmediyse tekrar deneyin.")
        elif isinstance(exc, (requests.exceptions.ConnectionError, requests.exceptions.Timeout)):
            st.warning("🌐 İnternet / sunucu bağlantısında geçici bir sorun oldu. Biraz sonra sayfayı yenileyip tekrar deneyin.")
        else:
            st.error(f"⚠️ '{self.ad}' bölümünde beklenmeyen bir sorun oluştu. Sayfayı yenileyip tekrar deneyin; devam ederse bu mesajı iletin.\n\nTeknik detay: `{type(exc).__name__}: {metin[:300]}`")
        return True

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
    """'42,707' ya da '1.234,56' gibi Türkçe ondalıklı bir metni float'a çevirir, boş/bozuksa 0.0 döner.
    NOT: pandas boş hücreleri None/NaN'a çevirebiliyor — 'nan' metni float('nan') ile SESSİZCE
    geçerli bir sayıya (NaN) dönüştüğü için bunu en başta ayrıca eleyip 0.0 döndürüyoruz, yoksa
    NaN toplamlara sessizce karışıp (Örn. tutar toplamları) tüm sonucu NaN yapabiliyordu."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
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
    # Sarıçam/Karaçam kendi başlarına ayrı türler olarak isteniyor, bu yüzden genel
    # "ÇAM" kontrolünden ÖNCE gelmeleri lazım (yoksa "SARIÇAM" içindeki "ÇAM" alt
    # dizesi genel kurala takılıp hepsi düz "Çam" altında toplanırdı).
    ("SARIÇAM", "Sarıçam"), ("KARAÇAM", "Karaçam"), ("ÇAM", "Çam"),
    ("KÖKNAR", "Köknar"), ("GÖKNAR", "Köknar"), ("LADİN", "Ladin"),
    ("KAYIN", "Kayın"), ("MEŞE", "Meşe"), ("KAVAK", "Kavak"), ("GÜRGEN", "Gürgen"),
    ("DİŞBUDAK", "Dişbudak"), ("SEDİR", "Sedir"), ("KESTANE", "Kestane"), ("CEVİZ", "Ceviz"),
]

def agac_turu_cikar(cinsi_metni):
    """'3.Sn.Nb.Kl. Sarıçam Tomruk' gibi bir metinden sınıf/kalite kodlarını görmezden gelip
    ana ağaç türünü çıkarır. Sarıçam ve Karaçam kendi başlarına ayrı gösterilir; diğer
    çam alt türleri (Kızılçam, Fıstıçam vb.) genel 'Çam' altında toplanır."""
    t = tr_upper(cinsi_metni)
    for anahtar, etiket in _AGAC_TURU_ANAHTARLARI:
        if anahtar in t:
            return etiket
    temiz = str(cinsi_metni).strip()
    return temiz if temiz else "Bilinmeyen"

def urun_tipi_cikar(cinsi_metni):
    """Cinsi metninde 'Maden Direği' gibi tomruktan farklı bir ürün tipi geçiyorsa onu
    döndürür (yoksa boş string) — aynı tür+boy'un tomruğu ile maden direği fiyat/talep
    açısından farklı olduğu için özet panellerinde ayrı gösterilmesi isteniyor."""
    t = tr_upper(cinsi_metni)
    if "MADEN DİR" in t:
        return "Maden Direği"
    return ""

def tl_formatla(deger):
    s = f"{deger:,.0f}"
    s = s.replace(",", "§").replace(".", ",").replace("§", ".")
    return f"{s} ₺"

def m3_formatla(deger):
    s = f"{deger:,.1f}"
    s = s.replace(",", "§").replace(".", ",").replace("§", ".")
    return f"{s} m³"

_TR_AYLAR = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
_TR_GUNLER = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]

def tarih_tr_formatla(ts):
    """pandas Timestamp -> '28 Eylül 2026 Pazartesi' gibi Türkçe okunaklı tarih."""
    return f"{ts.day} {_TR_AYLAR[ts.month - 1]} {ts.year} {_TR_GUNLER[ts.weekday()]}"

def siralanabilir_yap(df, sayi_kolonlari=None, tarih_kolonlari=None):
    """Google Sheets'ten ham metin olarak gelen Türkçe ondalıklı sayıları ve
    dd.mm.yyyy tarihlerini gerçek sayı/tarih tipine çevirir. Bunu yapmazsak
    Streamlit tablosunda bir sütun başlığına tıklayıp sıralatmak METİN
    (alfabetik) sıralaması yapıyordu — Örn. '10.09.2026' tarihi '2.06.2026'dan
    önce geliyordu, ya da '9,5' değeri '10,2'den büyük çıkıyordu. Görünüm,
    çağıran taraftaki column_config (siralama_column_config) ile Türkçe/
    dd.mm.yyyy kalmaya devam ediyor, sadece alttaki tip düzeliyor."""
    df = df.copy()
    for kol in (sayi_kolonlari or []):
        if kol in df.columns:
            df[kol] = df[kol].apply(sayi_parse)
    for kol in (tarih_kolonlari or []):
        if kol in df.columns:
            df[kol] = pd.to_datetime(df[kol], format='%d.%m.%Y', errors='coerce')
    return df

def yeniden_eskiye(df, tarih_kolonu):
    """Tabloyu en yeni tarih en üstte olacak şekilde sıralar. Aynı tarihli satırlarda
    Sheets'e sonradan eklenen (daha yeni kayıt) üstte kalır. Tarihi okunamayanlar en alta."""
    ters = df.iloc[::-1]
    if df.empty or tarih_kolonu not in df.columns:
        return ters
    tarih = df[tarih_kolonu]
    if not pd.api.types.is_datetime64_any_dtype(tarih):
        tarih = pd.to_datetime(tarih.astype(str).str.strip(), format='%d.%m.%Y', errors='coerce')
    return ters.assign(_SiraTarih=tarih).sort_values("_SiraTarih", ascending=False, kind="stable", na_position="last").drop(columns="_SiraTarih")

def siralama_column_config(sayi_format_kolonlari=None, tarih_kolonlari=None):
    """sayi_format_kolonlari: {'Kolon Adı': '%.2f'} gibi bir sözlük. Değerler
    gerçek sayı/tarih tipine çevrildikten sonra bu, görünümü eskisi gibi
    (Türkçe / dd.mm.yyyy) tutmak için kullanılıyor."""
    cfg = {}
    for kol, fmt in (sayi_format_kolonlari or {}).items():
        cfg[kol] = st.column_config.NumberColumn(kol, format=fmt)
    for kol in (tarih_kolonlari or []):
        cfg[kol] = st.column_config.DateColumn(kol, format="DD.MM.YYYY")
    return cfg

@st.cache_resource(show_spinner=False)
def _drive_senkron_hafizasi():
    """Sunucu genelinde (tüm kullanıcılar için ortak) sekme -> son yazılan içerik hash'i."""
    return {}

def nakliye_drive_senkronize(spreadsheet, df):
    """Nakliyesi tamamlanmış partileri ana Drive dosyasında 'Nakliye_Tümü' ve her
    İşletme + İhale Tarihi kombinasyonu için ayrı bir sekmede günceller — Excel indirmeye
    gerek kalmadan Drive dosyası her zaman güncel dursun diye.
    Sadece içeriği gerçekten değişen sekmeler yazılıyor ve hafıza tüm kullanıcılar için
    ortak: eskiden her yeni ziyaretçide TÜM sekmeler baştan yazılıyordu (ihale başına
    2-3 istek), bu da Google'ın dakikalık istek sınırını dolduruyordu."""
    if df.empty:
        return False

    yazilacak_df = df.copy()
    if "Fatura" in yazilacak_df.columns:
        yazilacak_df["Fatura"] = yazilacak_df["Fatura"].map(lambda v: "EVET" if v else "")

    hedefler = {"Nakliye_Tümü": yazilacak_df}
    if "İşletme" in yazilacak_df.columns and "İhale Tarihi" in yazilacak_df.columns:
        for (isletme_adi, tarih), grup in yazilacak_df.groupby(["İşletme", "İhale Tarihi"]):
            sekme_adi = f"Nak_{isletme_adi}_{tarih}".strip() or "Bilinmeyen"
            for ch in ['\\', '/', '*', '[', ']', ':', '?']:
                sekme_adi = sekme_adi.replace(ch, '-')
            hedefler[sekme_adi[:40]] = grup

    hafiza = _drive_senkron_hafizasi()
    degisenler = {}
    for baslik, grup_df in hedefler.items():
        h = hashlib.md5(grup_df.to_csv(index=False).encode("utf-8")).hexdigest()
        if hafiza.get(baslik) != h:
            degisenler[baslik] = (grup_df, h)
    if not degisenler:
        return False

    mevcut_sekmeler = {ws.title: ws for ws in spreadsheet.worksheets()}  # tek istek
    yazilacaklar = []
    for baslik, (grup_df, h) in degisenler.items():
        satirlar = [grup_df.columns.tolist()] + grup_df.astype(str).values.tolist()
        gereken_satir, gereken_sutun = len(satirlar) + 5, len(grup_df.columns) + 2
        ws = mevcut_sekmeler.get(baslik)
        if ws is None:
            spreadsheet.add_worksheet(title=baslik, rows=str(gereken_satir), cols=str(gereken_sutun))
        elif ws.row_count < len(satirlar) or ws.col_count < len(grup_df.columns):
            ws.resize(rows=max(ws.row_count, gereken_satir), cols=max(ws.col_count, gereken_sutun))
        yazilacaklar.append((baslik, satirlar))

    # Tüm değişen sekmeler tek "temizle" + tek "yaz" isteğiyle güncelleniyor.
    def _aralik(baslik):
        return "'" + baslik.replace("'", "''") + "'"
    spreadsheet.values_batch_clear(body={"ranges": [_aralik(b) for b, _ in yazilacaklar]})
    spreadsheet.values_batch_update({
        "valueInputOption": "USER_ENTERED",
        "data": [{"range": _aralik(b) + "!A1", "values": satirlar} for b, satirlar in yazilacaklar],
    })
    for baslik, (_, h) in degisenler.items():
        hafiza[baslik] = h
    return True

def nakliyeci_cari_ekle(spreadsheet, nakliyeci_adi, kayitlar):
    """Bir nakliyecinin kendi 'cari' sekmesine bu seferki taşıdığı partileri ekler (İşletme,
    tarih, cinsi/boy, çekilen miktar, ücret, hesaplanan toplam nakliye), ardından farklı seferler
    görsel olarak birbirinden ayırt edilsin diye bir boş satır bırakır. Sekme yoksa otomatik
    oluşturulur. 'Yarısı Bireysel'/'Tevfikatlı Dip Rakam'/'Fatura No' muhasebecinin elle
    dolduracağı sütunlar — hem Google Sheets üzerinden hem indirilen Excel'den düzenlenebilir."""
    if not nakliyeci_adi or not kayitlar:
        return
    sekme_adi = f"Cari_{nakliyeci_adi}".strip()
    for ch in ['\\', '/', '*', '[', ']', ':', '?']:
        sekme_adi = sekme_adi.replace(ch, '-')
    sekme_adi = sekme_adi[:40]
    baslik = ["Tarih", "İşletme", "İhale Tarihi", "Parti No", "Cinsi", "Boy", "Çekilen Miktar (m³)", "Nakliye Ücreti (TL/m³)", "Not", "Toplam Nakliye", "Bireysel", "Tevfikatlı Dip Rakam", "Fatura No"]
    try:
        ws = spreadsheet.worksheet(sekme_adi)
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=sekme_adi, rows="200", cols=str(len(baslik)))
        ws.append_row(baslik)

    # Eski sekmelerde bu yeni sütunlar olmayabilir — eksikse tamamla (Kasa_Takip'teki gibi).
    if ws.col_count < len(baslik):
        ws.add_cols(len(baslik) - ws.col_count)
    mevcut_baslik = ws.row_values(1)
    for i, h in enumerate(baslik):
        if i >= len(mevcut_baslik) or mevcut_baslik[i] != h:
            ws.update_cell(1, i + 1, h)

    genisletilmis_kayitlar = []
    for kayit in kayitlar:
        miktar = sayi_parse(kayit[6]) if len(kayit) > 6 else 0.0
        ucret = sayi_parse(kayit[7]) if len(kayit) > 7 else 0.0
        toplam_nakliye = round(miktar * ucret, 2)
        genisletilmis_kayitlar.append(list(kayit) + [toplam_nakliye, "", "", ""])

    ws.append_rows(genisletilmis_kayitlar, value_input_option='USER_ENTERED')
    ws.append_row([""] * len(baslik))

@st.cache_data(ttl=3600, show_spinner=False)
def ogm_erisilebilir():
    """OGM sitesi Türkiye dışındaki IP'leri tamamen engelliyor (test edildi: 12 ülkeden
    zaman aşımı, İstanbul'dan 0,2 sn). Streamlit Cloud ABD'de olduğu için orada link çekme
    hiç çalışmıyor; kullanıcıyı 20 sn bekletip hata vermek yerine baştan söylüyoruz.
    Site Türkiye'deki bir bilgisayarda (Yağız'ın Mac'i) çalışınca True döner."""
    try:
        requests.get("https://esatis.ogm.gov.tr/", headers={'User-Agent': 'Mozilla/5.0'}, verify=False, timeout=6)
        return True
    except Exception:
        return False

def ogm_getir(url, timeout=20, deneme=3):
    """OGM'den sayfa/PDF indirir; bağlantı hatası, zaman aşımı ya da sunucu hatasında (5xx)
    kısa beklemelerle 3 kez dener. OGM arka arkaya gelen isteklerde ara sıra cevap vermiyor —
    eskiden tek başarısız istek o partinin miktarını/kuturunu sessizce 0 yapıyordu."""
    import time
    son_hata = None
    for i in range(deneme):
        try:
            cevap = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, verify=False, timeout=timeout)
            if cevap.status_code >= 500:
                raise requests.exceptions.HTTPError(f"OGM sunucu hatası {cevap.status_code}", response=cevap)
            if cevap.status_code == 200:
                return cevap
            if cevap.status_code in (429,):
                raise requests.exceptions.HTTPError("OGM çok fazla istek dedi (429)", response=cevap)
            return cevap  # 404 vb.: tekrar denemenin anlamı yok, çağıran taraf kontrol ediyor
        except requests.exceptions.RequestException as e:
            son_hata = e
            print(f"[UYARI] OGM isteği başarısız ({i + 1}/{deneme}): {url} — {e}", file=sys.stderr)
            if i < deneme - 1:
                time.sleep(1.5 * (i + 1))
    raise son_hata

class OgmSayfaHatasi(Exception):
    pass

def ogm_link_duzelt(metin):
    """Yapıştırılan metinden OGM ihale numarasını bulup sonuç sayfası linkini döndürür, bulamazsa None."""
    t = str(metin or "").strip()
    if not t:
        return None
    m = re.search(r'ihale/(\d+)', t) or re.fullmatch(r'(\d{3,})', t)
    if not m:
        return None
    return f"https://esatis.ogm.gov.tr/ihale/{m.group(1)}/sonuc"

st.set_page_config(page_title="Kereste İhale & Maliyet Sistemi", layout="wide")

# --- GOOGLE SHEETS BAĞLANTISI ---
# Streamlit her tıklamada (filtre, checkbox vs.) tüm dosyayı baştan çalıştırıyor.
# Eskiden her seferinde bağlantı + tüm sekmeler yeniden okunuyordu (~15-20 okuma),
# bu da Google'ın dakikalık okuma sınırını (429 Quota exceeded) aşıyordu.
# Artık bağlantı bir kez kuruluyor, okumalar da kısa süre önbellekte tutuluyor;
# herhangi bir yazma işleminden sonra önbellek temizleniyor ki veri hep güncel kalsın.
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

@st.cache_resource(show_spinner=False)
def sheets_baglan():
    if os.path.exists("credentials.json"):
        creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    else:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)

    client = gspread.authorize(creds)
    # Google cevap vermezse sayfa sonsuza kadar donmasın.
    client.set_timeout(30)
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
        kasa_sheet = spreadsheet.add_worksheet(title="Kasa_Takip", rows="100", cols="20")
        kasa_sheet.append_row(["İşletme", "İhale Tarihi", "Parti No", "Cinsi", "Boy", "Miktar", "Birim Fiyat", "Taksitli Tutar", "Nakit Tutar", "Son Ödeme Tarihi", "Durum", "Not", "Nakliye Durumu", "Nakliye Notu", "Alan Firma", "Çekilen Miktar", "Fatura", "Nakliye Ücreti", "Nakliyeci", "OGM Fatura"])

    try:
        nakliyeci_sheet = spreadsheet.worksheet("Nakliyeciler")
    except:
        nakliyeci_sheet = spreadsheet.add_worksheet(title="Nakliyeciler", rows="100", cols="1")
        nakliyeci_sheet.append_row(["Nakliyeci Adı"])

    return spreadsheet, {
        "sheet": sheet,
        "takip": takip_sheet,
        "kasa": kasa_sheet,
        "nakliyeci": nakliyeci_sheet,
    }

@st.cache_data(ttl=300, show_spinner=False)
def _tum_sekmeleri_oku():
    """5 sekmenin hepsini TEK istekte okur (eskiden 5 ayrı istek sırayla gidiyordu).
    Uygulamadan yapılan her yazma önbelleği hemen temizliyor; 5 dakikalık süre sadece
    Google Sheets'te elle yapılan değişiklikler için — onlar için de 'Verileri Yenile' butonu var."""
    spreadsheet, sekmeler = sheets_baglan()
    anahtarlar = list(sekmeler)
    araliklar = ["'" + sekmeler[a].title.replace("'", "''") + "'" for a in anahtarlar]
    cevap = spreadsheet.values_batch_get(araliklar)
    return {
        a: gspread.utils.fill_gaps(vr.get("values", []))
        for a, vr in zip(anahtarlar, cevap.get("valueRanges", []))
    }

def _sekme_oku(anahtar):
    return _tum_sekmeleri_oku()[anahtar]

class OnbellekliSekme:
    """Worksheet sarmalayıcı: okumalar önbellekten gelir, yazma olunca önbellek temizlenir."""
    _YAZMA_METOTLARI = {
        "append_row", "append_rows", "update_cell", "update_cells", "update",
        "batch_update", "delete_rows", "insert_row", "insert_rows", "clear",
        "add_cols", "add_rows", "resize", "format",
    }

    def __init__(self, anahtar, ws):
        self._anahtar = anahtar
        self._ws = ws

    def get_all_values(self):
        return [list(r) for r in _sekme_oku(self._anahtar)]

    def col_values(self, n):
        degerler = [r[n - 1] if len(r) >= n else "" for r in _sekme_oku(self._anahtar)]
        while degerler and degerler[-1] == "":
            degerler.pop()
        return degerler

    def row_values(self, n):
        satirlar = _sekme_oku(self._anahtar)
        if n > len(satirlar):
            return []
        satir = list(satirlar[n - 1])
        while satir and satir[-1] == "":
            satir.pop()
        return satir

    def taze_satirlar(self):
        """Önbelleği atlayıp Sheets'ten en güncel hali okur — yazmadan hemen önce, ekranda
        görülen satırın bu arada (başka biri/elle düzenleme ile) değişip değişmediğini kontrol için."""
        _tum_sekmeleri_oku.clear()
        return self.get_all_values()

    def __getattr__(self, ad):
        deger = getattr(self._ws, ad)
        if ad in self._YAZMA_METOTLARI and callable(deger):
            def _yaz_ve_temizle(*args, **kwargs):
                try:
                    return deger(*args, **kwargs)
                finally:
                    _tum_sekmeleri_oku.clear()
            return _yaz_ve_temizle
        return deger

try:
    spreadsheet, _sekmeler = sheets_baglan()
    sheet = OnbellekliSekme("sheet", _sekmeler["sheet"])
    takip_sheet = OnbellekliSekme("takip", _sekmeler["takip"])
    kasa_sheet = OnbellekliSekme("kasa", _sekmeler["kasa"])
    nakliyeci_sheet = OnbellekliSekme("nakliyeci", _sekmeler["nakliyeci"])
    # İlk okumayı burada yap ki kota hatası olursa sayfanın ortasında değil, en üstte yakalansın.
    _tum_sekmeleri_oku()
    sheets_baglantisi = True
except Exception as e:
    sheets_baglantisi = False
    hata_mesaji = e

_baslik_col, _yenile_col = st.columns([5, 1])
with _baslik_col:
    st.title("🌲 Kereste İhale & Maliyet Takip Sistemi")
with _yenile_col:
    st.write("")
    if st.button("🔄 Verileri Yenile", help="Google Sheets'te elle bir değişiklik yaptıysanız hemen görmek için basın.", use_container_width=True):
        _tum_sekmeleri_oku.clear()
        st.rerun()

if not sheets_baglantisi:
    if "429" in str(hata_mesaji) or "Quota exceeded" in str(hata_mesaji):
        st.error("Google Sheets kısa süreliğine çok fazla istek aldı. 1 dakika bekleyip sayfayı yenileyin.")
    else:
        st.error(f"Google Sheets'e bağlanılamadı! Hata: {hata_mesaji}")
    if st.button("🔄 Tekrar Dene"):
        st.rerun()
    st.stop()

def satirlar_degismedi_mi(taze_veri, gorulen_veri, satir_nolari):
    """Yazmadan önce: ekranda görülen satırlar Sheets'teki güncel haliyle aynı mı?
    Aradan biri satır silmiş/düzenlemişse ya da aynı partiyi başka biri işlemişse
    yanlış satıra yazmamak için işlemi durdururuz."""
    for n in satir_nolari:
        if n - 1 >= len(taze_veri) or n - 1 >= len(gorulen_veri):
            return False
        a = [str(x).strip() for x in taze_veri[n - 1]]
        b = [str(x).strip() for x in gorulen_veri[n - 1]]
        uzunluk = max(len(a), len(b))
        a += [""] * (uzunluk - len(a))
        b += [""] * (uzunluk - len(b))
        if a != b:
            return False
    return True

TABLO_DEGISTI_MESAJI = "🔄 Bu kayıt siz ekrandayken değişmiş (başka biri işlem yapmış ya da tablo elle düzenlenmiş). Yanlış satıra yazılmasın diye işlem yapılmadı — sayfa yenilendi, lütfen kontrol edip tekrar deneyin."

bildirimleri_goster()

# Site açılınca ilk (varsayılan) sekme Geçmiş Alımlar; Yeni İhale Çek en sağda.
tab_gecmis, tab_odeme, tab_nakliye, tab_radar, tab_islem = st.tabs([
    "📊 Geçmiş Alımlar",
    "💳 Kasa & Ödeme Takibi",
    "🚚 Nakliye Takibi",
    "🔔 İhale Radarı",
    "📥 Yeni İhale Çek",
])

# --- İHALE ÇEKME SEKMESİ ---
# Her sekme bağımsız bir 'fragment': içindeki bir tıklama/filtre sadece o sekmeyi
# yeniden çalıştırıyor (eskiden 5 sekmenin hepsi baştan çiziliyor, site soluklaşıp kasıyordu).
# Kayıt sonrası st.rerun() yine tüm sayfayı yeniliyor ki diğer sekmeler de güncellensin.
@st.fragment
def _sekme_islem():
    with guvenli_bolum("Yeni İhale Çek"):
        kalici_bildirimleri_goster("islem")
        st.subheader("📥 Yeni İhale Ekle / Çek")
        islem_turu = st.radio("İşlem Türü Seçin:", ["🔗 OGM Sonuç Linkinden Toplu Çek (Bot)", "📄 İhale Öncesi PDF'den Hesapla (yakında)"])

        if islem_turu == "🔗 OGM Sonuç Linkinden Toplu Çek (Bot)" and not ogm_erisilebilir():
            st.warning(
                "🇹🇷 **Link ile ihale çekme bu sunucuda çalışmıyor.** OGM sitesi Türkiye dışından gelen "
                "bağlantıları engelliyor, bu internet sitesinin sunucusu ise yurtdışında.\n\n"
                "İhale çekmek için siteyi **Yağız'ın bilgisayarından** açın: kereste klasöründeki "
                "**'Siteyi Bilgisayarda Aç'** dosyasına çift tıklayın. Oradan çekilen ihaleler aynı "
                "Google Sheets'e yazılır ve buradaki sitede de görünür (en geç 5 dakika içinde ya da "
                "sağ üstteki '🔄 Verileri Yenile' ile hemen)."
            )
        elif islem_turu == "🔗 OGM Sonuç Linkinden Toplu Çek (Bot)":
            ihale_linki_ham = st.text_area(
                "OGM İhale Sonuç Linki — birden fazla ihale için linklerin arasına boşluk bırakın (ya da alt alta yapıştırın)",
                placeholder="Örn: https://esatis.ogm.gov.tr/ihale/207249/sonuc https://esatis.ogm.gov.tr/ihale/207250/sonuc",
                height=80,
            )
            # Yanlış/eksik yapıştırılan linkler (sonunda /sonuc yok, sadece ihale numarası
            # girilmiş vb.) için ihale numarasını yakalayıp doğru sonuç sayfası linkini kuruyoruz.
            linkler = []
            gecersiz_parcalar = []
            for _parca in re.split(r'[\s,;]+', ihale_linki_ham.strip()):
                if not _parca:
                    continue
                _link = ogm_link_duzelt(_parca)
                if _link is None:
                    gecersiz_parcalar.append(_parca)
                elif _link not in linkler:
                    linkler.append(_link)
            if gecersiz_parcalar:
                st.warning("⚠️ Şunlar geçerli bir OGM ihale linki gibi görünmüyor, atlanacak: " + ", ".join(f"`{g[:60]}`" for g in gecersiz_parcalar) + " — link 'esatis.ogm.gov.tr/ihale/NUMARA/...' şeklinde olmalı (ya da sadece ihale numarasını yazabilirsiniz).")
            if len(linkler) > 1:
                st.caption(f"🔗 {len(linkler)} ihale linki algılandı — hepsi sırayla çekilecek.")

            # --- LİNK YAPIŞTIRILINCA TESPİT EDİLEN YER ÖNİZLEMESİ ---
            if linkler:
                _onizleme_hafiza = st.session_state.setdefault("_onizleme_hafiza", {})
                for _link_no, ihale_linki in enumerate(linkler, 1):
                    if ihale_linki not in _onizleme_hafiza:
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
                        _onizleme_hafiza[ihale_linki] = (onizleme_yer, onizleme_hata)
                    onizleme_yer, onizleme_hata = _onizleme_hafiza[ihale_linki]

                    _on_ek = f"**{_link_no}.** " if len(linkler) > 1 else ""
                    if onizleme_yer:
                        st.caption(f"{_on_ek}📍 Tespit edilen yer: **{onizleme_yer}**")
                    elif onizleme_hata == "timeout":
                        st.warning(f"{_on_ek}⏱️ OGM sunucusu 15 saniye içinde cevap vermedi (site yavaş olabilir). 'Kazandıklarımızı Çek ve Kaydet' yine de denenebilir, 3 kez tekrar deniyor.")
                    else:
                        st.warning(f"{_on_ek}⚠️ Linkten yer bilgisi tespit edilemedi (link hatalı olabilir ya da ihale henüz sonuçlanmamış).")
            # -------------------------------------------------------------

            if st.button("Kazandıklarımızı Çek ve Kaydet", type="primary", use_container_width=True):
                if not linkler:
                    st.warning("Lütfen en az bir geçerli OGM sonuç linki yapıştırın.")
                else:
                    with st.spinner("Taktik devrede, OGM taranıyor... Lütfen bekleyin..."):
                        # --- ANA VERİTABANI MÜKERRER KONTROLÜ (İŞLETME + PARTİ) ---
                        mevcut_gecmis = sheet.taze_satirlar()
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

                        for _link_no, ihale_linki in enumerate(linkler, 1):
                            if len(linkler) > 1:
                                st.markdown(f"---\n**🔗 {_link_no}/{len(linkler)}. ihale:** {ihale_linki}")
                            try:
                                res = ogm_getir(ihale_linki, timeout=20)
                                if res.status_code != 200:
                                    raise OgmSayfaHatasi(res.status_code)
                                soup = BeautifulSoup(res.text, 'html.parser')

                                isletme_text = "Bilinmeyen İşletme"
                                isletme_match = re.search(r'([A-ZÇĞİÖŞÜ\s]+(?:OİM|OBM))', soup.text)
                                if isletme_match:
                                    isletme_text = isletme_kisalt(isletme_match.group(1))

                                # --- CLAUDE TAKTİĞİ (DATA-MILLIS OKUMA) KESİN ÇÖZÜMÜ ---
                                genel_ihale_tarihi = "Tarih Bulunamadı"
                            
                                tarih_elementleri = soup.find_all(attrs={"data-millis": True})
                                for el in tarih_elementleri:
                                    try:
                                        millis = int(el.get('data-millis'))
                                        if millis > 1000000000000: 
                                            genel_ihale_tarihi = datetime.fromtimestamp(millis / 1000.0, tz=TR_TZ).strftime('%d.%m.%Y')
                                            break 
                                    except:
                                        continue

                                if genel_ihale_tarihi == "Tarih Bulunamadı":
                                    bugun_str = simdi().strftime("%d.%m.%Y")
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

                                                # --- MİKTAR ve ADET: doğrudan sonuç tablosundan ---
                                                # Sonuç sayfasında miktar tam değeriyle duruyor (<span data-value="5.355">).
                                                # Eskiden miktar sadece PDF'ten okunuyordu; PDF o an inmezse ya da biçimi
                                                # değişirse (m3 → m³) miktar sessizce 0 yazılıyordu.
                                                parti_notlari = []  # bu partide ters giden her şey, kullanıcıya gösterilecek
                                                _miktar_span = row.find(attrs={"data-value": True})
                                                if _miktar_span:
                                                    try:
                                                        miktar_float = float(_miktar_span["data-value"])
                                                    except (ValueError, TypeError):
                                                        miktar_float = 0.0
                                                tablo_adet = 0
                                                _adet_metni = re.sub(r'\D', '', cols[3].get_text(strip=True)) if len(cols) > 3 else ""
                                                if _adet_metni:
                                                    tablo_adet = int(_adet_metni)

                                                # --- KUTUR ve BASKIN BOY: İstif Ebat Listesi PDF'inden ---
                                                t_adet = 0
                                                detay_a = row.find('a', href=True)
                                                if not detay_a:
                                                    parti_notlari.append("detay sayfası linki bulunamadı (kutur PDF'ten okunamadı)")
                                                else:
                                                    detay_linki = urljoin(ihale_linki, detay_a['href'])
                                                    pdf_isim = None
                                                    try:
                                                        d_res = ogm_getir(detay_linki, timeout=15)
                                                        d_soup = BeautifulSoup(d_res.text, 'html.parser')

                                                        pdf_link = None
                                                        for a_tag in d_soup.find_all('a', href=True):
                                                            if 'pdf' in a_tag['href'].lower() or 'ebat' in a_tag.text.lower():
                                                                pdf_link = urljoin(detay_linki, a_tag['href'])
                                                                break

                                                        if not pdf_link:
                                                            parti_notlari.append("detay sayfasında İstif Ebat Listesi PDF'i bulunamadı")
                                                        else:
                                                            p_res = ogm_getir(pdf_link, timeout=20)
                                                            # Benzersiz dosya adı: aynı anda iki kullanıcı/sekme bot
                                                            # çalıştırırsa birbirinin PDF'ini bozmasın diye.
                                                            pdf_isim = f"temp_bot_{uuid.uuid4().hex}.pdf"
                                                            with open(pdf_isim, "wb") as f:
                                                                f.write(p_res.content)

                                                            with pdfplumber.open(pdf_isim) as pdf:
                                                                pdf_text = pdf.pages[0].extract_text() or ""

                                                                # Miktar tablodan okunamadıysa PDF başlığındaki "Miktar (m³) : 5,355" (m3 de olabilir)
                                                                if miktar_float == 0.0:
                                                                    m3_match = re.search(r'Miktar\s*\(m(?:3|³)\)\s*[:\-]?\s*([\d\.,]+)', pdf_text, re.IGNORECASE)
                                                                    if m3_match:
                                                                        try:
                                                                            miktar_float = float(m3_match.group(1).replace('.', '').replace(',', '.'))
                                                                        except ValueError:
                                                                            pass

                                                                caplar = []
                                                                boy_adet = {}
                                                                # Uzun listeler birden fazla sayfaya taşabiliyor — tüm sayfaların tablosu okunuyor.
                                                                for _sayfa in pdf.pages:
                                                                    table_pdf = _sayfa.extract_table()
                                                                    if not table_pdf or len(table_pdf) < 2:
                                                                        continue
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
                                                                        # "Toplam" özet satırını atla (boy boş, adette genel toplam yazar).
                                                                        if p_row[boy_idx] is None or str(p_row[boy_idx]).strip() == '' or 'toplam' in str(p_row[0]).strip().lower():
                                                                            continue
                                                                        adet_match = re.search(r'\d+', str(p_row[adet_idx]).strip().replace('.', ''))
                                                                        if not adet_match:
                                                                            continue
                                                                        adet = int(adet_match.group())
                                                                        t_adet += adet

                                                                        b_match = re.search(r'[\d\.]+', str(p_row[boy_idx]).strip().replace(',', '.'))
                                                                        if b_match:
                                                                            try:
                                                                                s_b = str(float(b_match.group()))
                                                                                t_boy = s_b[:-2] if s_b.endswith('.0') else s_b
                                                                                boy_adet[t_boy] = boy_adet.get(t_boy, 0) + adet
                                                                            except ValueError:
                                                                                pass

                                                                        c_match = re.search(r'[\d\.]+', str(p_row[cap_idx]).strip().replace(',', '.'))
                                                                        if c_match:
                                                                            try:
                                                                                cap = float(c_match.group())
                                                                                if cap > 0:
                                                                                    caplar.extend([cap] * adet)
                                                                            except ValueError:
                                                                                pass

                                                                # %80 KURALI: partinin %80'i tek boydaysa boy o kabul edilir
                                                                if t_adet > 0:
                                                                    for b_deg, b_ad in boy_adet.items():
                                                                        if (b_ad / t_adet) >= 0.80:
                                                                            hesaplanan_boy = b_deg
                                                                            break

                                                                if caplar:
                                                                    hesaplanan_kutur = round(sum(caplar) / len(caplar), 2)
                                                                else:
                                                                    parti_notlari.append("PDF'teki çap tablosu okunamadı")
                                                    except requests.exceptions.RequestException as e:
                                                        print(f"[HATA] Parti {parti_no} detay/PDF indirilemedi: {e}", file=sys.stderr)
                                                        parti_notlari.append(f"OGM'den detay sayfası/PDF 3 denemede de indirilemedi ({type(e).__name__})")
                                                    except Exception as e:
                                                        print(f"[HATA] Parti {parti_no} PDF işlenemedi: {e}", file=sys.stderr)
                                                        parti_notlari.append(f"PDF işlenemedi ({type(e).__name__}: {str(e)[:80]})")
                                                    finally:
                                                        # Hata çıksa bile geçici dosya diskte unutulmasın.
                                                        if pdf_isim and os.path.exists(pdf_isim):
                                                            os.remove(pdf_isim)

                                                # Kutur PDF'ten okunamadıysa miktar + adet + boydan tahmini hesap (silindir formülü).
                                                if hesaplanan_kutur == 0.0:
                                                    _adet_hesap = t_adet or tablo_adet
                                                    try:
                                                        _h_boy = float(str(hesaplanan_boy).replace(',', '.'))
                                                    except ValueError:
                                                        _h_boy = 0.0
                                                    if _adet_hesap > 0 and miktar_float > 0 and _h_boy > 0:
                                                        import math
                                                        hesaplanan_kutur = round(math.sqrt((miktar_float * 40000) / (math.pi * _h_boy * _adet_hesap)), 2)
                                                        parti_notlari.append(f"kutur PDF'ten okunamadığı için miktar/adet/boydan tahmini hesaplandı ({hesaplanan_kutur:g} cm)")
                                                    else:
                                                        parti_notlari.append("kutur hesaplanamadı, 0 kaydedildi")

                                                if miktar_float == 0.0:
                                                    parti_notlari.append("MİKTAR okunamadı, 0 kaydedildi")
                                                if fiyat_int == 0:
                                                    parti_notlari.append("FİYAT (verdiğimiz pey) okunamadı, 0 kaydedildi")

                                                yeni_satir = [satir_ihale_tarihi, isletme_text, alan_firma, str(parti_no), cins, str(hesaplanan_boy), float(round(miktar_float, 3)), float(round(hesaplanan_kutur, 2)), "", int(fiyat_int), ""]  # Mesafe ve Nakliye Ücreti sütunları artık kullanılmıyor, boş bırakılıyor (sütun sırası kaymasın diye)
                                                eklenecek_satirlar.append(yeni_satir)
                                                mevcut_gecmis_set.add(kayit_id_bot)
                                                if parti_notlari:
                                                    supheli_partiler.append(f"Parti {parti_no} ({cins}, boy {hesaplanan_boy}): " + "; ".join(parti_notlari))

                                if len(eklenecek_satirlar) > 0:
                                    sheet.append_rows(eklenecek_satirlar, value_input_option='USER_ENTERED')
                                    st.success(f"🎉 Helal olsun! {len(eklenecek_satirlar)} adet yeni ihale işlendi! (Zaten kayıtlı olan {atlanan_adet} parti atlandı).")

                                    if isletme_text == "Bilinmeyen İşletme":
                                        st.error("❌ Sayfada işletme adı (… OİM/OBM) bulunamadı — partiler 'Bilinmeyen İşletme' olarak kaydedildi. Google Sheets'te İşletme sütununu elle düzeltin.")
                                    if genel_ihale_tarihi == "Tarih Bulunamadı":
                                        st.error("❌ Sayfada ihale tarihi bulunamadı — Tarih sütununa 'Tarih Bulunamadı' yazıldı. Google Sheets'te elle düzeltin.")
                                    if supheli_partiler:
                                        st.warning(f"⚠️ {len(supheli_partiler)} partide sorun var — kayıt yapıldı ama aşağıdaki değerleri Google Sheets'te kontrol edin:\n\n" + "\n".join(f"- {p}" for p in supheli_partiler))
                                    else:
                                        st.caption("✔️ Tüm partilerin miktar, fiyat, boy ve kutur bilgisi sorunsuz okundu.")
                                elif atlanan_adet > 0:
                                    st.warning(f"Bu sayfadaki kazandığımız {atlanan_adet} partinin tümü zaten veritabanında var, o yüzden yeniden eklenmedi (Mükerrer koruması devrede).")
                                elif not dogru_tablo:
                                    st.error("Bu sayfada ihale sonuç tablosu bulunamadı. İhale henüz sonuçlanmamış olabilir ya da link başka bir sayfaya ait — linki kontrol edip tekrar deneyin.")
                                else:
                                    st.error("Sayfa tarandı ancak firmalarımızın kazandığı herhangi bir parti bulunamadı.")
                            except OgmSayfaHatasi as e:
                                st.error(f"🌐 OGM bu ihale sayfasını açmadı (kod {e.args[0]}). İhale numarası yanlış olabilir ya da OGM sitesinde geçici bir sorun var — linki kontrol edip biraz sonra tekrar deneyin.")
                            except requests.exceptions.Timeout:
                                st.error("⏱️ OGM sunucusu 20 saniye içinde cevap vermedi. Bu genelde OGM'nin sitesi yavaş çalıştığında ya da Streamlit Cloud'un sunucu adresini geçici olarak yavaşlattığında olur — bir kaç dakika sonra tekrar dene. Sürekli oluyorsa bana söyle, başka bir çözüm bulalım.")
                            except requests.exceptions.RequestException as e:
                                print(f"[HATA] OGM bağlantısı: {e}", file=sys.stderr)
                                st.error("🌐 OGM sunucusuna bağlanılamadı. İnternet bağlantısını ya da OGM sitesinin açık olup olmadığını kontrol edip biraz sonra tekrar deneyin.")
                            # Diğer hatalar (Sheets kotası vb.) sekmenin genel güvenlik ağında anlaşılır mesajla gösteriliyor.
                            # Bot tekrar çalıştırılırsa mükerrer koruması zaten kaydedilmiş partileri atlıyor.

        else:
            st.info("📄 İhale öncesi PDF'den hesaplama özelliği henüz hazır değil. Şimdilik ihale sonuçlandıktan sonra yukarıdaki 'OGM Sonuç Linkinden Toplu Çek' seçeneğini kullanın.")

with tab_islem:
    _sekme_islem()


# --- GEÇMİŞ ALIMLAR ---
# Her sekme bağımsız bir 'fragment': içindeki bir tıklama/filtre sadece o sekmeyi
# yeniden çalıştırıyor (eskiden 5 sekmenin hepsi baştan çiziliyor, site soluklaşıp kasıyordu).
# Kayıt sonrası st.rerun() yine tüm sayfayı yeniliyor ki diğer sekmeler de güncellensin.
@st.fragment
def _sekme_gecmis():
    with guvenli_bolum("Geçmiş Alımlar"):
        kalici_bildirimleri_goster("gecmis")
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

                        # turkce_sayiyi_duzelt hücre hücre int/float/orijinal metin döndürür;
                        # bir sütunun TÜM değerleri sayısalsa sütunu gerçek float dtype'a
                        # çeviriyoruz — yoksa sütun "object" (karışık int/float) kalıp
                        # Streamlit tablosunda başlığa tıklayınca metin gibi (alfabetik)
                        # sıralanıyordu (Örn. 9,5 > 10,2 yanlış çıkıyordu).
                        for col in df.columns:
                            _dolu = df[col].dropna()
                            if not _dolu.empty and _dolu.apply(lambda v: isinstance(v, (int, float))).all():
                                df[col] = pd.to_numeric(df[col], errors='coerce')

                        # Tarih sütunu da gerçek tarih tipine çevrilmezse aynı şekilde
                        # alfabetik sıralanıp (10.09.2026, 2.06.2026'dan önce gelir gibi) yanlış sonuç verirdi.
                        if "Tarih" in df.columns:
                            _tarih_ham = df["Tarih"].astype(str).str.strip()
                            df["Tarih"] = pd.to_datetime(_tarih_ham, format='%d.%m.%Y', errors='coerce')
                            _bozuk_tarih = df[df["Tarih"].isna()]
                            if not _bozuk_tarih.empty:
                                _ornekler = ", ".join(f"{r.get('İşletme', '?')} Parti {r.get('Parti No', '?')} ('{_tarih_ham[i]}')" for i, r in _bozuk_tarih.head(8).iterrows())
                                st.warning(f"⚠️ {len(_bozuk_tarih)} satırın tarihi okunamadı (GG.AA.YYYY olmalı) — bu satırlar tarih filtresinde ve sıralamada yanlış yerde görünür. Google Sheets'te (Sayfa1) düzeltin: {_ornekler}")

                        # Mesafe ve nakliye ücreti artık kullanılmıyor — Sheets'teki eski veriye
                        # dokunmadan sadece ekrandan/filtreden/Excel'den kaldırıyoruz.
                        df = df.drop(columns=[c for c in ["Mesafe (KM)", "Nakliye Ücreti (TL/m³)"] if c in df.columns])

                        df = yeniden_eskiye(df, "Tarih").reset_index(drop=True)
                        df_filtered = df.copy()

                        st.markdown("##### 🔍 Tabloyu Filtrele")
                        with st.expander("Filtreleri Göster / Gizle", expanded=False):
                            # Tarih aralığı: boş bırakılan taraf sınırsız sayılıyor. Varsayılanı
                            # verinin ilk/son tarihi yapmıyoruz — yoksa sonradan eklenen yeni
                            # ihaleler eski "son tarih" yüzünden sessizce filtre dışında kalırdı.
                            if "Tarih" in df.columns and pd.api.types.is_datetime64_any_dtype(df["Tarih"]):
                                _tarih_col1, _tarih_col2, _ = st.columns([1, 1, 1])
                                with _tarih_col1:
                                    _tarih_bas = st.date_input("Tarih — başlangıç", value=None, format="DD.MM.YYYY", key="gecmis_tarih_bas")
                                with _tarih_col2:
                                    _tarih_bit = st.date_input("Tarih — bitiş", value=None, format="DD.MM.YYYY", key="gecmis_tarih_bit")
                                if _tarih_bas and _tarih_bit and _tarih_bas > _tarih_bit:
                                    st.warning("Başlangıç tarihi bitişten sonra olamaz — tarihleri kontrol edin.")
                                if _tarih_bas:
                                    df_filtered = df_filtered[df_filtered["Tarih"] >= pd.Timestamp(_tarih_bas)]
                                if _tarih_bit:
                                    df_filtered = df_filtered[df_filtered["Tarih"] <= pd.Timestamp(_tarih_bit)]

                            num_columns = 3
                            filter_cols = st.columns(num_columns)

                            # Parti No / Toplam m³ / Ort. Kutur gibi her satırda farklı olan sayılarla
                            # filtrelemek anlamsız — bunlar tabloda duruyor ama filtre kutusu yok.
                            _filtresiz = {"Parti No", "Toplam m³", "Ort. Kutur (cm)"}
                            _filtre_kolonlari = [c for c in df.columns if c not in _filtresiz and not pd.api.types.is_datetime64_any_dtype(df[c])]
                            for i, col_name in enumerate(_filtre_kolonlari):

                                unique_values = df[col_name].dropna().unique().tolist()
                                try: unique_values.sort()
                                except TypeError: unique_values.sort(key=lambda x: str(x))

                                selected_values = filter_cols[i % num_columns].multiselect(
                                    label=f"{col_name}",
                                    options=unique_values,
                                    default=[],
                                    placeholder="Seçin...",
                                )

                                if selected_values:
                                    df_filtered = df_filtered[df_filtered[col_name].isin(selected_values)]

                        st.dataframe(
                            df_filtered,
                            column_config=siralama_column_config(tarih_kolonlari=["Tarih"]),
                            use_container_width=True,
                        )
                        
                        # Filtrelenen partilerin ortalama m³ fiyatı — sadece ihale fiyatı (m³ Teklifimiz),
                        # nakliye dahil değil. Ana değer miktara göre ağırlıklı: 50 m³'lük parti
                        # 5 m³'lükten daha çok etkiliyor (toplam ödenen / toplam m³ = gerçek ortalama).
                        if "m³ Teklifimiz (TL)" in df_filtered.columns and "Toplam m³" in df_filtered.columns and not df_filtered.empty:
                            _fiyat = pd.to_numeric(df_filtered["m³ Teklifimiz (TL)"], errors="coerce")
                            _m3 = pd.to_numeric(df_filtered["Toplam m³"], errors="coerce")
                            _gecerli = (_fiyat > 0) & (_m3 > 0)
                            _toplam_m3 = float(_m3[_gecerli].sum())
                            if _toplam_m3 > 0:
                                _agirlikli_ort = float((_fiyat[_gecerli] * _m3[_gecerli]).sum() / _toplam_m3)
                                _basit_ort = float(_fiyat[_gecerli].mean())
                                _ort_col1, _ort_col2, _ort_col3 = st.columns(3)
                                _ort_col1.metric(
                                    "💵 Ortalama m³ Fiyatı (ihale)", tl_formatla(_agirlikli_ort),
                                    help="Filtrelenen partilerde m³ başına ödenen ortalama ihale fiyatı, miktara göre ağırlıklı (toplam ihale bedeli ÷ toplam m³). Nakliye dahil değil.",
                                )
                                _ort_col2.metric("📦 Toplam Miktar", m3_formatla(_toplam_m3))
                                _ort_col3.metric("🧾 Toplam İhale Bedeli", tl_formatla(_agirlikli_ort * _toplam_m3))
                                st.caption(f"Basit ortalama (her parti eşit sayılırsa): {tl_formatla(_basit_ort)} / m³ — {int(_gecerli.sum())} parti üzerinden.")

                        col1, col2 = st.columns(2)
                        with col1:
                            st.caption(f"Filtrelenmiş Sonuç: **{len(df_filtered)}** / Toplam: **{len(df)}** adet ihale gösteriliyor.")
                        with col2:
                            # CSV yerine Excel: Türkçe Excel CSV'yi ';' ile ayırıyor ve UTF-8'i
                            # tanımıyor — dosya tek sütunda, ş/ğ/İ harfleri bozuk açılıyordu.
                            _indir_df = df_filtered.copy()
                            if "Tarih" in _indir_df.columns:
                                _indir_df["Tarih"] = _indir_df["Tarih"].dt.strftime('%d.%m.%Y')
                            _gecmis_excel = io.BytesIO()
                            with pd.ExcelWriter(_gecmis_excel, engine='openpyxl') as writer:
                                _indir_df.to_excel(writer, sheet_name='Geçmiş Alımlar', index=False)
                            st.download_button(
                                label="📥 Süzülmüş Tabloyu İndir (Excel)",
                                data=_gecmis_excel.getvalue(),
                                file_name=f"ihale_gecmisi_{simdi().strftime('%Y%m%d')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            )
                    else:
                        st.info("Henüz kaydedilmiş geçmiş bir ihale bulunmuyor.")
                except Exception as e:
                    print(f"[HATA] Geçmiş Alımlar: {e}", file=sys.stderr)
                    if "429" in str(e) or "Quota exceeded" in str(e):
                        st.warning("⏳ Google Sheets şu an çok yoğun. 1 dakika bekleyip sayfayı yenileyin.")
                    else:
                        st.error(f"Geçmiş alımlar gösterilirken bir sorun oluştu. Sayfayı yenileyip tekrar deneyin; devam ederse bu mesajı iletin. (Teknik detay: {type(e).__name__}: {e})")

with tab_gecmis:
    _sekme_gecmis()


# --- KASA VE ÖDEME TAKİP SEKMESİ ---
# Her sekme bağımsız bir 'fragment': içindeki bir tıklama/filtre sadece o sekmeyi
# yeniden çalıştırıyor (eskiden 5 sekmenin hepsi baştan çiziliyor, site soluklaşıp kasıyordu).
# Kayıt sonrası st.rerun() yine tüm sayfayı yeniliyor ki diğer sekmeler de güncellensin.
@st.fragment
def _sekme_odeme():
    with guvenli_bolum("Kasa & Ödeme Takibi"):
        kalici_bildirimleri_goster("odeme")
        st.subheader("💳 Kasa ve Son Ödeme Tarihi Takibi")

        if sheets_baglantisi:

            # --- TABLO SÜTUN ONARICI ---
            kasa_data = kasa_sheet.get_all_values()
            headers = kasa_data[0] if len(kasa_data) > 0 else []
            ideal_headers = ["İşletme", "İhale Tarihi", "Parti No", "Cinsi", "Boy", "Miktar", "Birim Fiyat", "Taksitli Tutar", "Nakit Tutar", "Son Ödeme Tarihi", "Durum", "Not", "Nakliye Durumu", "Nakliye Notu", "Alan Firma", "Çekilen Miktar", "Fatura", "Nakliye Ücreti", "Nakliyeci", "OGM Fatura"]

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

                with st.expander("📊 Bekleyen Satışların Özeti"):
                    _odeme_tutar = 0.0
                    if not df_bekleyen.empty:
                        for _c in ["Taksitli Tutar", "Nakit Tutar"]:
                            if _c in df_bekleyen.columns:
                                _odeme_tutar += float(df_bekleyen[_c].apply(sayi_parse).sum())

                    col_ozet1, col_ozet2 = st.columns([1, 2])
                    with col_ozet1:
                        st.metric("💰 Bekleyen Toplam Tutar", tl_formatla(_odeme_tutar))
                        st.metric("📦 Bekleyen Parti Sayısı", len(df_bekleyen))
                    with col_ozet2:
                        if not df_bekleyen.empty and "Cinsi" in df_bekleyen.columns and "Miktar" in df_bekleyen.columns:
                            # Türe göre değil, tür+boy kombinasyonuna göre kırıyoruz (Örn. "Çam 3" ve
                            # "Çam 4" ayrı gösterilsin) — ama sınıf/kalite koduna kadar inmiyoruz.
                            _boy_kolonu = df_bekleyen["Boy"].astype(str).str.strip().replace("", "?") if "Boy" in df_bekleyen.columns else "?"
                            _tur_ozet = (
                                df_bekleyen.assign(
                                    _AgacTuru=df_bekleyen["Cinsi"].apply(agac_turu_cikar),
                                    _Boy=_boy_kolonu,
                                    _UrunTipi=df_bekleyen["Cinsi"].apply(urun_tipi_cikar),
                                    _MiktarSayi=df_bekleyen["Miktar"].apply(sayi_parse),
                                )
                                .groupby(["_AgacTuru", "_Boy", "_UrunTipi"])["_MiktarSayi"].sum()
                                .sort_values(ascending=False)
                            )
                            if not _tur_ozet.empty:
                                st.markdown("**Ağaç Türü ve Boya Göre Bekleyen Miktar**")
                                _tur_cols = st.columns(min(len(_tur_ozet), 4))
                                for i, ((tur, boy, urun_tipi), miktar) in enumerate(_tur_ozet.items()):
                                    etiket = f"{tur} {boy}" + (f" ({urun_tipi})" if urun_tipi else "")
                                    _tur_cols[i % len(_tur_cols)].metric(f"🌲 {etiket}", m3_formatla(miktar))
                        else:
                            st.caption("Ağaç türü kırılımı için henüz veri yok.")

                if not df_bekleyen.empty and "Son Ödeme Tarihi" in df_bekleyen.columns:
                    with st.expander("📅 Güne Göre Ödeme Takvimi (Toplu Görünüm)"):
                        _gunluk_df = df_bekleyen.assign(
                            _GunTarih=pd.to_datetime(df_bekleyen["Son Ödeme Tarihi"], format='%d.%m.%Y', errors='coerce'),
                            _Taksitli=df_bekleyen["Taksitli Tutar"].apply(sayi_parse) if "Taksitli Tutar" in df_bekleyen.columns else 0.0,
                            _Nakit=df_bekleyen["Nakit Tutar"].apply(sayi_parse) if "Nakit Tutar" in df_bekleyen.columns else 0.0,
                        ).dropna(subset=["_GunTarih"])

                        _gunluk_ozet = _gunluk_df.groupby("_GunTarih")[["_Taksitli", "_Nakit"]].sum().sort_index()

                        if _gunluk_ozet.empty:
                            st.caption("Gösterilecek ödeme tarihi bulunamadı.")
                        else:
                            _bugun_ts = pd.Timestamp(simdi().date())
                            _gun_cols = st.columns(3)
                            for i, (gun, satir) in enumerate(_gunluk_ozet.iterrows()):
                                _gecikti_mi = gun < _bugun_ts
                                with _gun_cols[i % 3]:
                                    with st.container(border=True):
                                        baslik = ("🔴 " if _gecikti_mi else "🗓️ ") + tarih_tr_formatla(gun)
                                        st.markdown(f"**{baslik}**")
                                        st.markdown(f"Taksitli: **{tl_formatla(satir['_Taksitli'])}**")
                                        st.markdown(f"Nakit: **{tl_formatla(satir['_Nakit'])}**")
                                        st.caption(f"Toplam: {tl_formatla(satir['_Taksitli'] + satir['_Nakit'])}")

                st.markdown("---")
                st.markdown("### ⏳ Son Ödeme Tarihi Yaklaşanlar (Tarih Sıralı)")

                if not df_bekleyen.empty:
                    df_bekleyen['Tarih_Formatli'] = pd.to_datetime(df_bekleyen['Son Ödeme Tarihi'].astype(str).str.strip(), format='%d.%m.%Y', errors='coerce')
                    _bozuk_odeme = df_bekleyen[df_bekleyen['Tarih_Formatli'].isna()]
                    if not _bozuk_odeme.empty:
                        st.warning(f"⚠️ {len(_bozuk_odeme)} partinin son ödeme tarihi okunamadı — GECİKTİ uyarısı ve ödeme takvimi bu partiler için ÇALIŞMAZ. Google Sheets'te (Kasa_Takip) düzeltin: " + ", ".join(f"{r['İşletme']} Parti {r['Parti No']} ('{r['Son Ödeme Tarihi']}')" for _, r in _bozuk_odeme.head(8).iterrows()))
                    df_bekleyen = df_bekleyen.sort_values(by='Tarih_Formatli', ascending=True)

                    # Son ödeme tarihi geçmiş ama hâlâ ödenmemiş partiler gözden kaçmasın diye
                    # "Durum" sütununda büyük harfle GECİKTİ yazılıyor ve satır kırmızıya boyanıyor.
                    bugun_ts = pd.Timestamp(simdi().date())
                    df_bekleyen['_Gecikti'] = df_bekleyen['Tarih_Formatli'] < bugun_ts

                    gorsel_kolonlar_kasa = [c for c in ["İşletme", "Alan Firma", "İhale Tarihi", "Parti No", "Cinsi", "Boy", "Miktar", "Birim Fiyat", "Taksitli Tutar", "Nakit Tutar", "Son Ödeme Tarihi", "Durum"] if c in df_bekleyen.columns]
                    gorsel_df_kasa = df_bekleyen[gorsel_kolonlar_kasa].copy()
                    gorsel_df_kasa = siralanabilir_yap(
                        gorsel_df_kasa,
                        sayi_kolonlari=["Parti No", "Miktar", "Birim Fiyat", "Taksitli Tutar", "Nakit Tutar"],
                        tarih_kolonlari=["İhale Tarihi", "Son Ödeme Tarihi"],
                    )
                    if "Durum" in gorsel_df_kasa.columns:
                        gorsel_df_kasa.loc[df_bekleyen['_Gecikti'].values, "Durum"] = "🔴 GECİKTİ"

                    # NOT: pandas Styler (satırı kırmızıya boyama) ile column_config (tarih/sayı
                    # biçimlendirme + doğru sıralama) Streamlit'te birlikte çalışmıyor — Styler
                    # kullanınca column_config sessizce yok sayılıp tarihler ham ISO, sayılar 6
                    # haneli ondalık görünüyordu. Sıralama/biçim önceliği olduğu için Styler'ı
                    # kaldırdık; GECİKTİ hâlâ "Durum" sütununda kırmızı emoji + büyük harfle duruyor.
                    st.dataframe(
                        gorsel_df_kasa,
                        column_config=siralama_column_config(
                            sayi_format_kolonlari={"Parti No": "%d", "Miktar": "%.3f", "Birim Fiyat": "%.2f", "Taksitli Tutar": "%.2f", "Nakit Tutar": "%.2f"},
                            tarih_kolonlari=["İhale Tarihi", "Son Ödeme Tarihi"],
                        ),
                        use_container_width=True,
                    )

                    df_bekleyen = df_bekleyen.drop(columns=['Tarih_Formatli', '_Gecikti'])

                    st.markdown("### ✅ Ödemeyi Gerçekleştir ve Listeden Sil")

                    col_secim, col_not, col_btn = st.columns([2, 2, 1])

                    with col_secim:
                        secenekler = []
                        for idx, row in df_bekleyen.iterrows():
                            firma_etiket = f" [{row['Alan Firma']}]" if row.get('Alan Firma') else ""
                            secenekler.append(f"Satır {row['SheetRow']} | {row['İşletme']}{firma_etiket} - Parti No: {row['Parti No']} - Taksitli: {row['Taksitli Tutar']} ₺ - Nakit: {row['Nakit Tutar']} ₺")

                        secilen_islemler = st.multiselect("Ödemesi Yapılan Parti(ler)i Seç — birden fazla seçilebilir", secenekler, placeholder="Parti seçin...", key="odeme_multiselect")

                    with col_not:
                        islem_notu = st.text_input("Satış / Ödeme Notu Ekle", placeholder="Örn: Ziraat Kartından Nakit İndirimli Çekildi")

                    with col_btn:
                        st.write("")
                        st.write("")
                        if st.button("💳 Ödendi Olarak İşaretle", type="primary", use_container_width=True):
                            if not secilen_islemler:
                                st.warning("Önce en az bir parti seçin.")
                            else:
                                gercek_satir_nolari = [int(x.split("|")[0].replace("Satır", "").strip()) for x in secilen_islemler]
                                durum_col_num = headers.index("Durum") + 1
                                not_col_num = headers.index("Not") + 1

                                with st.spinner("Ödeme Google Sheets'e işleniyor..."):
                                    if not satirlar_degismedi_mi(kasa_sheet.taze_satirlar(), kasa_data, gercek_satir_nolari):
                                        bildir(TABLO_DEGISTI_MESAJI, "warning")
                                        st.rerun()
                                    # Seçilen tüm partiler TEK istekte yazılıyor — yarıda bağlantı koparsa
                                    # bazıları ödendi bazıları ödenmedi diye yarım kalmasın.
                                    _odeme_hucreleri = []
                                    for _satir_no in gercek_satir_nolari:
                                        _odeme_hucreleri.append(gspread.Cell(_satir_no, durum_col_num, "ÖDENDİ"))
                                        _odeme_hucreleri.append(gspread.Cell(_satir_no, not_col_num, islem_notu))
                                    kasa_sheet.update_cells(_odeme_hucreleri, value_input_option='USER_ENTERED')

                                    bildir(f"✅ {len(gercek_satir_nolari)} partinin ödemesi işlendi ve arşive aktarıldı!")
                                    st.rerun()
                else:
                    st.success("🎉 Mükemmel! Şu an ödeme bekleyen hiçbir parti bulunmuyor. Kasa tertemiz!")

                st.markdown("---")
                st.markdown("### ✅ Ödemesi Gerçekleşen (Arşiv) Partiler")

                df_odenen = df_kasa[df_kasa['_DurumTemiz'] == "ÖDENDİ"].copy()
                if not df_odenen.empty:
                    df_odenen = yeniden_eskiye(df_odenen, "İhale Tarihi").reset_index(drop=True)
                    # OGM'nin satış faturası geldi mi? (Nakliye sekmesindeki "Fatura" nakliyecinin
                    # faturası — bu ayrı bir sütun: "OGM Fatura".)
                    if "OGM Fatura" in df_odenen.columns:
                        df_odenen["OGM Fatura"] = df_odenen["OGM Fatura"].astype(str).str.strip().apply(tr_upper) == "EVET"
                    else:
                        df_odenen["OGM Fatura"] = False
                    arsiv_kolonlar_kasa = [c for c in ["İşletme", "Alan Firma", "İhale Tarihi", "Parti No", "Cinsi", "Boy", "Miktar", "Birim Fiyat", "Taksitli Tutar", "Nakit Tutar", "Son Ödeme Tarihi", "Durum", "Not", "OGM Fatura"] if c in df_odenen.columns]
                    gorsel_df_odenen = siralanabilir_yap(
                        df_odenen[arsiv_kolonlar_kasa],
                        sayi_kolonlari=["Parti No", "Miktar", "Birim Fiyat", "Taksitli Tutar", "Nakit Tutar"],
                        tarih_kolonlari=["İhale Tarihi", "Son Ödeme Tarihi"],
                    )
                    _odenen_column_config = siralama_column_config(
                        sayi_format_kolonlari={"Parti No": "%d", "Miktar": "%.3f", "Birim Fiyat": "%.2f", "Taksitli Tutar": "%.2f", "Nakit Tutar": "%.2f"},
                        tarih_kolonlari=["İhale Tarihi", "Son Ödeme Tarihi"],
                    )
                    _odenen_column_config["OGM Fatura"] = st.column_config.CheckboxColumn("🧾 OGM Faturası Geldi mi?", help="OGM'den bu partinin faturası geldiğinde tikleyin.")
                    # Anahtar içeriğe bağlı (nakliye faturasındaki gibi): satırlar kayınca eski bir tik
                    # başka bir partiye uygulanmasın.
                    _ogm_fatura_key = "ogm_fatura_editor_" + hashlib.md5(
                        df_odenen[["SheetRow"] + arsiv_kolonlar_kasa].to_csv(index=False).encode("utf-8")
                    ).hexdigest()[:12]
                    duzenlenen_odenen = st.data_editor(
                        gorsel_df_odenen,
                        column_config=_odenen_column_config,
                        disabled=[c for c in arsiv_kolonlar_kasa if c != "OGM Fatura"],
                        hide_index=True,
                        use_container_width=True,
                        key=_ogm_fatura_key,
                    )

                    if "OGM Fatura" in headers:
                        _ogm_fatura_col = headers.index("OGM Fatura") + 1
                        _ogm_fatura_hucreleri = []
                        for i in range(len(gorsel_df_odenen)):
                            if bool(gorsel_df_odenen.iloc[i]["OGM Fatura"]) != bool(duzenlenen_odenen.iloc[i]["OGM Fatura"]):
                                _ogm_fatura_hucreleri.append(gspread.Cell(
                                    int(df_odenen.iloc[i]["SheetRow"]), _ogm_fatura_col,
                                    "EVET" if duzenlenen_odenen.iloc[i]["OGM Fatura"] else "",
                                ))
                        if _ogm_fatura_hucreleri:
                            if satirlar_degismedi_mi(kasa_sheet.taze_satirlar(), kasa_data, [c.row for c in _ogm_fatura_hucreleri]):
                                kasa_sheet.update_cells(_ogm_fatura_hucreleri, value_input_option='USER_ENTERED')
                                bildir("✅ OGM fatura durumu kaydedildi.")
                            else:
                                bildir(TABLO_DEGISTI_MESAJI, "warning")
                            st.rerun()
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

                        # Mükerrer kontrolü önbellekten değil Sheets'in en güncel halinden —
                        # iki kişi aynı partileri arka arkaya yapıştırırsa çift kayıt oluşmasın.
                        _kasa_taze = kasa_sheet.taze_satirlar()
                        mevcut_kayitlar = set()
                        if len(_kasa_taze) > 1:
                            for row in _kasa_taze[1:]:
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
                        kasa_sorunlari = []

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

                                _sorunlar = []
                                if bulunan_boy == "-":
                                    _sorunlar.append("boy Geçmiş Alımlar'da bulunamadı ('-' yazıldı — önce bu ihaleyi Yeni İhale Çek ile çekin ya da elle düzeltin)")
                                if miktar == 0.0:
                                    _sorunlar.append("miktar okunamadı (0)")
                                if taksitli_tutar == 0.0:
                                    _sorunlar.append("taksitli tutar okunamadı (0)")
                                if birim_fiyat == 0.0:
                                    _sorunlar.append("birim fiyat okunamadı (0)")

                                if kayit_id not in mevcut_kayitlar:
                                    if _sorunlar:
                                        kasa_sorunlari.append(f"{isletme} Parti {parti_no}: " + "; ".join(_sorunlar))
                                    yeni_kayitlar.append([isletme, ihale_tarihi, parti_no, cinsi, bulunan_boy, miktar, birim_fiyat, taksitli_tutar, nakit_tutar, son_tarih, "BEKLİYOR", "", "", "", bulunan_firma, "", "", "", ""])
                                    mevcut_kayitlar.add(kayit_id)
                                    eklenen_adet += 1
                            except Exception as e:
                                islenemedi_count += 1
                                son_hata = f"{type(e).__name__}: {e}"
                                print(f"[HATA] Kasa yapıştırma satırı işlenemedi: {son_hata}", file=sys.stderr)
                                continue
                                
                        if yeni_kayitlar:
                            kasa_sheet.append_rows(yeni_kayitlar, value_input_option='USER_ENTERED')
                            kalici_bildir("odeme", f"🎉 {eklenen_adet} adet parti ({firma_secimi}) Kasaya eklendi.", "success")
                            if kasa_sorunlari:
                                kalici_bildir("odeme", f"⚠️ Eklenen partilerden {len(kasa_sorunlari)} tanesinde eksik/okunamayan bilgi var — Google Sheets'te (Kasa_Takip) kontrol edin:\n\n" + "\n".join(f"- {x}" for x in kasa_sorunlari))
                            if islenemedi_count:
                                kalici_bildir("odeme", f"❌ Yapıştırılan metinde {islenemedi_count} parti satırı okunamadı ve kasaya EKLENMEDİ (son hata: `{son_hata}`). Bu partileri kontrol edip tekrar yapıştırın.", "error")
                            st.rerun()
                        elif found_count == 0:
                            st.error("❌ Yapıştırılan metinde tanınabilir hiçbir parti satırı bulunamadı. OGM 'Parti Satış' ekranındaki tabloyu (başlıklar dahil) tam olarak kopyaladığınızdan emin olun.")
                        elif islenemedi_count == found_count:
                            st.error(f"❌ {found_count} satır regex ile yakalandı ama hiçbiri işlenemedi (kasaya yazılmadı). Son hata: `{son_hata}`. Bu, 'zaten kayıtlı' değil gerçek bir işleme hatası — lütfen bu hata mesajını ilet.")
                        else:
                            st.warning("⚠️ Yeni parti bulunamadı. Kopyaladığınız verideki ihaleler zaten kasada mevcut.")
                else:
                    st.warning("Lütfen boş kutuya tabloyu yapıştırın.")

with tab_odeme:
    _sekme_odeme()


# --- NAKLİYE TAKİP SEKMESİ ---
# Her sekme bağımsız bir 'fragment': içindeki bir tıklama/filtre sadece o sekmeyi
# yeniden çalıştırıyor (eskiden 5 sekmenin hepsi baştan çiziliyor, site soluklaşıp kasıyordu).
# Kayıt sonrası st.rerun() yine tüm sayfayı yeniliyor ki diğer sekmeler de güncellensin.
@st.fragment
def _sekme_nakliye():
    with guvenli_bolum("Nakliye Takibi"):
        kalici_bildirimleri_goster("nakliye")
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

                with st.expander("📊 Nakliyesi Bekleyen Partilerin Özeti"):
                    _nakliye_tutar = 0.0
                    if not df_bekleyen_nakliye.empty:
                        for _c in ["Taksitli Tutar", "Nakit Tutar"]:
                            if _c in df_bekleyen_nakliye.columns:
                                _nakliye_tutar += float(df_bekleyen_nakliye[_c].apply(sayi_parse).sum())

                    col_nak1, col_nak2 = st.columns([1, 2])
                    with col_nak1:
                        st.metric("💰 Depodaki Malın Değeri", tl_formatla(_nakliye_tutar))
                        st.metric("📦 Bekleyen Parti Sayısı", len(df_bekleyen_nakliye))
                    with col_nak2:
                        if not df_bekleyen_nakliye.empty and "Cinsi" in df_bekleyen_nakliye.columns:
                            _boy_kolonu_nak = df_bekleyen_nakliye["Boy"].astype(str).str.strip().replace("", "?") if "Boy" in df_bekleyen_nakliye.columns else "?"
                            _tur_ozet_nak = (
                                df_bekleyen_nakliye.assign(
                                    _AgacTuru=df_bekleyen_nakliye["Cinsi"].apply(agac_turu_cikar),
                                    _Boy=_boy_kolonu_nak,
                                    _UrunTipi=df_bekleyen_nakliye["Cinsi"].apply(urun_tipi_cikar),
                                )
                                .groupby(["_AgacTuru", "_Boy", "_UrunTipi"])["Kalan Miktar"].sum()
                                .sort_values(ascending=False)
                            )
                            if not _tur_ozet_nak.empty:
                                st.markdown("**Ağaç Türü ve Boya Göre Depoda Kalan Miktar**")
                                _tur_cols_nak = st.columns(min(len(_tur_ozet_nak), 4))
                                for i, ((tur, boy, urun_tipi), miktar) in enumerate(_tur_ozet_nak.items()):
                                    etiket = f"{tur} {boy}" + (f" ({urun_tipi})" if urun_tipi else "")
                                    _tur_cols_nak[i % len(_tur_cols_nak)].metric(f"🌲 {etiket}", m3_formatla(miktar))
                        else:
                            st.caption("Ağaç türü kırılımı için henüz veri yok.")

                st.markdown("---")
                st.markdown("### 📦 Depodan Çekilmeyi Bekleyen Partiler (Ödemesi Yapılmış)")

                if not df_bekleyen_nakliye.empty:
                    df_bekleyen_nakliye['Tarih_Formatli'] = pd.to_datetime(df_bekleyen_nakliye['Son Ödeme Tarihi'], format='%d.%m.%Y', errors='coerce')
                    df_bekleyen_nakliye = yeniden_eskiye(df_bekleyen_nakliye.drop(columns=['Tarih_Formatli']), "İhale Tarihi")

                    gorsel_kolonlar = [c for c in ["İşletme", "Alan Firma", "İhale Tarihi", "Parti No", "Cinsi", "Boy", "Miktar", "Kalan Miktar", "Nakliye Durumu", "Birim Fiyat", "Taksitli Tutar", "Nakit Tutar", "Son Ödeme Tarihi"] if c in df_bekleyen_nakliye.columns]
                    gorsel_df_nakliye_bekleyen = siralanabilir_yap(
                        df_bekleyen_nakliye[gorsel_kolonlar],
                        sayi_kolonlari=["Parti No", "Miktar", "Birim Fiyat", "Taksitli Tutar", "Nakit Tutar"],
                        tarih_kolonlari=["İhale Tarihi", "Son Ödeme Tarihi"],
                    )
                    st.dataframe(
                        gorsel_df_nakliye_bekleyen,
                        column_config=siralama_column_config(
                            sayi_format_kolonlari={"Parti No": "%d", "Miktar": "%.3f", "Kalan Miktar": "%.3f", "Birim Fiyat": "%.2f", "Taksitli Tutar": "%.2f", "Nakit Tutar": "%.2f"},
                            tarih_kolonlari=["İhale Tarihi", "Son Ödeme Tarihi"],
                        ),
                        use_container_width=True,
                    )

                    secenekler_nakliye = []
                    for idx, row in df_bekleyen_nakliye.iterrows():
                        firma_etiket = f" [{row['Alan Firma']}]" if row.get('Alan Firma') else ""
                        secenekler_nakliye.append(f"Satır {row['SheetRow']} | {row['İşletme']}{firma_etiket} - Parti No: {row['Parti No']} - {row.get('Cinsi', '')} - Kalan: {row['Kalan Miktar']:g} m³ / Toplam: {row['_ToplamM3']:g} m³")

                    st.markdown("### 🚚 Bu Seferki Çekimi Kaydet")
                    st.caption("Bu seferki taşımada hangi partilerden çekildiğini seç (birden fazla olabilir). Bir parti tamamen, diğeri yarım kalmış olabilir — seçtiğin her parti için ayrı ayrı miktar gireceksin, varsayılan olarak kalanın tamamı gelir.")

                    secilenler_nakliye = st.multiselect("Bu Seferki Taşımada Hangi Partilerden Çekildi?", secenekler_nakliye, key="nakliye_multiselect", placeholder="Parti seçin...")

                    if secilenler_nakliye:
                        girilen_miktarlar = {}
                        girilen_ucretler = {}
                        for secim in list(secilenler_nakliye):
                            satir_no = int(secim.split("|")[0].replace("Satır", "").strip())
                            _eslesen = df_bekleyen_nakliye[df_bekleyen_nakliye['SheetRow'] == satir_no]
                            if _eslesen.empty:
                                secilenler_nakliye.remove(secim)
                                continue
                            satir_bilgi = _eslesen.iloc[0]
                            kalan = float(satir_bilgi['Kalan Miktar'])
                            col_miktar, col_ucret = st.columns(2)
                            with col_miktar:
                                girilen_miktarlar[satir_no] = st.number_input(
                                    f"{satir_bilgi['İşletme']} - Parti {satir_bilgi['Parti No']} — bu seferki çekilen (m³, kalan: {kalan:g})",
                                    min_value=0.0, max_value=max(kalan, 0.01), value=kalan, step=1.0,
                                    key=f"nakliye_miktar_{satir_no}",
                                )
                            with col_ucret:
                                varsayilan_ucret = 0.0
                                girilen_ucretler[satir_no] = st.number_input(
                                    f"{satir_bilgi['İşletme']} - Parti {satir_bilgi['Parti No']} — m³ başına nakliye ücreti (TL)",
                                    min_value=0.0, value=varsayilan_ucret, step=50.0,
                                    key=f"nakliye_ucret_{satir_no}",
                                    help="Nakliyecinin bu parti için m³ başına aldığı ücret — nakliyeci carisindeki 'Toplam Nakliye' bununla hesaplanır.",
                                )

                        nakliyeci_secenekler = nakliyeci_sheet.col_values(1)[1:] if sheets_baglantisi else []
                        YENI_NAKLIYECI_ETIKET = "➕ Yeni Nakliyeci Ekle..."
                        nakliyeci_secim = st.selectbox("Kim Getirdi?", nakliyeci_secenekler + [YENI_NAKLIYECI_ETIKET], key="nakliyeci_secim")
                        if nakliyeci_secim == YENI_NAKLIYECI_ETIKET:
                            secilen_nakliyeci = st.text_input("Yeni Nakliyeci Adı", placeholder="Örn: Ünal Ercan", key="yeni_nakliyeci_adi").strip()
                        else:
                            secilen_nakliyeci = nakliyeci_secim

                        nakliye_notu = st.text_input("Nakliye Notu (hangi araç / ekstra bilgi)", placeholder="Örn: Kamyonla çekildi", key="nakliye_notu_input")

                        if st.button("🚚 Bu Seferki Çekimi Kaydet", type="primary", use_container_width=True):
                            if "Nakliye Durumu" not in nakliye_headers or "Nakliye Notu" not in nakliye_headers or "Çekilen Miktar" not in nakliye_headers:
                                st.error("Kasa_Takip sayfasında gerekli sütunlar bulunamadı. Kasa & Ödeme sekmesini bir kez açıp tekrar dene.")
                            elif all(m <= 0 for m in girilen_miktarlar.values()):
                                st.warning("En az bir parti için 0'dan büyük miktar gir.")
                            elif not secilen_nakliyeci and nakliyeci_secim == YENI_NAKLIYECI_ETIKET:
                                st.warning("Yeni nakliyecinin adını yazın.")
                            else:
                                with st.spinner("Nakliye bilgisi Google Sheets'e işleniyor..."):
                                    _secilen_satirlar = [int(x.split("|")[0].replace("Satır", "").strip()) for x in secilenler_nakliye]
                                    if not satirlar_degismedi_mi(kasa_sheet.taze_satirlar(), nakliye_kasa_data, _secilen_satirlar):
                                        bildir(TABLO_DEGISTI_MESAJI, "warning")
                                        st.rerun()
                                    nakliye_durum_col = nakliye_headers.index("Nakliye Durumu") + 1
                                    nakliye_not_col = nakliye_headers.index("Nakliye Notu") + 1
                                    cekilen_col = nakliye_headers.index("Çekilen Miktar") + 1
                                    ucret_col = nakliye_headers.index("Nakliye Ücreti") + 1 if "Nakliye Ücreti" in nakliye_headers else None
                                    nakliyeci_col = nakliye_headers.index("Nakliyeci") + 1 if "Nakliyeci" in nakliye_headers else None
                                    bugun_str = simdi().strftime("%d.%m.%Y")
                                    ozet = []
                                    yazilacak_hucreler = []
                                    cari_kayitlar = []

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
                                        girilen_ucret = girilen_ucretler.get(satir_no, 0.0)

                                        eski_not = str(satir_bilgi.get("Nakliye Notu", "") or "").strip()
                                        yeni_not_parcasi = f"{bugun_str}: {girilen:g} m³ çekildi" + (f" ({nakliye_notu})" if nakliye_notu else "")
                                        guncel_not = f"{eski_not} | {yeni_not_parcasi}" if eski_not else yeni_not_parcasi

                                        yazilacak_hucreler.append(gspread.Cell(satir_no, nakliye_durum_col, yeni_durum))
                                        yazilacak_hucreler.append(gspread.Cell(satir_no, nakliye_not_col, guncel_not))
                                        yazilacak_hucreler.append(gspread.Cell(satir_no, cekilen_col, yeni_cekilen))
                                        if ucret_col:
                                            # Muhasebeci burayı düzeltmişse (gerçek fatura tabloda öngörülenden
                                            # farklıysa), bu partinin kaydı güncel/gerçek ücretle kalsın.
                                            yazilacak_hucreler.append(gspread.Cell(satir_no, ucret_col, girilen_ucret))
                                        if nakliyeci_col and secilen_nakliyeci:
                                            eski_nakliyeci = str(satir_bilgi.get("Nakliyeci", "") or "").strip()
                                            guncel_nakliyeci = f"{eski_nakliyeci} | {secilen_nakliyeci}" if eski_nakliyeci and eski_nakliyeci != secilen_nakliyeci else secilen_nakliyeci
                                            yazilacak_hucreler.append(gspread.Cell(satir_no, nakliyeci_col, guncel_nakliyeci))

                                        durum_metni = "tamamen çekildi" if yeni_durum == "NAKLİYE YAPILDI" else f"{yeni_kalan:g} m³ kaldı"
                                        ozet.append(f"{satir_bilgi['İşletme']} Parti {satir_bilgi['Parti No']}: {durum_metni}")

                                        if secilen_nakliyeci:
                                            cari_kayitlar.append([
                                                bugun_str, satir_bilgi['İşletme'], satir_bilgi.get('İhale Tarihi', ''),
                                                satir_bilgi['Parti No'], satir_bilgi.get('Cinsi', ''), satir_bilgi.get('Boy', ''),
                                                girilen, girilen_ucret, nakliye_notu,
                                            ])

                                    # Tüm partiler için tüm hücreler TEK API çağrısıyla yazılıyor — birden
                                    # fazla parti seçiliyken yarıda bağlantı kopması bazı partileri
                                    # güncellenmiş bazılarını güncellenmemiş bırakmasın diye.
                                    if yazilacak_hucreler:
                                        kasa_sheet.update_cells(yazilacak_hucreler, value_input_option='USER_ENTERED')

                                    if secilen_nakliyeci:
                                        try:
                                            if secilen_nakliyeci not in nakliyeci_secenekler:
                                                nakliyeci_sheet.append_row([secilen_nakliyeci])
                                            nakliyeci_cari_ekle(spreadsheet, secilen_nakliyeci, cari_kayitlar)
                                        except Exception as e:
                                            print(f"[HATA] Nakliyeci carisi yazılamadı: {e}", file=sys.stderr)
                                            kalici_bildir("nakliye", f"❌ Çekim Kasa'ya kaydedildi AMA '{secilen_nakliyeci}' nakliyecisinin cari sekmesine yazılamadı ({type(e).__name__}: {str(e)[:120]}). Google Sheets'te 'Cari_{secilen_nakliyeci}' sekmesine bu seferki çekimi elle ekleyin:\n\n" + "\n".join(f"- {k[1]} Parti {k[3]}: {k[6]:g} m³, {k[7]:g} TL/m³" for k in cari_kayitlar), "error")

                                    bildir("✅ Kaydedildi:\n\n" + "\n".join(f"- {o}" for o in ozet))
                                    st.rerun()
                else:
                    st.success("🎉 Depoda bekleyen (ödemesi yapılmış ama henüz çekilmemiş) parti yok!")

                st.markdown("---")
                st.markdown("### 🚛 Nakliyesi Tamamlanmış (Arşiv) Partiler")
                df_nakliye_tamam = df_odemesi_biten[df_odemesi_biten['_NakliyeTemiz'] == "NAKLİYE YAPILDI"].copy()
                if not df_nakliye_tamam.empty:
                    df_nakliye_tamam = yeniden_eskiye(df_nakliye_tamam, "İhale Tarihi").reset_index(drop=True)
                    if "Fatura" in df_nakliye_tamam.columns:
                        df_nakliye_tamam['Fatura'] = df_nakliye_tamam["Fatura"].astype(str).str.strip().apply(tr_upper) == "EVET"
                    else:
                        df_nakliye_tamam['Fatura'] = False

                    arsiv_kolonlar = [c for c in ["İşletme", "Alan Firma", "İhale Tarihi", "Parti No", "Cinsi", "Boy", "Miktar", "Birim Fiyat", "Taksitli Tutar", "Nakit Tutar", "Son Ödeme Tarihi", "Nakliye Ücreti", "Nakliyeci", "Nakliye Notu", "Fatura"] if c in df_nakliye_tamam.columns]
                    gorsel_arsiv = df_nakliye_tamam[arsiv_kolonlar].copy()
                    gorsel_arsiv = siralanabilir_yap(
                        gorsel_arsiv,
                        sayi_kolonlari=["Parti No", "Miktar", "Birim Fiyat", "Taksitli Tutar", "Nakit Tutar", "Nakliye Ücreti"],
                        tarih_kolonlari=["İhale Tarihi", "Son Ödeme Tarihi"],
                    )

                    _arsiv_column_config = siralama_column_config(
                        sayi_format_kolonlari={"Parti No": "%d", "Miktar": "%.3f", "Birim Fiyat": "%.2f", "Taksitli Tutar": "%.2f", "Nakit Tutar": "%.2f", "Nakliye Ücreti": "%.2f"},
                        tarih_kolonlari=["İhale Tarihi", "Son Ödeme Tarihi"],
                    )
                    _arsiv_column_config["Fatura"] = st.column_config.CheckboxColumn("📄 Fatura Geldi mi?", help="Muhasebeci fatura geldiğinde burayı tikleyip geçecek.")

                    # Anahtar tablonun içeriğine bağlı: veri değişince (yeni parti arşive düşünce)
                    # tablo sıfırdan çiziliyor. Sabit anahtarda, önceki bir tik "3. satır" diye
                    # hatırlanıp satırlar kaydığında BAŞKA bir partinin faturasına yazılabiliyordu.
                    _fatura_editor_key = "fatura_editor_" + hashlib.md5(
                        df_nakliye_tamam[["SheetRow"] + arsiv_kolonlar].to_csv(index=False).encode("utf-8")
                    ).hexdigest()[:12]
                    duzenlenen_arsiv = st.data_editor(
                        gorsel_arsiv,
                        column_config=_arsiv_column_config,
                        disabled=[c for c in arsiv_kolonlar if c != "Fatura"],
                        hide_index=True,
                        use_container_width=True,
                        key=_fatura_editor_key,
                    )

                    if "Fatura" in nakliye_headers:
                        fatura_col = nakliye_headers.index("Fatura") + 1
                        fatura_hucreleri = []
                        for i in range(len(gorsel_arsiv)):
                            eski_deger = gorsel_arsiv.iloc[i]['Fatura']
                            yeni_deger = duzenlenen_arsiv.iloc[i]['Fatura']
                            if bool(eski_deger) != bool(yeni_deger):
                                satir_no = int(df_nakliye_tamam.iloc[i]['SheetRow'])
                                fatura_hucreleri.append(gspread.Cell(satir_no, fatura_col, "EVET" if yeni_deger else ""))
                        if fatura_hucreleri:
                            if satirlar_degismedi_mi(kasa_sheet.taze_satirlar(), nakliye_kasa_data, [c.row for c in fatura_hucreleri]):
                                kasa_sheet.update_cells(fatura_hucreleri, value_input_option='USER_ENTERED')
                                bildir("✅ Fatura durumu kaydedildi.")
                            else:
                                bildir(TABLO_DEGISTI_MESAJI, "warning")
                            st.rerun()
                    else:
                        st.caption("⚠️ 'Fatura' sütunu henüz sayfada yok — Kasa & Ödeme sekmesini bir kez açıp tekrar dene, otomatik eklenecek.")
                else:
                    st.caption("Henüz nakliyesi tamamlanmış bir parti bulunmuyor.")

                # --- İHALE BAZLI TAM TABLO (Excel + Drive) ---
                # Yukarıdaki arşiv tablosu SADECE tamamen çekilmiş partileri gösteriyor. Ama
                # muhasebeciye/işletmeye "bu ihalede 4 parti aldık, 2'si çekildi" gibi TÜM
                # tabloyu göstermek için ödemesi yapılmış HER partiyi (çekilsin ya da çekilmesin)
                # durum etiketiyle birlikte ayrı bir tabloda tutuyoruz.
                if not df_odemesi_biten.empty:
                    df_ihale_ozet = yeniden_eskiye(df_odemesi_biten, "İhale Tarihi").reset_index(drop=True)

                    def _nakliye_durumu_ozetle(v):
                        v = tr_upper(str(v).strip())
                        if v == "NAKLİYE YAPILDI":
                            return "✅ Tamamlandı"
                        elif v == "KISMİ ÇEKİLDİ":
                            return "🟡 Kısmi Çekildi"
                        else:
                            return "🔴 Eksik (Çekilmedi)"

                    df_ihale_ozet["Nakliye Durumu Özeti"] = df_ihale_ozet["_NakliyeTemiz"].apply(_nakliye_durumu_ozetle)
                    if "Fatura" in df_ihale_ozet.columns:
                        df_ihale_ozet["Fatura"] = df_ihale_ozet["Fatura"].astype(str).str.strip().apply(tr_upper) == "EVET"
                    else:
                        df_ihale_ozet["Fatura"] = False

                    ihale_ozet_kolonlar = [c for c in ["İşletme", "Alan Firma", "İhale Tarihi", "Parti No", "Cinsi", "Boy", "Miktar", "Kalan Miktar", "Nakliye Durumu Özeti", "Birim Fiyat", "Taksitli Tutar", "Nakit Tutar", "Son Ödeme Tarihi", "Nakliye Ücreti", "Nakliyeci", "Nakliye Notu", "Fatura"] if c in df_ihale_ozet.columns]
                    df_ihale_ozet = df_ihale_ozet[ihale_ozet_kolonlar]

                    st.markdown("---")
                    st.markdown("### 📋 İhale Bazlı Tam Tablo (Çekilen + Eksik Tüm Partiler)")
                    st.caption("Bir ihalede aldığımız partilerin hepsi burada — çekilmemiş olanlar da 'Eksik' etiketiyle görünür, sadece tamamlananları değil.")
                    # NOT: burada Drive/Excel'e giden df_ihale_ozet değil, sadece ekrandaki
                    # gösterimi düzgün sıralansın diye ayrı bir kopya (gorsel_ihale_ozet)
                    # sayı/tarih tipine çevriliyor — Drive/Excel'deki metin biçimi bozulmasın diye.
                    gorsel_ihale_ozet = siralanabilir_yap(
                        df_ihale_ozet,
                        sayi_kolonlari=["Parti No", "Miktar", "Kalan Miktar", "Birim Fiyat", "Taksitli Tutar", "Nakit Tutar", "Nakliye Ücreti"],
                        tarih_kolonlari=["İhale Tarihi", "Son Ödeme Tarihi"],
                    )
                    st.dataframe(
                        gorsel_ihale_ozet,
                        column_config=siralama_column_config(
                            sayi_format_kolonlari={"Parti No": "%d", "Miktar": "%.3f", "Kalan Miktar": "%.3f", "Birim Fiyat": "%.2f", "Taksitli Tutar": "%.2f", "Nakit Tutar": "%.2f", "Nakliye Ücreti": "%.2f"},
                            tarih_kolonlari=["İhale Tarihi", "Son Ödeme Tarihi"],
                        ),
                        use_container_width=True,
                    )

                    st.markdown("#### ☁️ Ana Drive Dosyasına Otomatik Aktarım")
                    st.caption("Bu tablo, 'Kereste_İhale_Sistemi' dosyasında 'Nakliye_Tümü' sekmesine ve her ihale (İşletme + İhale Tarihi) için kendi ayrı sekmesine otomatik olarak işleniyor — indirmene gerek yok, Drive'da hep güncel duruyor.")
                    try:
                        with st.spinner("Drive'daki sekmeler kontrol ediliyor..."):
                            _senkron_oldu = nakliye_drive_senkronize(spreadsheet, df_ihale_ozet)
                        if _senkron_oldu:
                            st.success("✅ Drive'daki 'Nakliye_Tümü' ve ihale bazlı sekmeler güncellendi.")
                        else:
                            st.caption("☁️ Drive sekmeleri zaten güncel.")
                    except Exception as e:
                        # Drive aktarımı yan işlem — başarısız olursa sayfanın geri kalanı etkilenmesin,
                        # bir sonraki açılışta (değişen sekmeler hafızada işaretlenmediği için) tekrar denenir.
                        print(f"[HATA] Drive senkronizasyonu: {e}", file=sys.stderr)
                        traceback.print_exc(file=sys.stderr)
                        st.warning(f"☁️ Drive'daki 'Nakliye_Tümü' ve ihale sekmeleri şu an güncellenemedi ({type(e).__name__}). Sayfa bir sonraki açılışta otomatik tekrar deneyecek — aşağıdaki Excel indirme her zaman çalışır.")

                    # --- İŞLETME + İHALE TARİHİ KOMBİNASYONUNA GÖRE AYRI SEKMELİ EXCEL İNDİRME (opsiyonel, ekstra) ---
                    # Aynı yer (Örn. ALADAĞ) farklı tarihlerde birden fazla ihale olabilir; bunları
                    # tek sekmede birleştirmiyoruz, her ihale (yer + tarih) kendi sekmesinde ayrı duruyor.
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        df_ihale_ozet.to_excel(writer, sheet_name='Tümü', index=False)
                        if "İşletme" in df_ihale_ozet.columns and "İhale Tarihi" in df_ihale_ozet.columns:
                            for (isletme_adi, tarih), grup in df_ihale_ozet.groupby(['İşletme', 'İhale Tarihi']):
                                sheet_adi = f"{isletme_adi}_{tarih}".strip() or 'Bilinmeyen'
                                for ch in ['\\', '/', '*', '[', ']', ':', '?']:
                                    sheet_adi = sheet_adi.replace(ch, '-')
                                sheet_adi = sheet_adi[:31]
                                grup.to_excel(writer, sheet_name=sheet_adi, index=False)
                        elif "İşletme" in df_ihale_ozet.columns:
                            for isletme_adi, grup in df_ihale_ozet.groupby('İşletme'):
                                sheet_adi = str(isletme_adi).strip() or 'Bilinmeyen'
                                for ch in ['\\', '/', '*', '[', ']', ':', '?']:
                                    sheet_adi = sheet_adi.replace(ch, '-')
                                sheet_adi = sheet_adi[:31]
                                grup.to_excel(writer, sheet_name=sheet_adi, index=False)

                    st.download_button(
                        "📥 Nakliye Arşivini Excel Olarak İndir (Her İhale Ayrı Sekmede)",
                        data=excel_buffer.getvalue(),
                        file_name=f"nakliye_arsiv_{simdi().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
            else:
                st.info("Kasa henüz boş ya da veri okunamadı. Önce '💳 Kasa & Ödeme Takibi' sekmesinden veri ekle.")

with tab_nakliye:
    _sekme_nakliye()


# --- RADAR VE BİLDİRİM BÖLÜMÜ ---
# Her sekme bağımsız bir 'fragment': içindeki bir tıklama/filtre sadece o sekmeyi
# yeniden çalıştırıyor (eskiden 5 sekmenin hepsi baştan çiziliyor, site soluklaşıp kasıyordu).
# Kayıt sonrası st.rerun() yine tüm sayfayı yeniliyor ki diğer sekmeler de güncellensin.
@st.fragment
def _sekme_radar():
    with guvenli_bolum("İhale Radarı"):
        if not RADAR_AKTIF:
            st.info("🔕 İhale Radarı geçici olarak kapalı — şu an tarama yapılmıyor ve mail gönderilmiyor.")
            return
        st.subheader("🎯 Bölge Radarı & Mail Testi (PDF + RÖNTGEN Modu)")

        if sheets_baglantisi:
            mevcut_liste = takip_sheet.col_values(1)[1:]

            st.markdown("### 1️⃣ Takip Edilecek İşletmeleri Gir")
            col_ekle, col_bos = st.columns([2, 1])
            with col_ekle:
                yeni_isletme = st.text_input("Bölge Ekle", placeholder="Örn: ZONGULDAK VEYA MENGEN")
                if st.button("Listeye Ekle"):
                    # tr_upper: Python'un .upper()'ı "bilecik"i "BILECIK" (noktasız I) yapıyordu,
                    # OGM'deki "BİLECİK" ile hiç eşleşmediği için radar o bölgeyi sessizce kaçırıyordu.
                    _yeni = tr_upper(yeni_isletme.strip())
                    if not _yeni:
                        st.warning("Önce bölge adını yazın.")
                    elif _yeni in [tr_upper(x.strip()) for x in mevcut_liste]:
                        st.info(f"'{_yeni}' zaten takip listesinde.")
                    else:
                        takip_sheet.append_row([_yeni])
                        bildir(f"✅ '{_yeni}' takip listesine eklendi.")
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
                        bugun = simdi().strftime("%d.%m.%Y")
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

                                        tr_tz = TR_TZ

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

with tab_radar:
    _sekme_radar()
