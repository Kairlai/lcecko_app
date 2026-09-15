import base64
from datetime import datetime, timedelta
import io
from io import BytesIO
import os
import pandas as pd
import qrcode
import requests
import streamlit as st

FILE_PATH = "objednavky_lcecko.csv"
BANK_ACCOUNT = "123456789"  # Doplňte číslo účtu L-Céčka
BANK_CODE = "0800"       # Doplňte kód banky

GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", "")
GITHUB_REPO = st.secrets.get("GITHUB_REPO", "")
ADMIN_PASSWORD = st.secrets.get("ADMIN_PASSWORD", "heslo1234")

st.set_page_config(page_title="L-Céčko | Objednávkový systém", layout="wide", page_icon="🥗")

# Vizuální styl L-Céčka
st.markdown("""
    <style>
    .stApp {
        background-color: #f7f9fb;
    }
    .main-header {
        text-align: center;
        color: #2c3e50;
        margin-bottom: 20px;
    }
    div[data-testid="stColumn"] {
        background: #ffffff;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
    }
    </style>
""", unsafe_allow_html=True)

def get_headers():
    return {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json"
    }

def nacti_objednavky():
    if not GITHUB_TOKEN or not GITHUB_REPO:
        if not os.path.exists(FILE_PATH):
            df_empty = pd.DataFrame(columns=[
                "ID", "Datum_Vytvoreni", "Jmeno", "Prijmeni", "Telefon", "Email", 
                "Adresa", "Poznamka_Kuryr", "Kategorie", "Program", "Delka", 
                "Datum_Od", "Cena_Celkem", "Stav_Platby"
            ])
            df_empty.to_csv(FILE_PATH, index=False)
        return pd.read_csv(FILE_PATH), None

    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{FILE_PATH}"
    res = requests.get(url, headers=get_headers())
    if res.status_code == 200:
        data = res.json()
        sha = data["sha"]
        content_str = base64.b64decode(data["content"]).decode("utf-8")
        return pd.read_csv(io.StringIO(content_str)), sha
    else:
        df_empty = pd.DataFrame(columns=[
            "ID", "Datum_Vytvoreni", "Jmeno", "Prijmeni", "Telefon", "Email", 
            "Adresa", "Poznamka_Kuryr", "Kategorie", "Program", "Delka", 
            "Datum_Od", "Cena_Celkem", "Stav_Platby"
        ])
        return df_empty, None

def uloz_objednavky(df, sha=None):
    if not GITHUB_TOKEN or not GITHUB_REPO:
        df.to_csv(FILE_PATH, index=False)
        return True

    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{FILE_PATH}"
    csv_buffer = io.StringIO()
    df.to_csv(csv_buffer, index=False)
    content_b64 = base64.b64encode(csv_buffer.getvalue().encode("utf-8")).decode("utf-8")

    payload = {"message": "Nova objednavka L-Cecko", "content": content_b64}
    if sha:
        payload["sha"] = sha
    res = requests.put(url, headers=get_headers(), json=payload)
    return res.status_code in [200, 201]

df_orders, current_sha = nacti_objednavky()

# Navigace v postranní liště
st.sidebar.markdown("## 🥗 L-Céčko Plzeň")
rezim = st.sidebar.radio("Navigace:", ["🛒 Objednávka pro zákazníka", "🔐 Správa pro majitelku"])

# ---------------------------------------------------------
# 1. ZÁKAZNICKÝ OBJEDNÁVKOVÝ FORMULÁŘ
# ---------------------------------------------------------
if rezim == "🛒 Objednávka pro zákazníka":
    st.markdown("<h1 class='main-header'>🥗 Objednávkový formulář L-Céčko</h1>", unsafe_allow_html=True)
    st.write("Vyberte si stravovací program nebo sezónní nabídku a my se postaráme o zbytek.")
    
    col_menu, col_user = st.columns([1.2, 1])
    
    with col_menu:
        st.subheader("1. Výběr z nabídky")
        kategorie = st.radio("Kategorie nabídky:", ["Krabičkové diety", "Snídaňové balíčky", "Vánoční cukroví"], horizontal=True)
        
        vybrany_program = ""
        delka_trvani = ""
        cena_za_jednotku = 0
        pocet_jednotek = 1
        
        if kategorie == "Krabičkové diety":
            vybrany_program = st.selectbox("Výběr programu *", [
                "Krabičková strava – redukční program pro ženy (5–5 000 kJ)",
                "Krabičková strava – redukční program pro muže (7–7 500 kJ)",
                "Low Carb",
                "Bezlepkový program",
                "Vegetariánský program"
            ])
            delka_trvani = st.select_slider("Délka programu:", options=["1 týden (5 dní)", "2 týdny (10 dní)", "1 měsíc (20 dní)"])
            ceník = {"1 týden (5 dní)": 1850, "2 týdny (10 dní)": 3500, "1 měsíc (20 dní)": 6800}
            cena_za_jednotku = ceník[delka_trvani]
            
        elif kategorie == "Snídaňové balíčky":
            vybrany_program = st.selectbox("Typ snídaní:", [
                "Fit Snídaně (Kaše, Smoothie, Pečivo)",
                "Gurmán Snídaně (Vajíčka, Sýry, Uzeniny)"
            ])
            delka_trvani = st.select_slider("Délka:", options=["5 dní", "10 dní", "20 dní"])
            ceník_snidani = {"5 dní": 650, "10 dní": 1200, "20 dní": 2200}
            cena_za_jednotku = ceník_snidani[delka_trvani]
            
        elif kategorie == "Vánoční cukroví":
            vybrany_program = "Tradiční domácí vánoční cukroví (Mix)"
            delka_trvani = st.selectbox("Balení:", ["0.5 kg", "1.0 kg", "2.0 kg"])
            pocet_jednotek = st.number_input("Počet balení:", min_value=1, max_value=10, value=1)
            ceník_cukrovi = {"0.5 kg": 490, "1.0 kg": 890, "2.0 kg": 1690}
            cena_za_jednotku = ceník_cukrovi[delka_trvani] * pocet_jednotek

        datum_od = st.date_input("Požadované datum doručení / od:", value=datetime.today() + timedelta(days=2))

    with col_user:
        st.subheader("2. Doručovací údaje & Platba")
        jmeno = st.text_input("Jméno:").strip()
        prijmeni = st.text_input("Příjmení:").strip()
        telefon = st.text_input("Telefonní číslo:").strip()
        email = st.text_input("E-mail:").strip()
        adresa = st.text_input("Adresa doručení (Ulice, č.p., Město):").strip()
        poznamka = st.text_input("Poznámka pro kurýra (zvonek, patro...):").strip()
        
        st.divider()
        st.markdown(f"### **Celková cena: {cena_za_jednotku} Kč**")
        
        if st.button("Odeslat a vygenerovat QR platbu", type="primary", use_container_width=True):
            if not all([jmeno, prijmeni, telefon, email, adresa]):
                st.warning("⚠️ Prosím, vyplňte všechny kontaktní i doručovací údaje.")
            else:
                nove_id = 1 if df_orders.empty else int(df_orders["ID"].max()) + 1
                nova_objednavka = pd.DataFrame([{
                    "ID": nove_id,
                    "Datum_Vytvoreni": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "Jmeno": jmeno,
                    "Prijmeni": prijmeni,
                    "Telefon": telefon,
                    "Email": email,
                    "Adresa": adresa,
                    "Poznamka_Kuryr": poznamka,
                    "Kategorie": kategorie,
                    "Program": vybrany_program,
                    "Delka": delka_trvani,
                    "Datum_Od": datum_od.strftime("%Y-%m-%d"),
                    "Cena_Celkem": cena_za_jednotku,
                    "Stav_Platby": "Čeká na platbu"
                }])
                
                df_aktualni = pd.concat([df_orders, nova_objednavka], ignore_index=True)
                if uloz_objednavky(df_aktualni, current_sha):
                    st.success("🎉 Objednávka byla úspěšně přijata!")
                    
                    # Generování české QR platby (SPD format)
                    spd_str = f"SPD*1.0*ACC:{BANK_ACCOUNT}/{BANK_CODE}*AM:{cena_za_jednotku:.2f}*CC:CZK*X-VS:{nove_id}*MSG:L-Cecko ID {nove_id}"
                    qr_img = qrcode.make(spd_str)
                    buf = BytesIO()
                    qr_img.save(buf, format="PNG")
                    
                    st.markdown("---")
                    st.markdown("### 💳 Podklady pro platbu převodem")
                    col_qr, col_info = st.columns([1, 1.5])
                    with col_qr:
                        st.image(buf.getvalue(), caption="Naskenujte v banking aplikaci", width=200)
                    with col_info:
                        st.write(f"**Číslo účtu:** {BANK_ACCOUNT}/{BANK_CODE}")
                        st.write(f"**Částka:** {cena_za_jednotku} Kč")
                        st.write(f"**Variabilní symbol:** {nove_id}")
                        st.write(f"**Zpráva:** L-Cecko ID {nove_id}")
                else:
                    st.error("❌ Chyba při ukládání objednávky. Zkuste to prosím znovu.")

# ---------------------------------------------------------
# 2. SPRÁVA PRO MAJITELKU A KUCHYŇ
# ---------------------------------------------------------
else:
    st.markdown("<h1 class='main-header'>🔐 Interní dispečink L-Céčka</h1>", unsafe_allow_html=True)
    heslo = st.text_input("Zadejte přístupové heslo:", type="password")
    
    if heslo == ADMIN_PASSWORD:
        st.success("Přístup schválen.")
        
        tab1, tab2, tab3 = st.tabs(["📊 Přehled objednávek", "👨‍🍳 Výkaz pro Kuchyň", "🚗 Seznam pro Kurýra"])
        
        with tab1:
            st.subheader("Kompletní databáze objednávek")
            st.dataframe(df_orders, use_container_width=True, hide_index=True)
            
            if not df_orders.empty:
                csv_data = df_orders.to_csv(index=False).encode("utf-8-sig")
                st.download_button("📥 Stáhnout objednávky do Excelu", data=csv_data, file_name="lcecko_objednavky.csv", mime="text/csv")
                
        with tab2:
            st.subheader("👨‍🍳 Souhrn porcí pro kuchyň")
            if not df_orders.empty:
                vybrany_den = st.date_input("Zobrazit souhrn k datu:", value=datetime.today())
                df_kuchyn = df_orders[df_orders["Kategorie"] == "Krabičkové diety"]
                st.write(f"**Přehled menu k uvaření pro:** {vybrany_den.strftime('%d.%m.%Y')}")
                
                souhrn = df_kuchyn["Program"].value_counts().reset_index()
                souhrn.columns = ["Stravovací Program", "Počet Porcí"]
                st.table(souhrn)
            else:
                st.info("Zatím žádné objednávky v databázi.")
                
        with tab3:
            st.subheader("🚗 Rozvozový arch pro kurýra")
            if not df_orders.empty:
                rozvoz_df = df_orders[["Jmeno", "Prijmeni", "Telefon", "Adresa", "Poznamka_Kuryr", "Program"]]
                st.dataframe(rozvoz_df, use_container_width=True, hide_index=True)
            else:
                st.info("Zatím žádné adresy k rozvozu.")
    elif heslo:
        st.error("❌ Nesprávné heslo.")
