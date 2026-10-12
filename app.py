import base64
from datetime import datetime, timedelta
import io
from io import BytesIO
import os
import random
import urllib.parse
import pandas as pd
import qrcode
import requests
import json
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

FILE_PATH = "objednavky_lcecko.csv"
SETTINGS_PATH = "nastaveni_lcecko.json"
LOGO_PATH = "logo.jpg" 

# --- BANKOVNÍ ÚDAJE L-CÉČKA ---
BANK_ACCOUNT = "4610796083"
BANK_CODE = "0800"
BANK_IBAN = "CZ2208000000004610796083"

GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", "")
GITHUB_REPO = st.secrets.get("GITHUB_REPO", "")
ADMIN_PASSWORD = st.secrets.get("ADMIN_PASSWORD", "heslo1234")

# --- NASTAVENÍ E-MAILU ZE SECRETS ---
EMAIL_SENDER = st.secrets.get("EMAIL_SENDER", "")
EMAIL_PASSWORD = st.secrets.get("EMAIL_PASSWORD", "")
EMAIL_RECEIVER = st.secrets.get("EMAIL_RECEIVER", "")
SMTP_SERVER = st.secrets.get("SMTP_SERVER", "smtp.centrum.cz")
SMTP_PORT = int(st.secrets.get("SMTP_PORT", 465))

st.set_page_config(page_title="L-Céčko | Objednávkový systém", layout="wide", page_icon="🥗")

# Vizuální styl L-Céčka
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    [data-testid="stHeaderActionElements"] {visibility: hidden;}
    header {background-color: transparent !important;}
    footer {visibility: hidden;}
    
    h1, h2, h3, .main-header {
        color: #4A7833 !important;
    }

    div[data-testid="stColumn"] {
        background: #ffffff;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        border-top: 5px solid #4A7833;
    }
    
    div.stButton > button:first-child {
        background-color: #4A7833;
        color: white;
        border: none;
        border-radius: 8px;
        font-weight: bold;
    }
    div.stButton > button:first-child:hover {
        background-color: #385a26;
        color: white;
    }
    
    .summary-card {
        background-color: #e6f0e1;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #4A7833;
        margin-bottom: 20px;
    }
    
    .filter-box {
        background-color: #f8f9fa;
        padding: 15px 15px 5px 15px;
        border-radius: 8px;
        border: 1px solid #dee2e6;
        margin-bottom: 20px;
    }
    
    /* STYLY PRO BĚŽÍCÍ TITULKY */
    .marquee-container {
        width: 100%;
        overflow: hidden;
        background-color: #ffeb3b;
        color: #333;
        padding: 8px 0;
        font-weight: bold;
        border-radius: 5px;
        margin-bottom: 15px;
        white-space: nowrap;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }
    .marquee-content {
        display: inline-block;
        padding-left: 100%;
        animation: marquee 20s linear infinite;
    }
    @keyframes marquee {
        0%   { transform: translate(0, 0); }
        100% { transform: translate(-100%, 0); }
    }
    </style>
""", unsafe_allow_html=True)

def odeslat_email_upozorneni(id_obj, jmeno, prijmeni, adresa, program, delka, datum_od, cena, poznamka):
    if not EMAIL_SENDER or not EMAIL_PASSWORD or not EMAIL_RECEIVER:
        print("E-mailové údaje chybí v konfiguraci.")
        return False
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_SENDER
        msg['To'] = EMAIL_RECEIVER
        msg['Subject'] = f"🥗 Nová objednávka #{id_obj} - {jmeno} {prijmeni}"
        
        body = f"""Dobrý den,

přes objednávkový formulář L-Céčka právě dorazila nová objednávka!

Zákazník: {jmeno} {prijmeni}
Adresa pro kurýra: {adresa}
Poznámka pro kurýra / alergie: {poznamka if poznamka else '-'}

Menu / Program: {program}
Varianta / Délka: {delka}
Od kdy doručovat: {datum_od}
Celková cena k úhradě: {cena:,.0f} Kč

Objednávku si můžete detailně zkontrolovat v dispečinku aplikace.

Hezký den,
Váš systém L-Céčko
"""
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        if SMTP_PORT == 465:
            server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, timeout=5)
        else:
            server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=5)
            server.starttls()
            
        server.login(EMAIL_SENDER, EMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        print(f"Chyba pri odesilani e-mailu: {e}")
        return False

def get_headers():
    return {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json"
    }

def nacti_nastaveni():
    vychozi = {
        "snidane": True, 
        "cukrovi": True, 
        "zakazkova": True, 
        "min_ks_catering": 1,
        "bezici_text": "🎄 Přijímáme objednávky na Vánoční cukroví do 31.10.2026! 🎄 Dotazy / Mimořádné přání: 775 045 554 nebo formulář na www.l-cecko.cz"
    }
    if not GITHUB_TOKEN or not GITHUB_REPO:
        if not os.path.exists(SETTINGS_PATH):
            with open(SETTINGS_PATH, "w") as f:
                json.dump(vychozi, f)
        with open(SETTINGS_PATH, "r") as f:
            return json.load(f), None

    try:
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{SETTINGS_PATH}"
        res = requests.get(url, headers=get_headers(), timeout=5)
        if res.status_code == 200:
            data = res.json()
            sha = data["sha"]
            content_str = base64.b64decode(data["content"]).decode("utf-8")
            loaded = json.loads(content_str)
            for k, v in vychozi.items():
                if k not in loaded:
                    loaded[k] = v
            return loaded, sha
    except Exception as e:
        print(f"Chyba při načítání nastavení z GitHubu: {e}")
    
    return vychozi, None

def uloz_nastaveni(nastaveni_dict, sha=None):
    obsah = json.dumps(nastaveni_dict)
    if not GITHUB_TOKEN or not GITHUB_REPO:
        with open(SETTINGS_PATH, "w") as f:
            f.write(obsah)
        return True

    try:
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{SETTINGS_PATH}"
        content_b64 = base64.b64encode(obsah.encode("utf-8")).decode("utf-8")
        payload = {"message": "Aktualizace nastavení", "content": content_b64}
        if sha:
            payload["sha"] = sha
        res = requests.put(url, headers=get_headers(), json=payload, timeout=5)
        return res.status_code in [200, 201]
    except Exception as e:
        print(f"Chyba při ukládání nastavení na GitHub: {e}")
        return False

def nacti_objednavky():
    default_columns = [
        "ID", "Datum_Vytvoreni", "Jmeno", "Prijmeni", "Telefon", "Email", 
        "Adresa", "Poznamka_Kuryr", "Kategorie", "Program", "Delka", 
        "Datum_Od", "Cena_Celkem", "Stav_Platby"
    ]
    
    if not GITHUB_TOKEN or not GITHUB_REPO:
        if not os.path.exists(FILE_PATH):
            df_empty = pd.DataFrame(columns=default_columns)
            df_empty.to_csv(FILE_PATH, index=False)
            return df_empty, None
        return pd.read_csv(FILE_PATH), None

    try:
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{FILE_PATH}"
        res = requests.get(url, headers=get_headers(), timeout=5)
        if res.status_code == 200:
            data = res.json()
            sha = data["sha"]
            content_str = base64.b64decode(data["content"]).decode("utf-8")
            return pd.read_csv(io.StringIO(content_str)), sha
    except Exception as e:
        print(f"Chyba při načítání objednávek z GitHubu: {e}")
        
    df_empty = pd.DataFrame(columns=default_columns)
    return df_empty, None

def uloz_objednavky(df, sha=None):
    if not GITHUB_TOKEN or not GITHUB_REPO:
        df.to_csv(FILE_PATH, index=False)
        return True

    try:
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{FILE_PATH}"
        csv_buffer = io.StringIO()
        df.to_csv(csv_buffer, index=False)
        content_b64 = base64.b64encode(csv_buffer.getvalue().encode("utf-8")).decode("utf-8")

        payload = {"message": "Aktualizace objednavek L-Cecko", "content": content_b64}
        if sha:
            payload["sha"] = sha
        res = requests.put(url, headers=get_headers(), json=payload, timeout=5)
        return res.status_code in [200, 201]
    except Exception as e:
        print(f"Chyba při ukládání objednávek na GitHub: {e}")
        return False

def vytvor_profi_excel(df, titulek="Objednávky"):
    wb = Workbook()
    ws = wb.active
    ws.title = titulek
    
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
        
    green_fill = PatternFill(start_color="4A7833", end_color="4A7833", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    thin_border = Border(
        left=Side(style='thin', color='DDDDDD'),
        right=Side(style='thin', color='DDDDDD'),
        top=Side(style='thin', color='DDDDDD'),
        bottom=Side(style='thin', color='DDDDDD')
    )
    
    for cell in ws[1]:
        cell.fill = green_fill
        cell.font = white_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            cell.border = thin_border
            cell.alignment = Alignment(vertical='center')
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        ws.column_dimensions[col_letter].width = (max_length + 2)
        
    ws.auto_filter.ref = ws.dimensions
    
    output = BytesIO()
    wb.save(output)
    return output.getvalue()

def detekuj_alergie(text):
    """Skenuje poznámky a hledá klíčová slova k alergiím."""
    if pd.isna(text):
        return ""
    text_low = str(text).lower()
    klicova_slova = ["alerg", "lepek", "bezlep", "laktóz", "laktoz", "ořech", "mlék", "mlek", "česnek", "cibul", "ryb"]
    for slovo in klicova_slova:
        if slovo in text_low:
            return "🚨 ALERGIE/VÝJIMKA!"
    return ""

def vygeneruj_html_uctenku(id_obj, jmeno, prijmeni, adresa, telefon, program, delka, datum, cena):
    """Generátor krásné účtenky pro zákazníka."""
    html_obsah = f"""
    <!DOCTYPE html>
    <html lang="cs">
    <head>
    <meta charset="UTF-8">
    <title>Potvrzení objednávky L-Céčko</title>
    <style>
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; max-width: 600px; margin: 0 auto; padding: 40px 20px; background-color: #f9f9f9; }}
        .uctenka {{ background: #fff; padding: 30px; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); border-top: 6px solid #4A7833; }}
        .header {{ text-align: center; border-bottom: 2px dashed #ddd; padding-bottom: 20px; margin-bottom: 20px; }}
        .header h1 {{ color: #4A7833; margin: 0 0 5px 0; font-size: 28px; }}
        .header p {{ color: #777; margin: 0; font-size: 14px; }}
        .section-title {{ font-size: 14px; color: #888; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 10px; border-bottom: 1px solid #eee; padding-bottom: 5px; }}
        .info-row {{ display: flex; justify-content: space-between; margin-bottom: 8px; font-size: 15px; }}
        .info-row strong {{ font-weight: 600; }}
        .total-box {{ background-color: #e6f0e1; padding: 15px; border-radius: 8px; margin-top: 25px; text-align: center; }}
        .total-box h2 {{ margin: 0; color: #4A7833; font-size: 24px; }}
        .payment-info {{ margin-top: 25px; font-size: 14px; background: #fff8e1; padding: 15px; border-radius: 8px; border-left: 4px solid #ffc107; }}
        .footer {{ text-align: center; margin-top: 30px; font-size: 13px; color: #999; }}
    </style>
    </head>
    <body>
        <div class="uctenka">
            <div class="header">
                <h1>🥗 L-Céčko Plzeň</h1>
                <p>Potvrzení přijetí objednávky #{id_obj}</p>
            </div>
            
            <div class="section-title">Údaje zákazníka</div>
            <div class="info-row"><span>Zákazník:</span> <strong>{jmeno} {prijmeni}</strong></div>
            <div class="info-row"><span>Adresa:</span> <strong>{adresa}</strong></div>
            <div class="info-row"><span>Telefon:</span> <strong>{telefon}</strong></div>
            <br>
            
            <div class="section-title">Souhrn objednávky</div>
            <div class="info-row"><span>Program:</span> <strong>{program}</strong></div>
            <div class="info-row"><span>Varianta:</span> <strong>{delka}</strong></div>
            <div class="info-row"><span>Start rozvozu:</span> <strong>{datum}</strong></div>
            
            <div class="total-box">
                <span style="font-size: 14px; color: #666; text-transform: uppercase;">Celkem k úhradě</span>
                <h2>{cena} Kč</h2>
            </div>
            
            <div class="payment-info">
                <strong>💳 Pokyny k platbě převodem:</strong><br><br>
                Číslo účtu: <strong>{BANK_ACCOUNT}/{BANK_CODE}</strong><br>
                IBAN: <strong>{BANK_IBAN}</strong><br>
                Variabilní symbol: <strong>{id_obj}</strong><br>
                Částka: <strong>{cena} Kč</strong>
            </div>
            
            <div class="footer">
                Děkujeme, že jste si vybrali L-Céčko!<br>Těšíme se na Vás.
            </div>
        </div>
    </body>
    </html>
    """
    return html_obsah

def spust_gastro_oslavu():
    food_emojis = ['🥗', '🍎', '🥦', '🥕', '🥑', '🥪', '🍗', '🍅', '🥒', '🍳', '🥩', '🍲']
    html_str = """
    <style>
    @keyframes fallAndSpin {
        0% { top: -10vh; transform: rotate(0deg) scale(1); opacity: 1; }
        100% { top: 110vh; transform: rotate(360deg) scale(1.2); opacity: 0; }
    }
    .food-emoji {
        position: fixed;
        z-index: 99999;
        font-size: 2.5rem;
        pointer-events: none;
        animation-name: fallAndSpin;
        animation-timing-function: linear;
        animation-fill-mode: forwards;
    }
    </style>
    """
    for _ in range(40):
        emoji = random.choice(food_emojis)
        left = random.uniform(0, 100)
        delay = random.uniform(0, 1.5)
        duration = random.uniform(2.5, 4.5)
        html_str += f'<div class="food-emoji" style="left: {left}vw; animation-duration: {duration}s; animation-delay: {delay}s;">{emoji}</div>'
        
    st.markdown(html_str, unsafe_allow_html=True)

def vytvor_whatsapp_odkaz(telefon):
    tel_cisty = "".join([c for c in str(telefon) if c.isdigit() or c == '+'])
    if not tel_cisty.startswith('+'):
        if tel_cisty.startswith('00'):
            tel_cisty = '+' + tel_cisty[2:]
        else:
            tel_cisty = '+420' + tel_cisty
    
    tel_cisty_url = tel_cisty.replace('+', '')
    zprava = "Dobrý den, tady kurýr L-Céčko. Za cca 10 minut jsem u Vás s krabičkami! 🥗"
    zprava_url = urllib.parse.quote(zprava)
    return f"https://wa.me/{tel_cisty_url}?text={zprava_url}"


df_orders, current_sha = nacti_objednavky()
nastaveni_app, sha_nastaveni = nacti_nastaveni()

if os.path.exists(LOGO_PATH):
    st.sidebar.image(LOGO_PATH, width=140)

st.sidebar.markdown("## 🥗 L-Céčko Plzeň")
rezim = st.sidebar.radio("Navigace:", ["🛒 Objednávka pro zákazníka", "🔐 Správa"])

# ---------------------------------------------------------
# 1. ZÁKAZNICKÝ OBJEDNÁVKOVÝ FORMULÁŘ
# ---------------------------------------------------------
if rezim == "🛒 Objednávka pro zákazníka":
    
    # BEZPEČNÉ VYČIŠTĚNÍ FORMULÁŘE PŘED JEHO VYKRLESLENÍM
    if st.session_state.get("need_clear_form", False):
        for k in ["input_jmeno", "input_prijmeni", "input_telefon", "input_email", "input_ulice", "input_cp", "input_psc", "input_poznamka"]:
            st.session_state[k] = ""
        st.session_state["input_mesto"] = "Plzeň"
        for k in list(st.session_state.keys()):
            if k.startswith("slane_") or k.startswith("sladke_"):
                st.session_state[k] = 0
        if "cukrovi_vaha" in st.session_state:
            st.session_state["cukrovi_vaha"] = 1.0
        st.session_state["need_clear_form"] = False

    # BĚŽÍCÍ TITULKY
    bezici_text = nastaveni_app.get("bezici_text", "")
    if bezici_text.strip():
        st.markdown(f"""
            <div class="marquee-container">
                <div class="marquee-content">{bezici_text}</div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("<h1 class='main-header'>🥗 Objednávkový formulář L-Céčko</h1>", unsafe_allow_html=True)
    st.write("Vyberte si stravovací program nebo sezónní nabídku a my se postaráme o zbytek.")
    
    # POTVRZENÍ ODESLANÉ OBJEDNÁVKY
    if "potvrzeni_objednavky" in st.session_state:
        p = st.session_state["potvrzeni_objednavky"]
        spust_gastro_oslavu()
        
        if p["email_uspech"]:
            st.success(f"🎉 Objednávka #{p['id']} byla úspěšně přijata! Děkujeme.")
        else:
            st.success(f"🎉 Objednávka #{p['id']} byla úspěšně uložena do systému! (Upozornění na e-mail se nepodařilo odeslat).")
            
        c_d1, c_d2 = st.columns([1, 1])
        with c_d1:
            st.download_button(
                label="📥 Stáhnout shrnutí objednávky", 
                data=p["html_uctenka"], 
                file_name=f"Objednavka_LCecko_{p['id']}.html", 
                mime="text/html",
                type="secondary",
                key="btn_download_success"
            )
            
        st.markdown("### 💳 Podklady pro platbu převodem")
        col_qr, col_info = st.columns([1, 1.5])
        with col_qr:
            st.image(p["qr_bytes"], caption="Naskenujte v banking aplikaci", width=200)
        with col_info:
            st.write(f"**Číslo účtu:** {BANK_ACCOUNT}/{BANK_CODE}")
            st.write(f"**IBAN:** {BANK_IBAN}")
            st.write(f"**Částka:** {p['cena']} Kč")
            st.write(f"**Variabilní symbol:** {p['id']}")
            st.write(f"**Zpráva:** L-Cecko ID {p['id']}")
            
        if st.button("❌ Zavřít toto potvrzení", key="btn_close_confirmation"):
            del st.session_state["potvrzeni_objednavky"]
            st.rerun()
            
        st.divider()

    col_menu, col_user = st.columns([1.2, 1])
    
    with col_menu:
        st.subheader("1. Výběr z nabídky")
        
        moznosti_kategorii = ["Krabičkové diety"]
        if nastaveni_app.get("zakazkova", True):
            moznosti_kategorii.append("Zakázková výroba / Catering")
        if nastaveni_app.get("snidane", True):
            moznosti_kategorii.append("Snídaňové balíčky")
        if nastaveni_app.get("cukrovi", True):
            moznosti_kategorii.append("Vánoční cukroví")
            
        kategorie = st.radio("Kategorie nabídky:", moznosti_kategorii, horizontal=True)
        
        vybrany_program = ""
        delka_trvani = ""
        cena_za_jednotku = 0
        pocet_jednotek = 1
        celkem_ks_catering = 0
        
        if kategorie == "Krabičkové diety":
            vybrany_program = st.selectbox("Výběr programu *", [
                "Krabičková strava – redukční program pro ženy (5 000–5 500 kJ)",
                "Krabičková strava – redukční program pro muže (7 000–7 500 kJ)",
                "Low Carb"
            ])
            delka_trvani = st.radio("Délka programu:", ["Týdenní program (5 dní)", "Měsíční program (20 dní)"])
            
            # --- ZDE UPRAVTE CENY PRO KRABIČKOVÉ DIETY ---
            ceny_krabicky = {
                "Krabičková strava – redukční program pro ženy (5 000–5 500 kJ)": {
                    "Týdenní program (5 dní)": 1800, # Upravit cenu zde
                    "Měsíční program (20 dní)": 7200 # Upravit cenu zde
                },
                "Krabičková strava – redukční program pro muže (7 000–7 500 kJ)": {
                    "Týdenní program (5 dní)": 2050, # Upravit cenu zde
                    "Měsíční program (20 dní)": 8200 # Upravit cenu zde
                },
                "Low Carb": {
                    "Týdenní program (5 dní)": 2200, # Upravit cenu zde
                    "Měsíční program (20 dní)": 8800 # Upravit cenu zde
                }
            }
            cena_za_jednotku = ceny_krabicky[vybrany_program][delka_trvani]

        elif kategorie == "Zakázková výroba / Catering":
            min_ks_catering = int(nastaveni_app.get("min_ks_catering", 1))
            if min_ks_catering > 1:
                st.info(f"💡 Minimální celkový odběr pro catering je **{min_ks_catering} ks/kg** (můžete libovolně kombinovat různé položky).")
            
            # --- ZDE UPRAVTE CENY PRO SLANÝ CATERING ---
            ceny_slane = {
                "Kanapky": 30, # Upravit cenu zde
                "Chlebíček": 42, # Upravit cenu zde
                "Obložené ruské vejce": 59, # Upravit cenu zde
                "Obložený máslový croissant": 69, # Upravit cenu zde
                "Plněné bagety velké": 149, # Upravit cenu zde
                "Plněná mini bageta": 45, # Upravit cenu zde
                "Topinka s vejci": 35, # Upravit cenu zde
                "Smaženka": 39, # Upravit cenu zde
                "Mini burgery": 69, # Upravit cenu zde
                "Obložený bagel": 89, # Upravit cenu zde
                "Bramborák": 59, # Upravit cenu zde
                "Racio obložený chléb": 69, # Upravit cenu zde
                "Bezlepkové chlebíčky": 69, # Upravit cenu zde
                "Obložené mísy": 890, # Upravit cenu zde
                "Mini rautové řízky (kg)": 799, # Upravit cenu zde
                "Pomazánky mix (kg)": 280, # Upravit cenu zde
                "Vajíčkový salát s rajčaty (kg)": 310, # Upravit cenu zde
                "Zeleninový salát se sýrem (kg)": 260, # Upravit cenu zde
                "Bramborový salát (kg)": 270 # Upravit cenu zde
            }
            
            # --- ZDE UPRAVTE CENY PRO SLADKÝ CATERING ---
            ceny_sladke = {
                "Bábovka celá": 280, # Upravit cenu zde
                "Sladký závin (ks)": 45, # Upravit cenu zde
                "Kynuté koláče": 40, # Upravit cenu zde
                "Strouhaný koláč": 59, # Upravit cenu zde
                "Hraběnčiny řezy": 59, # Upravit cenu zde
                "Koláč s ovocem a drobenkou": 45, # Upravit cenu zde
                "Pavlova": 79, # Upravit cenu zde
                "Palačinka": 29, # Upravit cenu zde
                "Vdolečky": 39, # Upravit cenu zde
                "Řecká svačinka s granolou a ovocem": 79, # Upravit cenu zde
                "Maromka pohár": 89, # Upravit cenu zde
                "Makový bezlepkový dort": 89, # Upravit cenu zde
                "Jahodové řezy bezlepek": 89, # Upravit cenu zde
                "Cupcake": 59, # Upravit cenu zde
                "Mini plněné košíčky": 35, # Upravit cenu zde
                "Bezlepková roláda": 79, # Upravit cenu zde
                "Čokoládový dort": 89, # Upravit cenu zde
                "Lotus pohár": 79, # Upravit cenu zde
                "Jahody se šlehačkou": 69 # Upravit cenu zde
            }
            
            vybrane_polozky = []
            celkova_cena_catering = 0
            
            tab_slane, tab_sladke = st.tabs(["🥨 Slané občerstvení", "🍰 Sladké občerstvení"])
            
            with tab_slane:
                col_s1, col_s2 = st.columns(2)
                for idx, (polozka, cena) in enumerate(ceny_slane.items()):
                    col = col_s1 if idx % 2 == 0 else col_s2
                    jednotka = "Kč/kg" if "(kg)" in polozka else "Kč/ks"
                    ks = col.number_input(f"{polozka} ({cena} {jednotka})", min_value=0, max_value=200, value=0, key=f"slane_{idx}")
                    if ks > 0:
                        vybrane_polozky.append(f"{ks}x {polozka}")
                        celkova_cena_catering += ks * cena
                        celkem_ks_catering += ks
                        
            with tab_sladke:
                col_sl1, col_sl2 = st.columns(2)
                for idx, (polozka, cena) in enumerate(ceny_sladke.items()):
                    col = col_sl1 if idx % 2 == 0 else col_sl2
                    jednotka = "Kč/kg" if "(kg)" in polozka else "Kč/ks"
                    ks = col.number_input(f"{polozka} ({cena} {jednotka})", min_value=0, max_value=200, value=0, key=f"sladke_{idx}")
                    if ks > 0:
                        vybrane_polozky.append(f"{ks}x {polozka}")
                        celkova_cena_catering += ks * cena
                        celkem_ks_catering += ks
            
            vybrany_program = ", ".join(vybrane_polozky) if vybrane_polozky else "Žádná položka nevybrána"
            delka_trvani = "Zakázková výroba"
            cena_za_jednotku = celkova_cena_catering
            
        elif kategorie == "Snídaňové balíčky":
            vybrany_program = st.selectbox("Typ snídaní:", [
                "Fit Snídaně (Kaše, Smoothie, Pečivo)",
                "Gurmán Snídaně (Vajíčka, Sýry, Uzeniny)"
            ])
            delka_trvani = st.select_slider("Délka:", options=["5 dní", "10 dní", "20 dní"])
            
            # --- ZDE UPRAVTE CENY PRO SNÍDANĚ ---
            ceník_snidani = {"5 dní": 650, "10 dní": 1200, "20 dní": 2200} # Upravit ceny zde
            cena_za_jednotku = ceník_snidani[delka_trvani]

        elif kategorie == "Vánoční cukroví":
            st.markdown("### 🎄 FIT Cukroví")
            st.write("Mix tradičních druhů ve zdravější verzi (Vanilkové rohlíčky, Mandlové rohlíčky, Pekanové pracny, Košíčky, Linecké, vosí úly...)")
            
            vaha_kg = st.number_input("Počet kil (kg):", min_value=0.5, max_value=20.0, value=1.0, step=0.5, key="cukrovi_vaha")
            
            vybrany_program = "FIT Cukroví"
            delka_trvani = f"{vaha_kg} kg"
            
            # --- ZDE UPRAVTE CENU ZA KG CUKROVÍ ---
            cena_za_jednotku = int(1400 * vaha_kg) # Upravit cenu za kg (nyní 1400) zde

        datum_od = st.date_input("Požadované datum doručení / od:", value=datetime.today() + timedelta(days=2))

    with col_user:
        st.subheader("2. Doručovací údaje & Platba")
        
        col_jmeno, col_prijmeni = st.columns(2)
        jmeno = col_jmeno.text_input("Jméno *", key="input_jmeno").strip()
        prijmeni = col_prijmeni.text_input("Příjmení *", key="input_prijmeni").strip()
        
        col_tel, col_email = st.columns(2)
        telefon = col_tel.text_input("Telefonní číslo *", key="input_telefon").strip()
        email = col_email.text_input("E-mail *", key="input_email").strip()
        
        st.markdown("<p style='font-size: 14px; font-weight: bold; margin-bottom: 0;'>Adresa pro kurýra</p>", unsafe_allow_html=True)
        
        col_ulice, col_cp = st.columns([2, 1])
        ulice = col_ulice.text_input("Ulice (případně obec) *", key="input_ulice").strip()
        cp = col_cp.text_input("Číslo popisné *", key="input_cp").strip()
        
        col_mesto, col_mesto_psc = st.columns([2, 1])
        mesto = col_mesto.text_input("Město *", value="Plzeň", key="input_mesto").strip()
        psc = col_mesto_psc.text_input("PSČ (nepovinné):", key="input_psc").strip()
        
        poznamka = st.text_input("Poznámka pro kurýra (alergie, zvonek, patro...):", key="input_poznamka").strip()
        
        st.divider()
        
        st.markdown(f"""
            <div class='summary-card'>
                <h4 style='color: #4A7833; margin: 0;'>Vaše objednávka</h4>
                <p style='margin: 5px 0 0 0; font-size: 14px;'>{vybrany_program} ({delka_trvani})</p>
                <h3 style='margin: 10px 0 0 0;'>Celková cena: {cena_za_jednotku} Kč</h3>
            </div>
        """, unsafe_allow_html=True)
        
        if st.button("Odeslat a vygenerovat QR platbu", type="primary", use_container_width=True):
            if not all([jmeno, prijmeni, telefon, email, ulice, cp, mesto]):
                st.warning("⚠️ Prosím, vyplňte všechny osobní údaje a celou adresu.")
            elif kategorie == "Zakázková výroba / Catering" and cena_za_jednotku == 0:
                st.warning("⚠️ Vyberte prosím alespoň 1 položku ze slaného nebo sladkého občerstvení.")
            elif kategorie == "Zakázková výroba / Catering" and celkem_ks_catering < int(nastaveni_app.get('min_ks_catering', 1)):
                min_k = int(nastaveni_app.get('min_ks_catering', 1))
                st.error(f"❌ Minimální celkový odběr pro catering je {min_k} ks/kg (máte vybráno {celkem_ks_catering}). Prosím, přidejte další položky.")
            else:
                with st.spinner('Odesílám objednávku do kuchyně... 👩‍🍳'):
                    psc_text = f", {psc}" if psc else ""
                    adresa_komplet = f"{ulice} {cp}, {mesto}{psc_text}"
                    
                    nove_id = 1 if df_orders.empty else int(df_orders["ID"].max()) + 1
                    nova_objednavka = pd.DataFrame([{
                        "ID": nove_id,
                        "Datum_Vytvoreni": datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "Jmeno": jmeno,
                        "Prijmeni": prijmeni,
                        "Telefon": telefon,
                        "Email": email,
                        "Adresa": adresa_komplet,
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
                        email_uspech = odeslat_email_upozorneni(nove_id, jmeno, prijmeni, adresa_komplet, vybrany_program, delka_trvani, datum_od.strftime("%Y-%m-%d"), cena_za_jednotku, poznamka)
                        
                        html_uctenka = vygeneruj_html_uctenku(
                            nove_id, jmeno, prijmeni, adresa_komplet, telefon, 
                            vybrany_program, delka_trvani, datum_od.strftime("%d.%m.%Y"), cena_za_jednotku
                        )
                        
                        spd_str = f"SPD*1.0*ACC:{BANK_IBAN}*AM:{cena_za_jednotku:.2f}*CC:CZK*X-VS:{nove_id}*MSG:L-Cecko ID {nove_id}"
                        qr_img = qrcode.make(spd_str)
                        buf = BytesIO()
                        qr_img.save(buf, format="PNG")
                        
                        st.session_state["potvrzeni_objednavky"] = {
                            "id": nove_id,
                            "email_uspech": email_uspech,
                            "html_uctenka": html_uctenka,
                            "qr_bytes": buf.getvalue(),
                            "cena": cena_za_jednotku
                        }
                        
                        # NASTAVENÍ PŘÍZNAKU PRO BEZPEČNÉ VYČIŠTĚNÍ FORMULÁŘE PŘI PŘÍŠTÍM RERUNU
                        st.session_state["need_clear_form"] = True
                            
                        st.rerun()
                    else:
                        st.error("❌ Chyba při ukládání objednávky na GitHub. Zkuste to prosím znovu.")

# ---------------------------------------------------------
# 2. SPRÁVA PRO MAJITELKU A KUCHYŇ
# ---------------------------------------------------------
