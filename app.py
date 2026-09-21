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
BANK_ACCOUNT = "123456789"  # Doplňte číslo účtu L-Céčka
BANK_CODE = "0800"       # Doplňte kód banky

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
    if not GITHUB_TOKEN or not GITHUB_REPO:
        if not os.path.exists(SETTINGS_PATH):
            with open(SETTINGS_PATH, "w") as f:
                json.dump({"snidane": True, "cukrovi": True, "zakazkova": True}, f)
        with open(SETTINGS_PATH, "r") as f:
            return json.load(f), None

    try:
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{SETTINGS_PATH}"
        res = requests.get(url, headers=get_headers(), timeout=5)
        if res.status_code == 200:
            data = res.json()
            sha = data["sha"]
            content_str = base64.b64decode(data["content"]).decode("utf-8")
            return json.loads(content_str), sha
    except Exception as e:
        print(f"Chyba při načítání nastavení z GitHubu: {e}")
    
    return {"snidane": True, "cukrovi": True, "zakazkova": True}, None

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
    st.markdown("<h1 class='main-header'>🥗 Objednávkový formulář L-Céčko</h1>", unsafe_allow_html=True)
    st.write("Vyberte si stravovací program nebo sezónní nabídku a my se postaráme o zbytek.")
    
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
        
        if kategorie == "Krabičkové diety":
            vybrany_program = st.selectbox("Výběr programu *", [
                "Krabičková strava – redukční program pro ženy (5 000–5 500 kJ)",
                "Krabičková strava – redukční program pro muže (7 000–7 500 kJ)",
                "Low Carb"
            ])
            delka_trvani = st.radio("Délka programu:", ["Týdenní program (5 dní)", "Měsíční program (20 dní)"])
            
            ceny_krabicky = {
                "Krabičková strava – redukční program pro ženy (5 000–5 500 kJ)": {
                    "Týdenní program (5 dní)": 1800,
                    "Měsíční program (20 dní)": 7200
                },
                "Krabičková strava – redukční program pro muže (7 000–7 500 kJ)": {
                    "Týdenní program (5 dní)": 2050,
                    "Měsíční program (20 dní)": 8200
                },
                "Low Carb": {
                    "Týdenní program (5 dní)": 2200,
                    "Měsíční program (20 dní)": 8800
                }
            }
            cena_za_jednotku = ceny_krabicky[vybrany_program][delka_trvani]

        elif kategorie == "Zakázková výroba / Catering":
            ceny_slane = {
                "Kanapky": 30,
                "Chlebíček": 42,
                "Obložené ruské vejce": 59,
                "Obložený máslový croissant": 69,
                "Plněné bagety velké": 149,
                "Plněná mini bageta": 45,
                "Topinka s vejci": 35,
                "Smaženka": 39,
                "Mini burgery": 69,
                "Obložený bagel": 89,
                "Bramborák": 59,
                "Racio obložený chléb": 69,
                "Bezlepkové chlebíčky": 69,
                "Obložené mísy": 890,
                "Mini rautové řízky (kg)": 799,
                "Pomazánky mix (kg)": 280,
                "Vajíčkový salát s rajčaty (kg)": 310,
                "Zeleninový salát se sýrem (kg)": 260,
                "Bramborový salát (kg)": 270
            }
            
            ceny_sladke = {
                "Bábovka celá": 280,
                "Sladký závin (ks)": 45,
                "Kynuté koláče": 40,
                "Strouhaný koláč": 59,
                "Hraběnčiny řezy": 59,
                "Koláč s ovocem a drobenkou": 45,
                "Pavlova": 79,
                "Palačinka": 29,
                "Vdolečky": 39,
                "Řecká svačinka s granolou a ovocem": 79,
                "Maromka pohár": 89,
                "Makový bezlepkový dort": 89,
                "Jahodové řezy bezlepek": 89,
                "Cupcake": 59,
                "Mini plněné košíčky": 35,
                "Bezlepková roláda": 79,
                "Čokoládový dort": 89,
                "Lotus pohár": 79,
                "Jahody se šlehačkou": 69
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
                        jednotka_str = "kg" if "(kg)" in polozka else "ks"
                        vybrane_polozky.append(f"{ks}x {polozka}")
                        celkova_cena_catering += ks * cena
                        
            with tab_sladke:
                col_sl1, col_sl2 = st.columns(2)
                for idx, (polozka, cena) in enumerate(ceny_sladke.items()):
                    col = col_sl1 if idx % 2 == 0 else col_sl2
                    jednotka = "Kč/kg" if "(kg)" in polozka else "Kč/ks"
                    ks = col.number_input(f"{polozka} ({cena} {jednotka})", min_value=0, max_value=200, value=0, key=f"sladke_{idx}")
                    if ks > 0:
                        vybrane_polozky.append(f"{ks}x {polozka}")
                        celkova_cena_catering += ks * cena
            
            vybrany_program = ", ".join(vybrane_polozky) if vybrane_polozky else "Žádná položka nevybrána"
            delka_trvani = "Zakázková výroba"
            cena_za_jednotku = celkova_cena_catering
            
        elif kategorie == "Snídaňové balíčky":
            vybrany_program = st.selectbox("Typ snídaní:", [
                "Fit Snídaně (Kaše, Smoothie, Pečivo)",
                "Gurmán Snídaně (Vajíčka, Sýry, Uzeniny)"
            ])
            delka_trvani = st.select_slider("Délka:", options=["5 dní", "10 dní", "20 dní"])
            ceník_snidani = {"5 dní": 650, "10 dní": 1200, "20 dní": 2200}
            cena_za_jednotku = ceník_snidani[delka_trvani]

        elif kategorie == "Vánoční cukroví":
            vybrany_program = "FIT Cukroví (1 kg)"
            
            st.markdown("### 🎄 FIT Cukroví (1 kg)")
            st.write("Mix tradičních druhů ve zdravější verzi (Vanilkové rohlíčky, Mandlové rohlíčky, Pekanové pracny, Košíčky, Linecké, vosí úly...)")
            st.info("💡 Pokud máte zájem o bezlepkovou variantu (1 890 Kč / kg), napište nám to prosím do 'Poznámky pro kurýra' v dalším kroku objednávky.")
            
            pocet_jednotek = st.number_input("Počet balení (ks):", min_value=1, max_value=20, value=1)
            
            cena_za_jednotku = 1400 * pocet_jednotek
            delka_trvani = "Vánoční cukroví"

        datum_od = st.date_input("Požadované datum doručení / od:", value=datetime.today() + timedelta(days=2))

    with col_user:
        st.subheader("2. Doručovací údaje & Platba")
        
        col_jmeno, col_prijmeni = st.columns(2)
        jmeno = col_jmeno.text_input("Jméno:").strip()
        prijmeni = col_prijmeni.text_input("Příjmení:").strip()
        
        col_tel, col_email = st.columns(2)
        telefon = col_tel.text_input("Telefonní číslo:").strip()
        email = col_email.text_input("E-mail:").strip()
        
        st.markdown("<p style='font-size: 14px; font-weight: bold; margin-bottom: 0;'>Adresa pro kurýra</p>", unsafe_allow_html=True)
        
        col_ulice, col_cp = st.columns([2, 1])
        ulice = col_ulice.text_input("Ulice (případně obec):").strip()
        cp = col_cp.text_input("Číslo popisné:").strip()
        
        col_mesto, col_psc = st.columns([2, 1])
        mesto = col_mesto.text_input("Město:", value="Plzeň").strip()
        psc = col_psc.text_input("PSČ (nepovinné):").strip()
        
        poznamka = st.text_input("Poznámka pro kurýra (alergie, zvonek, patro...):").strip()
        
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
                        # Zavolání funkce pro odeslání e-mailu
                        email_uspech = odeslat_email_upozorneni(nove_id, jmeno, prijmeni, adresa_komplet, vybrany_program, delka_trvani, datum_od.strftime("%Y-%m-%d"), cena_za_jednotku, poznamka)
                        
                        spust_gastro_oslavu()
                        
                        if email_uspech:
                            st.success("🎉 Objednávka byla úspěšně přijata! Děkujeme.")
                        else:
                            st.success("🎉 Objednávka byla úspěšně uložena do systému! (Upozornění na e-mail se bohužel nepodařilo odeslat, zkontrolujte nastavení SMTP serveru).")
                        
                        html_uctenka = vygeneruj_html_uctenku(
                            nove_id, jmeno, prijmeni, adresa_komplet, telefon, 
                            vybrany_program, delka_trvani, datum_od.strftime("%d.%m.%Y"), cena_za_jednotku
                        )
                        st.download_button(
                            label="📥 Stáhnout shrnutí objednávky", 
                            data=html_uctenka, 
                            file_name=f"Objednavka_LCecko_{nove_id}.html", 
                            mime="text/html",
                            type="secondary"
                        )
                        
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
                        st.error("❌ Chyba při ukládání objednávky na GitHub. Zkuste to prosím znovu.")

# ---------------------------------------------------------
# 2. SPRÁVA PRO MAJITELKU A KUCHYŇ
# ---------------------------------------------------------
else:
    st.markdown("<h1 class='main-header'>🔐 Interní dispečink L-Céčka</h1>", unsafe_allow_html=True)
    
    url_heslo = st.query_params.get("heslo", "")
    
    if url_heslo == ADMIN_PASSWORD:
        heslo = url_heslo
    else:
        heslo = st.text_input("Zadejte přístupové heslo:", type="password")
    
    if heslo == ADMIN_PASSWORD:
        st.success("✅ Přístup schválen.")
        
        if not df_orders.empty:
            st.markdown("### 📈 Finanční přehled")
            
            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
            
            obrat_zaplaceno = df_orders[df_orders['Stav_Platby'] == 'Zaplaceno']['Cena_Celkem'].sum()
            obrat_ceka = df_orders[df_orders['Stav_Platby'] == 'Čeká na platbu']['Cena_Celkem'].sum()
            pocet_ceka = len(df_orders[df_orders['Stav_Platby'] == 'Čeká na platbu'])
            pocet_celkem = len(df_orders)
            
            obrat_zaplaceno_str = f"{obrat_zaplaceno:,.0f} Kč".replace(',', ' ')
            obrat_ceka_str = f"{obrat_ceka:,.0f} Kč".replace(',', ' ')
            
            col_m1.metric("💰 Zaplaceno (Obrat)", obrat_zaplaceno_str)
            col_m2.metric("⌛ Očekávané platby", obrat_ceka_str)
            col_m3.metric("⚠️ Nezkontrolované obj.", f"{pocet_ceka} ks")
            col_m4.metric("📦 Celkem objednávek", f"{pocet_celkem} ks")
            st.divider()
        
        tab1, tab2, tab3, tab4 = st.tabs(["📊 Správa databáze", "👨‍🍳 Výkaz pro Kuchyň", "🚗 Seznam pro Kurýra", "⚙️ Nastavení nabídky"])
        
        with tab1:
            st.subheader("Kompletní databáze objednávek")
            if not df_orders.empty:
                st.markdown("<div class='filter-box'>", unsafe_allow_html=True)
                st.markdown("**🔍 Filtrovat objednávky:**")
                col_f1, col_f2, col_f3 = st.columns(3)
                
                search_query = col_f1.text_input("Hledat (Jméno, Adresa, ID...)", "").lower()
                
                dostupne_stavy = list(df_orders['Stav_Platby'].unique())
                vybrane_stavy = col_f2.multiselect("Stav platby:", dostupne_stavy, default=dostupne_stavy)
                
                dostupna_data = sorted(list(df_orders['Datum_Od'].unique()))
                vybrane_datum = col_f3.selectbox("Datum doručení (Od):", ["Všechna data"] + dostupna_data)
                st.markdown("</div>", unsafe_allow_html=True)

                df_filtered = df_orders.copy()
                
                df_filtered.insert(1, "⚠️ POZOR", df_filtered["Poznamka_Kuryr"].apply(detekuj_alergie))
                
                if search_query:
                    mask = df_filtered.apply(lambda row: row.astype(str).str.lower().str.contains(search_query).any(), axis=1)
                    df_filtered = df_filtered[mask]
                
                if vybrane_stavy:
                    df_filtered = df_filtered[df_filtered['Stav_Platby'].isin(vybrane_stavy)]
                    
                if vybrane_datum != "Všechna data":
                    df_filtered = df_filtered[df_filtered['Datum_Od'] == vybrane_datum]

                st.info(f"💡 Zobrazeno {len(df_filtered)} objednávek. Pro smazání zaškrtněte políčko **Smazat 🗑️** a uložte změny.")
                
                if "Smazat 🗑️" not in df_filtered.columns:
                    df_filtered.insert(0, "Smazat 🗑️", False)

                edited_df = st.data_editor(
                    df_filtered,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Smazat 🗑️": st.column_config.CheckboxColumn(
                            "Smazat 🗑️",
                            help="Zaškrtněte pro trvalé smazání objednávky",
                            default=False,
                        ),
                        "Stav_Platby": st.column_config.SelectboxColumn(
                            "Stav Platby",
                            options=["Čeká na platbu", "Zaplaceno", "Stornováno"],
                            required=True,
                        ),
                        "⚠️ POZOR": st.column_config.TextColumn("⚠️ POZOR", disabled=True)
                    }
                )
                
                if st.button("💾 Uložit změny", type="primary"):
                    radky_ke_smazani = edited_df[edited_df["Smazat 🗑️"] == True]["ID"].tolist()
                    df_aktualni = df_orders[~df_orders["ID"].isin(radky_ke_smazani)].copy()
                    
                    for index, row in edited_df[edited_df["Smazat 🗑️"] == False].iterrows():
                        df_aktualni.loc[df_aktualni["ID"] == row["ID"], "Stav_Platby"] = row["Stav_Platby"]
                    
                    if uloz_objednavky(df_aktualni, current_sha):
                        st.success("✅ Změny byly uloženy!")
                        st.rerun()
                    else:
                        st.error("❌ Chyba při ukládání změn do databáze.")
                
                st.divider()
                
                export_df = edited_df.drop(columns=["Smazat 🗑️"], errors="ignore")
                excel_data = vytvor_profi_excel(export_df, titulek="Vyfiltrované objednávky")
                st.download_button(
                    label="📥 Stáhnout aktuální zobrazení do Excelu", 
                    data=excel_data, 
                    file_name=f"lcecko_export_{datetime.now().strftime('%d_%m')}.xlsx", 
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("Zatím žádné objednávky v databázi.")
                
        with tab2:
            st.subheader("👨‍🍳 Souhrn porcí pro kuchyň")
            if not df_orders.empty:
                vybrany_den = st.date_input("Zobrazit souhrn k datu:", value=datetime.today())
                
                df_kuchyn = df_orders[df_orders["Datum_Od"] == vybrany_den.strftime("%Y-%m-%d")].copy()
                
                st.write(f"**Přehled menu a výroby pro:** {vybrany_den.strftime('%d.%m.%Y')}")
                
                if not df_kuchyn.empty:
                    souhrn = df_kuchyn["Program"].value_counts().reset_index()
                    souhrn.columns = ["Stravovací Program / Produkty", "Počet / Ks"]
                    st.table(souhrn)
                    
                    st.markdown("#### ⚠️ Zvláštní požadavky a alergie na tento den:")
                    df_kuchyn["Alergie"] = df_kuchyn["Poznamka_Kuryr"].apply(detekuj_alergie)
                    alergici = df_kuchyn[df_kuchyn["Alergie"] != ""]
                    if not alergici.empty:
                        st.dataframe(alergici[["Jmeno", "Prijmeni", "Program", "Poznamka_Kuryr"]], hide_index=True)
                    else:
                        st.success("Dnes nejsou hlášeny žádné speciální požadavky ani alergie.")
                    
                    st.divider()
                    
                    excel_kuchyn = vytvor_profi_excel(souhrn, titulek="Vyroba_Kuchyn")
                    st.download_button(
                        label="👨‍🍳 Stáhnout běžný výkaz (Excel)", 
                        data=excel_kuchyn, 
                        file_name=f"kuchyn_vyroba_{vybrany_den.strftime('%d_%m')}.xlsx", 
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:
                    st.info(f"Na den {vybrany_den.strftime('%d.%m.%Y')} nejsou naplánované žádné objednávky.")
            else:
                st.info("Zatím žádné objednávky v databázi.")
                
        with tab3:
            st.subheader("🚗 Rozvozový arch pro kurýra")
            
            with st.expander("📱 Přímý odkaz pro kurýra (Bez zadávání hesla)"):
                st.write("""
                **Jak to funguje:**
                Pošlete kurýrovi speciální odkaz níže. Když si ho uloží do telefonu na plochu, aplikace ho **vždy přihlásí automaticky** bez zadávání hesla!
                """)
                base_app_url = st.text_input("Vložte sem základní adresu aplikace (např. https://lcecko.streamlit.app):")
                if base_app_url:
                    cisty_odkaz = base_app_url.rstrip('/')
                    tajny_odkaz = f"{cisty_odkaz}/?heslo={ADMIN_PASSWORD}"
                    st.code(tajny_odkaz, language="text")
                    
                    qr_app = qrcode.make(tajny_odkaz)
                    buf_app = BytesIO()
                    qr_app.save(buf_app, format="PNG")
                    st.image(buf_app.getvalue(), width=180, caption="Naskenovat v telefonu pro přímý přístup")

            st.divider()

            if not df_orders.empty:
                vybrany_den_kuryr = st.date_input("Rozvozy na den:", value=datetime.today(), key="kuryr_den")
                df_kuryr = df_orders[df_orders["Datum_Od"] == vybrany_den_kuryr.strftime("%Y-%m-%d")].copy()
                
                if not df_kuryr.empty:
                    adresy = df_kuryr["Adresa"].dropna().tolist()
                    if adresy:
                        adresy_mapy = adresy[:10] if len(adresy) > 10 else adresy
                        lomitka_adresy = "/".join([urllib.parse.quote(a) for a in adresy_mapy])
                        google_maps_url = f"https://www.google.com/maps/dir/{lomitka_adresy}"
                        
                        st.markdown(f"""
                            <a href="{google_maps_url}" target="_blank" style="background-color: #4285F4; color: white; padding: 10px 20px; border-radius: 8px; text-decoration: none; font-weight: bold; display: inline-block; margin-bottom: 20px; border: 1px solid #357ae8;">
                                📍 Naplánovat trasu na dnešek v Google Mapách
                            </a>
                        """, unsafe_allow_html=True)
                        if len(adresy) > 10:
                            st.caption("⚠️ Zobrazená trasa na mapě obsahuje prvních 10 zastávek (limit prohlížeče).")

                    rozvoz_df = df_kuryr[["Jmeno", "Prijmeni", "Telefon", "Adresa", "Poznamka_Kuryr", "Program", "Stav_Platby"]].copy()
                    rozvoz_df["Napsat zákazníkovi 💬"] = rozvoz_df["Telefon"].apply(vytvor_whatsapp_odkaz)
                    
                    st.data_editor(
                        rozvoz_df,
                        use_container_width=True,
                        hide_index=True,
                        disabled=True,
                        column_config={
                            "Napsat zákazníkovi 💬": st.column_config.LinkColumn(
                                "WhatsApp",
                                display_text="Otevřít chat 💬"
                            )
                        }
                    )
                    
                    excel_kuryr = vytvor_profi_excel(rozvoz_df.drop(columns=["Napsat zákazníkovi 💬"]), titulek="Rozvozy")
                    st.download_button(
                        label="🚗 Stáhnout arch pro kurýra (Excel)", 
                        data=excel_kuryr, 
                        file_name=f"kuryr_rozvozy_{vybrany_den_kuryr.strftime('%d_%m')}.xlsx", 
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.info(f"Na den {vybrany_den_kuryr.strftime('%d.%m.%Y')} nejsou žádné rozvozy.")
            else:
                st.info("Zatím žádné adresy k rozvozu.")
                
        with tab4:
            st.subheader("⚙️ Zapínání a vypínání částí nabídky")
            st.write("Pomocí těchto přepínačů můžete jednoduše skrýt určité sekce z webu pro zákazníky, když je zrovna nenabízíte.")
            
            zobrazovat_zakazkova = st.checkbox("Zobrazovat sekci 'Zakázková výroba / Catering' pro zákazníky", value=nastaveni_app.get("zakazkova", True))
            zobrazovat_snidane = st.checkbox("Zobrazovat sekci 'Snídaňové balíčky' pro zákazníky", value=nastaveni_app.get("snidane", True))
            zobrazovat_cukrovi = st.checkbox("Zobrazovat sekci 'Vánoční cukroví' pro zákazníky", value=nastaveni_app.get("cukrovi", True))
            
            if st.button("Uložit nastavení webu", type="primary"):
                nove_nastaveni = {
                    "zakazkova": zobrazovat_zakazkova,
                    "snidane": zobrazovat_snidane,
                    "cukrovi": zobrazovat_cukrovi
                }
                if uloz_nastaveni(nove_nastaveni, sha_nastaveni):
                    st.success("✅ Nastavení nabídky bylo úspěšně uloženo!")
                    st.rerun()
                else:
                    st.error("❌ Při ukládání došlo k chybě.")
    elif heslo:
        st.error("❌ Nesprávné heslo.")

# ---------------------------------------------------------
# PATIČKA APLIKACE
# ---------------------------------------------------------
st.markdown(
    """
    <hr style="margin-top: 50px; margin-bottom: 10px; border: 0; border-top: 1px solid #d3d3d3;">
    <div style="text-align: center; color: #666666; font-size: 0.85rem; padding-bottom: 20px;">
        Aplikaci vytvořil Pavel Kouba®
    </div>
    """,
    unsafe_allow_html=True
)
