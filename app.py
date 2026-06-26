"""
=============================================================================
 PRIVATE WEALTH ANALYTICS - PORTFOLIO IMPLEMENTATION ENGINE
=============================================================================
 Live drift monitoring and trade-order generation.
 Clean fintech SaaS surface: teal-emerald accent, quiet slate text,
 pill status tags, SVG iconography (no emoji), custom HTML result tables.
=============================================================================
"""

import os
import sys
import shutil
import subprocess
from datetime import datetime

import pandas as pd
import streamlit as st

# yfinance + plotly are required. If missing, offer an in-app install path.
try:
    import yfinance as yf
    import plotly.express as px
except ModuleNotFoundError as e:
    st.set_page_config(page_title="Portfolio Implementation Engine", layout="wide")
    st.error(f"Missing required library: {e.name}")
    st.write(
        "If the terminal install isn't working, install the libraries directly "
        "from here, then refresh the page."
    )
    if st.button("Install required libraries"):
        with st.spinner("Installing… this can take up to a minute."):
            try:
                subprocess.check_call(
                    [sys.executable, "-m", "pip", "install", "plotly", "yfinance"]
                )
                st.success("Installed. Refresh the page (press R) to continue.")
            except Exception as err:
                st.error(f"Installation failed: {err}")
    st.stop()


# --------------------------------------------------------------------------- #
# CONFIGURATION
# --------------------------------------------------------------------------- #
SAVED_MANDATES_DIR = "saved_mandates"
PRICE_CACHE_TTL = 300

os.makedirs(SAVED_MANDATES_DIR, exist_ok=True)

# --- CLOUD AUTO-LOADER ---
# Automatically pulls any mandate files from the root GitHub folder into the app's memory
for f in os.listdir("."):
    if f.lower().endswith((".csv", ".xlsx", ".xls")) and "mandate" in f.lower():
        dest = os.path.join(SAVED_MANDATES_DIR, f)
        if not os.path.exists(dest):
            shutil.copy(f, dest)

st.set_page_config(
    page_title="Portfolio Implementation Engine",
    layout="wide",
    initial_sidebar_state="expanded",
)


# --------------------------------------------------------------------------- #
# DESIGN TOKENS
# --------------------------------------------------------------------------- #
BRAND = "#0EA37F"
BRAND_HOVER = "#0C8A6B"
INK = "#0F1B2D"
MUTED = "#64748B"
BORDER = "#E8EDF2"
SURFACE = "#FFFFFF"
CANVAS = "#F7F9FB"
POS = "#0B8A5B"
NEG = "#C8363C"
CASH_COLOR = "#94A3B8"
UNCLASS_COLOR = "#CBD5E1"

CATEGORY_PALETTE = [
    "#0EA37F", "#3B82F6", "#F59E0B", "#8B5CF6", "#EC4899",
    "#06B6D4", "#84CC16", "#F43F5E", "#A855F7", "#0D9488",
]

# Native-market currency symbols for the per-security price column.
CCY_SYMBOLS = {"USD": "US$", "CAD": "C$", "EUR": "€", "GBP": "£", "JPY": "¥", "CHF": "CHF "}


# --------------------------------------------------------------------------- #
# SVG ICONS (inline, stroke = currentColor) — no emoji
# --------------------------------------------------------------------------- #
def _svg(body, size=18):
    return (
        f'<svg viewBox="0 0 24 24" width="{size}" height="{size}" fill="none" '
        f'stroke="currentColor" stroke-width="1.75" stroke-linecap="round" '
        f'stroke-linejoin="round">{body}</svg>'
    )

ICONS = {
    # Header brand: sliders / equalizer (tuning allocations)
    "sliders": _svg('<line x1="4" y1="21" x2="4" y2="13"/><line x1="4" y1="9" x2="4" y2="3"/>'
                    '<line x1="12" y1="21" x2="12" y2="12"/><line x1="12" y1="8" x2="12" y2="3"/>'
                    '<line x1="20" y1="21" x2="20" y2="15"/><line x1="20" y1="11" x2="20" y2="3"/>'
                    '<line x1="2" y1="11" x2="6" y2="11"/><line x1="10" y1="10" x2="14" y2="10"/>'
                    '<line x1="18" y1="13" x2="22" y2="13"/>'),
    # Current allocation: donut ring
    "donut": _svg('<circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="3.4"/>'),
    # Target allocation: crosshair / bullseye (distinct from donut)
    "target": _svg('<circle cx="12" cy="12" r="8.5"/><circle cx="12" cy="12" r="3.3"/>'
                   '<line x1="12" y1="1.5" x2="12" y2="5"/><line x1="12" y1="19" x2="12" y2="22.5"/>'
                   '<line x1="1.5" y1="12" x2="5" y2="12"/><line x1="19" y1="12" x2="22.5" y2="12"/>'),
    "wallet": _svg('<rect x="3" y="6" width="18" height="13" rx="2.5"/>'
                   '<path d="M3 10h18"/><circle cx="16.5" cy="14" r="1.1"/>'),
    "flow": _svg('<polyline points="7 8 3 12 7 16"/><line x1="3" y1="12" x2="21" y2="12"/>'
                 '<polyline points="17 8 21 12 17 16"/>'),
    "coins": _svg('<circle cx="8.5" cy="8.5" r="4.8"/><circle cx="15.5" cy="15.5" r="4.8"/>'),
    "banknote": _svg('<rect x="2.5" y="6" width="19" height="12" rx="2.5"/>'
                     '<circle cx="12" cy="12" r="2.4"/><path d="M6 12h.01M18 12h.01"/>'),
    "list": _svg('<line x1="8.5" y1="7" x2="20" y2="7"/><line x1="8.5" y1="12" x2="20" y2="12"/>'
                 '<line x1="8.5" y1="17" x2="20" y2="17"/><circle cx="4.5" cy="7" r="1.1"/>'
                 '<circle cx="4.5" cy="12" r="1.1"/><circle cx="4.5" cy="17" r="1.1"/>'),
    "exchange": _svg('<polyline points="7 5 7 19"/><polyline points="4 16 7 19 10 16"/>'
                     '<polyline points="17 19 17 5"/><polyline points="14 8 17 5 20 8"/>'),
    "upload": _svg('<path d="M12 15V4"/><polyline points="8 8 12 4 16 8"/>'
                   '<path d="M4 15v3a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-3"/>', size=26),
}


# --------------------------------------------------------------------------- #
# BILINGUAL DICTIONARY
# --------------------------------------------------------------------------- #
LANG = {
    "en": {
        "header_title": "Portfolio Implementation Engine",
        "header_desc": "Automated drift monitoring & trade-order generation",
        "sidebar_config": "System configuration",
        "mandate_spec": "1. Mandate specification",
        "upload_mandate": "Upload IPS mandate (CSV or Excel)",
        "active_db": "Active database file",
        "target_mandate": "Target mandate",
        "tactical_cash": "2. Tactical cash flow",
        "deposit_withdraw": "Deposit (+) / Withdrawal (−)",
        "main_title": "Implementation & order generation",
        "main_tracking": "Currently tracking",
        "client_holdings": "Client holdings",
        "upload_holdings": "Upload client holdings (CSV or Excel)",
        "empty_title": "Upload client holdings to begin",
        "waiting_holdings": "Add a file with Ticker and Quantity columns to run the live analysis.",
        "curr_val": "Current portfolio value",
        "cash_flow": "Tactical cash flow",
        "cash_equiv": "Cash & equivalents",
        "post_cash": "Est. residual cash",
        "alloc_current": "Current allocation",
        "alloc_target": "Target allocation",
        "legend_title": "Allocation breakdown & legend",
        "analysis_title": "Holdings & drift analysis",
        "orders_title": "Execution orders",
        "status_ok": "On target",
        "status_breach": "Action required",
        "no_trades": "Portfolio is within tolerance. No trades required.",
        "export_btn": "Export trade ticket (CSV)",
        "footer": "Realized by Samari Gbian",
        "fetching_prices": "Fetching live market data and FX rates…",
        "view_mandate": "View mandate targets",
        "howto_title": "How to use this engine",
        "col_ticker": "Ticker",
        "col_asset": "Asset class",
        "col_price": "Price",
        "col_current": "Current",
        "col_target": "Target",
        "col_drift": "Drift",
        "col_status": "Status",
        "col_action": "Action",
        "col_shares": "Shares",
        "col_est_value": "Est. value",
        "no_priced": "No live prices could be retrieved for any holding. Check your tickers and try again.",
        "missing_note": "No live price for: {items}. These positions are excluded, so weights reflect only the priced holdings.",
        "offmandate_note": "Held but not in this mandate (0% target, flagged for full sell): {items}.",
        "targetsum_note": "Target weights for this mandate sum to {pct:.1f}%, not 100%.",
        "howto_text": """
**1. Configure (sidebar).** Choose the interface **language** and the **base currency**. The base currency is used for portfolio-level figures only — portfolio value, cash flow, cash, and residual. Individual security prices are always shown in the currency of the market where each one trades.

**2. Upload an IPS mandate.** A CSV or Excel file describing one or more target models. The built-in scanner recognizes common column names and maps them automatically. It looks for:
- **Mandate_Name** — the model's name (one file can hold several mandates)
- **Ticker** — the security symbol (e.g. `AAPL`, `RY.TO`, `CASH_USD`)
- **Asset_Class** — e.g. US Equity, Fixed Income, Cash
- **Target_Weight_Percent** — the target weight. Either percent (`20`) or a fraction (`0.20`) — both are detected and displayed as %.
- **Drift_Band_Percent** — the tolerance in percentage points (`3` means a position can drift ±3 pts before a trade triggers)

Uploaded mandates are saved to a local `saved_mandates/` folder and remembered next time.

**3. Select the database file and the target mandate** from the dropdowns. The "View mandate targets" panel shows the active model's weights.

**4. (Optional) Tactical cash flow.** Enter a deposit (+) or withdrawal (−) in the base currency. It is added to the pool *before* target dollars are sized, so fresh cash is deployed across the model.

**5. Upload client holdings.** A CSV or Excel with **Ticker** and **Quantity**. Represent cash as `CASH_USD` (or `CASH_CAD`), using the dollar amount as the quantity.

---

**Reading the results**
- **Metric cards** — portfolio value, your cash flow, current cash weight, and the estimated residual cash *after* the proposed trades (all in base currency).
- **Donut charts + legend** — current vs. target allocation by asset class. Slices too small to label are summarized, with exact figures, in the **Allocation breakdown & legend** dropdown.
- **Holdings & drift analysis** — live price (native market currency), current vs. target weight, and **drift in percentage points**. A red **Action required** tag means the position has moved outside its drift band.
- **Execution orders** — the whole-share **BUY / SELL** orders needed to return to target, with an estimated value in base currency. Export the ticket as CSV.

**How the numbers work**
- **Live pricing** tries the fast quote first, then the standard price fields, then the most recent close.
- **FX** — non-base prices (e.g. Canadian `.TO` listings vs. US listings) are converted into the base currency *only* for portfolio totals and weights; the price column stays native.
- **Band-aware rebalancing** — only positions outside their drift band generate trades, so the book isn't churned for tiny deviations. A position in the mandate but not yet held is bought from zero; a position held but absent from the mandate is flagged for a full sell.
- **Residual cash** is mechanical: starting cash + your cash flow − the net cost of the generated trades. With balanced buy/sell targets it lands near the model's cash target.

*No management-fee logic is applied anywhere in this engine.*
""",
    },
    "fr": {
        "header_title": "Moteur d'implémentation de portefeuille",
        "header_desc": "Surveillance de la dérive et génération d'ordres",
        "sidebar_config": "Configuration du système",
        "mandate_spec": "1. Spécification du mandat",
        "upload_mandate": "Téléverser le mandat (CSV ou Excel)",
        "active_db": "Fichier de base de données actif",
        "target_mandate": "Mandat cible",
        "tactical_cash": "2. Flux de trésorerie tactique",
        "deposit_withdraw": "Dépôt (+) / Retrait (−)",
        "main_title": "Implémentation et génération d'ordres",
        "main_tracking": "Suivi actuel",
        "client_holdings": "Positions du client",
        "upload_holdings": "Téléverser les positions (CSV ou Excel)",
        "empty_title": "Téléversez les positions pour commencer",
        "waiting_holdings": "Ajoutez un fichier avec les colonnes Ticker et Quantité pour lancer l'analyse.",
        "curr_val": "Valeur actuelle",
        "cash_flow": "Flux de trésorerie",
        "cash_equiv": "Trésorerie et équivalents",
        "post_cash": "Encaisse résiduelle est.",
        "alloc_current": "Allocation actuelle",
        "alloc_target": "Allocation cible",
        "legend_title": "Détail de la répartition et légende",
        "analysis_title": "Analyse des positions et de la dérive",
        "orders_title": "Ordres d'exécution",
        "status_ok": "Cible atteinte",
        "status_breach": "Action requise",
        "no_trades": "Le portefeuille est dans les marges. Aucune transaction requise.",
        "export_btn": "Exporter le bordereau (CSV)",
        "footer": "Réalisé par Samari Gbian",
        "fetching_prices": "Récupération des données de marché et des taux de change…",
        "view_mandate": "Voir les cibles du mandat",
        "howto_title": "Comment utiliser ce moteur",
        "col_ticker": "Symbole",
        "col_asset": "Classe d'actif",
        "col_price": "Prix",
        "col_current": "Actuel",
        "col_target": "Cible",
        "col_drift": "Dérive",
        "col_status": "Statut",
        "col_action": "Ordre",
        "col_shares": "Actions",
        "col_est_value": "Valeur est.",
        "no_priced": "Aucun prix en direct n'a pu être récupéré. Vérifiez vos symboles et réessayez.",
        "missing_note": "Aucun prix pour : {items}. Ces positions sont exclues; les pondérations ne reflètent que les titres évalués.",
        "offmandate_note": "Détenu mais absent de ce mandat (cible 0 %, marqué pour vente totale) : {items}.",
        "targetsum_note": "Les pondérations cibles de ce mandat totalisent {pct:.1f} %, et non 100 %.",
        "howto_text": """
**1. Configuration (barre latérale).** Choisissez la **langue** et la **devise de base**. La devise de base ne s'applique qu'aux valeurs globales du portefeuille — valeur, flux de trésorerie, encaisse et solde résiduel. Le prix de chaque titre est toujours affiché dans la devise du marché où il se négocie.

**2. Téléversez un mandat (IPS).** Un fichier CSV ou Excel décrivant un ou plusieurs modèles cibles. Le lecteur intégré reconnaît les noms de colonnes courants et les associe automatiquement. Il recherche :
- **Mandate_Name** — le nom du modèle (un fichier peut contenir plusieurs mandats)
- **Ticker** — le symbole du titre (ex. `AAPL`, `RY.TO`, `CASH_USD`)
- **Asset_Class** — ex. Actions US, Titres à revenu fixe, Trésorerie
- **Target_Weight_Percent** — la pondération cible. En pourcentage (`20`) ou en fraction (`0.20`) — les deux sont détectés et affichés en %.
- **Drift_Band_Percent** — la marge en points de pourcentage (`3` = dérive de ±3 pts avant déclenchement)

Les mandats téléversés sont sauvegardés dans un dossier local `saved_mandates/` et mémorisés.

**3. Sélectionnez le fichier et le mandat cible** dans les menus. Le panneau « Voir les cibles du mandat » affiche les pondérations du modèle actif.

**4. (Optionnel) Flux de trésorerie tactique.** Saisissez un dépôt (+) ou un retrait (−) dans la devise de base. Il est ajouté au total *avant* le calcul des montants cibles, afin de déployer les liquidités dans le modèle.

**5. Téléversez les positions du client.** Un CSV ou Excel avec **Ticker** et **Quantité**. Représentez la trésorerie par `CASH_USD` (ou `CASH_CAD`), le montant servant de quantité.

---

**Lecture des résultats**
- **Cartes d'indicateurs** — valeur du portefeuille, flux de trésorerie, pondération actuelle de la trésorerie, et encaisse résiduelle estimée *après* les transactions (en devise de base).
- **Graphiques en anneau + légende** — répartition actuelle vs cible par classe d'actif. Les petites tranches sont détaillées, avec leurs valeurs exactes, dans le menu **Détail de la répartition et légende**.
- **Analyse des positions et de la dérive** — prix en direct (devise native), pondération actuelle vs cible, et **dérive en points de pourcentage**. Une étiquette rouge **Action requise** signale une position hors de sa marge.
- **Ordres d'exécution** — les ordres d'**ACHAT / VENTE** en actions entières pour revenir à la cible, avec une valeur estimée en devise de base. Exportez le bordereau en CSV.

**Fonctionnement des calculs**
- **Prix en direct** : cotation rapide d'abord, puis les champs de prix standards, puis la dernière clôture.
- **Change** : les prix hors devise de base (ex. titres canadiens `.TO` vs américains) sont convertis dans la devise de base *uniquement* pour les totaux et pondérations; la colonne des prix reste native.
- **Rééquilibrage selon la marge** : seules les positions hors marge génèrent des ordres. Une position du mandat non encore détenue est achetée à partir de zéro; une position détenue mais absente du mandat est marquée pour vente totale.
- **Encaisse résiduelle** : mécanique — trésorerie de départ + flux de trésorerie − coût net des transactions générées.

*Aucune logique de frais de gestion n'est appliquée dans ce moteur.*
""",
    },
}


# --------------------------------------------------------------------------- #
# CSS — fintech SaaS surface
# --------------------------------------------------------------------------- #
custom_css = f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    #MainMenu {{visibility: hidden;}}
    footer {{visibility: hidden;}}
    [data-testid="stHeader"] {{background: transparent !important;}}
    .block-container {{ padding-top: 2rem !important; padding-bottom: 2rem !important; max-width: 1240px; }}

    html, body, [class*="css"] {{
        font-family: 'Inter', system-ui, sans-serif !important;
        color: {INK};
    }}
    .stApp {{ background: {CANVAS}; }}
    h1, h2, h3, h4 {{ font-weight: 700 !important; letter-spacing: -0.02em !important; color: {INK} !important; }}

    [data-testid="stSidebar"] {{ background: {SURFACE} !important; border-right: 1px solid {BORDER} !important; }}
    [data-testid="stSidebar"] .stSelectbox label, [data-testid="stSidebar"] .stNumberInput label {{
        font-size: 0.8rem !important; color: {MUTED} !important; font-weight: 500 !important;
    }}

    /* Form controls (sidebar + main) styled to match the card surface */
    div[data-baseweb="select"] > div {{
        background: {SURFACE} !important; border: 1px solid {BORDER} !important;
        border-radius: 8px !important; min-height: 40px !important;
    }}
    div[data-baseweb="select"] > div:hover {{ border-color: #CBD5E1 !important; }}
    div[data-baseweb="select"] > div:focus-within {{
        border-color: {BRAND} !important; box-shadow: 0 0 0 3px rgba(14,163,127,0.12) !important;
    }}
    div[data-baseweb="input"], div[data-baseweb="base-input"] {{
        background: {SURFACE} !important; border: 1px solid {BORDER} !important; border-radius: 8px !important;
    }}
    div[data-baseweb="input"]:focus-within {{
        border-color: {BRAND} !important; box-shadow: 0 0 0 3px rgba(14,163,127,0.12) !important;
    }}
    [data-testid="stNumberInput"] input {{ background: transparent !important; color: {INK} !important; }}
    [data-testid="stNumberInput"] button {{
        background: #F8FAFC !important; border: 1px solid {BORDER} !important; color: {INK} !important;
    }}
    [data-testid="stFileUploaderDropzone"] {{
        background: {SURFACE} !important; border: 1px solid {BORDER} !important; border-radius: 10px !important;
    }}

    /* Buttons -> teal */
    [data-testid="stDownloadButton"] > button {{
        background: {BRAND} !important; color: #FFFFFF !important; border: none !important;
        border-radius: 8px !important; padding: 10px 18px !important; font-weight: 600 !important;
        width: 100% !important; transition: all 0.18s ease !important; box-shadow: 0 1px 2px rgba(14,163,127,0.25);
    }}
    [data-testid="stDownloadButton"] > button:hover {{
        background: {BRAND_HOVER} !important; transform: translateY(-1px);
        box-shadow: 0 4px 10px -2px rgba(14,163,127,0.35);
    }}
    .stButton > button {{
        background: {BRAND} !important; color:#fff !important; border:none !important;
        border-radius:8px !important; font-weight:600 !important;
    }}

    /* Brand header */
    .brand {{ display:flex; align-items:center; gap:13px; padding-bottom:18px; margin-bottom:6px; border-bottom:1px solid {BORDER}; }}
    .brand-mark {{ width:42px; height:42px; border-radius:11px; background:#EAF6F1; color:{BRAND};
        display:flex; align-items:center; justify-content:center; flex:none; }}
    .brand-title {{ font-size:1.4rem; font-weight:700; color:{INK}; letter-spacing:-0.025em; line-height:1.1; }}
    .brand-sub {{ font-size:0.82rem; color:{MUTED}; margin-top:2px; }}

    /* Section headers */
    .sec-head {{ display:flex; align-items:center; gap:9px; margin:2px 0 12px; }}
    .sec-head .ico {{ color:{BRAND}; display:inline-flex; }}
    .sec-head .txt {{ font-weight:600; font-size:1.0rem; color:{INK}; letter-spacing:-0.01em; }}

    /* Metric cards */
    .metric-row {{ display:flex; gap:14px; margin:8px 0 24px; flex-wrap:nowrap; }}
    .metric-card {{ flex:1; min-width:0; background:{SURFACE}; border:1px solid {BORDER}; border-radius:12px;
        padding:16px 18px; display:flex; align-items:center; gap:14px; box-shadow:0 1px 2px rgba(16,24,40,0.04); }}
    .metric-card .ico {{ flex:none; width:38px; height:38px; border-radius:9px; background:#EAF6F1; color:{BRAND};
        display:flex; align-items:center; justify-content:center; }}
    .metric-label {{ font-size:0.78rem; color:{MUTED}; font-weight:500; margin-bottom:3px;
        white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
    .metric-value {{ font-size:1.45rem; font-weight:600; color:{INK}; font-variant-numeric:tabular-nums;
        white-space:nowrap; letter-spacing:-0.02em; }}
    @media (max-width:820px) {{ .metric-row {{ flex-wrap:wrap; }} .metric-card {{ min-width:45%; }} }}

    /* Empty state */
    .empty-state {{ background:{SURFACE}; border:1px dashed #D2DCE6; border-radius:14px; padding:42px 24px;
        text-align:center; box-shadow:0 1px 2px rgba(16,24,40,0.03); margin-bottom:18px; }}
    .empty-state .ico {{ color:{BRAND}; display:inline-flex; margin-bottom:10px; }}
    .empty-title {{ font-weight:600; color:{INK}; font-size:1.02rem; }}
    .empty-sub {{ color:{MUTED}; font-size:0.86rem; margin-top:4px; }}

    /* Notices */
    .notice {{ padding:10px 14px; border:1px solid; border-radius:10px; font-size:0.84rem; margin:4px 0 16px; line-height:1.45; }}

    /* Result tables */
    .table-wrap {{ border:1px solid {BORDER}; border-radius:12px; overflow:hidden; background:{SURFACE};
        box-shadow:0 1px 2px rgba(16,24,40,0.04); }}
    .pf-table {{ width:100%; border-collapse:collapse; font-size:0.86rem; table-layout:auto; }}
    .pf-table thead th {{ background:#F8FAFC; color:{MUTED}; font-weight:600; font-size:0.72rem;
        text-transform:uppercase; letter-spacing:0.03em; padding:11px 14px; border-bottom:1px solid {BORDER}; white-space:nowrap; }}
    .pf-table tbody td {{ padding:11px 14px; border-bottom:1px solid #F1F5F9; color:{INK}; font-variant-numeric:tabular-nums; }}
    .pf-table tbody tr:last-child td {{ border-bottom:none; }}
    .pf-table tbody tr:hover td {{ background:#FAFBFC; }}
    .ticker-cell {{ font-weight:600; }}
    .muted-cell {{ color:{MUTED}; }}

    /* Pills */
    .pill {{ display:inline-block; padding:2px 10px; border-radius:999px; font-size:0.72rem; font-weight:600; white-space:nowrap; }}
    .pill-buy {{ background:#E7F6EF; color:{POS}; }}
    .pill-sell {{ background:#FCEBEC; color:{NEG}; }}
    .pill-action {{ background:#FCEBEC; color:{NEG}; }}
    .pill-ok {{ background:#EEF2F6; color:{MUTED}; }}

    /* Legend table */
    .legend-table {{ width:100%; border-collapse:collapse; font-size:0.85rem; }}
    .legend-table th {{ text-align:left; color:{MUTED}; font-weight:500; padding:7px 8px; border-bottom:1px solid {BORDER}; }}
    .legend-table th.num, .legend-table td.num {{ text-align:right; font-variant-numeric:tabular-nums; }}
    .legend-table td {{ padding:8px; border-bottom:1px solid #F1F5F9; color:{INK}; }}
    .legend-table tr:last-child td {{ border-bottom:none; }}
    .dot {{ display:inline-block; width:10px; height:10px; border-radius:50%; margin-right:9px; vertical-align:middle; }}

    /* Footer */
    .custom-footer {{ margin-top:54px; padding-top:20px; border-top:1px solid {BORDER};
        font-size:0.8rem; color:#9CA3AF; text-align:center; }}
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)


# --------------------------------------------------------------------------- #
# STATE
# --------------------------------------------------------------------------- #
if "lang" not in st.session_state:
    st.session_state.lang = "en"
if "currency" not in st.session_state:
    st.session_state.currency = "CAD"

t = LANG[st.session_state.lang]


# --------------------------------------------------------------------------- #
# RENDER HELPERS
# --------------------------------------------------------------------------- #
def section_header(icon_key, text):
    st.markdown(
        f'<div class="sec-head"><span class="ico">{ICONS[icon_key]}</span>'
        f'<span class="txt">{text}</span></div>',
        unsafe_allow_html=True,
    )

def notice(text, kind="warn"):
    palette = {
        "warn": ("#FEF6E7", "#92610A", "#F5D58A"),
        "info": ("#EEF4FF", "#1E40AF", "#C3D4FB"),
    }
    bg, fg, bd = palette[kind]
    st.markdown(
        f'<div class="notice" style="background:{bg};color:{fg};border-color:{bd}">{text}</div>',
        unsafe_allow_html=True,
    )

def pill(text, kind):
    return f'<span class="pill pill-{kind}">{text}</span>'

def html_table(headers, rows, aligns):
    head = "".join(
        f'<th class="{"num" if a == "right" else ""}">{h}</th>'
        for h, a in zip(headers, aligns)
    )
    body = ""
    for row in rows:
        cells = "".join(
            f'<td style="text-align:{a}">{c}</td>' for c, a in zip(row, aligns)
        )
        body += f"<tr>{cells}</tr>"
    return (
        f'<div class="table-wrap"><table class="pf-table">'
        f"<thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>"
    )

def ccy_symbol(ccy):
    return CCY_SYMBOLS.get(str(ccy).upper(), f"{ccy} ")

def render_instructions(expanded):
    with st.expander(t["howto_title"], expanded=expanded):
        st.markdown(t["howto_text"])

def render_footer():
    st.markdown(f'<div class="custom-footer">{t["footer"]}</div>', unsafe_allow_html=True)

def build_color_map(classes):
    cmap, i = {}, 0
    for c in sorted(classes, key=lambda x: str(x)):
        cl = str(c).strip().lower()
        if "cash" in cl:
            cmap[c] = CASH_COLOR
        elif cl in ("unclassified", "nan", ""):
            cmap[c] = UNCLASS_COLOR
        else:
            cmap[c] = CATEGORY_PALETTE[i % len(CATEGORY_PALETTE)]
            i += 1
    return cmap

def make_donut(df, value_col, color_map):
    df = df.copy()
    total = df[value_col].sum()
    df["pct"] = (df[value_col] / total * 100.0) if total else 0.0
    df["lbl"] = df["pct"].apply(lambda p: f"{p:.1f}%" if p >= 6 else "")

    light_fills = {CASH_COLOR, UNCLASS_COLOR, "#F59E0B", "#84CC16"}
    label_colors = [
        INK if color_map.get(c, CASH_COLOR) in light_fills else "#FFFFFF"
        for c in df["Asset_Class"]
    ]

    fig = px.pie(
        df, values=value_col, names="Asset_Class", hole=0.62,
        color="Asset_Class", color_discrete_map=color_map,
    )
    fig.update_traces(
        text=df["lbl"], textinfo="text", textposition="inside",
        insidetextorientation="horizontal", sort=False,
        marker=dict(line=dict(color="#FFFFFF", width=2)),
        hovertemplate="%{label}: %{percent}<extra></extra>",
        textfont=dict(color=label_colors, size=13),
    )
    fig.update_layout(
        showlegend=False, height=292, margin=dict(t=6, b=6, l=6, r=6),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, sans-serif", size=13, color=INK),
    )
    return fig


# --------------------------------------------------------------------------- #
# DATA HELPERS
# --------------------------------------------------------------------------- #
@st.cache_data(ttl=PRICE_CACHE_TTL, show_spinner=False)
def fetch_live_price(ticker):
    """
    Return (price, currency). Price via fast_info -> .info fields -> recent close.
    Currency from yfinance metadata, falling back to a suffix guess.
    """
    ccy = None
    try:
        tk_obj = yf.Ticker(ticker)
        price = None
        try:
            fi = tk_obj.fast_info
            price = fi.get("last_price")
            ccy = fi.get("currency")
        except Exception:
            pass
        if not (price and price > 0):
            try:
                info = tk_obj.info or {}
                ccy = ccy or info.get("currency")
                for key in ("regularMarketPrice", "currentPrice", "previousClose"):
                    if info.get(key) and info.get(key) > 0:
                        price = float(info[key])
                        break
            except Exception:
                pass
        if not (price and price > 0):
            hist = tk_obj.history(period="5d")
            if not hist.empty:
                price = float(hist["Close"].dropna().iloc[-1])
        if price and price > 0:
            if not ccy:
                ccy = "CAD" if ticker.endswith((".TO", ".V", ".CN", ".NE")) else "USD"
            return float(price), str(ccy).upper()
        return None, ccy
    except Exception:
        return None, ccy

@st.cache_data(ttl=3600, show_spinner=False)
def fetch_fx_rates():
    """Live CAD/USD conversion rates with sensible fallbacks."""
    try:
        cad_usd = yf.Ticker("CADUSD=X").fast_info.get("last_price") or 0.73
        usd_cad = yf.Ticker("USDCAD=X").fast_info.get("last_price") or 1.37
        return float(cad_usd), float(usd_cad)
    except Exception:
        return 0.73, 1.37

def is_cash_ticker(tk):
    return "CASH" in tk or tk == "NBC200"


# --------------------------------------------------------------------------- #
# THE SCANNER (reads messy mandate / holdings files, maps aliased columns)
# --------------------------------------------------------------------------- #
def read_tabular(path_or_file, expected_type="mandate"):
    name = getattr(path_or_file, "name", str(path_or_file)).lower()
    if name.endswith((".xlsx", ".xls")):
        raw_dfs = pd.read_excel(path_or_file, header=None, sheet_name=None)
    else:
        raw_dfs = {"Sheet1": pd.read_csv(path_or_file, header=None)}

    valid_dfs = []
    for sheet_name, raw_df in raw_dfs.items():
        if raw_df.empty:
            continue

        first_cell = str(raw_df.iloc[0, 0]).strip()
        fallback_name = first_cell if (len(first_cell) > 3 and first_cell != "nan") else sheet_name

        header_idx = -1
        for i, row in raw_df.iterrows():
            row_str = " ".join(str(v).lower() for v in row.values)
            if any(k in row_str for k in ("ticker", "symbol", "weight", "quantity")):
                header_idx = i
                break
        if header_idx == -1:
            continue

        df = raw_df.iloc[header_idx + 1:].copy()
        df.columns = raw_df.iloc[header_idx].values
        df = df.dropna(how="all")

        aliases = {
            "Mandate_Name": ["mandate", "mandatename", "strategy", "portfolio", "model", "name", "mandat"],
            "Ticker": ["ticker", "symbol", "sym", "asset", "symbole", "ticker/identifier"],
            "Asset_Class": ["assetclass", "class", "sector", "category", "type", "categorie", "classe"],
            "Target_Weight_Percent": ["target", "targetweight", "weight", "targetpercent", "allocation", "cible", "poids", "target%"],
            "Drift_Band_Percent": ["drift", "driftband", "tolerance", "band", "driftpercent", "derive", "marge", "drift%", "driftband(+/-)"],
            "Quantity": ["quantity", "qty", "shares", "units", "position", "amount", "quantite", "currentquantity(units)"],
        }
        rename_map = {}
        for col in df.columns:
            clean = str(col).lower().replace(" ", "").replace("_", "").replace("%", "").strip()
            for canonical, alist in aliases.items():
                if clean in alist:
                    rename_map[col] = canonical
                    break
        df = df.rename(columns=rename_map)

        if expected_type == "mandate" and "Mandate_Name" not in df.columns:
            df["Mandate_Name"] = fallback_name

        if "Ticker" in df.columns:
            df = df[~df["Ticker"].isna()]
            df = df[~df["Ticker"].astype(str).str.upper().str.contains("TOTAL")]
            df = df[~df["Ticker"].astype(str).str.startswith("*")]

        for col in ("Target_Weight_Percent", "Drift_Band_Percent", "Quantity"):
            if col in df.columns:
                df[col] = (
                    df[col].astype(str)
                    .str.replace("%", "").str.replace(",", ".").str.replace(" ", "")
                    .str.replace("\xa0", "").str.strip()
                )
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

        required = (["Mandate_Name", "Ticker", "Target_Weight_Percent"]
                    if expected_type == "mandate" else ["Ticker", "Quantity"])
        if all(c in df.columns for c in required):
            valid_dfs.append(df)

    if not valid_dfs:
        need = (["Mandate_Name", "Ticker", "Target_Weight_Percent"]
                if expected_type == "mandate" else ["Ticker", "Quantity"])
        st.error(f"Could not map columns in your file. Expected variations of: {', '.join(need)}.")
        st.stop()

    return pd.concat(valid_dfs, ignore_index=True)


# --------------------------------------------------------------------------- #
# BRAND HEADER  (sliders icon — distinct from the chart icons)
# --------------------------------------------------------------------------- #
st.markdown(
    f"""
    <div class="brand">
        <div class="brand-mark">{ICONS['sliders']}</div>
        <div>
            <div class="brand-title">{t['header_title']}</div>
            <div class="brand-sub">{t['header_desc']}</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# --------------------------------------------------------------------------- #
# SIDEBAR
# --------------------------------------------------------------------------- #
col_lang, col_curr = st.sidebar.columns(2)
new_lang = col_lang.selectbox("Language / Langue", ["English", "Français"],
                              index=0 if st.session_state.lang == "en" else 1)
st.session_state.lang = "en" if new_lang == "English" else "fr"
t = LANG[st.session_state.lang]

new_curr = col_curr.selectbox("Base currency", ["USD", "CAD"],
                              index=0 if st.session_state.currency == "USD" else 1)
st.session_state.currency = new_curr

st.sidebar.markdown("---")
st.sidebar.subheader(t["sidebar_config"])
st.sidebar.write("")

# 1. Mandate
st.sidebar.markdown(f"**{t['mandate_spec']}**")
uploaded_mandate = st.sidebar.file_uploader(t["upload_mandate"], type=["csv", "xlsx", "xls"])
if uploaded_mandate is not None:
    with open(os.path.join(SAVED_MANDATES_DIR, uploaded_mandate.name), "wb") as f:
        f.write(uploaded_mandate.getbuffer())

saved_mandates = sorted(
    f for f in os.listdir(SAVED_MANDATES_DIR)
    if f.lower().endswith((".csv", ".xlsx", ".xls"))
    and "holdings" not in f.lower() and "portfolio" not in f.lower()
)
if not saved_mandates:
    st.sidebar.error("No mandates found. Please upload one.")
    render_instructions(expanded=True)
    render_footer()
    st.stop()

selected_file = st.sidebar.selectbox(t["active_db"], saved_mandates)
mandate_df = read_tabular(os.path.join(SAVED_MANDATES_DIR, selected_file), "mandate")
mandate_names = sorted(mandate_df["Mandate_Name"].dropna().unique())
selected_mandate = st.sidebar.selectbox(t["target_mandate"], mandate_names)
active = mandate_df[mandate_df["Mandate_Name"] == selected_mandate].copy()

# --- Normalize target-weight scale: fractions (0.03) -> percent (3) ---------
active["Target_Weight_Percent"] = pd.to_numeric(active["Target_Weight_Percent"], errors="coerce").fillna(0.0)
if "Drift_Band_Percent" in active.columns:
    active["Drift_Band_Percent"] = pd.to_numeric(active["Drift_Band_Percent"], errors="coerce").fillna(0.0)
_tsum = float(active["Target_Weight_Percent"].sum())
_scale = 100.0 if (0 < _tsum <= 1.5) else 1.0   # sums to ~1 -> stored as fractions
if _scale != 1.0:
    active["Target_Weight_Percent"] = active["Target_Weight_Percent"] * _scale
    if "Drift_Band_Percent" in active.columns:
        active["Drift_Band_Percent"] = active["Drift_Band_Percent"] * _scale

with st.sidebar.expander(t["view_mandate"]):
    cols = [c for c in ["Ticker", "Asset_Class", "Target_Weight_Percent"] if c in active.columns]
    st.dataframe(
        active[cols], hide_index=True, use_container_width=True,
        column_config={
            "Ticker": st.column_config.TextColumn(t["col_ticker"]),
            "Asset_Class": st.column_config.TextColumn(t["col_asset"]),
            "Target_Weight_Percent": st.column_config.NumberColumn(t["col_target"], format="%.2f%%"),
        },
    )

st.sidebar.write("")

# 2. Cash flow
st.sidebar.markdown(f"**{t['tactical_cash']}**")
cash_flow = st.sidebar.number_input(
    f"{t['deposit_withdraw']} ({st.session_state.currency})",
    value=0.0, step=1000.0, format="%.2f",
)


# --------------------------------------------------------------------------- #
# MAIN — TITLE + HOLDINGS UPLOAD
# --------------------------------------------------------------------------- #
st.subheader(t["main_title"])
st.caption(f"{t['main_tracking']}: {selected_mandate}")
st.write("")
st.markdown(f"**{t['client_holdings']}**")
holdings_file = st.file_uploader(t["upload_holdings"], type=["csv", "xlsx", "xls"], key="client_file")

# --- AUTO-LOADER LOGIC FOR DEMO HOLDINGS ---
demo_files = sorted(f for f in os.listdir(DEMO_HOLDINGS_DIR) if f.lower().endswith((".csv", ".xlsx", ".xls")))

if holdings_file is not None:
    st.session_state.cached_holdings = read_tabular(holdings_file, "holdings")
elif demo_files:
    demo_path = os.path.join(DEMO_HOLDINGS_DIR, demo_files[0])
    st.session_state.cached_holdings = read_tabular(demo_path, "holdings")
    st.info(f"💡 Auto-loaded demo portfolio: **{demo_files[0]}**. Upload your own file to override this.")

if "cached_holdings" not in st.session_state or st.session_state.cached_holdings is None:
    st.caption(t["waiting_holdings"])
    render_instructions(expanded=True); render_footer(); st.stop()

holdings_df = st.session_state.cached_holdings.copy()
holdings_df["Ticker"] = holdings_df["Ticker"].astype(str).str.strip().str.upper()
holdings_df["Quantity"] = pd.to_numeric(holdings_df["Quantity"], errors="coerce").fillna(0.0)


# --------------------------------------------------------------------------- #
# BUILD UNIVERSE & MAPS
# --------------------------------------------------------------------------- #
active["Ticker"] = active["Ticker"].astype(str).str.strip().str.upper()
if "Drift_Band_Percent" not in active.columns:
    active["Drift_Band_Percent"] = 0.0
if "Asset_Class" not in active.columns:
    active["Asset_Class"] = "Unclassified"

target_map = active.set_index("Ticker")[["Asset_Class", "Target_Weight_Percent", "Drift_Band_Percent"]].to_dict("index")
holdings_map = dict(zip(holdings_df["Ticker"], holdings_df["Quantity"]))
holdings_class_map = (
    dict(zip(holdings_df["Ticker"], holdings_df["Asset_Class"]))
    if "Asset_Class" in holdings_df.columns else {}
)
universe = sorted(set(holdings_map) | set(target_map))


# --------------------------------------------------------------------------- #
# LIVE PRICING + FX
#   Native_Price -> shown in each security's own market currency
#   Base_Price   -> converted into the chosen base currency, used for all math
# --------------------------------------------------------------------------- #
cad_usd, usd_cad = fetch_fx_rates()
base = st.session_state.currency

def to_base_factor(native_ccy):
    n = str(native_ccy).upper()
    if n == base:
        return 1.0
    if n == "USD" and base == "CAD":
        return usd_cad
    if n == "CAD" and base == "USD":
        return cad_usd
    return 1.0   # unknown pair -> face value

native_price, native_ccy, base_price_map, missing_prices = {}, {}, {}, []
with st.spinner(t["fetching_prices"]):
    for tk in universe:
        if is_cash_ticker(tk):
            nccy = "USD" if "USD" in tk else ("CAD" if "CAD" in tk else base)
            np_ = 1.0
        else:
            np_, ccy = fetch_live_price(tk)
            if np_ is None:
                missing_prices.append(tk)
                continue
            nccy = ccy or ("CAD" if tk.endswith((".TO", ".V", ".CN", ".NE")) else "USD")
        native_price[tk] = np_
        native_ccy[tk] = nccy
        base_price_map[tk] = np_ * to_base_factor(nccy)

priced = [tk for tk in universe if tk in base_price_map]


# --------------------------------------------------------------------------- #
# CORE CALCULATIONS
# --------------------------------------------------------------------------- #
def _tgt(tk, field, default=0.0):
    return target_map.get(tk, {}).get(field, default)

def resolve_asset_class(tk):
    if is_cash_ticker(tk):
        return "Cash Equivalent"
    c = _tgt(tk, "Asset_Class", None)
    if c and str(c).strip().lower() not in ("", "nan", "unclassified"):
        return c
    h = holdings_class_map.get(tk)
    if h and str(h).strip().lower() not in ("", "nan"):
        return h
    return "Unclassified"

rows = []
for tk in priced:
    qty = float(holdings_map.get(tk, 0.0))
    bp = base_price_map[tk]
    rows.append({
        "Ticker": tk, "Quantity": qty,
        "Native_Price": native_price[tk], "Native_Currency": native_ccy[tk],
        "Base_Price": bp, "Market_Value": qty * bp,
    })
analysis = pd.DataFrame(rows)

if missing_prices:
    notice(t["missing_note"].format(items=", ".join(missing_prices)), "warn")

if analysis.empty:
    notice(t["no_priced"], "warn")
    render_instructions(expanded=True)
    render_footer()
    st.stop()

total_value = float(analysis["Market_Value"].sum())
total_to_allocate = total_value + cash_flow

analysis["Asset_Class"] = analysis["Ticker"].apply(resolve_asset_class)
analysis["Current_Weight"] = (analysis["Market_Value"] / total_value * 100.0) if total_value else 0.0
analysis["Target_Weight"] = analysis["Ticker"].apply(lambda tk: _tgt(tk, "Target_Weight_Percent"))
analysis["Drift_Band"] = analysis["Ticker"].apply(lambda tk: _tgt(tk, "Drift_Band_Percent"))
analysis["Drift"] = analysis["Current_Weight"] - analysis["Target_Weight"]
analysis["Breach"] = analysis["Drift"].abs() > analysis["Drift_Band"]
analysis = analysis.sort_values("Market_Value", ascending=False).reset_index(drop=True)

cash_mask = (
    analysis["Asset_Class"].str.lower().str.contains("cash", na=False)
    | analysis["Ticker"].str.contains("CASH", na=False)
)
cash_equiv_value = float(analysis.loc[cash_mask, "Market_Value"].sum())
cash_equiv_pct = (cash_equiv_value / total_value * 100.0) if total_value else 0.0

# Trade generation (band-aware). Base_Price keeps shares + dollars consistent.
orders = []
for _, r in analysis.iterrows():
    tk = r["Ticker"]
    if is_cash_ticker(tk) or not r["Breach"]:
        continue
    target_dollar = r["Target_Weight"] / 100.0 * total_to_allocate
    trade_dollar = target_dollar - r["Market_Value"]
    p = r["Base_Price"]
    shares = round(trade_dollar / p) if p else 0
    if shares == 0:
        continue
    orders.append({
        "Ticker": tk,
        "Action": "BUY" if shares > 0 else "SELL",
        "Shares_to_Trade": int(abs(shares)),
        "Estimated_Dollar_Value": round(abs(shares) * p, 2),
        "_signed_dollar": shares * p,
    })
orders_df = pd.DataFrame(orders)
net_security_flow = float(orders_df["_signed_dollar"].sum()) if not orders_df.empty else 0.0
projected_cash = cash_equiv_value + cash_flow - net_security_flow


# --------------------------------------------------------------------------- #
# DASHBOARD — METRIC CARDS  (all base currency)
# --------------------------------------------------------------------------- #
curr_sym = "$" if st.session_state.currency == "USD" else "C$"

st.markdown(
    f"""
    <div class="metric-row">
        <div class="metric-card"><div class="ico">{ICONS['wallet']}</div>
            <div><div class="metric-label">{t['curr_val']}</div>
            <div class="metric-value">{curr_sym}{total_value:,.0f}</div></div></div>
        <div class="metric-card"><div class="ico">{ICONS['flow']}</div>
            <div><div class="metric-label">{t['cash_flow']}</div>
            <div class="metric-value">{curr_sym}{cash_flow:,.0f}</div></div></div>
        <div class="metric-card"><div class="ico">{ICONS['coins']}</div>
            <div><div class="metric-label">{t['cash_equiv']}</div>
            <div class="metric-value">{cash_equiv_pct:.1f}%</div></div></div>
        <div class="metric-card" title="Cash remaining after the proposed trades settle">
            <div class="ico">{ICONS['banknote']}</div>
            <div><div class="metric-label">{t['post_cash']}</div>
            <div class="metric-value">{curr_sym}{projected_cash:,.0f}</div></div></div>
    </div>
    """,
    unsafe_allow_html=True,
)

# Validation notices
target_sum = float(active["Target_Weight_Percent"].sum())
if abs(target_sum - 100.0) > 0.5:
    notice(t["targetsum_note"].format(pct=target_sum), "warn")

off_mandate = [tk for tk in holdings_map if tk not in target_map and not is_cash_ticker(tk)]
if off_mandate:
    notice(t["offmandate_note"].format(items=", ".join(off_mandate)), "info")


# --------------------------------------------------------------------------- #
# DASHBOARD — DONUT CHARTS + LEGEND
# --------------------------------------------------------------------------- #
curr_pie = analysis[analysis["Current_Weight"] > 0].groupby("Asset_Class")["Current_Weight"].sum().reset_index()
tgt_pie = active[active["Target_Weight_Percent"] > 0].groupby("Asset_Class")["Target_Weight_Percent"].sum().reset_index()

color_map = build_color_map(set(curr_pie["Asset_Class"]) | set(tgt_pie["Asset_Class"]))

col_pie1, col_pie2 = st.columns(2)
with col_pie1:
    with st.container(border=True):
        section_header("donut", t["alloc_current"])
        if not curr_pie.empty:
            st.plotly_chart(make_donut(curr_pie, "Current_Weight", color_map),
                            use_container_width=True, config={"displayModeBar": False})
with col_pie2:
    with st.container(border=True):
        section_header("target", t["alloc_target"])
        if not tgt_pie.empty:
            st.plotly_chart(make_donut(tgt_pie, "Target_Weight_Percent", color_map),
                            use_container_width=True, config={"displayModeBar": False})

# Full breakdown (with exact % for the small slices) inside a dropdown.
cur_map = dict(zip(curr_pie["Asset_Class"], curr_pie["Current_Weight"]))
tgt_map = dict(zip(tgt_pie["Asset_Class"], tgt_pie["Target_Weight_Percent"]))
legend_classes = sorted(set(cur_map) | set(tgt_map), key=lambda c: -cur_map.get(c, 0.0))

legend_rows = ""
for cls in legend_classes:
    color = color_map.get(cls, CASH_COLOR)
    legend_rows += (
        f'<tr><td><span class="dot" style="background:{color}"></span>{cls}</td>'
        f'<td class="num">{cur_map.get(cls, 0.0):.1f}%</td>'
        f'<td class="num">{tgt_map.get(cls, 0.0):.1f}%</td></tr>'
    )
legend_html = (
    f'<table class="legend-table"><thead><tr>'
    f'<th>{t["col_asset"]}</th><th class="num">{t["col_current"]}</th>'
    f'<th class="num">{t["col_target"]}</th></tr></thead>'
    f"<tbody>{legend_rows}</tbody></table>"
)
with st.expander(t["legend_title"]):
    st.markdown(legend_html, unsafe_allow_html=True)

st.write("")


# --------------------------------------------------------------------------- #
# DASHBOARD — RESULT TABLES (full width, stacked, custom HTML)
# --------------------------------------------------------------------------- #
section_header("list", t["analysis_title"])
h_rows = []
for _, r in analysis.iterrows():
    breach = bool(r["Breach"])
    drift_html = (
        f'<span style="color:{NEG if breach else MUTED};'
        f'font-weight:{600 if breach else 400}">{r["Drift"]:+.2f}</span>'
    )
    status_html = pill(t["status_breach"], "action") if breach else pill(t["status_ok"], "ok")
    h_rows.append([
        f'<span class="ticker-cell">{r["Ticker"]}</span>',
        f'<span class="muted-cell">{r["Asset_Class"]}</span>',
        f'{ccy_symbol(r["Native_Currency"])}{r["Native_Price"]:,.2f}',
        f'{r["Current_Weight"]:.2f}%',
        f'{r["Target_Weight"]:.2f}%',
        drift_html,
        status_html,
    ])
st.markdown(
    html_table(
        [t["col_ticker"], t["col_asset"], t["col_price"], t["col_current"],
         t["col_target"], t["col_drift"], t["col_status"]],
        h_rows,
        ["left", "left", "right", "right", "right", "right", "left"],
    ),
    unsafe_allow_html=True,
)

st.write("")

section_header("exchange", t["orders_title"])
if orders_df.empty:
    st.markdown(
        f'<div class="table-wrap"><div style="padding:18px 16px;color:{MUTED};'
        f'font-size:0.88rem">{t["no_trades"]}</div></div>',
        unsafe_allow_html=True,
    )
else:
    o_rows = []
    for _, r in orders_df.iterrows():
        action_html = pill(r["Action"], "buy" if r["Action"] == "BUY" else "sell")
        o_rows.append([
            f'<span class="ticker-cell">{r["Ticker"]}</span>',
            action_html,
            f'{int(r["Shares_to_Trade"]):,}',
            f'{curr_sym}{r["Estimated_Dollar_Value"]:,.0f}',
        ])
    st.markdown(
        html_table(
            [t["col_ticker"], t["col_action"], t["col_shares"], t["col_est_value"]],
            o_rows,
            ["left", "left", "right", "right"],
        ),
        unsafe_allow_html=True,
    )
    st.write("")
    export_df = orders_df[["Ticker", "Action", "Shares_to_Trade", "Estimated_Dollar_Value"]].copy()
    dl_col = st.columns([1, 2.2])[0]
    with dl_col:
        st.download_button(
            t["export_btn"],
            data=export_df.to_csv(index=False).encode("utf-8"),
            file_name=f"trade_ticket_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
            use_container_width=True,
        )


# --------------------------------------------------------------------------- #
# INSTRUCTIONS + FOOTER  (also shown on the empty screen above)
# --------------------------------------------------------------------------- #
st.markdown("---")
render_instructions(expanded=False)
render_footer()
