"""
=============================================================================
 PRIVATE WEALTH ANALYTICS - PORTFOLIO IMPLEMENTATION ENGINE
=============================================================================
 Live drift monitoring, trade-order generation, and tax-aware impact.

 THREE MODES, auto-detected from the holdings file:
   1. Legacy / aggregate     -> Ticker, Quantity only.
        Whole book vs. the selected mandate. Donuts + 2 tables + cash flow.
   2. Aggregate + tax (Tier 0)-> adds Account_Type + Book_Cost_CAD.
        Same aggregate view, plus per-sell realized gain / estimated tax,
        with sells routed registered-first across the book.
   3. Per-account mandates    -> adds Account + Mandate_Name.
        Each account is rebalanced INDEPENDENTLY to its own mandate, shown as
        stacked per-account blocks with a consolidated tax summary on top and
        per-account cash inputs. (Registered-first routing does not apply here
        because each account self-corrects to its own target.)

 Canada tax model: average-cost ACB (not lots); gains computed in CAD;
 inclusion rate configurable (50% default). No management-fee logic anywhere.
=============================================================================
"""

import os
import sys
import subprocess
import unicodedata
from datetime import datetime

import pandas as pd
import streamlit as st

try:
    import yfinance as yf
    import plotly.express as px
except ModuleNotFoundError as e:
    st.set_page_config(page_title="Portfolio Implementation Engine", layout="wide")
    st.error(f"Missing required library: {e.name}")
    st.write("If the terminal install isn't working, install the libraries here, then refresh.")
    if st.button("Install required libraries"):
        with st.spinner("Installing…"):
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", "plotly", "yfinance"])
                st.success("Installed. Refresh the page (press R) to continue.")
            except Exception as err:
                st.error(f"Installation failed: {err}")
    st.stop()


# --------------------------------------------------------------------------- #
# CONFIG
# --------------------------------------------------------------------------- #
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SAVED_MANDATES_DIR = os.path.join(BASE_DIR, "saved_mandates")
DEMO_HOLDINGS = os.path.join(BASE_DIR, "demo_holdings", "Client_Targeted_Rebalance_v1.xlsx")
PRICE_CACHE_TTL = 300
os.makedirs(SAVED_MANDATES_DIR, exist_ok=True)
st.set_page_config(page_title="Portfolio Implementation Engine", layout="wide",
                   initial_sidebar_state="expanded")

REGISTERED_TYPES = {"TFSA", "RRSP", "RRIF", "FHSA", "RESP", "LIRA", "LIF",
                    "LRSP", "RDSP", "DPSP", "RPP", "RLIF", "PRPP"}

BRAND, BRAND_HOVER = "#B0894F", "#967333"          # muted brass accent
NAVY, NAVY_SOFT, NAVY_LINE = "#102A43", "#1C3A57", "#284A6B"
INK, MUTED, BORDER = "#16293D", "#5E6E7F", "#E5E1D8"
SURFACE, CANVAS = "#FFFFFF", "#F4F2EC"
POS, NEG, AMBER = "#1C7A4D", "#B23A3A", "#8A6A33"
CASH_COLOR, UNCLASS_COLOR = "#9AA7B4", "#CBD3DB"
CATEGORY_PALETTE = ["#B0894F", "#2C5282", "#4A7C7C", "#9C6B4A", "#5B7553",
                    "#8C6BA8", "#C0884F", "#3D6B8E", "#A8557A", "#6B7280"]
CCY_SYMBOLS = {"USD": "US$", "CAD": "C$", "EUR": "€", "GBP": "£", "JPY": "¥", "CHF": "CHF "}


# --------------------------------------------------------------------------- #
# SVG ICONS — no emoji
# --------------------------------------------------------------------------- #
def _svg(body, size=18):
    return (f'<svg viewBox="0 0 24 24" width="{size}" height="{size}" fill="none" '
            f'stroke="currentColor" stroke-width="1.75" stroke-linecap="round" '
            f'stroke-linejoin="round">{body}</svg>')

ICONS = {
    "sliders": _svg('<line x1="4" y1="21" x2="4" y2="13"/><line x1="4" y1="9" x2="4" y2="3"/>'
                    '<line x1="12" y1="21" x2="12" y2="12"/><line x1="12" y1="8" x2="12" y2="3"/>'
                    '<line x1="20" y1="21" x2="20" y2="15"/><line x1="20" y1="11" x2="20" y2="3"/>'
                    '<line x1="2" y1="11" x2="6" y2="11"/><line x1="10" y1="10" x2="14" y2="10"/>'
                    '<line x1="18" y1="13" x2="22" y2="13"/>'),
    "donut": _svg('<circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="3.4"/>'),
    "target": _svg('<circle cx="12" cy="12" r="8.5"/><circle cx="12" cy="12" r="3.3"/>'
                   '<line x1="12" y1="1.5" x2="12" y2="5"/><line x1="12" y1="19" x2="12" y2="22.5"/>'
                   '<line x1="1.5" y1="12" x2="5" y2="12"/><line x1="19" y1="12" x2="22.5" y2="12"/>'),
    "wallet": _svg('<rect x="3" y="6" width="18" height="13" rx="2.5"/>'
                   '<path d="M3 10h18"/><circle cx="16.5" cy="14" r="1.1"/>'),
    "flow": _svg('<polyline points="7 8 3 12 7 16"/><line x1="3" y1="12" x2="21" y2="12"/>'
                 '<polyline points="17 8 21 12 17 16"/>'),
    "coins": _svg('<circle cx="12" cy="12" r="8.6"/><line x1="12" y1="7" x2="12" y2="17"/>'
                  '<path d="M15 7.9H10.4a2.3 2.3 0 0 0 0 4.6h3.2a2.3 2.3 0 0 1 0 4.6H8.4"/>'),
    "banknote": _svg('<rect x="2.5" y="6" width="19" height="12" rx="2.5"/>'
                     '<circle cx="12" cy="12" r="2.4"/><path d="M6 12h.01M18 12h.01"/>'),
    "list": _svg('<line x1="8.5" y1="7" x2="20" y2="7"/><line x1="8.5" y1="12" x2="20" y2="12"/>'
                 '<line x1="8.5" y1="17" x2="20" y2="17"/><circle cx="4.5" cy="7" r="1.1"/>'
                 '<circle cx="4.5" cy="12" r="1.1"/><circle cx="4.5" cy="17" r="1.1"/>'),
    "exchange": _svg('<polyline points="7 5 7 19"/><polyline points="4 16 7 19 10 16"/>'
                     '<polyline points="17 19 17 5"/><polyline points="14 8 17 5 20 8"/>'),
    "percent": _svg('<line x1="19" y1="5" x2="5" y2="19"/><circle cx="6.5" cy="6.5" r="2.4"/>'
                    '<circle cx="17.5" cy="17.5" r="2.4"/>'),
    "folder": _svg('<path d="M4 7a2 2 0 0 1 2-2h3l2 2h7a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V7z"/>'),
    "layers": _svg('<path d="M12 3 3 7.5l9 4.5 9-4.5L12 3z"/><path d="M3 12l9 4.5L21 12"/>'
                   '<path d="M3 16.5 12 21l9-4.5"/>'),
    "upload": _svg('<path d="M12 15V4"/><polyline points="8 8 12 4 16 8"/>'
                   '<path d="M4 15v3a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-3"/>', size=26),
}


# --------------------------------------------------------------------------- #
# BILINGUAL DICTIONARY
# --------------------------------------------------------------------------- #
LANG = {
    "en": {
        "header_title": "Portfolio Implementation Engine",
        "header_desc": "Drift monitoring, tax-aware trades & per-account mandates",
        "sidebar_config": "System configuration",
        "mandate_spec": "1. Mandate specification",
        "upload_mandate": "Upload IPS mandate (CSV or Excel)",
        "active_db": "Active database file",
        "target_mandate": "Default / fallback mandate",
        "tactical_cash": "2. Tactical cash flow",
        "deposit_withdraw": "Deposit (+) / Withdrawal (−)",
        "tax_assumptions": "3. Tax assumptions",
        "tax_help": "Used when your holdings include Account_Type and Book_Cost_CAD.",
        "marginal_rate": "Client marginal rate, ordinary income (%)",
        "inclusion_rate": "Capital gains inclusion rate (%)",
        "main_title": "Implementation & order generation",
        "main_tracking": "Currently tracking",
        "client_holdings": "Client holdings",
        "upload_holdings": "Upload client holdings (CSV or Excel)",
        "empty_title": "Upload client holdings to begin",
        "waiting_holdings": "Add a file with Ticker and Quantity columns to run the live analysis.",
        "demo_banner": "Showing a sample portfolio so you can explore the engine. Upload your own holdings above to replace it.",
        "curr_val": "Current portfolio value",
        "household_value": "Household value",
        "cash_flow": "Tactical cash flow",
        "cash_equiv": "Cash & equivalents",
        "post_cash": "Est. residual cash",
        "alloc_current": "Current allocation",
        "alloc_target": "Target allocation",
        "legend_title": "Allocation breakdown & legend",
        "analysis_title": "Holdings & drift analysis",
        "orders_title": "Execution orders",
        "taximpact_title": "Tax impact of this rebalance",
        "taximpact_hh": "Consolidated tax impact (household)",
        "accounts_title": "Holdings by account",
        "per_acct_cash": "Tactical cash flow — per account",
        "per_acct_blocks": "Per-account rebalance",
        "combined_ticket": "Combined household trade ticket",
        "approach_label": "Rebalancing approach",
        "approach_per": "Each account → its own mandate",
        "approach_whole": "Whole household → one mandate (tax-routed)",
        "assign_title": "Mandate assignment",
        "view_account": "Select account to view in detail",
        "cash_on_main": "Per-account cash inputs are on the main page →",
        "status_ok": "On target",
        "status_breach": "Action required",
        "no_trades": "Within tolerance — no trades required.",
        "export_btn": "Export trade ticket (CSV)",
        "export_combined": "Export combined household ticket (CSV)",
        "footer": "Realized by Samari Gbian",
        "fetching_prices": "Fetching live market data and FX rates…",
        "view_mandate": "View mandate targets",
        "howto_title": "How to use this engine",
        "col_ticker": "Ticker", "col_asset": "Asset class", "col_price": "Price",
        "col_current": "Current", "col_target": "Target", "col_drift": "Drift",
        "col_status": "Status", "col_action": "Action", "col_shares": "Shares",
        "col_est_value": "Est. value", "col_gain": "Realized gain (CAD)", "col_tax": "Est. tax (CAD)",
        "col_account": "Account", "col_type": "Type", "col_acb": "ACB / unit (CAD)",
        "col_mv": "Market value", "col_unreal": "Unrealized gain (CAD)", "col_mandate": "Mandate",
        "m_net_gain": "Net realized gain (CAD)", "m_est_tax": "Estimated tax (CAD)",
        "m_eff_rate": "Effective tax rate", "m_accounts": "Accounts",
        "lbl_gain": "Realized gain", "lbl_tax": "Est. tax", "lbl_value": "Account value",
        "tax_caption": "Sells are drawn registered-first (no tax), then taxable accounts (losses before gains). "
                       "Capital losses offset gains. Gains are computed in CAD regardless of display currency.",
        "tax_caption_pa": "Each account is rebalanced independently to its own mandate; realized gains are computed "
                          "on that account's taxable sells against its average ACB. The consolidated total nets "
                          "losses against gains across accounts (assumes one taxpayer).",
        "tax_hint": "Add Account, Account_Type, and Book_Cost_CAD columns to unlock realized-gain and tax analysis. "
                    "Add a Mandate_Name column to rebalance each account to its own mandate.",
        "no_priced": "No live prices could be retrieved for any holding. Check your tickers and try again.",
        "no_priced_acct": "{acct}: no live prices for its holdings — skipped.",
        "missing_note": "No live price for: {items}. These positions are excluded; weights reflect only priced holdings.",
        "offmandate_note": "Held but not in this mandate (0% target → full sell): {items}.",
        "targetsum_note": "Target weights for {name} sum to {pct:.1f}%, not 100%.",
        "mandate_missing": "Account {acct}: mandate '{name}' not found in the database; using fallback '{fb}'.",
        "howto_text": """
**1. Configure (sidebar).** Language, **base currency** (portfolio-level figures only — security prices stay native), and under **Tax assumptions** the marginal rate and inclusion rate.

**2. Upload an IPS mandate database.** CSV/Excel with `Mandate_Name`, `Ticker`, `Asset_Class`, `Target_Weight_Percent` (percent `20` or fraction `0.20`), `Drift_Band_Percent`. One file can hold many mandates. Saved to `saved_mandates/`.

**3. Pick the default/fallback mandate** (used in the simple modes, and for any account whose named mandate isn't found).

**4. Cash flow.** Simple modes use the sidebar input; per-account mode gives each account its own field.

**5. Upload client holdings.** The columns you include decide the mode:
- **`Ticker`, `Quantity`** → classic aggregate view (donuts + tables) against the selected mandate.
- **+ `Account_Type`, `Book_Cost_CAD`** → adds realized-gain and estimated-tax on every sell (sells routed registered-first across the book). Book cost is the total CAD adjusted cost base; leave blank for registered accounts and cash.
- **+ `Account`** (and optionally `Account_Type`) → **per-account mode**. Each account is rebalanced **independently to its own mandate**, which you pick in the **Mandate assignment** dropdowns inside the engine — no need to put it in the file. (If you do include a `Mandate_Name` column, it just pre-fills those dropdowns.) Shown as stacked per-account blocks with a consolidated tax summary, plus a toggle to instead treat the whole household as one mandate.

Cash is `CASH_USD` / `CASH_CAD` with the dollar amount as quantity.

---

**How the tax math works**
- **Average-cost ACB** — Canada pools identical property; the gain is measured against the average cost, not US-style lots.
- **Registered-first** (aggregate mode only) — sells come from registered accounts first, then taxable, losses before gains.
- **Per-account mode** — each account self-corrects to its own target, so routing doesn't apply; each taxable account's sells realize gains against its own ACB, and the household total nets them.
- **Inclusion** — taxable gain = realized gain × inclusion rate (50% default); tax = taxable gain × marginal rate.

*No management-fee logic anywhere.*
""",
    },
    "fr": {
        "header_title": "Moteur d'implémentation de portefeuille",
        "header_desc": "Dérive, transactions fiscalement avisées et mandats par compte",
        "sidebar_config": "Configuration du système",
        "mandate_spec": "1. Spécification du mandat",
        "upload_mandate": "Téléverser le mandat (CSV ou Excel)",
        "active_db": "Fichier de base de données actif",
        "target_mandate": "Mandat par défaut / de repli",
        "tactical_cash": "2. Flux de trésorerie tactique",
        "deposit_withdraw": "Dépôt (+) / Retrait (−)",
        "tax_assumptions": "3. Hypothèses fiscales",
        "tax_help": "Utilisé si les positions contiennent Account_Type et Book_Cost_CAD.",
        "marginal_rate": "Taux marginal du client, revenu ordinaire (%)",
        "inclusion_rate": "Taux d'inclusion des gains en capital (%)",
        "main_title": "Implémentation et génération d'ordres",
        "main_tracking": "Suivi actuel",
        "client_holdings": "Positions du client",
        "upload_holdings": "Téléverser les positions (CSV ou Excel)",
        "empty_title": "Téléversez les positions pour commencer",
        "waiting_holdings": "Ajoutez un fichier avec les colonnes Ticker et Quantité.",
        "demo_banner": "Affichage d'un portefeuille exemple pour explorer le moteur. Téléversez vos propres positions ci-dessus pour le remplacer.",
        "curr_val": "Valeur actuelle",
        "household_value": "Valeur du ménage",
        "cash_flow": "Flux de trésorerie",
        "cash_equiv": "Trésorerie et équivalents",
        "post_cash": "Encaisse résiduelle est.",
        "alloc_current": "Allocation actuelle",
        "alloc_target": "Allocation cible",
        "legend_title": "Détail de la répartition et légende",
        "analysis_title": "Analyse des positions et de la dérive",
        "orders_title": "Ordres d'exécution",
        "taximpact_title": "Impact fiscal de ce rééquilibrage",
        "taximpact_hh": "Impact fiscal consolidé (ménage)",
        "accounts_title": "Positions par compte",
        "per_acct_cash": "Flux de trésorerie — par compte",
        "per_acct_blocks": "Rééquilibrage par compte",
        "combined_ticket": "Bordereau consolidé du ménage",
        "approach_label": "Approche de rééquilibrage",
        "approach_per": "Chaque compte → son propre mandat",
        "approach_whole": "Ménage entier → un seul mandat (optimisé fiscalement)",
        "assign_title": "Attribution des mandats",
        "view_account": "Choisir le compte à afficher en détail",
        "cash_on_main": "Les entrées de trésorerie par compte sont sur la page principale →",
        "status_ok": "Cible atteinte",
        "status_breach": "Action requise",
        "no_trades": "Dans les marges — aucune transaction requise.",
        "export_btn": "Exporter le bordereau (CSV)",
        "export_combined": "Exporter le bordereau consolidé (CSV)",
        "footer": "Réalisé par Samari Gbian",
        "fetching_prices": "Récupération des données de marché et des taux de change…",
        "view_mandate": "Voir les cibles du mandat",
        "howto_title": "Comment utiliser ce moteur",
        "col_ticker": "Symbole", "col_asset": "Classe d'actif", "col_price": "Prix",
        "col_current": "Actuel", "col_target": "Cible", "col_drift": "Dérive",
        "col_status": "Statut", "col_action": "Ordre", "col_shares": "Actions",
        "col_est_value": "Valeur est.", "col_gain": "Gain réalisé (CAD)", "col_tax": "Impôt est. (CAD)",
        "col_account": "Compte", "col_type": "Type", "col_acb": "PBR / unité (CAD)",
        "col_mv": "Valeur marchande", "col_unreal": "Gain non réalisé (CAD)", "col_mandate": "Mandat",
        "m_net_gain": "Gain réalisé net (CAD)", "m_est_tax": "Impôt estimé (CAD)",
        "m_eff_rate": "Taux d'imposition effectif", "m_accounts": "Comptes",
        "lbl_gain": "Gain réalisé", "lbl_tax": "Impôt est.", "lbl_value": "Valeur du compte",
        "tax_caption": "Les ventes sont prélevées d'abord dans les comptes enregistrés (sans impôt), puis imposables "
                       "(pertes avant gains). Les pertes compensent les gains. Calcul en CAD quelle que soit la devise.",
        "tax_caption_pa": "Chaque compte est rééquilibré indépendamment selon son mandat; les gains sont calculés sur les "
                          "ventes imposables de ce compte selon son PBR moyen. Le total consolidé compense pertes et gains "
                          "entre comptes (un seul contribuable).",
        "tax_hint": "Ajoutez Account, Account_Type et Book_Cost_CAD pour l'analyse des gains et de l'impôt. "
                    "Ajoutez Mandate_Name pour rééquilibrer chaque compte selon son propre mandat.",
        "no_priced": "Aucun prix en direct n'a pu être récupéré. Vérifiez vos symboles.",
        "no_priced_acct": "{acct} : aucun prix pour ses positions — ignoré.",
        "missing_note": "Aucun prix pour : {items}. Ces positions sont exclues; pondérations sur les titres évalués.",
        "offmandate_note": "Détenu mais absent du mandat (cible 0 % → vente totale) : {items}.",
        "targetsum_note": "Les pondérations cibles de {name} totalisent {pct:.1f} %, et non 100 %.",
        "mandate_missing": "Compte {acct} : mandat « {name} » introuvable; repli sur « {fb} ».",
        "howto_text": """
**1. Configuration (barre latérale).** Langue, **devise de base** (valeurs globales seulement — les prix restent natifs), et sous **Hypothèses fiscales** le taux marginal et le taux d'inclusion.

**2. Téléversez une base de mandats.** CSV/Excel avec `Mandate_Name`, `Ticker`, `Asset_Class`, `Target_Weight_Percent` (pourcentage `20` ou fraction `0.20`), `Drift_Band_Percent`. Un fichier peut contenir plusieurs mandats.

**3. Choisissez le mandat par défaut / de repli.**

**4. Flux de trésorerie.** Les modes simples utilisent la barre latérale; le mode par compte donne un champ par compte.

**5. Téléversez les positions.** Les colonnes incluses déterminent le mode :
- **`Ticker`, `Quantité`** → vue agrégée classique (anneaux + tableaux).
- **+ `Account_Type`, `Book_Cost_CAD`** → gain réalisé et impôt estimé sur chaque vente (ventes prélevées comptes enregistrés d'abord).
- **+ `Account`** (et facultativement `Account_Type`) → **mode par compte**. Chaque compte est rééquilibré **indépendamment selon son mandat**, que vous choisissez dans les listes **Attribution des mandats** du moteur — inutile de l'inscrire dans le fichier. (Une colonne `Mandate_Name`, si présente, ne fait que préremplir ces listes.) En blocs empilés avec un résumé fiscal consolidé, plus une bascule pour traiter tout le ménage selon un seul mandat.

Trésorerie : `CASH_USD` / `CASH_CAD`.

---

**Calcul fiscal**
- **PBR moyen** — le Canada regroupe les biens identiques; gain mesuré sur le coût moyen.
- **Comptes enregistrés d'abord** (mode agrégé) — ventes des comptes enregistrés en premier, pertes avant gains.
- **Mode par compte** — chaque compte se corrige seul; chaque compte imposable réalise ses gains selon son PBR, le ménage les compense.
- **Inclusion** — gain imposable = gain réalisé × taux d'inclusion (50 %); impôt = gain imposable × taux marginal.

*Aucune logique de frais de gestion.*
""",
    },
}


# --------------------------------------------------------------------------- #
# CSS
# --------------------------------------------------------------------------- #
custom_css = f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Lora:wght@500;600;700&display=swap');
    #MainMenu {{visibility: hidden;}} footer {{visibility: hidden;}}
    [data-testid="stHeader"] {{background: transparent !important;}}
    .block-container {{ padding-top: 1.6rem !important; padding-bottom: 2rem !important; max-width: 1240px; }}
    html, body, [class*="css"] {{ font-family: 'Inter', system-ui, sans-serif !important; color: {INK}; }}
    .stApp {{ background: {CANVAS}; }}
    h1,h2,h3,h4 {{ font-family: 'Lora', Georgia, serif !important; font-weight: 600 !important; letter-spacing: -0.01em !important; color: {INK} !important; }}

    /* Sidebar — deep navy */
    [data-testid="stSidebar"] {{ background: {NAVY} !important; border-right: 1px solid {NAVY_LINE} !important; }}
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {{ color:#EAF0F6 !important; }}
    [data-testid="stSidebar"] .stMarkdown, [data-testid="stSidebar"] .stMarkdown p {{ color:#C6D3E0 !important; }}
    [data-testid="stSidebar"] .stMarkdown strong {{ color:#EAF0F6 !important; }}
    [data-testid="stSidebar"] label {{ color:#A9BACB !important; font-size:0.8rem !important; font-weight:500 !important; }}
    [data-testid="stSidebar"] [data-testid="stCaptionContainer"] {{ color:#92A6BC !important; }}
    [data-testid="stSidebar"] hr {{ border-color:{NAVY_LINE} !important; }}
    [data-testid="stSidebar"] [data-testid="stExpander"] {{ border:1px solid {NAVY_LINE} !important; border-radius:8px !important; background:{NAVY_SOFT} !important; }}
    [data-testid="stSidebar"] [data-testid="stExpander"] summary {{ color:#C6D3E0 !important; }}
    [data-testid="stSidebar"] [data-testid="stExpander"] p {{ color:#C6D3E0 !important; }}

    /* Inputs (sidebar + main) — light fields, brass focus */
    div[data-baseweb="select"] > div {{ background:{SURFACE} !important; border:1px solid #D7D2C7 !important; border-radius:8px !important; min-height:40px !important; color:{INK} !important; }}
    div[data-baseweb="select"] > div:hover {{ border-color:{BRAND} !important; }}
    div[data-baseweb="select"] > div:focus-within {{ border-color:{BRAND} !important; box-shadow:0 0 0 3px rgba(176,137,79,0.18) !important; }}
    div[data-baseweb="select"] div {{ color:{INK}; }}
    div[data-baseweb="input"], div[data-baseweb="base-input"] {{ background:{SURFACE} !important; border:1px solid #D7D2C7 !important; border-radius:8px !important; }}
    div[data-baseweb="input"]:focus-within {{ border-color:{BRAND} !important; box-shadow:0 0 0 3px rgba(176,137,79,0.18) !important; }}
    [data-testid="stNumberInput"] input {{ background:transparent !important; color:{INK} !important; }}
    [data-testid="stNumberInput"] button {{ background:#F4F2EC !important; border:1px solid #D7D2C7 !important; color:{INK} !important; }}
    .stSelectbox label, .stNumberInput label, .stRadio label {{ color:{MUTED}; }}
    .stRadio [role="radiogroup"] label {{ color:{INK} !important; }}

    /* Buttons — brass */
    [data-testid="stDownloadButton"] > button {{ background:{BRAND} !important; color:#fff !important; border:none !important; border-radius:8px !important; padding:10px 18px !important; font-weight:600 !important; width:100% !important; transition:all 0.18s ease !important; box-shadow:0 1px 2px rgba(16,42,67,0.18); }}
    [data-testid="stDownloadButton"] > button:hover {{ background:{BRAND_HOVER} !important; transform:translateY(-1px); }}
    .stButton > button {{ background:{BRAND} !important; color:#fff !important; border:none !important; border-radius:8px !important; font-weight:600 !important; }}

    /* File uploader — single, prominent, on-brand dropzone */
    [data-testid="stFileUploaderDropzone"] {{ background:{SURFACE} !important; border:1.5px dashed #C9C2B2 !important; border-radius:14px !important; padding:30px 24px !important; box-shadow:0 1px 2px rgba(16,42,67,0.04); transition:border-color 0.18s ease; }}
    [data-testid="stFileUploaderDropzone"]:hover {{ border-color:{BRAND} !important; }}
    /* green upload arrow — only the instructions icon, not the file chip */
    [data-testid="stFileUploaderDropzoneInstructions"] svg {{ width:38px !important; height:38px !important; color:{POS} !important; fill:{POS} !important; }}
    /* brass Browse button only (secondary button), never the delete button */
    [data-testid="stFileUploaderDropzone"] button[kind="secondary"],
    [data-testid="stFileUploaderDropzone"] [data-testid="stBaseButton-secondary"],
    [data-testid="stFileUploaderDropzone"] [data-testid="baseButton-secondary"] {{ background:{BRAND} !important; color:#fff !important; border:none !important; border-radius:8px !important; font-weight:600 !important; }}
    /* uploaded-file chip + delete (x) stay quiet and neutral */
    [data-testid="stFileUploaderFile"] svg, [data-testid="stFileUploaderFileData"] svg {{ color:{MUTED} !important; fill:none !important; }}
    [data-testid="stFileUploaderDeleteBtn"], [data-testid="stFileUploaderDeleteBtn"] button {{ background:transparent !important; box-shadow:none !important; border:none !important; }}
    [data-testid="stFileUploaderDeleteBtn"] svg {{ color:{MUTED} !important; fill:none !important; width:18px !important; height:18px !important; }}

    /* Brand header — navy bar */
    .brand {{ display:flex; align-items:center; gap:14px; background:{NAVY}; padding:18px 22px; border-radius:14px; margin-bottom:22px; box-shadow:0 6px 20px rgba(16,42,67,0.18); }}
    .brand-mark {{ width:44px; height:44px; border-radius:11px; background:rgba(176,137,79,0.20); color:{BRAND}; display:flex; align-items:center; justify-content:center; flex:none; border:1px solid rgba(176,137,79,0.45); }}
    .brand-title {{ font-family:'Lora', Georgia, serif; font-size:1.45rem; font-weight:600; color:#F4F6F9; letter-spacing:-0.01em; line-height:1.15; }}
    .brand-sub {{ font-size:0.84rem; color:#9FB2C6; margin-top:3px; }}

    .sec-head {{ display:flex; align-items:center; gap:9px; margin:4px 0 13px; }}
    .sec-head .ico {{ color:{BRAND}; display:inline-flex; }}
    .sec-head .txt {{ font-weight:600; font-size:1.02rem; color:{INK}; letter-spacing:-0.01em; }}

    /* Metric cards */
    .metric-row {{ display:flex; gap:14px; margin:8px 0 24px; flex-wrap:nowrap; }}
    .metric-card {{ flex:1; min-width:0; background:{SURFACE}; border:1px solid {BORDER}; border-radius:12px; padding:16px 18px; display:flex; align-items:center; gap:14px; box-shadow:0 2px 8px rgba(16,42,67,0.06); }}
    .metric-card .ico {{ flex:none; width:38px; height:38px; border-radius:9px; background:#F3ECDD; color:{BRAND}; display:flex; align-items:center; justify-content:center; }}
    .metric-label {{ font-size:0.78rem; color:{MUTED}; font-weight:500; margin-bottom:3px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
    .metric-value {{ font-size:1.45rem; font-weight:600; color:{INK}; font-variant-numeric:tabular-nums; white-space:nowrap; letter-spacing:-0.01em; }}
    @media (max-width:820px) {{ .metric-row {{ flex-wrap:wrap; }} .metric-card {{ min-width:45%; }} }}

    .notice {{ padding:10px 14px; border:1px solid; border-radius:10px; font-size:0.84rem; margin:4px 0 16px; line-height:1.45; }}

    /* Tables */
    .table-wrap {{ border:1px solid {BORDER}; border-radius:12px; overflow:hidden; background:{SURFACE}; box-shadow:0 2px 8px rgba(16,42,67,0.05); }}
    .pf-table {{ width:100%; border-collapse:collapse; font-size:0.86rem; table-layout:auto; }}
    .pf-table thead th {{ background:#FAF8F2; color:{MUTED}; font-weight:600; font-size:0.72rem; text-transform:uppercase; letter-spacing:0.03em; padding:11px 14px; border-bottom:1px solid {BORDER}; white-space:nowrap; }}
    .pf-table tbody td {{ padding:11px 14px; border-bottom:1px solid #F1EFE8; color:{INK}; font-variant-numeric:tabular-nums; }}
    .pf-table tbody tr:last-child td {{ border-bottom:none; }}
    .pf-table tbody tr:hover td {{ background:#FBFAF6; }}
    .ticker-cell {{ font-weight:600; }} .muted-cell {{ color:{MUTED}; }}

    .pill {{ display:inline-block; padding:2px 10px; border-radius:999px; font-size:0.72rem; font-weight:600; white-space:nowrap; }}
    .pill-buy {{ background:#E6F1EA; color:{POS}; }} .pill-sell {{ background:#F7E8E8; color:{NEG}; }}
    .pill-action {{ background:#F7E8E8; color:{NEG}; }} .pill-ok {{ background:#ECEAE3; color:{MUTED}; }}
    .pill-reg {{ background:#E7EEF6; color:#2C5282; }} .pill-tax {{ background:#F3EBDA; color:{AMBER}; }}

    .legend-table {{ width:100%; border-collapse:collapse; font-size:0.85rem; }}
    .legend-table th {{ text-align:left; color:{MUTED}; font-weight:500; padding:7px 8px; border-bottom:1px solid {BORDER}; }}
    .legend-table th.num, .legend-table td.num {{ text-align:right; font-variant-numeric:tabular-nums; }}
    .legend-table td {{ padding:8px; border-bottom:1px solid #F1EFE8; color:{INK}; }}
    .legend-table tr:last-child td {{ border-bottom:none; }}
    .dot {{ display:inline-block; width:10px; height:10px; border-radius:50%; margin-right:9px; vertical-align:middle; }}
    .tax-cap {{ color:{MUTED}; font-size:0.78rem; margin-top:8px; line-height:1.45; }}

    .acct-head {{ display:flex; align-items:center; justify-content:space-between; gap:12px; flex-wrap:wrap; background:#FAF8F2; border:1px solid {BORDER}; border-radius:12px; padding:13px 16px; margin:6px 0 12px; }}
    .acct-head .left {{ display:flex; align-items:center; gap:10px; }}
    .acct-name {{ font-family:'Lora', Georgia, serif; font-weight:600; font-size:1.05rem; color:{INK}; }}
    .acct-meta {{ font-size:0.82rem; color:{MUTED}; font-variant-numeric:tabular-nums; }}
    .acct-tax {{ font-size:0.84rem; color:{INK}; font-variant-numeric:tabular-nums; margin:8px 0 2px; }}
    .acct-tax b {{ font-weight:600; }}
    hr.block-div {{ border:none; border-top:1px solid {BORDER}; margin:26px 0 18px; }}

    .custom-footer {{ margin-top:54px; padding-top:20px; border-top:1px solid {BORDER}; font-size:0.8rem; color:#9AA7B4; text-align:center; }}
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
# SMALL HELPERS
# --------------------------------------------------------------------------- #
def section_header(icon_key, text):
    st.markdown(f'<div class="sec-head"><span class="ico">{ICONS[icon_key]}</span>'
                f'<span class="txt">{text}</span></div>', unsafe_allow_html=True)

def notice(text, kind="warn"):
    pal = {"warn": ("#FEF6E7", "#92610A", "#F5D58A"), "info": ("#EEF4FF", "#1E40AF", "#C3D4FB")}
    bg, fg, bd = pal[kind]
    st.markdown(f'<div class="notice" style="background:{bg};color:{fg};border-color:{bd}">{text}</div>',
                unsafe_allow_html=True)

def pill(text, kind):
    return f'<span class="pill pill-{kind}">{text}</span>'

def html_table(headers, rows, aligns):
    head = "".join(f'<th class="{"num" if a == "right" else ""}">{h}</th>' for h, a in zip(headers, aligns))
    body = "".join("<tr>" + "".join(f'<td style="text-align:{a}">{c}</td>' for c, a in zip(r, aligns)) + "</tr>" for r in rows)
    return f'<div class="table-wrap"><table class="pf-table"><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>'

def ccy_symbol(ccy):
    return CCY_SYMBOLS.get(str(ccy).upper(), f"{ccy} ")

def signed_cad(v):
    return f'{"−" if v < 0 else ""}C${abs(v):,.0f}'

def render_instructions(expanded):
    with st.expander(t["howto_title"], expanded=expanded):
        st.markdown(t["howto_text"])

def render_footer():
    st.markdown(f'<div class="custom-footer">{t["footer"]}</div>', unsafe_allow_html=True)

def is_cash_ticker(tk):
    return "CASH" in tk or tk == "NBC200"

def is_taxable_account(atype):
    a = "".join(ch for ch in str(atype).upper() if ch.isalnum())
    return a not in REGISTERED_TYPES

def build_color_map(classes):
    cmap, i = {}, 0
    for c in sorted(classes, key=lambda x: str(x)):
        cl = str(c).strip().lower()
        if "cash" in cl:
            cmap[c] = CASH_COLOR
        elif cl in ("unclassified", "nan", ""):
            cmap[c] = UNCLASS_COLOR
        else:
            cmap[c] = CATEGORY_PALETTE[i % len(CATEGORY_PALETTE)]; i += 1
    return cmap

def make_donut(df, value_col, color_map):
    df = df.copy()
    total = df[value_col].sum()
    df["pct"] = (df[value_col] / total * 100.0) if total else 0.0
    df["lbl"] = df["pct"].apply(lambda p: f"{p:.1f}%" if p >= 6 else "")
    light = {CASH_COLOR, UNCLASS_COLOR, "#F59E0B", "#84CC16"}
    label_colors = [INK if color_map.get(c, CASH_COLOR) in light else "#FFFFFF" for c in df["Asset_Class"]]
    fig = px.pie(df, values=value_col, names="Asset_Class", hole=0.62, color="Asset_Class", color_discrete_map=color_map)
    fig.update_traces(text=df["lbl"], textinfo="text", textposition="inside", insidetextorientation="horizontal",
                      sort=False, marker=dict(line=dict(color="#FFFFFF", width=2)),
                      hovertemplate="%{label}: %{percent}<extra></extra>", textfont=dict(color=label_colors, size=13))
    fig.update_layout(showlegend=False, height=292, margin=dict(t=6, b=6, l=6, r=6),
                      paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                      font=dict(family="Inter, sans-serif", size=13, color=INK))
    return fig


# --------------------------------------------------------------------------- #
# DATA HELPERS
# --------------------------------------------------------------------------- #
@st.cache_data(ttl=PRICE_CACHE_TTL, show_spinner=False)
def fetch_live_price(ticker):
    ccy = None
    try:
        o = yf.Ticker(ticker)
        price = None
        try:
            fi = o.fast_info
            price = fi.get("last_price"); ccy = fi.get("currency")
        except Exception:
            pass
        if not (price and price > 0):
            try:
                info = o.info or {}
                ccy = ccy or info.get("currency")
                for k in ("regularMarketPrice", "currentPrice", "previousClose"):
                    if info.get(k) and info.get(k) > 0:
                        price = float(info[k]); break
            except Exception:
                pass
        if not (price and price > 0):
            hist = o.history(period="5d")
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
    try:
        cad_usd = yf.Ticker("CADUSD=X").fast_info.get("last_price") or 0.73
        usd_cad = yf.Ticker("USDCAD=X").fast_info.get("last_price") or 1.37
        return float(cad_usd), float(usd_cad)
    except Exception:
        return 0.73, 1.37


# --------------------------------------------------------------------------- #
# SCANNER
# --------------------------------------------------------------------------- #
def _norm(x):
    """Reduce a label to letters/digits only: kills case, spaces, punctuation,
    accents, and invisible characters (NBSP, zero-width) so matching is reliable."""
    s = unicodedata.normalize("NFKD", str(x)).lower()
    return "".join(ch for ch in s if ch.isalnum())


def read_tabular(path_or_file, expected_type="mandate"):
    name = getattr(path_or_file, "name", str(path_or_file)).lower()
    raw_dfs = (pd.read_excel(path_or_file, header=None, sheet_name=None)
               if name.endswith((".xlsx", ".xls")) else {"Sheet1": pd.read_csv(path_or_file, header=None)})
    valid = []
    for sheet, raw in raw_dfs.items():
        if raw.empty:
            continue
        first = str(raw.iloc[0, 0]).strip()
        fallback = first if (len(first) > 3 and first != "nan") else sheet
        # --- robust header detection + column mapping -------------------------
        # Each field: (canonical, exact normalized aliases, substring tokens).
        # Order = priority, so specific fields claim a column before generic ones.
        FIELD_ALIASES = [
            ("Account_Type", {"accounttype", "registration", "registrationtype", "plantype",
                              "typecompte", "accountcategory"}, ("accounttype", "registration", "plantype")),
            ("Book_Cost_CAD", {"bookcost", "bookcostcad", "acb", "acbcad", "acbtotal", "adjustedcostbase",
                               "adjustedcostbasecad", "costbasis", "bookvalue", "prixderevient", "coutdebase"},
             ("bookcost", "costbasis", "adjustedcost", "acb")),
            ("Drift_Band_Percent", {"drift", "driftband", "driftbandpercent", "tolerance", "band",
                                    "driftpercent", "derive", "marge"}, ("drift", "tolerance")),
            ("Target_Weight_Percent", {"target", "targetweight", "weight", "targetweightpercent",
                                       "targetpercent", "allocation", "cible", "poids"},
             ("targetweight", "target", "weight", "allocation")),
            ("Mandate_Name", {"mandate", "mandatename", "strategy", "model", "mandat"}, ("mandate", "strategy")),
            ("Asset_Class", {"assetclass", "class", "sector", "category", "categorie", "classe", "assettype"},
             ("assetclass", "sector")),
            ("Quantity", {"quantity", "qty", "shares", "units", "position", "amount", "quantite",
                          "currentquantityunits", "sharecount", "numshares", "unitsheld"},
             ("quantity", "shares", "units")),
            ("Ticker", {"ticker", "symbol", "sym", "asset", "symbole", "tickeridentifier", "cusip",
                        "isin", "identifier"}, ("ticker", "symbol", "cusip", "isin")),
            ("Account", {"account", "accountname", "accountid", "acct", "compte", "accountnumber",
                         "accountno"}, ("account",)),
        ]
        exact_lookup = {a: canon for canon, exacts, _ in FIELD_ALIASES for a in exacts}

        def _is_label(v):
            n = _norm(v)
            return (n in exact_lookup) or any(tok in n for _, _, toks in FIELD_ALIASES for tok in toks)

        # Header = first row carrying >= 2 recognizable column labels (skips titles).
        h = next((i for i, row in raw.iterrows()
                  if sum(_is_label(v) for v in row.values) >= 2), -1)
        if h == -1:
            continue
        df = raw.iloc[h + 1:].copy()
        df.columns = raw.iloc[h].values
        df = df.dropna(how="all")

        norms = {col: _norm(col) for col in df.columns}
        rmap, claimed = {}, set()
        for col in df.columns:                       # exact match wins
            c = exact_lookup.get(norms[col])
            if c and c not in claimed:
                rmap[col] = c; claimed.add(c)
        for col in df.columns:                       # then substring, by priority
            if col in rmap:
                continue
            for canon, _, toks in FIELD_ALIASES:
                if canon not in claimed and any(tok in norms[col] for tok in toks):
                    rmap[col] = canon; claimed.add(canon); break
        df = df.rename(columns=rmap)
        if expected_type == "mandate" and "Mandate_Name" not in df.columns:
            df["Mandate_Name"] = fallback
        if "Ticker" in df.columns:
            df = df[~df["Ticker"].isna()]
            df = df[~df["Ticker"].astype(str).str.upper().str.contains("TOTAL")]
            df = df[~df["Ticker"].astype(str).str.startswith("*")]
        for col in ("Target_Weight_Percent", "Drift_Band_Percent", "Quantity", "Book_Cost_CAD"):
            if col in df.columns:
                df[col] = (df[col].astype(str).str.replace("%", "").str.replace(",", ".")
                           .str.replace(" ", "").str.replace("\xa0", "").str.strip())
                df[col] = pd.to_numeric(df[col], errors="coerce")
                if col != "Book_Cost_CAD":
                    df[col] = df[col].fillna(0.0)
        required = (["Mandate_Name", "Ticker", "Target_Weight_Percent"]
                    if expected_type == "mandate" else ["Ticker", "Quantity"])
        if all(c in df.columns for c in required):
            valid.append(df)
    if not valid:
        need = (["Mandate_Name", "Ticker", "Target_Weight_Percent"]
                if expected_type == "mandate" else ["Ticker", "Quantity"])
        st.error(f"Could not map columns in your file. Expected variations of: {', '.join(need)}.")
        st.stop()
    return pd.concat(valid, ignore_index=True)


# --------------------------------------------------------------------------- #
# BRAND HEADER + SIDEBAR
# --------------------------------------------------------------------------- #
st.markdown(f"""<div class="brand"><div class="brand-mark">{ICONS['sliders']}</div>
    <div><div class="brand-title">{t['header_title']}</div>
    <div class="brand-sub">{t['header_desc']}</div></div></div>""", unsafe_allow_html=True)

col_lang, col_curr = st.sidebar.columns(2)
new_lang = col_lang.selectbox("Language / Langue", ["English", "Français"],
                              index=0 if st.session_state.lang == "en" else 1)
st.session_state.lang = "en" if new_lang == "English" else "fr"
t = LANG[st.session_state.lang]
st.session_state.currency = col_curr.selectbox("Base currency", ["USD", "CAD"],
                                               index=0 if st.session_state.currency == "USD" else 1)

st.sidebar.markdown("---")
st.sidebar.subheader(t["sidebar_config"])
st.sidebar.write("")
st.sidebar.markdown(f"**{t['mandate_spec']}**")
uploaded_mandate = st.sidebar.file_uploader(t["upload_mandate"], type=["csv", "xlsx", "xls"])
if uploaded_mandate is not None:
    with open(os.path.join(SAVED_MANDATES_DIR, uploaded_mandate.name), "wb") as f:
        f.write(uploaded_mandate.getbuffer())

saved_mandates = sorted(f for f in os.listdir(SAVED_MANDATES_DIR)
                        if f.lower().endswith((".csv", ".xlsx", ".xls"))
                        and "holdings" not in f.lower() and "portfolio" not in f.lower())
if not saved_mandates:
    st.sidebar.error("No mandates found. Please upload one.")
    render_instructions(expanded=True); render_footer(); st.stop()

selected_file = st.sidebar.selectbox(t["active_db"], saved_mandates)
mandate_df = read_tabular(os.path.join(SAVED_MANDATES_DIR, selected_file), "mandate")
mandate_names = sorted(mandate_df["Mandate_Name"].dropna().unique())
selected_mandate = st.sidebar.selectbox(t["target_mandate"], mandate_names)


def prepare_mandate(name):
    """Filter, clean, default, and normalize a mandate (fraction weights -> percent)."""
    m = mandate_df[mandate_df["Mandate_Name"] == name].copy()
    m["Ticker"] = m["Ticker"].astype(str).str.strip().str.upper()
    m["Target_Weight_Percent"] = pd.to_numeric(m["Target_Weight_Percent"], errors="coerce").fillna(0.0)
    if "Drift_Band_Percent" not in m.columns:
        m["Drift_Band_Percent"] = 0.0
    else:
        m["Drift_Band_Percent"] = pd.to_numeric(m["Drift_Band_Percent"], errors="coerce").fillna(0.0)
    if "Asset_Class" not in m.columns:
        m["Asset_Class"] = "Unclassified"
    s = float(m["Target_Weight_Percent"].sum())
    if 0 < s <= 1.5:
        m["Target_Weight_Percent"] *= 100.0
        m["Drift_Band_Percent"] *= 100.0
    return m


with st.sidebar.expander(t["view_mandate"]):
    _m = prepare_mandate(selected_mandate)
    cols = [c for c in ["Ticker", "Asset_Class", "Target_Weight_Percent"] if c in _m.columns]
    st.dataframe(_m[cols], hide_index=True, use_container_width=True,
                 column_config={"Ticker": st.column_config.TextColumn(t["col_ticker"]),
                                "Asset_Class": st.column_config.TextColumn(t["col_asset"]),
                                "Target_Weight_Percent": st.column_config.NumberColumn(t["col_target"], format="%.2f%%")})

st.sidebar.write("")
st.sidebar.markdown(f"**{t['tactical_cash']}**")
# In per-account mode the cash is entered per account on the main page, so the
# single sidebar input would be redundant — hide it then.
_cached = st.session_state.get("cached_holdings")
_acct_cols = (_cached is not None) and any(c in _cached.columns for c in ("Account", "Account_Type"))
_per_labels = (LANG["en"]["approach_per"], LANG["fr"]["approach_per"])
_per_active = _acct_cols and (st.session_state.get("approach", LANG["en"]["approach_per"]) in _per_labels)
if _per_active:
    st.sidebar.caption(t["cash_on_main"])
    cash_flow_global = 0.0
else:
    cash_flow_global = st.sidebar.number_input(f"{t['deposit_withdraw']} ({st.session_state.currency})",
                                               value=0.0, step=1000.0, format="%.2f")
st.sidebar.write("")
st.sidebar.markdown(f"**{t['tax_assumptions']}**")
with st.sidebar.expander(t["tax_assumptions"], expanded=False):
    st.caption(t["tax_help"])
    marginal_rate = st.number_input(t["marginal_rate"], value=53.53, min_value=0.0, max_value=100.0, step=0.5)
    inclusion_rate = st.number_input(t["inclusion_rate"], value=50.0, min_value=0.0, max_value=100.0, step=5.0)
incl, marg = inclusion_rate / 100.0, marginal_rate / 100.0


# --------------------------------------------------------------------------- #
# MAIN — TITLE + HOLDINGS UPLOAD
# --------------------------------------------------------------------------- #
st.subheader(t["main_title"])
st.caption(f"{t['main_tracking']}: {selected_mandate}")
st.write("")
st.markdown(f"**{t['client_holdings']}**")
holdings_file = st.file_uploader(t["upload_holdings"], type=["csv", "xlsx", "xls"], key="client_file")
if holdings_file is not None:
    st.session_state.cached_holdings = read_tabular(holdings_file, "holdings")
    st.session_state.using_demo = False
elif "cached_holdings" not in st.session_state:
    # First visit, nothing uploaded — auto-load the bundled demo portfolio so the
    # engine shows a live analysis immediately. An upload replaces it.
    if os.path.exists(DEMO_HOLDINGS):
        st.session_state.cached_holdings = read_tabular(DEMO_HOLDINGS, "holdings")
        st.session_state.using_demo = True
    else:
        st.session_state.cached_holdings = None

if st.session_state.get("cached_holdings") is None:
    st.caption(t["waiting_holdings"])
    render_instructions(expanded=True); render_footer(); st.stop()

if st.session_state.get("using_demo"):
    notice(t["demo_banner"], "info")

holdings_df = st.session_state.cached_holdings.copy()
holdings_df["Ticker"] = holdings_df["Ticker"].astype(str).str.strip().str.upper()
holdings_df["Quantity"] = pd.to_numeric(holdings_df["Quantity"], errors="coerce").fillna(0.0)
if "Book_Cost_CAD" in holdings_df.columns:
    holdings_df["Book_Cost_CAD"] = pd.to_numeric(holdings_df["Book_Cost_CAD"], errors="coerce")

has_account = "Account" in holdings_df.columns
has_type = "Account_Type" in holdings_df.columns
has_acb = "Book_Cost_CAD" in holdings_df.columns
has_mandate_col = "Mandate_Name" in holdings_df.columns
tax_capable = has_type and has_acb
# Account mode is triggered by an Account (or Account_Type) column. Mandates are
# assigned in the UI; a Mandate_Name column, if present, just pre-fills them.
account_col = "Account" if has_account else ("Account_Type" if has_type else None)
account_mode = account_col is not None


# --------------------------------------------------------------------------- #
# PRICING (once). Mandates are chosen in the UI, so price every ticker that any
# mandate could reference, plus everything held.
# --------------------------------------------------------------------------- #
mandate_tickers = set(str(x).strip().upper() for x in mandate_df["Ticker"].dropna())
universe = sorted(set(holdings_df["Ticker"]) | mandate_tickers)

cad_usd, usd_cad = fetch_fx_rates()
base = st.session_state.currency

def to_base_factor(ccy):
    n = str(ccy).upper()
    if n == base: return 1.0
    if n == "USD" and base == "CAD": return usd_cad
    if n == "CAD" and base == "USD": return cad_usd
    return 1.0

def to_cad_factor(ccy):
    n = str(ccy).upper()
    return usd_cad if n == "USD" else 1.0

native_price, native_ccy, base_price_map, cad_price_map, missing_prices = {}, {}, {}, {}, []
with st.spinner(t["fetching_prices"]):
    for tk in universe:
        if is_cash_ticker(tk):
            nccy = "USD" if "USD" in tk else ("CAD" if "CAD" in tk else base)
            np_ = 1.0
        else:
            np_, ccy = fetch_live_price(tk)
            if np_ is None:
                missing_prices.append(tk); continue
            nccy = ccy or ("CAD" if tk.endswith((".TO", ".V", ".CN", ".NE")) else "USD")
        native_price[tk] = np_; native_ccy[tk] = nccy
        base_price_map[tk] = np_ * to_base_factor(nccy)
        cad_price_map[tk] = np_ * to_cad_factor(nccy)

curr_sym = "$" if base == "USD" else "C$"
if missing_prices:
    notice(t["missing_note"].format(items=", ".join(missing_prices)), "warn")


# --------------------------------------------------------------------------- #
# CORE: compute one book (a set of holdings vs one mandate)
# --------------------------------------------------------------------------- #
def build_ledger(sub_df):
    lm = {}
    if "Account_Type" not in sub_df.columns:
        return lm
    for _, row in sub_df.iterrows():
        qty = float(row["Quantity"])
        if qty == 0:
            continue
        bc = row.get("Book_Cost_CAD", None)
        bc = None if (bc is None or pd.isna(bc)) else float(bc)
        lm.setdefault(row["Ticker"], []).append({
            "account": str(row.get("Account", "—")), "type": str(row.get("Account_Type", "")),
            "taxable": is_taxable_account(row.get("Account_Type", "")), "qty": qty,
            "acb_per_unit": (bc / qty) if (bc is not None and qty > 0) else None})
    return lm

def attribute_sell_gain(ledger_map, ticker, n_shares, cad_price):
    accts = ledger_map.get(ticker, [])
    def key(a):
        unreal = (cad_price - a["acb_per_unit"]) if a["acb_per_unit"] is not None else 0.0
        return (1 if a["taxable"] else 0, unreal)
    remaining, realized = float(n_shares), 0.0
    for a in sorted(accts, key=key):
        if remaining <= 0:
            break
        take = min(remaining, a["qty"])
        if take <= 0:
            continue
        if a["taxable"] and a["acb_per_unit"] is not None:
            realized += take * (cad_price - a["acb_per_unit"])
        remaining -= take
    return realized

def compute_book(sub_df, mandate_m, cash_amt, tax_on):
    target_map = mandate_m.set_index("Ticker")[["Asset_Class", "Target_Weight_Percent", "Drift_Band_Percent"]].to_dict("index")
    holdings_map = sub_df.groupby("Ticker")["Quantity"].sum().to_dict()
    class_map = dict(zip(sub_df["Ticker"], sub_df["Asset_Class"])) if "Asset_Class" in sub_df.columns else {}
    ledger_map = build_ledger(sub_df) if tax_on else {}
    uni = sorted(set(holdings_map) | set(target_map))
    priced = [tk for tk in uni if tk in base_price_map]
    if not priced:
        return None

    def res_class(tk):
        if is_cash_ticker(tk):
            return "Cash Equivalent"
        c = target_map.get(tk, {}).get("Asset_Class", None)
        if c and str(c).strip().lower() not in ("", "nan", "unclassified"):
            return c
        h = class_map.get(tk)
        if h and str(h).strip().lower() not in ("", "nan"):
            return h
        return "Unclassified"

    rows = []
    for tk in priced:
        qty = float(holdings_map.get(tk, 0.0))
        bp = base_price_map[tk]
        rows.append({"Ticker": tk, "Native_Price": native_price[tk], "Native_Currency": native_ccy[tk],
                     "Base_Price": bp, "Market_Value": qty * bp})
    a = pd.DataFrame(rows)
    total_value = float(a["Market_Value"].sum())
    total_to_alloc = total_value + cash_amt
    a["Asset_Class"] = a["Ticker"].apply(res_class)
    a["Current_Weight"] = (a["Market_Value"] / total_value * 100.0) if total_value else 0.0
    a["Target_Weight"] = a["Ticker"].apply(lambda tk: target_map.get(tk, {}).get("Target_Weight_Percent", 0.0))
    a["Drift_Band"] = a["Ticker"].apply(lambda tk: target_map.get(tk, {}).get("Drift_Band_Percent", 0.0))
    a["Drift"] = a["Current_Weight"] - a["Target_Weight"]
    a["Breach"] = a["Drift"].abs() > a["Drift_Band"]
    a = a.sort_values("Market_Value", ascending=False).reset_index(drop=True)

    cash_mask = (a["Asset_Class"].str.lower().str.contains("cash", na=False) | a["Ticker"].str.contains("CASH", na=False))
    cash_val = float(a.loc[cash_mask, "Market_Value"].sum())

    # --- initial whole-share orders, tracking each sleeve's post-trade value ---
    orders = {}        # ticker -> order dict
    post_val = {}      # non-cash priceable in-mandate sleeve -> post-trade market value
    tgt_val = {}       # sleeve -> target dollar
    px = {}            # sleeve -> price
    for _, r in a.iterrows():
        tk = r["Ticker"]; p = r["Base_Price"]
        if is_cash_ticker(tk):
            continue
        target_dollar = r["Target_Weight"] / 100.0 * total_to_alloc
        mv_after = r["Market_Value"]
        if r["Breach"]:
            shares = round((target_dollar - r["Market_Value"]) / p) if p else 0
            if shares != 0:
                gain = attribute_sell_gain(ledger_map, tk, abs(shares), cad_price_map[tk]) if (shares < 0 and tax_on) else 0.0
                orders[tk] = {"Ticker": tk, "Action": "BUY" if shares > 0 else "SELL",
                              "shares": int(abs(shares)), "signed": int(shares), "price": p, "gain": gain}
                mv_after = r["Market_Value"] + shares * p
        if r["Target_Weight"] > 0 and p:
            tgt_val[tk] = target_dollar; post_val[tk] = mv_after; px[tk] = p

    # --- residual sweep ---
    # Whole-share rounding and drift bands leave sleeves under target, stranding
    # cash. Deploy that leftover (anything above the mandate's intended cash floor
    # = its cash sleeves + any sum-to-<100% gap) into the most-underweight sleeves,
    # one share at a time, rounding each toward (not past) its target.
    net_flow0 = sum(o["signed"] * o["price"] for o in orders.values())
    proj_cash0 = cash_val + cash_amt - net_flow0
    cash_floor = total_to_alloc - sum(tgt_val.values())
    remaining = min(proj_cash0, proj_cash0 - cash_floor)
    guard = 0
    while remaining > 0 and tgt_val and guard < 300000:
        guard += 1
        cand = [tk for tk in tgt_val if px[tk] <= remaining and (tgt_val[tk] - post_val[tk]) >= 0.5 * px[tk]]
        if not cand:
            break
        tk = min(cand, key=lambda x: post_val[x] / tgt_val[x] if tgt_val[x] else 9e9)
        post_val[tk] += px[tk]; remaining -= px[tk]
        if tk in orders and orders[tk]["Action"] == "BUY":
            orders[tk]["shares"] += 1; orders[tk]["signed"] += 1
        else:
            orders[tk] = {"Ticker": tk, "Action": "BUY", "shares": 1, "signed": 1, "price": px[tk], "gain": 0.0}

    # --- finalize orders ---
    rows = []
    for o in orders.values():
        if o["shares"] == 0:
            continue
        rows.append({"Ticker": o["Ticker"], "Action": o["Action"], "Shares_to_Trade": o["shares"],
                     "Estimated_Dollar_Value": round(o["shares"] * o["price"], 2),
                     "Realized_Gain_CAD": o["gain"], "Est_Tax_CAD": max(o["gain"], 0.0) * incl * marg,
                     "_signed_dollar": o["signed"] * o["price"]})
    od = pd.DataFrame(rows)
    if not od.empty:
        od = od.sort_values(["Action", "Ticker"]).reset_index(drop=True)
    net_flow = float(od["_signed_dollar"].sum()) if not od.empty else 0.0
    net_gain = float(od["Realized_Gain_CAD"].sum()) if not od.empty else 0.0
    off_m = [tk for tk in holdings_map if tk not in target_map and not is_cash_ticker(tk)]
    return {"analysis": a, "orders": od, "total_value": total_value,
            "cash_val": cash_val, "cash_pct": (cash_val / total_value * 100.0) if total_value else 0.0,
            "projected_cash": cash_val + cash_amt - net_flow, "net_gain": net_gain, "off_mandate": off_m}


# --------------------------------------------------------------------------- #
# RENDER HELPERS (shared by both modes)
# --------------------------------------------------------------------------- #
def drift_table(a):
    rows = []
    for _, r in a.iterrows():
        br = bool(r["Breach"])
        drift_html = f'<span style="color:{NEG if br else MUTED};font-weight:{600 if br else 400}">{r["Drift"]:+.2f}</span>'
        status = pill(t["status_breach"], "action") if br else pill(t["status_ok"], "ok")
        rows.append([f'<span class="ticker-cell">{r["Ticker"]}</span>',
                     f'<span class="muted-cell">{r["Asset_Class"]}</span>',
                     f'{ccy_symbol(r["Native_Currency"])}{r["Native_Price"]:,.2f}',
                     f'{r["Current_Weight"]:.2f}%', f'{r["Target_Weight"]:.2f}%', drift_html, status])
    return html_table([t["col_ticker"], t["col_asset"], t["col_price"], t["col_current"],
                       t["col_target"], t["col_drift"], t["col_status"]], rows,
                      ["left", "left", "right", "right", "right", "right", "left"])

def orders_table(od, tax_on):
    rows = []
    for _, r in od.iterrows():
        ap = pill(r["Action"], "buy" if r["Action"] == "BUY" else "sell")
        row = [f'<span class="ticker-cell">{r["Ticker"]}</span>', ap,
               f'{int(r["Shares_to_Trade"]):,}', f'{curr_sym}{r["Estimated_Dollar_Value"]:,.0f}']
        if tax_on:
            if r["Action"] == "SELL":
                g = r["Realized_Gain_CAD"]
                gh = f'<span style="color:{POS}">{signed_cad(g)}</span>' if g < 0 else f'{signed_cad(g)}'
                th = (f'<span style="color:{NEG}">C${r["Est_Tax_CAD"]:,.0f}</span>'
                      if r["Est_Tax_CAD"] > 0 else '<span class="muted-cell">—</span>')
            else:
                gh = th = '<span class="muted-cell">—</span>'
            row += [gh, th]
        rows.append(row)
    headers = [t["col_ticker"], t["col_action"], t["col_shares"], t["col_est_value"]]
    aligns = ["left", "left", "right", "right"]
    if tax_on:
        headers += [t["col_gain"], t["col_tax"]]; aligns += ["right", "right"]
    return html_table(headers, rows, aligns)

def no_trades_box():
    return (f'<div class="table-wrap"><div style="padding:18px 16px;color:{MUTED};'
            f'font-size:0.88rem">{t["no_trades"]}</div></div>')


def allocation_charts(analysis, mandate_m):
    """Two donuts (current vs target) + a shared legend, for one book/account."""
    curr_pie = analysis[analysis["Current_Weight"] > 0].groupby("Asset_Class")["Current_Weight"].sum().reset_index()
    tgt_pie = mandate_m[mandate_m["Target_Weight_Percent"] > 0].groupby("Asset_Class")["Target_Weight_Percent"].sum().reset_index()
    cmap = build_color_map(set(curr_pie["Asset_Class"]) | set(tgt_pie["Asset_Class"]))
    c1, c2 = st.columns(2)
    with c1:
        with st.container(border=True):
            section_header("donut", t["alloc_current"])
            if not curr_pie.empty:
                st.plotly_chart(make_donut(curr_pie, "Current_Weight", cmap),
                                use_container_width=True, config={"displayModeBar": False})
    with c2:
        with st.container(border=True):
            section_header("target", t["alloc_target"])
            if not tgt_pie.empty:
                st.plotly_chart(make_donut(tgt_pie, "Target_Weight_Percent", cmap),
                                use_container_width=True, config={"displayModeBar": False})
    cur_map = dict(zip(curr_pie["Asset_Class"], curr_pie["Current_Weight"]))
    tgt_map = dict(zip(tgt_pie["Asset_Class"], tgt_pie["Target_Weight_Percent"]))
    classes = sorted(set(cur_map) | set(tgt_map), key=lambda c: -cur_map.get(c, 0.0))
    rows = "".join(f'<tr><td><span class="dot" style="background:{cmap.get(c, CASH_COLOR)}"></span>{c}</td>'
                   f'<td class="num">{cur_map.get(c, 0.0):.1f}%</td>'
                   f'<td class="num">{tgt_map.get(c, 0.0):.1f}%</td></tr>' for c in classes)
    with st.expander(t["legend_title"]):
        st.markdown(f'<table class="legend-table"><thead><tr><th>{t["col_asset"]}</th>'
                    f'<th class="num">{t["col_current"]}</th><th class="num">{t["col_target"]}</th></tr></thead>'
                    f'<tbody>{rows}</tbody></table>', unsafe_allow_html=True)


# --------------------------------------------------------------------------- #
# ===========================  RENDER: PER-ACCOUNT  ========================== #
# --------------------------------------------------------------------------- #
# Account mode: choose how to rebalance (per-account is the default).
if account_mode:
    approach = st.radio(t["approach_label"], [t["approach_per"], t["approach_whole"]],
                        horizontal=True, index=0, key="approach")
    do_per_account = (approach == t["approach_per"])
else:
    do_per_account = False

if do_per_account:
    accounts = list(dict.fromkeys(holdings_df[account_col].astype(str)))  # preserve order

    # Optional defaults from a Mandate_Name column (if the file has one).
    file_default = {}
    if has_mandate_col:
        for acct in accounts:
            sub = holdings_df[holdings_df[account_col].astype(str) == acct]
            nm = [n for n in sub["Mandate_Name"].dropna().astype(str) if n.strip() and n in mandate_names]
            if nm:
                file_default[acct] = nm[0]

    # Assign a mandate to each account, IN THE UI (one dropdown per account).
    section_header("layers", t["assign_title"])
    acct_mandate = {}
    with st.container(border=True):
        chunks = [accounts[i:i + 2] for i in range(0, len(accounts), 2)]
        for chunk in chunks:
            ccols = st.columns(len(chunk))
            for j, acct in enumerate(chunk):
                atype = (str(holdings_df[holdings_df[account_col].astype(str) == acct]["Account_Type"].iloc[0])
                         if has_type else "")
                lbl = acct + (f"  ·  {atype}" if (has_type and account_col != "Account_Type") else "")
                default_name = file_default.get(acct, selected_mandate)
                if default_name not in mandate_names:
                    default_name = selected_mandate
                acct_mandate[acct] = ccols[j].selectbox(lbl, mandate_names,
                                                        index=mandate_names.index(default_name), key=f"mand_{acct}")
    prepared = {n: prepare_mandate(n) for n in set(acct_mandate.values())}

    # Per-account cash inputs.
    with st.expander(t["per_acct_cash"], expanded=False):
        cash_by_acct = {}
        rows_of_accts = [accounts[i:i + 3] for i in range(0, len(accounts), 3)]
        for chunk in rows_of_accts:
            ccols = st.columns(len(chunk))
            for j, acct in enumerate(chunk):
                cash_by_acct[acct] = ccols[j].number_input(f"{acct} ({base})", value=0.0,
                                                            step=1000.0, format="%.2f", key=f"cash_{acct}")

    # Compute every account.
    results, combined_rows, hh_value, hh_gain = {}, [], 0.0, 0.0
    for acct in accounts:
        sub = holdings_df[holdings_df[account_col].astype(str) == acct]
        res = compute_book(sub, prepared[acct_mandate[acct]], cash_by_acct.get(acct, 0.0), tax_on=tax_capable)
        results[acct] = res
        if res:
            hh_value += res["total_value"]; hh_gain += res["net_gain"]
            for _, o in res["orders"].iterrows():
                combined_rows.append({"Account": acct, "Ticker": o["Ticker"], "Action": o["Action"],
                                      "Shares_to_Trade": o["Shares_to_Trade"],
                                      "Estimated_Dollar_Value": o["Estimated_Dollar_Value"],
                                      "Realized_Gain_CAD": o["Realized_Gain_CAD"], "Est_Tax_CAD": o["Est_Tax_CAD"]})

    hh_taxable = max(hh_gain, 0.0) * incl
    hh_tax = hh_taxable * marg
    hh_eff = (hh_tax / hh_gain * 100.0) if hh_gain > 0 else 0.0

    # Consolidated summary.
    if tax_capable:
        section_header("percent", t["taximpact_hh"])
        gcol = POS if hh_gain < 0 else INK
        st.markdown(f"""<div class="metric-row">
            <div class="metric-card"><div class="ico">{ICONS['wallet']}</div><div>
                <div class="metric-label">{t['household_value']}</div>
                <div class="metric-value">{curr_sym}{hh_value:,.0f}</div></div></div>
            <div class="metric-card"><div class="ico">{ICONS['coins']}</div><div>
                <div class="metric-label">{t['m_net_gain']}</div>
                <div class="metric-value" style="color:{gcol}">{signed_cad(hh_gain)}</div></div></div>
            <div class="metric-card"><div class="ico">{ICONS['percent']}</div><div>
                <div class="metric-label">{t['m_est_tax']}</div>
                <div class="metric-value" style="color:{NEG if hh_tax > 0 else INK}">C${hh_tax:,.0f}</div></div></div>
            <div class="metric-card"><div class="ico">{ICONS['banknote']}</div><div>
                <div class="metric-label">{t['m_eff_rate']}</div>
                <div class="metric-value">{hh_eff:.1f}%</div></div></div></div>""", unsafe_allow_html=True)
        st.markdown(f'<div class="tax-cap">{t["tax_caption_pa"]}</div>', unsafe_allow_html=True)
    else:
        st.markdown(f"""<div class="metric-row">
            <div class="metric-card"><div class="ico">{ICONS['wallet']}</div><div>
                <div class="metric-label">{t['household_value']}</div>
                <div class="metric-value">{curr_sym}{hh_value:,.0f}</div></div></div>
            <div class="metric-card"><div class="ico">{ICONS['folder']}</div><div>
                <div class="metric-label">{t['m_accounts']}</div>
                <div class="metric-value">{len(accounts)}</div></div></div></div>""", unsafe_allow_html=True)
        notice(t["tax_hint"], "info")

    # Select one account and show its detail (keeps a dense page manageable and
    # lets each table live in its own dropdown — expanders can't be nested, so a
    # selector rather than per-account expanders is what makes this work).
    section_header("layers", t["per_acct_blocks"])
    sel_acct = st.selectbox(t["view_account"], accounts, key="view_acct")
    res = results[sel_acct]
    sub = holdings_df[holdings_df[account_col].astype(str) == sel_acct]
    atype = str(sub["Account_Type"].iloc[0]) if has_type else ""
    taxable = is_taxable_account(atype) if has_type else False
    type_pill = pill(atype, "tax" if taxable else "reg") if atype else ""

    if res is None:
        st.markdown(f'<div class="acct-head"><div class="left"><span class="acct-name">{sel_acct}</span>'
                    f'{type_pill}</div></div>', unsafe_allow_html=True)
        notice(t["no_priced_acct"].format(acct=sel_acct), "warn")
    else:
        n_breach = int(res["analysis"]["Breach"].sum())
        st.markdown(f"""<div class="acct-head">
            <div class="left"><span class="acct-name">{sel_acct}</span>{type_pill}</div>
            <div class="acct-meta">{t['col_mandate']}: {acct_mandate[sel_acct]} &nbsp;·&nbsp;
                {t['lbl_value']}: {curr_sym}{res['total_value']:,.0f} &nbsp;·&nbsp;
                {n_breach} {t['status_breach'].lower()}</div></div>""", unsafe_allow_html=True)

        # Two allocation charts for this account.
        allocation_charts(res["analysis"], prepared[acct_mandate[sel_acct]])
        st.write("")

        # Each table in its own dropdown.
        with st.expander(t["analysis_title"], expanded=True):
            st.markdown(drift_table(res["analysis"]), unsafe_allow_html=True)
        with st.expander(t["orders_title"], expanded=True):
            if res["orders"].empty:
                st.markdown(no_trades_box(), unsafe_allow_html=True)
            else:
                st.markdown(orders_table(res["orders"], tax_capable), unsafe_allow_html=True)
                if tax_capable and taxable:
                    ag = float(res["orders"]["Realized_Gain_CAD"].sum())
                    at = max(ag, 0.0) * incl * marg
                    st.markdown(f'<div class="acct-tax"><b>{t["lbl_gain"]}:</b> {signed_cad(ag)} &nbsp;·&nbsp; '
                                f'<b>{t["lbl_tax"]}:</b> C${at:,.0f}</div>', unsafe_allow_html=True)
        if res["off_mandate"]:
            notice(t["offmandate_note"].format(items=", ".join(res["off_mandate"])), "info")


    # Combined household ticket download.
    if combined_rows:
        st.write("")
        section_header("exchange", t["combined_ticket"])
        comb = pd.DataFrame(combined_rows)
        export_cols = ["Account", "Ticker", "Action", "Shares_to_Trade", "Estimated_Dollar_Value"]
        if tax_capable:
            export_cols += ["Realized_Gain_CAD", "Est_Tax_CAD"]
        dl = st.columns([1, 2.2])[0]
        with dl:
            st.download_button(t["export_combined"], data=comb[export_cols].to_csv(index=False).encode("utf-8"),
                               file_name=f"household_ticket_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                               mime="text/csv", use_container_width=True)

    st.markdown("---")
    render_instructions(expanded=False)
    render_footer()
    st.stop()


# --------------------------------------------------------------------------- #
# ===================  RENDER: AGGREGATE (legacy + Tier 0)  ================== #
# --------------------------------------------------------------------------- #
mandate_m = prepare_mandate(selected_mandate)
res = compute_book(holdings_df, mandate_m, cash_flow_global, tax_on=tax_capable)
if res is None:
    notice(t["no_priced"], "warn"); render_instructions(expanded=True); render_footer(); st.stop()
analysis = res["analysis"]

st.markdown(f"""<div class="metric-row">
    <div class="metric-card"><div class="ico">{ICONS['wallet']}</div><div>
        <div class="metric-label">{t['curr_val']}</div>
        <div class="metric-value">{curr_sym}{res['total_value']:,.0f}</div></div></div>
    <div class="metric-card"><div class="ico">{ICONS['flow']}</div><div>
        <div class="metric-label">{t['cash_flow']}</div>
        <div class="metric-value">{curr_sym}{cash_flow_global:,.0f}</div></div></div>
    <div class="metric-card"><div class="ico">{ICONS['coins']}</div><div>
        <div class="metric-label">{t['cash_equiv']}</div>
        <div class="metric-value">{res['cash_pct']:.1f}%</div></div></div>
    <div class="metric-card"><div class="ico">{ICONS['banknote']}</div><div>
        <div class="metric-label">{t['post_cash']}</div>
        <div class="metric-value">{curr_sym}{res['projected_cash']:,.0f}</div></div></div></div>""",
    unsafe_allow_html=True)

target_sum = float(mandate_m["Target_Weight_Percent"].sum())
if abs(target_sum - 100.0) > 0.5:
    notice(t["targetsum_note"].format(name=selected_mandate, pct=target_sum), "warn")
if res["off_mandate"]:
    notice(t["offmandate_note"].format(items=", ".join(res["off_mandate"])), "info")

# Donuts + legend
curr_pie = analysis[analysis["Current_Weight"] > 0].groupby("Asset_Class")["Current_Weight"].sum().reset_index()
tgt_pie = mandate_m[mandate_m["Target_Weight_Percent"] > 0].groupby("Asset_Class")["Target_Weight_Percent"].sum().reset_index()
color_map = build_color_map(set(curr_pie["Asset_Class"]) | set(tgt_pie["Asset_Class"]))
c1, c2 = st.columns(2)
with c1:
    with st.container(border=True):
        section_header("donut", t["alloc_current"])
        if not curr_pie.empty:
            st.plotly_chart(make_donut(curr_pie, "Current_Weight", color_map), use_container_width=True, config={"displayModeBar": False})
with c2:
    with st.container(border=True):
        section_header("target", t["alloc_target"])
        if not tgt_pie.empty:
            st.plotly_chart(make_donut(tgt_pie, "Target_Weight_Percent", color_map), use_container_width=True, config={"displayModeBar": False})

cur_map = dict(zip(curr_pie["Asset_Class"], curr_pie["Current_Weight"]))
tgt_map = dict(zip(tgt_pie["Asset_Class"], tgt_pie["Target_Weight_Percent"]))
legend_classes = sorted(set(cur_map) | set(tgt_map), key=lambda c: -cur_map.get(c, 0.0))
legend_rows = "".join(f'<tr><td><span class="dot" style="background:{color_map.get(c, CASH_COLOR)}"></span>{c}</td>'
                      f'<td class="num">{cur_map.get(c, 0.0):.1f}%</td><td class="num">{tgt_map.get(c, 0.0):.1f}%</td></tr>'
                      for c in legend_classes)
with st.expander(t["legend_title"]):
    st.markdown(f'<table class="legend-table"><thead><tr><th>{t["col_asset"]}</th>'
                f'<th class="num">{t["col_current"]}</th><th class="num">{t["col_target"]}</th></tr></thead>'
                f'<tbody>{legend_rows}</tbody></table>', unsafe_allow_html=True)
st.write("")

# Holdings & drift
section_header("list", t["analysis_title"])
st.markdown(drift_table(analysis), unsafe_allow_html=True)
st.write("")

# Execution orders
section_header("exchange", t["orders_title"])
if res["orders"].empty:
    st.markdown(no_trades_box(), unsafe_allow_html=True)
else:
    st.markdown(orders_table(res["orders"], tax_capable), unsafe_allow_html=True)
    if tax_capable:
        st.markdown(f'<div class="tax-cap">{t["tax_caption"]}</div>', unsafe_allow_html=True)
    st.write("")
    export_cols = ["Ticker", "Action", "Shares_to_Trade", "Estimated_Dollar_Value"]
    if tax_capable:
        export_cols += ["Realized_Gain_CAD", "Est_Tax_CAD"]
    dl = st.columns([1, 2.2])[0]
    with dl:
        st.download_button(t["export_btn"], data=res["orders"][export_cols].to_csv(index=False).encode("utf-8"),
                           file_name=f"trade_ticket_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                           mime="text/csv", use_container_width=True)

# Tax impact (aggregate)
if tax_capable:
    st.write("")
    section_header("percent", t["taximpact_title"])
    ng = res["net_gain"]; tg = max(ng, 0.0) * incl; tx = tg * marg
    eff = (tx / ng * 100.0) if ng > 0 else 0.0
    gcol = POS if ng < 0 else INK
    st.markdown(f"""<div class="metric-row">
        <div class="metric-card"><div class="ico">{ICONS['coins']}</div><div>
            <div class="metric-label">{t['m_net_gain']}</div>
            <div class="metric-value" style="color:{gcol}">{signed_cad(ng)}</div></div></div>
        <div class="metric-card"><div class="ico">{ICONS['percent']}</div><div>
            <div class="metric-label">{t['m_est_tax']}</div>
            <div class="metric-value" style="color:{NEG if tx > 0 else INK}">C${tx:,.0f}</div></div></div>
        <div class="metric-card"><div class="ico">{ICONS['banknote']}</div><div>
            <div class="metric-label">{t['m_eff_rate']}</div>
            <div class="metric-value">{eff:.1f}%</div></div></div></div>""", unsafe_allow_html=True)
else:
    notice(t["tax_hint"], "info")

# Holdings by account (aggregate, if accounts present)
if has_account and has_type:
    section_header("folder", t["accounts_title"])
    sub = holdings_df[holdings_df["Quantity"] != 0].sort_values(["Account", "Ticker"])
    a_rows = []
    for _, row in sub.iterrows():
        tk = row["Ticker"]
        if tk not in base_price_map:
            continue
        qty = float(row["Quantity"])
        taxable = is_taxable_account(row.get("Account_Type", ""))
        bc = row.get("Book_Cost_CAD", None)
        bc = None if (bc is None or pd.isna(bc)) else float(bc)
        acb_u = (bc / qty) if (bc is not None and qty > 0) else None
        mv = qty * base_price_map[tk]
        tpill = pill(str(row.get("Account_Type", "")), "tax" if taxable else "reg")
        if taxable and acb_u is not None and not is_cash_ticker(tk):
            unreal = qty * (cad_price_map[tk] - acb_u)
            uh = (f'<span style="color:{POS}">{signed_cad(unreal)}</span>' if unreal >= 0
                  else f'<span style="color:{NEG}">{signed_cad(unreal)}</span>')
            ah = f'C${acb_u:,.2f}'
        else:
            uh, ah = '<span class="muted-cell">—</span>', '<span class="muted-cell">—</span>'
        a_rows.append([f'<span class="muted-cell">{row.get("Account", "—")}</span>', tpill,
                       f'<span class="ticker-cell">{tk}</span>', f'{qty:,.0f}', ah,
                       f'{curr_sym}{mv:,.0f}', uh])
    st.markdown(html_table([t["col_account"], t["col_type"], t["col_ticker"], t["col_shares"],
                            t["col_acb"], t["col_mv"], t["col_unreal"]], a_rows,
                           ["left", "left", "left", "right", "right", "right", "right"]), unsafe_allow_html=True)

st.markdown("---")
render_instructions(expanded=False)
render_footer()
