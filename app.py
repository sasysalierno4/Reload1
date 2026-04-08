"""
Wind3 Reload Dashboard — v4
Autore: Salvo (Ambassador A6 — Campania & Puglia)

Novità v4:
- Bug fix completi (filtro AND, divisione per zero, soglie hard-coded, ecc.)
- Anagrafica gerarchica dinamica: Zona → Area Manager → District Manager → Store
- Filtri multi-livello: Zona / AM / DM / categoria store / singolo store
- Vista "Mappa Zona" con tutti i negozi che competono
- Soglie davvero configurabili (anche in Excel/PDF)
- Cache più solida, format_func ottimizzato
"""

import streamlit as st
import pandas as pd
import io
import zipfile
import json as _json
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from fpdf import FPDF

# ─────────────────────────────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Wind3 Reload – Dashboard",
    page_icon="📶",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────
#  SESSION STATE
# ─────────────────────────────────────────────────────────────────
for k, v in {
    "selected_dm": None,
    "goto_detail": False,
    "goto_store_detail": False,
    "selected_store": None,
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ─────────────────────────────────────────────────────────────────
#  CSS
# ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@300;400;500;600;700;800&family=Barlow+Condensed:wght@600;700;800&display=swap');
:root {
    --blue:#003087; --blue-mid:#0050B3; --blue-light:#1976D2;
    --accent:#00A8E0; --red:#C62828; --orange:#E65100;
    --bg:#04101F; --card:#081828; --border:rgba(0,168,224,0.18);
    --text:#E8F0FE; --muted:#7B97BF;
}
* { font-family:'Barlow',sans-serif !important; box-sizing:border-box; }
html,body,[data-testid="stApp"] { background:var(--bg) !important; color:var(--text) !important; }
#MainMenu,footer,header,.stDeployButton { display:none !important; }
[data-testid="stToolbar"] { display:none !important; }
.block-container { padding:1.2rem 2rem 3rem !important; max-width:1500px !important; }
[data-testid="stSidebar"] { background:var(--card) !important; border-right:1px solid var(--border) !important; }
[data-testid="stSidebar"] * { color:var(--text) !important; }
.w3-header {
    background:linear-gradient(135deg,#001850 0%,#003087 60%,#005CB8 100%);
    border-radius:14px; padding:1.6rem 2.2rem; margin-bottom:1.2rem;
    border:1px solid rgba(0,168,224,0.3); position:relative; overflow:hidden;
}
.w3-header::after {
    content:''; position:absolute; bottom:0; left:0; right:0; height:3px;
    background:linear-gradient(90deg,var(--accent),var(--blue-light),transparent);
}
.w3-glow { position:absolute; top:-60px; right:-60px; width:300px; height:300px;
    background:radial-gradient(circle,rgba(0,168,224,0.12) 0%,transparent 70%); pointer-events:none; }
.w3-title { font-family:'Barlow Condensed',sans-serif !important; font-size:2.1rem; font-weight:800; color:white; line-height:1; }
.w3-title span { color:var(--accent); }
.w3-sub { font-size:0.78rem; color:rgba(255,255,255,0.5); letter-spacing:0.14em; text-transform:uppercase; margin-top:0.35rem; }
.w3-badge { display:inline-block; background:rgba(0,168,224,0.15); border:1px solid var(--accent);
    color:var(--accent); font-size:0.72rem; font-weight:700; letter-spacing:0.12em;
    padding:0.2rem 0.65rem; border-radius:100px; text-transform:uppercase; margin-top:0.6rem; margin-right:0.4rem; }
.kpi-row { display:flex; gap:1rem; margin-bottom:1.2rem; flex-wrap:wrap; }
.kpi-tile { flex:1; min-width:140px; background:var(--card); border:1px solid var(--border);
    border-radius:12px; padding:1rem 1.3rem; position:relative; overflow:hidden; }
.kpi-tile::before { content:''; position:absolute; top:0; left:0; right:0; height:3px; }
.kpi-tile.accent::before { background:var(--accent); }
.kpi-tile.green::before  { background:#4CAF50; }
.kpi-tile.red::before    { background:var(--red); }
.kpi-tile.orange::before { background:var(--orange); }
.kpi-label { font-size:0.7rem; font-weight:700; letter-spacing:0.13em; text-transform:uppercase; color:var(--muted); margin-bottom:0.3rem; }
.kpi-value { font-family:'Barlow Condensed',sans-serif !important; font-size:1.9rem; font-weight:800; line-height:1; color:var(--text); }
.kpi-value.red   { color:#EF5350; }
.kpi-value.green { color:#66BB6A; }
.kpi-value.accent { color:var(--accent); }
.kpi-sub { font-size:0.72rem; color:var(--muted); margin-top:0.2rem; }
.section-title { font-family:'Barlow Condensed',sans-serif !important; font-size:0.72rem; font-weight:700;
    letter-spacing:0.16em; text-transform:uppercase; color:var(--accent);
    margin-bottom:0.75rem; padding-bottom:0.4rem; border-bottom:1px solid var(--border); margin-top:0.5rem; }
.dm-card { background:var(--card); border:1px solid var(--border); border-radius:10px;
    padding:0.85rem 1.2rem; display:flex; align-items:center; gap:1.2rem; margin-bottom:0.5rem;
    transition:border-color 0.2s; position:relative; overflow:hidden; }
.dm-card:hover { border-color:rgba(0,168,224,0.4); }
.dm-card.critico    { border-left:4px solid var(--red); }
.dm-card.attenzione { border-left:4px solid var(--orange); }
.dm-card.ok         { border-left:4px solid #4CAF50; }
.dm-name { font-weight:700; font-size:0.9rem; color:var(--text); min-width:220px; flex:1; }
.dm-region { font-size:0.72rem; color:var(--muted); font-weight:400; }
.dm-kpi { text-align:center; min-width:80px; }
.dm-kpi-label { font-size:0.65rem; color:var(--muted); letter-spacing:0.1em; text-transform:uppercase; }
.dm-kpi-val { font-family:'Barlow Condensed',sans-serif !important; font-size:1.3rem; font-weight:800; line-height:1.1; }
.dm-kpi-val.red    { color:#EF5350; }
.dm-kpi-val.orange { color:#FFA726; }
.dm-kpi-val.green  { color:#66BB6A; }
.dm-kpi-val.white  { color:var(--text); }
.dm-trend { font-size:0.7rem; }
.dm-trend.up   { color:#66BB6A; }
.dm-trend.down { color:#EF5350; }
.dm-trend.flat { color:var(--muted); }
.dm-alerts { font-size:0.72rem; color:#EF5350; min-width:160px; }
.dm-forever { font-size:0.7rem; color:var(--muted); text-align:center; min-width:80px; }
.tree-am { background:var(--card); border:1px solid var(--border); border-left:4px solid var(--accent);
    border-radius:10px; padding:0.8rem 1.2rem; margin-top:0.8rem; }
.tree-am-name { font-family:'Barlow Condensed',sans-serif !important; font-size:1.1rem; font-weight:800;
    color:var(--accent); letter-spacing:0.04em; }
.tree-dm { background:rgba(0,168,224,0.04); border-left:3px solid var(--blue-light);
    padding:0.5rem 1rem; margin:0.3rem 0 0.3rem 1.2rem; border-radius:6px; font-size:0.85rem; }
.tree-dm-name { font-weight:700; color:var(--text); }
.tree-dm-meta { font-size:0.72rem; color:var(--muted); margin-top:0.1rem; }
.tree-store { font-size:0.78rem; color:var(--muted); margin-left:2.5rem; padding:0.15rem 0; }
[data-testid="stFileUploader"] { background:var(--card) !important; border:2px dashed var(--border) !important; border-radius:12px !important; padding:1rem !important; }
.stDownloadButton > button { background:linear-gradient(135deg,var(--blue-mid),var(--blue-light)) !important;
    color:white !important; border:none !important; border-radius:8px !important; font-weight:600 !important; padding:0.5rem 1.2rem !important; }
.stDownloadButton > button:hover { background:linear-gradient(135deg,var(--blue-light),var(--accent)) !important; }
.stTabs [data-baseweb="tab-list"] { background:var(--card) !important; border-radius:10px !important; border:1px solid var(--border) !important; gap:0 !important; }
.stTabs [data-baseweb="tab"] { background:transparent !important; color:var(--muted) !important; font-weight:600 !important; font-size:0.85rem !important; border-radius:8px !important; }
.stTabs [aria-selected="true"] { background:var(--blue-mid) !important; color:white !important; }
.stTabs [data-baseweb="tab-panel"] { background:transparent !important; padding:1rem 0 !important; }
[data-testid="stDataFrame"] { border-radius:10px; overflow:hidden; }
hr { border-color:var(--border) !important; margin:1.2rem 0 !important; }
.sidebar-section { font-family:'Barlow Condensed',sans-serif !important; font-size:0.7rem; font-weight:700;
    letter-spacing:0.15em; text-transform:uppercase; color:var(--accent);
    margin:1rem 0 0.5rem; padding-bottom:0.3rem; border-bottom:1px solid var(--border); }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────────────
def safe_pct(num, den):
    """Divisione protetta che restituisce 0.0 invece di inf/NaN."""
    try:
        if den is None or pd.isna(den) or den == 0:
            return 0.0
        return round(float(num) / float(den) * 100.0, 1)
    except Exception:
        return 0.0


def safe_str(v, default="—"):
    if v is None:
        return default
    try:
        if pd.isna(v):
            return default
    except Exception:
        pass
    return str(v)


def latin1(s):
    """
    Sanitizza testo per PDF FPDF (font core Helvetica = latin-1).
    Sostituisce caratteri non supportati con equivalenti ASCII.
    """
    if s is None:
        return ""
    s = str(s)
    repl = {
        "–": "-", "—": "-", "−": "-",
        "“": '"', "”": '"', "„": '"', "‟": '"',
        "‘": "'", "’": "'", "‚": ",", "‛": "'",
        "…": "...", "•": "*", "·": "-",
        "→": "->", "←": "<-", "↑": "^", "↓": "v",
        "▲": "^", "▼": "v", "♾️": "inf", "✓": "v", "✗": "x",
        "€": "EUR", "£": "GBP",
        "Δ": "Delta", "%": "%",
    }
    for k, v in repl.items():
        s = s.replace(k, v)
    # Tutto ciò che non è latin-1 → ?
    return s.encode("latin-1", errors="replace").decode("latin-1")


# ─────────────────────────────────────────────────────────────────
#  CORE LOGIC
# ─────────────────────────────────────────────────────────────────
REQUIRED_COLS = [
    "SHOP_CODE", "STORE", "COMPANY_NAME", "CITY", "PROVINCE_CODE",
    "STORE_TYPE", "REGION", "AMBASSADOR", "AREA_MANAGER", "DISTRICT_MANAGER",
    "TAM", "SR_PLUS_GROSS_SALES", "SR_PLUS_NET_SALES",
    "Activeforever", "TotalStore", "MONTH", "YEAR",
]

OPTIONAL_COLS = ["STORE_ADDRESS"]


def load_and_filter(file_bytes, cfg):
    """
    Carica il file Excel e filtra per ambassador/regioni.
    Bug fix v4:
      - filtro AND (non OR) tra ambassador e regioni
      - 'latest' calcolato DOPO il filtro per zona
      - controllo colonne mancanti
      - colonne opzionali create vuote se assenti
    """
    df = pd.read_excel(file_bytes, sheet_name=cfg["sheet_name"])

    # Check colonne obbligatorie
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"Colonne mancanti nel file: {', '.join(missing)}")

    # Colonne opzionali
    for c in OPTIONAL_COLS:
        if c not in df.columns:
            df[c] = ""

    valid = df[df["SHOP_CODE"].notna()].copy()

    # Applica filtro zona PRIMA di calcolare 'latest'
    amb_filter = cfg.get("ambassador_filter") or []
    reg_filter = cfg.get("regioni_filter") or []
    filter_mode = cfg.get("filter_mode", "AND")  # "AND" | "OR" | "AMB" | "REG"

    if filter_mode == "AND":
        m = pd.Series(True, index=valid.index)
        if amb_filter:
            m &= valid["AMBASSADOR"].isin(amb_filter)
        if reg_filter:
            m &= valid["REGION"].isin(reg_filter)
    elif filter_mode == "OR":
        m = pd.Series(False, index=valid.index)
        if amb_filter:
            m |= valid["AMBASSADOR"].isin(amb_filter)
        if reg_filter:
            m |= valid["REGION"].isin(reg_filter)
        if not amb_filter and not reg_filter:
            m = pd.Series(True, index=valid.index)
    elif filter_mode == "AMB":
        m = valid["AMBASSADOR"].isin(amb_filter) if amb_filter else pd.Series(True, index=valid.index)
    else:  # REG
        m = valid["REGION"].isin(reg_filter) if reg_filter else pd.Series(True, index=valid.index)

    zone = valid[m].copy()

    if zone.empty:
        raise ValueError(
            "Nessun dato per la zona selezionata. "
            "Controlla Ambassador / Regioni / modalità filtro."
        )

    # 'latest' calcolato sulla zona filtrata
    latest = zone.sort_values(["YEAR", "MONTH"], ascending=False).iloc[0]
    month, year = int(latest["MONTH"]), int(latest["YEAR"])

    current = zone[(zone["MONTH"] == month) & (zone["YEAR"] == year)].copy()

    prev_month = month - 1 if month > 1 else 12
    prev_year = year if month > 1 else year - 1
    prev = zone[(zone["MONTH"] == prev_month) & (zone["YEAR"] == prev_year)].copy()

    meta = {
        "month": month, "year": year,
        "prev_month": prev_month, "prev_year": prev_year,
    }
    return current, prev, meta


def compute_kpis(df, df_prev, cfg):
    """Aggrega per DM e calcola KPI + status. Bug fix: divisione protetta, soglie e flag forever davvero usati."""
    if df.empty:
        return pd.DataFrame()

    agg = df.groupby(["REGION", "AREA_MANAGER", "DISTRICT_MANAGER"], dropna=False).agg(
        TAM=("TAM", "sum"),
        Gross_Sales=("SR_PLUS_GROSS_SALES", "sum"),
        Net_Sales=("SR_PLUS_NET_SALES", "sum"),
        Forever_Active=("Activeforever", "sum"),
        Total_Stores=("TotalStore", "sum"),
        N_Negozi=("SHOP_CODE", "count"),
    ).reset_index()

    agg["Gross_%"] = agg.apply(lambda r: safe_pct(r["Gross_Sales"], r["TAM"]), axis=1)
    agg["Net_%"] = agg.apply(lambda r: safe_pct(r["Net_Sales"], r["TAM"]), axis=1)
    agg["Forever_%"] = agg.apply(lambda r: safe_pct(r["Forever_Active"], r["Total_Stores"]), axis=1)

    if df_prev is not None and not df_prev.empty:
        pa = df_prev.groupby("DISTRICT_MANAGER", dropna=False).agg(
            TAM_p=("TAM", "sum"),
            Gross_p=("SR_PLUS_GROSS_SALES", "sum"),
            Net_p=("SR_PLUS_NET_SALES", "sum"),
        ).reset_index()
        pa["Gross_%_p"] = pa.apply(lambda r: safe_pct(r["Gross_p"], r["TAM_p"]), axis=1)
        pa["Net_%_p"] = pa.apply(lambda r: safe_pct(r["Net_p"], r["TAM_p"]), axis=1)
        agg = agg.merge(pa[["DISTRICT_MANAGER", "Gross_%_p", "Net_%_p"]], on="DISTRICT_MANAGER", how="left")
        agg["Gross_Δ"] = (agg["Gross_%"] - agg["Gross_%_p"]).round(1)
        agg["Net_Δ"] = (agg["Net_%"] - agg["Net_%_p"]).round(1)
    else:
        agg["Gross_Δ"] = pd.NA
        agg["Net_Δ"] = pd.NA

    sg = cfg["soglia_critica_gross"]
    sn = cfg["soglia_critica_net"]
    sa_g = cfg.get("soglia_attenzione_gross", sg + 20)
    sa_n = cfg.get("soglia_attenzione_net", sn + 20)
    forever_zero_critico = cfg.get("forever_zero_critico", True)

    def status(row):
        alerts = []
        is_critical = False
        if row["Gross_%"] < sg:
            alerts.append(f"Gross {row['Gross_%']}% < {sg}%")
            is_critical = True
        if row["Net_%"] < sn:
            alerts.append(f"Net {row['Net_%']}% < {sn}%")
            is_critical = True
        if row["Forever_Active"] == 0:
            if forever_zero_critico:
                alerts.append("Forever = 0")
                is_critical = True
            else:
                alerts.append("Forever = 0 (warn)")
        if not alerts:
            return "ok", []
        if is_critical:
            return "critico", alerts
        # attenzione: sopra critica ma sotto attenzione
        if row["Gross_%"] < sa_g or row["Net_%"] < sa_n:
            return "attenzione", alerts
        return "attenzione", alerts

    statuses = agg.apply(status, axis=1)
    agg["status"] = [s[0] for s in statuses]
    agg["alerts"] = [s[1] for s in statuses]

    def tl(d):
        if d is None or pd.isna(d):
            return "—"
        if d > 0:
            return f"▲ +{d}%"
        if d < 0:
            return f"▼ {d}%"
        return "= 0%"

    agg["Gross_Trend"] = agg["Gross_Δ"].apply(tl)
    agg["Net_Trend"] = agg["Net_Δ"].apply(tl)
    return agg.sort_values("Gross_%", ascending=False).reset_index(drop=True)


def build_hierarchy(df):
    """
    Costruisce mappa gerarchica: AM → {DM → [stores]}.
    Restituisce dict per il rendering ad albero.
    """
    tree = {}
    for am, am_df in df.groupby("AREA_MANAGER", dropna=False):
        am_key = safe_str(am, "(senza AM)")
        tree[am_key] = {
            "regions": sorted(am_df["REGION"].dropna().unique().tolist()),
            "n_dm": am_df["DISTRICT_MANAGER"].nunique(),
            "n_stores": len(am_df),
            "tam": int(am_df["TAM"].sum()),
            "gross": int(am_df["SR_PLUS_GROSS_SALES"].sum()),
            "net": int(am_df["SR_PLUS_NET_SALES"].sum()),
            "dms": {},
        }
        tree[am_key]["gross_pct"] = safe_pct(tree[am_key]["gross"], tree[am_key]["tam"])
        tree[am_key]["net_pct"] = safe_pct(tree[am_key]["net"], tree[am_key]["tam"])
        for dm, dm_df in am_df.groupby("DISTRICT_MANAGER", dropna=False):
            dm_key = safe_str(dm, "(senza DM)")
            stores = []
            for _, s in dm_df.iterrows():
                stores.append({
                    "store": safe_str(s.get("STORE")),
                    "company": safe_str(s.get("COMPANY_NAME")),
                    "city": safe_str(s.get("CITY")),
                    "province": safe_str(s.get("PROVINCE_CODE")),
                    "type": safe_str(s.get("STORE_TYPE")),
                    "tam": int(s.get("TAM", 0) or 0),
                    "gross": int(s.get("SR_PLUS_GROSS_SALES", 0) or 0),
                    "net": int(s.get("SR_PLUS_NET_SALES", 0) or 0),
                    "forever": int(s.get("Activeforever", 0) or 0),
                })
            tree[am_key]["dms"][dm_key] = {
                "n_stores": len(stores),
                "tam": int(dm_df["TAM"].sum()),
                "gross": int(dm_df["SR_PLUS_GROSS_SALES"].sum()),
                "net": int(dm_df["SR_PLUS_NET_SALES"].sum()),
                "regions": sorted(dm_df["REGION"].dropna().unique().tolist()),
                "stores": stores,
            }
            tree[am_key]["dms"][dm_key]["gross_pct"] = safe_pct(
                tree[am_key]["dms"][dm_key]["gross"], tree[am_key]["dms"][dm_key]["tam"])
            tree[am_key]["dms"][dm_key]["net_pct"] = safe_pct(
                tree[am_key]["dms"][dm_key]["net"], tree[am_key]["dms"][dm_key]["tam"])
    return tree


def generate_message(row, meta):
    mese_str = datetime(meta["year"], meta["month"], 1).strftime("%B %Y").capitalize()

    def ts(d):
        if d is None or (isinstance(d, float) and pd.isna(d)):
            return ""
        if d > 0:
            return f"▲ +{d}%"
        if d < 0:
            return f"▼ {d}%"
        return "= 0%"

    fi = "🔴" if row["Forever_Active"] == 0 else "♾️"
    alerts = row.get("alerts", [])
    ab = "\n".join([f"⚠️ {a}" for a in alerts]) if alerts else "✅ Tutti i KPI nella norma"
    return f"""📊 *AVANZAMENTO RELOAD – {mese_str}*
👤 *{row['DISTRICT_MANAGER']}*
📍 {row['REGION']} | AM: {row['AREA_MANAGER']}
─────────────────────────
🔵 *Gross (Lordo):*   {int(row['Gross_Sales'])} / {int(row['TAM'])}  →  *{row['Gross_%']}%*  {ts(row.get('Gross_Δ'))}
🟢 *Net (Netto):*     {int(row['Net_Sales'])} / {int(row['TAM'])}  →  *{row['Net_%']}%*  {ts(row.get('Net_Δ'))}
{fi}  *Reload Forever:*  {int(row['Forever_Active'])} / {int(row['Total_Stores'])} store  →  *{row['Forever_%']}%*
─────────────────────────
🏪 Negozi: {int(row['N_Negozi'])}
{ab}"""


# ─────────────────────────────────────────────────────────────────
#  EXCEL / PDF EXPORT (con soglie configurabili — bug fix)
# ─────────────────────────────────────────────────────────────────
def _color_for_pct(val, soglia_critica, soglia_attenzione, RED, ORANGE, GREEN):
    try:
        v = float(val)
    except Exception:
        return None
    if v < soglia_critica:
        return RED
    if v < soglia_attenzione:
        return ORANGE
    return GREEN


def build_excel_riepilogo(kpis, meta, soglia_gross, soglia_net, sa_gross, sa_net, zona_label, regioni):
    wb = Workbook()
    ws = wb.active
    ws.title = "Riepilogo"
    BLUE = PatternFill("solid", fgColor="003087")
    RED = PatternFill("solid", fgColor="FFCDD2")
    ORANGE = PatternFill("solid", fgColor="FFE0B2")
    GREEN = PatternFill("solid", fgColor="C8E6C9")
    BD = Border(left=Side(style="thin"), right=Side(style="thin"),
                top=Side(style="thin"), bottom=Side(style="thin"))
    C = Alignment(horizontal="center", vertical="center")
    ws.merge_cells("A1:O1")
    tc = ws["A1"]
    reg_str = " & ".join(regioni) if regioni else "—"
    tc.value = f"AVANZAMENTO RELOAD – {meta['month']:02d}/{meta['year']} – {reg_str} ({zona_label})"
    tc.font = Font(bold=True, color="FFFFFF", size=12)
    tc.fill = BLUE
    tc.alignment = C
    ws.row_dimensions[1].height = 30
    headers = ["Regione", "Area Manager", "District Manager", "N Negozi", "TAM",
               "Gross Sales", "Gross %", "Gross Δ", "Net Sales", "Net %", "Net Δ",
               "Forever Active", "Total Stores", "Forever %", "Status"]
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=2, column=ci, value=h)
        c.font = Font(bold=True, color="FFFFFF", size=9)
        c.fill = PatternFill("solid", fgColor="0050B3")
        c.alignment = C
        c.border = BD
    ws.row_dimensions[2].height = 22

    def ds(v):
        try:
            if v is None or pd.isna(v):
                return "—"
        except Exception:
            pass
        return f"+{v}%" if v > 0 else f"{v}%"

    for ri, (_, row) in enumerate(kpis.iterrows(), 3):
        vals = [row["REGION"], row["AREA_MANAGER"], row["DISTRICT_MANAGER"],
                int(row["N_Negozi"]), int(row["TAM"]),
                int(row["Gross_Sales"]), row["Gross_%"], ds(row.get("Gross_Δ")),
                int(row["Net_Sales"]), row["Net_%"], ds(row.get("Net_Δ")),
                int(row["Forever_Active"]), int(row["Total_Stores"]),
                row["Forever_%"], row["status"].upper()]
        for ci, val in enumerate(vals, 1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.border = BD
            c.alignment = C
            if ci == 7:  # Gross %
                fill = _color_for_pct(row["Gross_%"], soglia_gross, sa_gross, RED, ORANGE, GREEN)
                if fill:
                    c.fill = fill
            if ci == 10:  # Net %
                fill = _color_for_pct(row["Net_%"], soglia_net, sa_net, RED, ORANGE, GREEN)
                if fill:
                    c.fill = fill
        ws.row_dimensions[ri].height = 18

    for col in ws.columns:
        ml = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(ml + 3, 38)
    ws.freeze_panes = "A3"
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def build_excel_dm(row, df_raw, meta, soglia_gross, soglia_net, sa_gross, sa_net):
    wb = Workbook()
    ws = wb.active
    ws.title = "Dettaglio Store"
    BLUE = PatternFill("solid", fgColor="003087")
    RED = PatternFill("solid", fgColor="FFCDD2")
    GREEN = PatternFill("solid", fgColor="C8E6C9")
    ORANGE = PatternFill("solid", fgColor="FFE0B2")
    BD = Border(left=Side(style="thin"), right=Side(style="thin"),
                top=Side(style="thin"), bottom=Side(style="thin"))
    C = Alignment(horizontal="center")
    ws.merge_cells("A1:J1")
    tc = ws["A1"]
    tc.value = f"{row['DISTRICT_MANAGER']} – {meta['month']:02d}/{meta['year']}"
    tc.font = Font(bold=True, color="FFFFFF", size=11)
    tc.fill = BLUE
    tc.alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 28
    ws["A2"] = "GROSS %"; ws["B2"] = f"{row['Gross_%']}%"
    ws["C2"] = "NET %"; ws["D2"] = f"{row['Net_%']}%"
    ws["E2"] = "FOREVER %"; ws["F2"] = f"{row['Forever_%']}%"
    ws["G2"] = "NEGOZI"; ws["H2"] = int(row["N_Negozi"])
    for cell in [ws["A2"], ws["C2"], ws["E2"], ws["G2"]]:
        cell.font = Font(bold=True, color="003087")
    ws.row_dimensions[2].height = 20

    sub = df_raw[df_raw["DISTRICT_MANAGER"] == row["DISTRICT_MANAGER"]].copy()
    sub["Gross_%"] = sub.apply(lambda r: safe_pct(r["SR_PLUS_GROSS_SALES"], r["TAM"]), axis=1)
    sub["Net_%"] = sub.apply(lambda r: safe_pct(r["SR_PLUS_NET_SALES"], r["TAM"]), axis=1)
    cols = ["STORE", "CITY", "PROVINCE_CODE", "STORE_TYPE", "TAM",
            "SR_PLUS_GROSS_SALES", "Gross_%", "SR_PLUS_NET_SALES", "Net_%", "Activeforever"]
    hdrs = ["Store", "Città", "Prov", "Tipo", "TAM", "Gross", "Gross%", "Net", "Net%", "Forever"]
    for ci, h in enumerate(hdrs, 1):
        c = ws.cell(row=4, column=ci, value=h)
        c.font = Font(bold=True, color="FFFFFF", size=9)
        c.fill = PatternFill("solid", fgColor="0050B3")
        c.alignment = C
        c.border = BD

    for ri, (_, sr) in enumerate(sub[cols].iterrows(), 5):
        for ci, val in enumerate(sr.values, 1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.border = BD
            c.alignment = C
            if ci == 7 and isinstance(val, (int, float)):
                fill = _color_for_pct(val, soglia_gross, sa_gross, RED, ORANGE, GREEN)
                if fill:
                    c.fill = fill
            if ci == 9 and isinstance(val, (int, float)):
                fill = _color_for_pct(val, soglia_net, sa_net, RED, ORANGE, GREEN)
                if fill:
                    c.fill = fill
        ws.row_dimensions[ri].height = 17

    for col in ws.columns:
        ml = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(ml + 3, 35)
    ws.freeze_panes = "A5"
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def build_pdf_riepilogo(kpis, meta, soglia_gross, soglia_net, zona_label, regioni):
    mese_str = datetime(meta["year"], meta["month"], 1).strftime("%B %Y").upper()
    pdf = FPDF()
    pdf.add_page(orientation="L")
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.set_fill_color(0, 48, 135)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 14)
    reg_str = " & ".join(regioni) if regioni else "—"
    pdf.cell(0, 12, latin1(f"WIND3 RELOAD – {mese_str} – {reg_str} ({zona_label})"),
             fill=True, ln=True, align="C")
    pdf.ln(3)
    pdf.set_font("Helvetica", "", 8)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 6,
             latin1(f"Soglie critiche: Gross < {soglia_gross}%  |  Net < {soglia_net}%  |  Forever = 0 -> CRITICO"),
             ln=True, align="C")
    pdf.ln(4)
    headers = ["District Manager", "Regione", "N.Neg", "TAM", "Gross", "Gross%",
               "Net", "Net%", "Forever%", "Trend G", "Trend N", "Status"]
    widths = [52, 26, 12, 16, 16, 18, 16, 18, 20, 18, 18, 20]
    pdf.set_fill_color(0, 80, 179)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 8)
    for h, w in zip(headers, widths):
        pdf.cell(w, 8, latin1(h), border=1, align="C", fill=True)
    pdf.ln()
    pdf.set_font("Helvetica", "", 7.5)
    for _, row in kpis.iterrows():
        if row["status"] == "critico":
            pdf.set_fill_color(255, 205, 210)
        elif row["status"] == "attenzione":
            pdf.set_fill_color(255, 224, 178)
        else:
            pdf.set_fill_color(240, 248, 255)
        pdf.set_text_color(30, 30, 30)

        def safe_delta(v):
            try:
                if v is None or pd.isna(v):
                    return "—"
            except Exception:
                pass
            return f"+{v}%" if v > 0 else f"{v}%"

        vals = [
            safe_str(row["DISTRICT_MANAGER"])[:28],
            safe_str(row["REGION"])[:14],
            str(int(row["N_Negozi"])),
            str(int(row["TAM"])),
            str(int(row["Gross_Sales"])),
            f"{row['Gross_%']}%",
            str(int(row["Net_Sales"])),
            f"{row['Net_%']}%",
            f"{row['Forever_%']}%",
            safe_delta(row.get("Gross_Δ")),
            safe_delta(row.get("Net_Δ")),
            row["status"].upper(),
        ]
        for v, w in zip(vals, widths):
            pdf.cell(w, 7, latin1(str(v)), border=1, align="C", fill=True)
        pdf.ln()
    pdf.ln(4)
    pdf.set_font("Helvetica", "I", 7)
    pdf.set_text_color(150, 150, 150)
    pdf.cell(0, 5, latin1(f"Generato il {datetime.now().strftime('%d/%m/%Y %H:%M')} – Wind3 Reload Dashboard {zona_label}"),
             align="C")
    return bytes(pdf.output())


# ─────────────────────────────────────────────────────────────────
#  SIDEBAR
# ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div class="w3-title" style="font-size:1.3rem">⚙️ Configurazione</div>', unsafe_allow_html=True)
    st.markdown("---")

    st.markdown('<div class="sidebar-section">🗺️ Zona</div>', unsafe_allow_html=True)
    zona_label = st.text_input("Ambassador code", "A6", help="Es: A1, A2, A6...")
    regioni_input = st.text_input("Regioni (separate da virgola)", "Campania, Puglia")
    regioni = [r.strip() for r in regioni_input.split(",") if r.strip()]
    filter_mode = st.selectbox(
        "Modalità filtro zona",
        ["AND", "OR", "AMB", "REG"],
        index=0,
        help=(
            "AND = ambassador E regione (consigliato)  •  "
            "OR = ambassador O regione  •  "
            "AMB = solo ambassador  •  "
            "REG = solo regioni"
        ),
    )

    st.markdown('<div class="sidebar-section">🎯 Soglie Critiche</div>', unsafe_allow_html=True)
    st.caption("Sotto queste soglie → 🔴 CRITICO")
    soglia_gross = st.slider("Gross % critica", 0, 100, 40, 1)
    soglia_net = st.slider("Net % critica", 0, 100, 37, 1)
    forever_zero_critico = st.checkbox("Forever = 0 → critico", value=True)

    st.markdown('<div class="sidebar-section">⚠️ Soglie Attenzione</div>', unsafe_allow_html=True)
    st.caption("Sotto queste soglie → 🟠 ATTENZIONE (sopra critica)")
    sa_gross = st.slider("Gross % attenzione", 0, 100, 60, 1)
    sa_net = st.slider("Net % attenzione", 0, 100, 50, 1)
    if sa_gross < soglia_gross:
        sa_gross = soglia_gross
    if sa_net < soglia_net:
        sa_net = soglia_net

    st.markdown('<div class="sidebar-section">📑 Foglio Excel</div>', unsafe_allow_html=True)
    sheet_name = st.text_input("Nome foglio", "Sales x Store")

    st.markdown("---")
    st.markdown(f"""
    <div style="font-size:0.78rem;color:#7B97BF;line-height:1.8">
    🔴 Gross critico: <b style="color:#EF5350">< {soglia_gross}%</b><br>
    🟠 Gross attenzione: <b style="color:#FFA726">< {sa_gross}%</b><br>
    🔴 Net critico: <b style="color:#EF5350">< {soglia_net}%</b><br>
    🟠 Net attenzione: <b style="color:#FFA726">< {sa_net}%</b><br>
    🗺️ Zona: <b style="color:#00A8E0">{zona_label}</b> – {', '.join(regioni)}<br>
    ⚙️ Filtro: <b>{filter_mode}</b>
    </div>
    """, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────
#  HEADER
# ─────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="w3-header">
    <div class="w3-glow"></div>
    <div class="w3-title">WIND<span>3</span> RELOAD</div>
    <div class="w3-sub">Sales Performance Dashboard · v4</div>
    <div class="w3-badge">📶 Ambassador {zona_label}</div>
    <div class="w3-badge">🗺️ {' & '.join(regioni) if regioni else '—'}</div>
    <div class="w3-badge">⚙️ Filtro {filter_mode}</div>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────
#  UPLOAD
# ─────────────────────────────────────────────────────────────────
uploaded = st.file_uploader("📎 Carica il file Excel di avanzamento", type=["xlsx"])

if uploaded is None:
    st.markdown("""
    <div style="text-align:center;padding:4rem 2rem;color:#7B97BF;">
        <div style="font-size:3rem;margin-bottom:1rem;">📂</div>
        <div style="font-family:'Barlow Condensed';font-size:1.3rem;font-weight:700;margin-bottom:0.5rem;">
            Carica il file Excel per iniziare
        </div>
        <div style="font-size:0.85rem;">Formato: <code>Avanzamento_DD-MM-YY__Sales_x_Store.xlsx</code></div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()


# ─────────────────────────────────────────────────────────────────
#  ANALISI
# ─────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_df(file_bytes, amb, reg_str, sheet_name, filter_mode):
    cfg = {
        "ambassador_filter": [amb] if amb else [],
        "regioni_filter": [r.strip() for r in reg_str.split(",") if r.strip()],
        "sheet_name": sheet_name,
        "filter_mode": filter_mode,
    }
    df, df_prev, meta = load_and_filter(io.BytesIO(file_bytes), cfg)
    return df, df_prev, meta


file_bytes = uploaded.read()

with st.spinner("Analisi in corso..."):
    try:
        df, df_prev, meta = load_df(file_bytes, zona_label, ",".join(regioni), sheet_name, filter_mode)
    except Exception as e:
        st.error(f"Errore nel caricamento: {e}")
        st.stop()

# KPI calcolati FUORI dalla cache (così le soglie cambiano subito)
cfg_kpi = {
    "soglia_critica_gross": soglia_gross,
    "soglia_critica_net": soglia_net,
    "soglia_attenzione_gross": sa_gross,
    "soglia_attenzione_net": sa_net,
    "forever_zero_critico": forever_zero_critico,
}
kpis = compute_kpis(df, df_prev, cfg_kpi)

if kpis.empty:
    st.warning("Nessun KPI calcolabile sui dati filtrati.")
    st.stop()

mese_str = datetime(meta["year"], meta["month"], 1).strftime("%B %Y").capitalize()
n_critici = (kpis["status"] == "critico").sum()
n_attenzione = (kpis["status"] == "attenzione").sum()
n_ok = (kpis["status"] == "ok").sum()
avg_gross = kpis["Gross_%"].mean()
avg_net = kpis["Net_%"].mean()

if st.session_state["selected_dm"] is None or st.session_state["selected_dm"] not in kpis["DISTRICT_MANAGER"].values:
    st.session_state["selected_dm"] = kpis["DISTRICT_MANAGER"].iloc[0]


# ─────────────────────────────────────────────────────────────────
#  KPI TILES
# ─────────────────────────────────────────────────────────────────
n_am = kpis["AREA_MANAGER"].nunique()
st.markdown(f"""
<div class="kpi-row">
    <div class="kpi-tile accent">
        <div class="kpi-label">📅 Periodo</div>
        <div class="kpi-value accent" style="font-size:1.3rem">{mese_str}</div>
        <div class="kpi-sub">{len(df)} store · {len(kpis)} DM · {n_am} AM</div>
    </div>
    <div class="kpi-tile {'red' if avg_gross<soglia_gross else 'green'}">
        <div class="kpi-label">📊 Gross medio</div>
        <div class="kpi-value {'red' if avg_gross<soglia_gross else 'green'}">{avg_gross:.1f}%</div>
        <div class="kpi-sub">Soglia: {soglia_gross}%</div>
    </div>
    <div class="kpi-tile {'red' if avg_net<soglia_net else 'green'}">
        <div class="kpi-label">📈 Net medio</div>
        <div class="kpi-value {'red' if avg_net<soglia_net else 'green'}">{avg_net:.1f}%</div>
        <div class="kpi-sub">Soglia: {soglia_net}%</div>
    </div>
    <div class="kpi-tile red">
        <div class="kpi-label">🔴 Critici</div>
        <div class="kpi-value red">{n_critici}</div>
        <div class="kpi-sub">su {len(kpis)} DM</div>
    </div>
    <div class="kpi-tile orange">
        <div class="kpi-label">⚠️ Attenzione</div>
        <div class="kpi-value" style="color:#FFA726">{n_attenzione}</div>
        <div class="kpi-sub">in fascia</div>
    </div>
    <div class="kpi-tile green">
        <div class="kpi-label">✅ OK</div>
        <div class="kpi-value green">{n_ok}</div>
        <div class="kpi-sub">sopra soglie</div>
    </div>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────
#  TABS
# ─────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs([
    "📋 Riepilogo DM",
    "📄 Dettaglio & Messaggi",
    "🌳 Mappa Zona",
    "⬇️ Download",
])


# ── TAB 1: RIEPILOGO DM ────────────────────────────────────────
with tab1:
    col_f1, col_f2, col_f3, col_f4 = st.columns([1, 1, 1, 2])
    with col_f1:
        filtro_status = st.selectbox("Status", ["Tutti", "🔴 Critici", "⚠️ Attenzione", "✅ OK"])
    with col_f2:
        filtro_regione = st.selectbox("Regione", ["Tutte"] + sorted(kpis["REGION"].dropna().unique().tolist()))
    with col_f3:
        store_types = ["Tutti"] + sorted(df["STORE_TYPE"].dropna().unique().tolist())
        filtro_tipo = st.selectbox("Tipo Store", store_types)
    with col_f4:
        filtro_am = st.selectbox(
            "Area Manager",
            ["Tutti"] + sorted(kpis["AREA_MANAGER"].dropna().unique().tolist()),
        )

    col_s1, col_s2 = st.columns([2, 1])
    with col_s1:
        sort_col = st.selectbox(
            "Ordina per",
            ["Gross_%", "Net_%", "Forever_%", "N_Negozi", "TAM", "DISTRICT_MANAGER"],
            format_func=lambda x: {
                "Gross_%": "Gross %", "Net_%": "Net %", "Forever_%": "Forever %",
                "N_Negozi": "N. Negozi", "TAM": "TAM", "DISTRICT_MANAGER": "Nome",
            }[x],
        )
    with col_s2:
        sort_asc = st.checkbox("Crescente", value=False)

    filtered = kpis.copy()
    if filtro_status == "🔴 Critici":
        filtered = filtered[filtered["status"] == "critico"]
    elif filtro_status == "⚠️ Attenzione":
        filtered = filtered[filtered["status"] == "attenzione"]
    elif filtro_status == "✅ OK":
        filtered = filtered[filtered["status"] == "ok"]
    if filtro_regione != "Tutte":
        filtered = filtered[filtered["REGION"] == filtro_regione]
    if filtro_am != "Tutti":
        filtered = filtered[filtered["AREA_MANAGER"] == filtro_am]

    if filtro_tipo != "Tutti":
        df_tipo = df[df["STORE_TYPE"] == filtro_tipo]
        at = df_tipo.groupby("DISTRICT_MANAGER").agg(
            TAM_t=("TAM", "sum"),
            Gross_t=("SR_PLUS_GROSS_SALES", "sum"),
            Net_t=("SR_PLUS_NET_SALES", "sum"),
            N_t=("SHOP_CODE", "count"),
        ).reset_index()
        at["Gross_%_t"] = at.apply(lambda r: safe_pct(r["Gross_t"], r["TAM_t"]), axis=1)
        at["Net_%_t"] = at.apply(lambda r: safe_pct(r["Net_t"], r["TAM_t"]), axis=1)
        filtered = filtered.merge(
            at[["DISTRICT_MANAGER", "Gross_%_t", "Net_%_t", "N_t"]],
            on="DISTRICT_MANAGER", how="inner",
        )

    filtered = filtered.sort_values(sort_col, ascending=sort_asc)

    st.markdown(
        f'<div class="section-title">District Manager — {len(filtered)} risultati · {sort_col} {"↑" if sort_asc else "↓"}</div>',
        unsafe_allow_html=True,
    )

    show_chart = st.checkbox("📊 Mostra grafico Gross/Net/Forever", value=False)
    if show_chart and not filtered.empty:
        import altair as alt
        cd = filtered[["DISTRICT_MANAGER", "Gross_%", "Net_%", "Forever_%"]].rename(
            columns={"DISTRICT_MANAGER": "DM", "Gross_%": "Gross", "Net_%": "Net", "Forever_%": "Forever"})
        cm = cd.melt("DM", var_name="KPI", value_name="Valore")
        chart = alt.Chart(cm).mark_bar().encode(
            x=alt.X("DM:N", sort=None, axis=alt.Axis(labelAngle=-45, labelLimit=120)),
            y=alt.Y("Valore:Q", title="%"),
            color=alt.Color("KPI:N", scale=alt.Scale(domain=["Gross", "Net", "Forever"],
                                                     range=["#1976D2", "#00A8E0", "#4CAF50"])),
            xOffset="KPI:N",
            tooltip=["DM", "KPI", "Valore"],
        ).properties(height=300).configure_view(strokeOpacity=0).configure(
            background="transparent"
        ).configure_axis(labelColor="#7B97BF", titleColor="#7B97BF", gridColor="#0C1F35"
                         ).configure_legend(labelColor="#E8F0FE", titleColor="#7B97BF")
        st.altair_chart(chart, use_container_width=True)

    def ckpi(val, s, sa):
        if val < s:
            return "red"
        if val < sa:
            return "orange"
        return "green"

    def thtml(d):
        if d is None or pd.isna(d):
            return '<span class="dm-trend flat">—</span>'
        if d > 0:
            return f'<span class="dm-trend up">▲ +{d}%</span>'
        if d < 0:
            return f'<span class="dm-trend down">▼ {d}%</span>'
        return '<span class="dm-trend flat">= 0%</span>'

    for _, row in filtered.iterrows():
        dm_name = row["DISTRICT_MANAGER"]
        alerts_txt = " | ".join(row["alerts"]) if row["alerts"] else ""
        f_icon = "🔴" if row["Forever_Active"] == 0 else "♾️"
        g_show = row.get("Gross_%_t", row["Gross_%"])
        n_show = row.get("Net_%_t", row["Net_%"])
        tipo_note = f" ({filtro_tipo})" if filtro_tipo != "Tutti" else ""

        col_card, col_btn = st.columns([11, 1])
        with col_card:
            st.markdown(f"""
            <div class="dm-card {row['status']}">
                <div class="dm-name">{dm_name}<div class="dm-region">{row['REGION']} · {row['AREA_MANAGER']}</div></div>
                <div class="dm-kpi">
                    <div class="dm-kpi-label">Gross{tipo_note}</div>
                    <div class="dm-kpi-val {ckpi(g_show,soglia_gross,sa_gross)}">{g_show}%</div>
                    {thtml(row.get('Gross_Δ'))}
                </div>
                <div class="dm-kpi">
                    <div class="dm-kpi-label">Net{tipo_note}</div>
                    <div class="dm-kpi-val {ckpi(n_show,soglia_net,sa_net)}">{n_show}%</div>
                    {thtml(row.get('Net_Δ'))}
                </div>
                <div class="dm-forever">
                    {f_icon} <strong>{int(row['Forever_Active'])}</strong>/{int(row['Total_Stores'])}<br>
                    <span style="font-size:0.8rem">{row['Forever_%']}%</span>
                </div>
                <div class="dm-alerts">{alerts_txt}</div>
            </div>""", unsafe_allow_html=True)
        with col_btn:
            if st.button("🏪", key=f"btn_{dm_name}", help=f"Vedi store di {dm_name}"):
                st.session_state["selected_dm"] = dm_name
                st.session_state["goto_detail"] = True
                st.rerun()

    if st.session_state.get("goto_detail"):
        st.info("👆 Vai al tab **📄 Dettaglio & Messaggi** per vedere i negozi del DM selezionato")
        st.session_state["goto_detail"] = False


# ── TAB 2: DETTAGLIO & MESSAGGI ────────────────────────────────
with tab2:
    sub_tab1, sub_tab2 = st.tabs(["👤 Per District Manager", "🔍 Cerca Store / Ragione Sociale"])

    # Precomputa mappa stato per DM (fix performance format_func O(n²))
    status_by_dm = dict(zip(kpis["DISTRICT_MANAGER"], kpis["status"]))

    def dm_label(x):
        s = status_by_dm.get(x, "ok")
        icon = "🔴" if s == "critico" else ("⚠️" if s == "attenzione" else "✅")
        return f"{icon} {x}"

    with sub_tab1:
        dm_list = kpis["DISTRICT_MANAGER"].tolist()
        saved_dm = st.session_state.get("selected_dm", dm_list[0])
        dm_idx = dm_list.index(saved_dm) if saved_dm in dm_list else 0

        dm_selected = st.selectbox(
            "Seleziona District Manager", dm_list, index=dm_idx, format_func=dm_label,
        )
        st.session_state["selected_dm"] = dm_selected
        row_sel = kpis[kpis["DISTRICT_MANAGER"] == dm_selected].iloc[0]

        col_msg, col_store = st.columns([1, 2])

        with col_msg:
            st.markdown('<div class="section-title">💬 Messaggio pronto</div>', unsafe_allow_html=True)
            msg = generate_message(row_sel, meta)
            st.text_area("msg", msg, height=320, label_visibility="collapsed")
            st.download_button(
                "⬇️ Scarica .txt", data=msg.encode("utf-8"),
                file_name=f"{dm_selected.replace(' ','_')}_{meta['month']:02d}_{meta['year']}.txt",
                mime="text/plain", key="dl_msg_dm",
            )

        with col_store:
            st.markdown('<div class="section-title">🏪 Dettaglio Store</div>', unsafe_allow_html=True)
            sub = df[df["DISTRICT_MANAGER"] == dm_selected].copy()
            sub["Gross_%"] = sub.apply(lambda r: safe_pct(r["SR_PLUS_GROSS_SALES"], r["TAM"]), axis=1)
            sub["Net_%"] = sub.apply(lambda r: safe_pct(r["SR_PLUS_NET_SALES"], r["TAM"]), axis=1)

            tipi_dm = ["Tutti"] + sorted(sub["STORE_TYPE"].dropna().unique().tolist())
            ft_dm = st.selectbox("Filtra per tipo store", tipi_dm, key="tipo_dm")
            if ft_dm != "Tutti":
                sub = sub[sub["STORE_TYPE"] == ft_dm]

            display_cols = {
                "STORE": "Store", "COMPANY_NAME": "Ragione Sociale", "CITY": "Città",
                "PROVINCE_CODE": "Prov", "STORE_TYPE": "Tipo", "TAM": "TAM",
                "SR_PLUS_GROSS_SALES": "Gross", "Gross_%": "Gross%",
                "SR_PLUS_NET_SALES": "Net", "Net_%": "Net%", "Activeforever": "Forever",
            }
            sub_display = sub[list(display_cols.keys())].rename(columns=display_cols)

            sc_store = st.selectbox("Ordina store per", ["Gross%", "Net%", "TAM", "Store", "Città"], key="sort_store")
            sa_store = st.checkbox("Crescente", value=True, key="asc_store")
            sub_display = sub_display.sort_values(sc_store, ascending=sa_store)

            def cg(v):
                if v < soglia_gross:
                    return "background-color:#FFCDD2;color:#B71C1C"
                if v < sa_gross:
                    return "background-color:#FFE0B2;color:#E65100"
                return "background-color:#C8E6C9;color:#1B5E20"

            def cn(v):
                if v < soglia_net:
                    return "background-color:#FFCDD2;color:#B71C1C"
                if v < sa_net:
                    return "background-color:#FFE0B2;color:#E65100"
                return "background-color:#C8E6C9;color:#1B5E20"

            styled = sub_display.style.map(cg, subset=["Gross%"]).map(cn, subset=["Net%"])
            st.dataframe(styled, use_container_width=True, height=350)

            buf_dm = build_excel_dm(row_sel, df, meta, soglia_gross, soglia_net, sa_gross, sa_net)
            st.download_button(
                "⬇️ Scarica Excel DM", data=buf_dm,
                file_name=f"{dm_selected.replace(' ','_')}_{meta['month']:02d}_{meta['year']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_excel_dm",
            )

    with sub_tab2:
        st.markdown('<div class="section-title">🔍 Ricerca Store / Ragione Sociale</div>', unsafe_allow_html=True)
        col_sq, col_st1, col_st2 = st.columns([3, 1, 1])
        with col_sq:
            search_query = st.text_input(
                "Cerca", placeholder="Es: MEDIAWORLD, BLOCKCHAIN SRL, VIA ROMA, NAPOLI...",
                label_visibility="collapsed",
            )
        with col_st1:
            ft_search = st.selectbox(
                "Tipo", ["Tutti"] + sorted(df["STORE_TYPE"].dropna().unique().tolist()),
                key="tipo_search", label_visibility="collapsed",
            )
        with col_st2:
            ft_region_search = st.selectbox(
                "Regione", ["Tutte"] + sorted(df["REGION"].dropna().unique().tolist()),
                key="reg_search", label_visibility="collapsed",
            )

        if search_query.strip():
            q = search_query.strip().upper()
            search_cols = ["STORE", "COMPANY_NAME", "CITY", "STORE_ADDRESS", "PROVINCE_CODE"]
            mask = pd.Series(False, index=df.index)
            for c in search_cols:
                if c in df.columns:
                    mask |= df[c].astype(str).str.upper().str.contains(q, na=False)
            results = df[mask].copy()
            if ft_search != "Tutti":
                results = results[results["STORE_TYPE"] == ft_search]
            if ft_region_search != "Tutte":
                results = results[results["REGION"] == ft_region_search]

            if results.empty:
                st.info(f"Nessun risultato per: {search_query}")
            else:
                results["Gross_%"] = results.apply(lambda r: safe_pct(r["SR_PLUS_GROSS_SALES"], r["TAM"]), axis=1)
                results["Net_%"] = results.apply(lambda r: safe_pct(r["SR_PLUS_NET_SALES"], r["TAM"]), axis=1)
                st.markdown(f"**{len(results)} store trovati**")
                rc = {
                    "AREA_MANAGER": "AM", "DISTRICT_MANAGER": "District Manager",
                    "STORE": "Store", "COMPANY_NAME": "Ragione Sociale",
                    "CITY": "Città", "PROVINCE_CODE": "Prov", "STORE_TYPE": "Tipo", "TAM": "TAM",
                    "SR_PLUS_GROSS_SALES": "Gross", "Gross_%": "Gross%",
                    "SR_PLUS_NET_SALES": "Net", "Net_%": "Net%", "Activeforever": "Forever",
                }
                rd = results[list(rc.keys())].rename(columns=rc).sort_values("Gross%")

                def cg2(v):
                    if v < soglia_gross:
                        return "background-color:#FFCDD2;color:#B71C1C"
                    if v < sa_gross:
                        return "background-color:#FFE0B2;color:#E65100"
                    return "background-color:#C8E6C9;color:#1B5E20"

                def cn2(v):
                    if v < soglia_net:
                        return "background-color:#FFCDD2;color:#B71C1C"
                    if v < sa_net:
                        return "background-color:#FFE0B2;color:#E65100"
                    return "background-color:#C8E6C9;color:#1B5E20"

                st.dataframe(
                    rd.style.map(cg2, subset=["Gross%"]).map(cn2, subset=["Net%"]),
                    use_container_width=True, height=400,
                )
        else:
            st.markdown(f"""
            <div style="text-align:center;padding:2.5rem;color:#7B97BF;">
                <div style="font-size:2rem;margin-bottom:0.5rem;">🔍</div>
                Digita nome store, ragione sociale, città, indirizzo o provincia<br>
                tra i tuoi <b style="color:#00A8E0">{len(df)}</b> store
            </div>""", unsafe_allow_html=True)


# ── TAB 3: MAPPA ZONA (gerarchia) ──────────────────────────────
with tab3:
    st.markdown(
        '<div class="section-title">🌳 Mappa gerarchica della zona</div>',
        unsafe_allow_html=True,
    )
    st.caption(
        f"Vista ad albero di tutti gli **Area Manager**, i **District Manager** e i "
        f"**negozi** che competono alla zona {zona_label} ({' & '.join(regioni) if regioni else '—'})."
    )

    col_h1, col_h2, col_h3 = st.columns([1, 1, 1])
    with col_h1:
        h_filter_am = st.selectbox(
            "Filtra AM", ["Tutti"] + sorted(df["AREA_MANAGER"].dropna().unique().tolist()),
            key="h_am",
        )
    with col_h2:
        h_filter_type = st.selectbox(
            "Filtra tipo store",
            ["Tutti"] + sorted(df["STORE_TYPE"].dropna().unique().tolist()),
            key="h_type",
        )
    with col_h3:
        h_show_stores = st.checkbox("Mostra elenco negozi", value=False, key="h_stores")

    df_h = df.copy()
    if h_filter_am != "Tutti":
        df_h = df_h[df_h["AREA_MANAGER"] == h_filter_am]
    if h_filter_type != "Tutti":
        df_h = df_h[df_h["STORE_TYPE"] == h_filter_type]

    tree = build_hierarchy(df_h)

    # Riassunto
    tot_am = len(tree)
    tot_dm = sum(len(v["dms"]) for v in tree.values())
    tot_stores = sum(v["n_stores"] for v in tree.values())
    st.markdown(f"""
    <div style="margin:0.6rem 0 1rem;font-size:0.85rem;color:#7B97BF">
        <b style="color:#00A8E0">{tot_am}</b> Area Manager ·
        <b style="color:#00A8E0">{tot_dm}</b> District Manager ·
        <b style="color:#00A8E0">{tot_stores}</b> negozi
    </div>
    """, unsafe_allow_html=True)

    # Esporta gerarchia in CSV
    rows_export = []
    for am_name, am_data in tree.items():
        for dm_name, dm_data in am_data["dms"].items():
            for s in dm_data["stores"]:
                rows_export.append({
                    "Area Manager": am_name,
                    "District Manager": dm_name,
                    "Store": s["store"],
                    "Ragione Sociale": s["company"],
                    "Città": s["city"],
                    "Prov": s["province"],
                    "Tipo": s["type"],
                    "TAM": s["tam"],
                    "Gross": s["gross"],
                    "Gross %": safe_pct(s["gross"], s["tam"]),
                    "Net": s["net"],
                    "Net %": safe_pct(s["net"], s["tam"]),
                    "Forever": s["forever"],
                })
    if rows_export:
        export_df = pd.DataFrame(rows_export)
        csv_buf = export_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Esporta mappa zona (CSV)", data=csv_buf,
            file_name=f"mappa_zona_{zona_label}_{meta['month']:02d}_{meta['year']}.csv",
            mime="text/csv",
        )

    # Render albero
    for am_name, am_data in tree.items():
        st.markdown(f"""
        <div class="tree-am">
            <div class="tree-am-name">👔 {am_name}</div>
            <div style="font-size:0.78rem;color:#7B97BF;margin-top:0.2rem">
                {', '.join(am_data['regions']) or '—'} ·
                <b style="color:#E8F0FE">{am_data['n_dm']}</b> DM ·
                <b style="color:#E8F0FE">{am_data['n_stores']}</b> store ·
                Gross <b style="color:#00A8E0">{am_data['gross_pct']}%</b> ·
                Net <b style="color:#00A8E0">{am_data['net_pct']}%</b>
            </div>
        </div>
        """, unsafe_allow_html=True)

        for dm_name, dm_data in am_data["dms"].items():
            st.markdown(f"""
            <div class="tree-dm">
                <div class="tree-dm-name">🧑‍💼 {dm_name}</div>
                <div class="tree-dm-meta">
                    {', '.join(dm_data['regions']) or '—'} ·
                    <b style="color:#E8F0FE">{dm_data['n_stores']}</b> store ·
                    TAM {dm_data['tam']} ·
                    Gross <b style="color:#00A8E0">{dm_data['gross_pct']}%</b> ·
                    Net <b style="color:#00A8E0">{dm_data['net_pct']}%</b>
                </div>
            </div>
            """, unsafe_allow_html=True)
            if h_show_stores:
                for s in dm_data["stores"]:
                    g = safe_pct(s["gross"], s["tam"])
                    color = "#EF5350" if g < soglia_gross else ("#FFA726" if g < sa_gross else "#66BB6A")
                    st.markdown(f"""
                    <div class="tree-store">
                        🏪 <b style="color:#E8F0FE">{s['store']}</b> · {s['city']} ({s['province']}) ·
                        <span style="color:#7B97BF">{s['type']}</span> ·
                        TAM {s['tam']} · Gross <span style="color:{color}"><b>{g}%</b></span> ·
                        Forever {s['forever']}
                    </div>
                    """, unsafe_allow_html=True)


# ── TAB 4: DOWNLOAD ────────────────────────────────────────────
with tab4:
    st.markdown('<div class="section-title">⬇️ Download pacchetto completo</div>', unsafe_allow_html=True)
    col_d1, col_d2, col_d3 = st.columns(3)

    with col_d1:
        st.markdown("**📊 Riepilogo Excel**")
        st.caption("Tutti i DM con KPI, trend, alert — colorato per soglie")
        buf_rie = build_excel_riepilogo(kpis, meta, soglia_gross, soglia_net, sa_gross, sa_net, zona_label, regioni)
        st.download_button(
            "⬇️ riepilogo_generale.xlsx", data=buf_rie,
            file_name=f"riepilogo_{zona_label}_{meta['month']:02d}_{meta['year']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with col_d2:
        st.markdown("**📄 Riepilogo PDF**")
        st.caption("Versione stampabile — con trend mese precedente")
        pdf_bytes = build_pdf_riepilogo(kpis, meta, soglia_gross, soglia_net, zona_label, regioni)
        st.download_button(
            "⬇️ riepilogo_generale.pdf", data=pdf_bytes,
            file_name=f"riepilogo_{zona_label}_{meta['month']:02d}_{meta['year']}.pdf",
            mime="application/pdf", use_container_width=True,
        )

    with col_d3:
        st.markdown("**💬 Tutti i messaggi .txt**")
        st.caption("Un messaggio per ogni DM, pronti da copiare su WhatsApp")
        sep = "=" * 60 + "\n"
        all_msgs = sep.join([generate_message(row, meta) + "\n\n" for _, row in kpis.iterrows()])
        st.download_button(
            "⬇️ messaggi_whatsapp.txt", data=all_msgs.encode("utf-8"),
            file_name=f"messaggi_{zona_label}_{meta['month']:02d}_{meta['year']}.txt",
            mime="text/plain", use_container_width=True,
        )

    st.markdown("---")
    st.markdown("**📦 ZIP — Excel singolo per ogni DM**")
    st.caption("Un file Excel per ciascun District Manager con dettaglio store")
    if st.button("🗜️ Genera ZIP", use_container_width=True):
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for _, row in kpis.iterrows():
                dm_buf = build_excel_dm(row, df, meta, soglia_gross, soglia_net, sa_gross, sa_net)
                safe = (row["DISTRICT_MANAGER"]
                        .replace("/", "_").replace("\\", "_").replace(" ", "_")[:40])
                zf.writestr(f"{safe}_{meta['month']:02d}_{meta['year']}.xlsx", dm_buf.read())
        zip_buf.seek(0)
        st.download_button(
            "⬇️ Scarica ZIP", data=zip_buf,
            file_name=f"wind3_reload_DM_{zona_label}_{meta['month']:02d}_{meta['year']}.zip",
            mime="application/zip", use_container_width=True,
        )

    st.markdown("---")
    st.markdown(f"""
    <div style="text-align:center;color:#7B97BF;font-size:0.75rem;padding:1rem;">
        Wind3 Reload Dashboard v4 · Ambassador {zona_label} · {mese_str}<br>
        Soglie — Gross critica: {soglia_gross}% / attenzione: {sa_gross}% ·
        Net critica: {soglia_net}% / attenzione: {sa_net}%
    </div>
    """, unsafe_allow_html=True)
