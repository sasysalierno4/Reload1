import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Wind3 Reload – Dashboard A6",
    page_icon="📶",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────
#  CUSTOM CSS
# ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@300;400;500;600;700;800&family=Barlow+Condensed:wght@600;700;800&display=swap');

:root {
    --blue:        #003087;
    --blue-mid:    #0050B3;
    --blue-light:  #1976D2;
    --accent:      #00A8E0;
    --red:         #C62828;
    --orange:      #E65100;
    --yellow:      #F9A825;
    --green:       #2E7D32;
    --bg:          #04101F;
    --card:        #081828;
    --card2:       #0C1F35;
    --border:      rgba(0,168,224,0.18);
    --text:        #E8F0FE;
    --muted:       #7B97BF;
}

* { font-family: 'Barlow', sans-serif !important; box-sizing: border-box; }

html, body, [data-testid="stApp"] {
    background: var(--bg) !important;
    color: var(--text) !important;
}
#MainMenu, footer, header, .stDeployButton { display: none !important; }
[data-testid="stToolbar"] { display: none !important; }

.block-container { padding: 1.5rem 2.5rem 3rem !important; max-width: 1440px !important; }

/* ── SIDEBAR ── */
[data-testid="stSidebar"] {
    background: var(--card) !important;
    border-right: 1px solid var(--border) !important;
}
[data-testid="stSidebar"] * { color: var(--text) !important; }
[data-testid="stSidebar"] .stSlider label,
[data-testid="stSidebar"] .stNumberInput label { color: var(--muted) !important; font-size: 0.8rem !important; }

/* ── HEADER ── */
.w3-header {
    background: linear-gradient(135deg, #001850 0%, #003087 60%, #005CB8 100%);
    border-radius: 14px;
    padding: 2rem 2.5rem;
    margin-bottom: 1.5rem;
    border: 1px solid rgba(0,168,224,0.3);
    position: relative;
    overflow: hidden;
}
.w3-header::after {
    content: '';
    position: absolute;
    bottom: 0; left: 0; right: 0;
    height: 3px;
    background: linear-gradient(90deg, var(--accent), var(--blue-light), transparent);
}
.w3-glow {
    position: absolute;
    top: -60px; right: -60px;
    width: 300px; height: 300px;
    background: radial-gradient(circle, rgba(0,168,224,0.12) 0%, transparent 70%);
    pointer-events: none;
}
.w3-title {
    font-family: 'Barlow Condensed', sans-serif !important;
    font-size: 2.2rem; font-weight: 800;
    letter-spacing: -0.01em; color: white; line-height: 1;
}
.w3-title span { color: var(--accent); }
.w3-sub {
    font-size: 0.78rem; color: rgba(255,255,255,0.5);
    letter-spacing: 0.14em; text-transform: uppercase; margin-top: 0.35rem;
}
.w3-badge {
    display: inline-block;
    background: rgba(0,168,224,0.15);
    border: 1px solid var(--accent);
    color: var(--accent);
    font-size: 0.72rem; font-weight: 700;
    letter-spacing: 0.12em; padding: 0.2rem 0.65rem;
    border-radius: 100px; text-transform: uppercase; margin-top: 0.6rem;
}

/* ── KPI TILES ── */
.kpi-row { display: flex; gap: 1rem; margin-bottom: 1.5rem; flex-wrap: wrap; }
.kpi-tile {
    flex: 1; min-width: 140px;
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 1.1rem 1.4rem;
    position: relative; overflow: hidden;
}
.kpi-tile::before {
    content: ''; position: absolute;
    top: 0; left: 0; right: 0; height: 3px;
}
.kpi-tile.accent::before { background: var(--accent); }
.kpi-tile.green::before  { background: #4CAF50; }
.kpi-tile.red::before    { background: var(--red); }
.kpi-tile.orange::before { background: var(--orange); }
.kpi-label {
    font-size: 0.7rem; font-weight: 700;
    letter-spacing: 0.13em; text-transform: uppercase;
    color: var(--muted); margin-bottom: 0.3rem;
}
.kpi-value {
    font-family: 'Barlow Condensed', sans-serif !important;
    font-size: 2rem; font-weight: 800; line-height: 1;
    color: var(--text);
}
.kpi-value.red    { color: #EF5350; }
.kpi-value.green  { color: #66BB6A; }
.kpi-value.accent { color: var(--accent); }
.kpi-sub { font-size: 0.72rem; color: var(--muted); margin-top: 0.2rem; }

/* ── SECTION TITLE ── */
.section-title {
    font-family: 'Barlow Condensed', sans-serif !important;
    font-size: 0.72rem; font-weight: 700;
    letter-spacing: 0.16em; text-transform: uppercase;
    color: var(--accent); margin-bottom: 0.75rem;
    padding-bottom: 0.4rem;
    border-bottom: 1px solid var(--border);
}

/* ── DM CARDS ── */
.dm-grid { display: flex; flex-direction: column; gap: 0.5rem; }
.dm-card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 0.9rem 1.2rem;
    display: flex; align-items: center; gap: 1.2rem;
    transition: border-color 0.2s;
    position: relative; overflow: hidden;
}
.dm-card:hover { border-color: rgba(0,168,224,0.4); }
.dm-card.critico { border-left: 4px solid var(--red); }
.dm-card.attenzione { border-left: 4px solid var(--orange); }
.dm-card.ok { border-left: 4px solid #4CAF50; }
.dm-name {
    font-weight: 700; font-size: 0.9rem;
    color: var(--text); min-width: 220px; flex: 1;
}
.dm-region { font-size: 0.72rem; color: var(--muted); font-weight: 400; }
.dm-kpi { text-align: center; min-width: 80px; }
.dm-kpi-label { font-size: 0.65rem; color: var(--muted); letter-spacing: 0.1em; text-transform: uppercase; }
.dm-kpi-val {
    font-family: 'Barlow Condensed', sans-serif !important;
    font-size: 1.3rem; font-weight: 800; line-height: 1.1;
}
.dm-kpi-val.red    { color: #EF5350; }
.dm-kpi-val.orange { color: #FFA726; }
.dm-kpi-val.green  { color: #66BB6A; }
.dm-kpi-val.white  { color: var(--text); }
.dm-trend { font-size: 0.7rem; }
.dm-trend.up   { color: #66BB6A; }
.dm-trend.down { color: #EF5350; }
.dm-trend.flat { color: var(--muted); }
.dm-alerts { font-size: 0.72rem; color: #EF5350; min-width: 180px; }
.dm-forever { font-size: 0.7rem; color: var(--muted); text-align: center; min-width: 90px; }

/* ── UPLOAD AREA ── */
[data-testid="stFileUploader"] {
    background: var(--card) !important;
    border: 2px dashed var(--border) !important;
    border-radius: 12px !important;
    padding: 1rem !important;
}
[data-testid="stFileUploader"]:hover {
    border-color: var(--accent) !important;
}

/* ── BUTTONS ── */
.stDownloadButton > button {
    background: linear-gradient(135deg, var(--blue-mid), var(--blue-light)) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    letter-spacing: 0.03em !important;
    padding: 0.5rem 1.2rem !important;
}
.stDownloadButton > button:hover {
    background: linear-gradient(135deg, var(--blue-light), var(--accent)) !important;
}

/* ── TABS ── */
.stTabs [data-baseweb="tab-list"] {
    background: var(--card) !important;
    border-radius: 10px !important;
    border: 1px solid var(--border) !important;
    gap: 0 !important;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    color: var(--muted) !important;
    font-weight: 600 !important;
    font-size: 0.85rem !important;
    border-radius: 8px !important;
}
.stTabs [aria-selected="true"] {
    background: var(--blue-mid) !important;
    color: white !important;
}
.stTabs [data-baseweb="tab-panel"] {
    background: transparent !important;
    padding: 1rem 0 !important;
}

/* ── DATAFRAME ── */
[data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }

/* ── DIVIDER ── */
hr { border-color: var(--border) !important; margin: 1.5rem 0 !important; }

/* ── SIDEBAR LABELS ── */
.sidebar-section {
    font-family: 'Barlow Condensed', sans-serif !important;
    font-size: 0.7rem; font-weight: 700;
    letter-spacing: 0.15em; text-transform: uppercase;
    color: var(--accent); margin: 1.2rem 0 0.5rem;
    padding-bottom: 0.3rem;
    border-bottom: 1px solid var(--border);
}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────
#  CORE LOGIC
# ─────────────────────────────────────────────────────────────────

def load_and_filter(file_bytes, cfg):
    df = pd.read_excel(file_bytes, sheet_name=cfg["sheet_name"])
    valid = df[df["SHOP_CODE"].notna()].copy()
    latest = valid.sort_values(["YEAR", "MONTH"], ascending=False).iloc[0]
    month, year = int(latest["MONTH"]), int(latest["YEAR"])

    current = valid[(valid["MONTH"] == month) & (valid["YEAR"] == year)].copy()

    # Filtro zona A6
    mask = pd.Series(False, index=current.index)
    mask |= current["AMBASSADOR"].isin(cfg["ambassador_filter"])
    mask |= current["REGION"].isin(cfg["regioni_filter"])
    current = current[mask]

    # Mese precedente
    prev_month = month - 1 if month > 1 else 12
    prev_year  = year if month > 1 else year - 1
    prev = valid[(valid["MONTH"] == prev_month) & (valid["YEAR"] == prev_year)].copy()
    mask_prev = pd.Series(False, index=prev.index)
    mask_prev |= prev["AMBASSADOR"].isin(cfg["ambassador_filter"])
    mask_prev |= prev["REGION"].isin(cfg["regioni_filter"])
    prev = prev[mask_prev]

    meta = {"month": month, "year": year, "prev_month": prev_month, "prev_year": prev_year}
    return current, prev, meta


def compute_kpis(df, df_prev, cfg):
    agg = df.groupby(["REGION", "AREA_MANAGER", "DISTRICT_MANAGER"]).agg(
        TAM=("TAM", "sum"),
        Gross_Sales=("SR_PLUS_GROSS_SALES", "sum"),
        Net_Sales=("SR_PLUS_NET_SALES", "sum"),
        Forever_Active=("Activeforever", "sum"),
        Total_Stores=("TotalStore", "sum"),
        N_Negozi=("SHOP_CODE", "count"),
    ).reset_index()

    agg["Gross_%"] = (agg["Gross_Sales"] / agg["TAM"] * 100).round(1)
    agg["Net_%"]   = (agg["Net_Sales"]   / agg["TAM"] * 100).round(1)
    agg["Forever_%"] = (agg["Forever_Active"] / agg["Total_Stores"] * 100).round(1)

    # Trend
    if not df_prev.empty:
        prev_agg = df_prev.groupby("DISTRICT_MANAGER").agg(
            TAM_p=("TAM", "sum"),
            Gross_p=("SR_PLUS_GROSS_SALES", "sum"),
            Net_p=("SR_PLUS_NET_SALES", "sum"),
        ).reset_index()
        prev_agg["Gross_%_p"] = (prev_agg["Gross_p"] / prev_agg["TAM_p"] * 100).round(1)
        prev_agg["Net_%_p"]   = (prev_agg["Net_p"]   / prev_agg["TAM_p"] * 100).round(1)
        agg = agg.merge(prev_agg[["DISTRICT_MANAGER","Gross_%_p","Net_%_p"]], on="DISTRICT_MANAGER", how="left")
        agg["Gross_Δ"] = (agg["Gross_%"] - agg["Gross_%_p"]).round(1)
        agg["Net_Δ"]   = (agg["Net_%"]   - agg["Net_%_p"]).round(1)
    else:
        agg["Gross_Δ"] = None
        agg["Net_Δ"]   = None

    # Semaforo per ogni DM
    def status(row):
        sg = cfg["soglia_critica_gross"]
        sn = cfg["soglia_critica_net"]
        alerts = []
        if row["Gross_%"] < sg:      alerts.append(f"Gross {row['Gross_%']}% < {sg}%")
        if row["Net_%"]   < sn:      alerts.append(f"Net {row['Net_%']}% < {sn}%")
        if row["Forever_Active"] == 0: alerts.append("Forever = 0")
        
        # Fasce attenzione
        fasce = cfg.get("fasce", [])
        for fascia in fasce:
            lo, hi, label = fascia
            if lo <= row["Gross_%"] < hi:
                alerts.append(f"Gross in fascia {label}")
                break

        if not alerts:
            return "ok", []
        # Se critico su almeno un KPI obbligatorio
        critico = row["Gross_%"] < sg or row["Net_%"] < sn or row["Forever_Active"] == 0
        return ("critico" if critico else "attenzione"), alerts

    agg["status"], agg["alerts"] = zip(*agg.apply(status, axis=1))
    return agg.sort_values("Gross_%", ascending=False).reset_index(drop=True)


def generate_message(row, meta):
    mese_str = datetime(meta["year"], meta["month"], 1).strftime("%B %Y").capitalize()

    def trend_str(delta):
        if delta is None or (hasattr(delta, '__class__') and delta.__class__.__name__ == 'float' and pd.isna(delta)):
            return ""
        if delta > 0: return f"▲ +{delta}%"
        if delta < 0: return f"▼ {delta}%"
        return "= 0%"

    g_trend = trend_str(row.get("Gross_Δ"))
    n_trend = trend_str(row.get("Net_Δ"))
    forever_icon = "🔴" if row["Forever_Active"] == 0 else "♾️"
    alerts = row.get("alerts", [])
    alert_block = "\n".join([f"⚠️ {a}" for a in alerts]) if alerts else "✅ Tutti i KPI nella norma"

    return f"""📊 *AVANZAMENTO RELOAD – {mese_str}*
👤 *{row['DISTRICT_MANAGER']}*
📍 {row['REGION']} | AM: {row['AREA_MANAGER']}
─────────────────────────
🔵 *Gross (Lordo):*   {row['Gross_Sales']} / {row['TAM']}  →  *{row['Gross_%']}%*  {g_trend}
🟢 *Net (Netto):*     {row['Net_Sales']} / {row['TAM']}  →  *{row['Net_%']}%*  {n_trend}
{forever_icon}  *Reload Forever:*  {int(row['Forever_Active'])} / {int(row['Total_Stores'])} store  →  *{row['Forever_%']}%*
─────────────────────────
🏪 Negozi: {row['N_Negozi']}
{alert_block}"""


def build_excel_riepilogo(kpis, meta):
    wb = Workbook()
    ws = wb.active
    ws.title = "Riepilogo"
    BLUE   = PatternFill("solid", fgColor="003087")
    RED    = PatternFill("solid", fgColor="FFCDD2")
    ORANGE = PatternFill("solid", fgColor="FFE0B2")
    GREEN  = PatternFill("solid", fgColor="C8E6C9")
    BORDER = Border(left=Side(style="thin"), right=Side(style="thin"),
                    top=Side(style="thin"),  bottom=Side(style="thin"))
    CENTER = Alignment(horizontal="center", vertical="center")

    # Titolo
    ws.merge_cells("A1:N1")
    tc = ws["A1"]
    tc.value = f"AVANZAMENTO RELOAD – {meta['month']:02d}/{meta['year']} – CAMPANIA & PUGLIA (A6)"
    tc.font  = Font(bold=True, color="FFFFFF", size=12)
    tc.fill  = BLUE
    tc.alignment = CENTER
    ws.row_dimensions[1].height = 30

    headers = ["Regione","Area Manager","District Manager","N Negozi","TAM",
               "Gross Sales","Gross %","Gross Δ","Net Sales","Net %","Net Δ",
               "Forever Active","Total Stores","Forever %","Status"]
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=2, column=ci, value=h)
        c.font = Font(bold=True, color="FFFFFF", size=9)
        c.fill = PatternFill("solid", fgColor="0050B3")
        c.alignment = CENTER
        c.border = BORDER
    ws.row_dimensions[2].height = 22

    for ri, (_, row) in enumerate(kpis.iterrows(), 3):
        vals = [
            row["REGION"], row["AREA_MANAGER"], row["DISTRICT_MANAGER"],
            row["N_Negozi"], row["TAM"],
            row["Gross_Sales"], row["Gross_%"],
            (f"+{row['Gross_Δ']}%" if pd.notna(row.get("Gross_Δ")) and row["Gross_Δ"] > 0
             else (f"{row['Gross_Δ']}%" if pd.notna(row.get("Gross_Δ")) else "—")),
            row["Net_Sales"], row["Net_%"],
            (f"+{row['Net_Δ']}%" if pd.notna(row.get("Net_Δ")) and row["Net_Δ"] > 0
             else (f"{row['Net_Δ']}%" if pd.notna(row.get("Net_Δ")) else "—")),
            row["Forever_Active"], row["Total_Stores"], row["Forever_%"],
            row["status"].upper(),
        ]
        for ci, val in enumerate(vals, 1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.border = BORDER
            c.alignment = CENTER
            if ci == 7:  # Gross%
                gp = row["Gross_%"]
                if gp < 40:   c.fill = RED
                elif gp < 60: c.fill = ORANGE
                else:          c.fill = GREEN
            if ci == 10:  # Net%
                np_ = row["Net_%"]
                if np_ < 37:  c.fill = RED
                elif np_ < 50: c.fill = ORANGE
                else:          c.fill = GREEN
        ws.row_dimensions[ri].height = 18

    for col in ws.columns:
        ml = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(ml + 3, 38)
    ws.freeze_panes = "A3"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def build_excel_dm(row, df_raw, meta):
    wb = Workbook()
    ws = wb.active
    ws.title = "Dettaglio Store"
    BLUE   = PatternFill("solid", fgColor="003087")
    RED    = PatternFill("solid", fgColor="FFCDD2")
    GREEN  = PatternFill("solid", fgColor="C8E6C9")
    ORANGE = PatternFill("solid", fgColor="FFE0B2")
    BORDER = Border(left=Side(style="thin"), right=Side(style="thin"),
                    top=Side(style="thin"),  bottom=Side(style="thin"))
    CENTER = Alignment(horizontal="center")

    # Titolo
    ws.merge_cells("A1:J1")
    tc = ws["A1"]
    tc.value = f"{row['DISTRICT_MANAGER']} – {meta['month']:02d}/{meta['year']}"
    tc.font = Font(bold=True, color="FFFFFF", size=11)
    tc.fill = BLUE
    tc.alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 28

    # KPI summary
    ws["A2"] = "GROSS %"; ws["B2"] = f"{row['Gross_%']}%"
    ws["C2"] = "NET %";   ws["D2"] = f"{row['Net_%']}%"
    ws["E2"] = "FOREVER %"; ws["F2"] = f"{row['Forever_%']}%"
    ws["G2"] = "NEGOZI"; ws["H2"] = int(row["N_Negozi"])
    for cell in [ws["A2"], ws["C2"], ws["E2"], ws["G2"]]:
        cell.font = Font(bold=True, color="003087")
    ws.row_dimensions[2].height = 20

    # Dettaglio store
    sub = df_raw[df_raw["DISTRICT_MANAGER"] == row["DISTRICT_MANAGER"]].copy()
    sub["Gross_%"] = (sub["SR_PLUS_GROSS_SALES"] / sub["TAM"] * 100).round(1)
    sub["Net_%"]   = (sub["SR_PLUS_NET_SALES"]   / sub["TAM"] * 100).round(1)
    cols = ["STORE","CITY","PROVINCE_CODE","STORE_TYPE","TAM",
            "SR_PLUS_GROSS_SALES","Gross_%","SR_PLUS_NET_SALES","Net_%","Activeforever"]
    hdrs = ["Store","Città","Prov","Tipo","TAM","Gross","Gross%","Net","Net%","Forever"]
    for ci, h in enumerate(hdrs, 1):
        c = ws.cell(row=4, column=ci, value=h)
        c.font = Font(bold=True, color="FFFFFF", size=9)
        c.fill = PatternFill("solid", fgColor="0050B3")
        c.alignment = CENTER
        c.border = BORDER

    for ri, (_, sr) in enumerate(sub[cols].iterrows(), 5):
        for ci, val in enumerate(sr.values, 1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.border = BORDER
            c.alignment = CENTER
            if ci == 7 and isinstance(val, float):
                c.fill = RED if val < 40 else (ORANGE if val < 60 else GREEN)
            if ci == 9 and isinstance(val, float):
                c.fill = RED if val < 37 else (ORANGE if val < 50 else GREEN)
        ws.row_dimensions[ri].height = 17

    for col in ws.columns:
        ml = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(ml + 3, 35)
    ws.freeze_panes = "A5"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─────────────────────────────────────────────────────────────────
#  SIDEBAR – CONFIGURAZIONE SOGLIE
# ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div class="w3-title" style="font-size:1.3rem">⚙️ Configurazione</div>', unsafe_allow_html=True)
    st.markdown("---")

    st.markdown('<div class="sidebar-section">🎯 Soglie Critiche</div>', unsafe_allow_html=True)
    st.caption("Sotto queste soglie → 🔴 CRITICO")
    soglia_gross = st.slider("Gross % minima", 0, 100, 40, 1, help="Default: 40%")
    soglia_net   = st.slider("Net % minima",   0, 100, 37, 1, help="Default: 37%")
    forever_zero_critico = st.checkbox("Forever = 0 → critico", value=True)

    st.markdown('<div class="sidebar-section">⚠️ Fasce di Attenzione</div>', unsafe_allow_html=True)
    st.caption("Evidenzia DM in range specifici")

    use_fasce = st.checkbox("Abilita fasce personalizzate", value=False)
    fasce = []
    if use_fasce:
        n_fasce = st.number_input("Numero fasce", 1, 5, 2)
        for i in range(int(n_fasce)):
            st.markdown(f"**Fascia {i+1}**")
            col1, col2 = st.columns(2)
            lo = col1.number_input(f"Da %", 0, 100, 40 + i*10, key=f"lo_{i}")
            hi = col2.number_input(f"A %",  0, 100, 50 + i*10, key=f"hi_{i}")
            label = st.text_input(f"Etichetta", f"Fascia {i+1}", key=f"lab_{i}")
            fasce.append((lo, hi, label))

    st.markdown('<div class="sidebar-section">🗺️ Zona</div>', unsafe_allow_html=True)
    zona_label = st.text_input("Ambassador code", "A6")
    regioni_input = st.text_input("Regioni (separate da virgola)", "Campania, Puglia")
    regioni = [r.strip() for r in regioni_input.split(",") if r.strip()]

    st.markdown("---")
    st.markdown('<div class="sidebar-section">📖 Soglie attive</div>', unsafe_allow_html=True)
    st.markdown(f"""
    - 🔴 Gross critico: **< {soglia_gross}%**
    - 🔴 Net critico: **< {soglia_net}%**
    - 🔴 Forever zero: **{'Sì' if forever_zero_critico else 'No'}**
    - 🗺️ Zona: **{zona_label}** – {', '.join(regioni)}
    """)

# ─────────────────────────────────────────────────────────────────
#  HEADER
# ─────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="w3-header">
    <div class="w3-glow"></div>
    <div class="w3-title">WIND<span>3</span> RELOAD</div>
    <div class="w3-sub">Sales Performance Dashboard</div>
    <div class="w3-badge">📶 Ambassador {zona_label} &nbsp;|&nbsp; {' & '.join(regioni)}</div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────
#  UPLOAD
# ─────────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "📎 Carica il file Excel di avanzamento",
    type=["xlsx"],
    help="File Sales x Store – formato standard Wind3 Reload"
)

if uploaded is None:
    st.markdown("""
    <div style="text-align:center; padding:4rem 2rem; color:#7B97BF;">
        <div style="font-size:3rem; margin-bottom:1rem;">📂</div>
        <div style="font-family:'Barlow Condensed'; font-size:1.3rem; font-weight:700; margin-bottom:0.5rem;">
            Carica il file Excel per iniziare
        </div>
        <div style="font-size:0.85rem;">
            Formato atteso: <code>Avanzamento_DD-MM-YY__Sales_x_Store.xlsx</code>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ─────────────────────────────────────────────────────────────────
#  ANALISI
# ─────────────────────────────────────────────────────────────────
cfg = {
    "ambassador_filter": [zona_label],
    "regioni_filter": regioni,
    "soglia_critica_gross": soglia_gross,
    "soglia_critica_net": soglia_net,
    "forever_zero_critico": forever_zero_critico,
    "fasce": fasce,
    "sheet_name": "Sales x Store",
}

with st.spinner("Analisi in corso..."):
    try:
        df, df_prev, meta = load_and_filter(uploaded, cfg)
        kpis = compute_kpis(df, df_prev, cfg)
    except Exception as e:
        st.error(f"Errore nel caricamento: {e}")
        st.stop()

mese_str = datetime(meta["year"], meta["month"], 1).strftime("%B %Y").capitalize()
n_critici    = (kpis["status"] == "critico").sum()
n_attenzione = (kpis["status"] == "attenzione").sum()
n_ok         = (kpis["status"] == "ok").sum()
avg_gross    = kpis["Gross_%"].mean()
avg_net      = kpis["Net_%"].mean()

# ─────────────────────────────────────────────────────────────────
#  KPI TILES
# ─────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="kpi-row">
    <div class="kpi-tile accent">
        <div class="kpi-label">📅 Periodo</div>
        <div class="kpi-value accent" style="font-size:1.4rem">{mese_str}</div>
        <div class="kpi-sub">{len(df)} store analizzati</div>
    </div>
    <div class="kpi-tile {'red' if avg_gross < soglia_gross else 'green'}">
        <div class="kpi-label">📊 Gross medio</div>
        <div class="kpi-value {'red' if avg_gross < soglia_gross else 'green'}">{avg_gross:.1f}%</div>
        <div class="kpi-sub">Soglia critica: {soglia_gross}%</div>
    </div>
    <div class="kpi-tile {'red' if avg_net < soglia_net else 'green'}">
        <div class="kpi-label">📈 Net medio</div>
        <div class="kpi-value {'red' if avg_net < soglia_net else 'green'}">{avg_net:.1f}%</div>
        <div class="kpi-sub">Soglia critica: {soglia_net}%</div>
    </div>
    <div class="kpi-tile red">
        <div class="kpi-label">🔴 Critici</div>
        <div class="kpi-value red">{n_critici}</div>
        <div class="kpi-sub">su {len(kpis)} District Manager</div>
    </div>
    <div class="kpi-tile orange">
        <div class="kpi-label">⚠️ Attenzione</div>
        <div class="kpi-value" style="color:#FFA726">{n_attenzione}</div>
        <div class="kpi-sub">fasce personalizzate</div>
    </div>
    <div class="kpi-tile green">
        <div class="kpi-label">✅ OK</div>
        <div class="kpi-value green">{n_ok}</div>
        <div class="kpi-sub">sopra le soglie</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────
#  TABS
# ─────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["📋 Riepilogo DM", "📄 Dettaglio & Messaggi", "⬇️ Download"])

# ── TAB 1: RIEPILOGO ────────────────────────────────────────────
with tab1:
    # Filtri rapidi
    col_f1, col_f2, col_f3 = st.columns([1, 1, 2])
    with col_f1:
        filtro_status = st.selectbox("Filtra per status", ["Tutti", "🔴 Critici", "⚠️ Attenzione", "✅ OK"])
    with col_f2:
        filtro_regione = st.selectbox("Filtra per regione", ["Tutte"] + sorted(kpis["REGION"].unique().tolist()))
    with col_f3:
        filtro_am = st.selectbox("Filtra per Area Manager", ["Tutti"] + sorted(kpis["AREA_MANAGER"].unique().tolist()))

    filtered = kpis.copy()
    if filtro_status == "🔴 Critici":    filtered = filtered[filtered["status"] == "critico"]
    elif filtro_status == "⚠️ Attenzione": filtered = filtered[filtered["status"] == "attenzione"]
    elif filtro_status == "✅ OK":        filtered = filtered[filtered["status"] == "ok"]
    if filtro_regione != "Tutte": filtered = filtered[filtered["REGION"] == filtro_regione]
    if filtro_am != "Tutti":      filtered = filtered[filtered["AREA_MANAGER"] == filtro_am]

    st.markdown(f'<div class="section-title">District Manager — {len(filtered)} risultati</div>', unsafe_allow_html=True)

    def color_kpi(val, soglia, tipo="gross"):
        s = soglia_gross if tipo == "gross" else soglia_net
        if val < s: return "red"
        if val >= s + 20: return "green"
        return "white"

    def trend_html(delta):
        if delta is None or pd.isna(delta): return '<span class="dm-trend flat">—</span>'
        if delta > 0: return f'<span class="dm-trend up">▲ +{delta}%</span>'
        if delta < 0: return f'<span class="dm-trend down">▼ {delta}%</span>'
        return '<span class="dm-trend flat">= 0%</span>'

    cards_html = '<div class="dm-grid">'
    for _, row in filtered.iterrows():
        status_cls = row["status"]
        alerts_txt = " | ".join(row["alerts"]) if row["alerts"] else "✅ OK"
        g_col = color_kpi(row["Gross_%"], soglia_gross, "gross")
        n_col = color_kpi(row["Net_%"],   soglia_net,   "net")
        f_icon = "🔴" if row["Forever_Active"] == 0 else "♾️"
        g_trend = trend_html(row.get("Gross_Δ"))
        n_trend = trend_html(row.get("Net_Δ"))

        cards_html += f"""
        <div class="dm-card {status_cls}">
            <div class="dm-name">
                {row['DISTRICT_MANAGER']}
                <div class="dm-region">{row['REGION']} · {row['AREA_MANAGER']}</div>
            </div>
            <div class="dm-kpi">
                <div class="dm-kpi-label">Gross</div>
                <div class="dm-kpi-val {g_col}">{row['Gross_%']}%</div>
                {g_trend}
            </div>
            <div class="dm-kpi">
                <div class="dm-kpi-label">Net</div>
                <div class="dm-kpi-val {n_col}">{row['Net_%']}%</div>
                {n_trend}
            </div>
            <div class="dm-forever">
                {f_icon} <strong>{int(row['Forever_Active'])}</strong>/{int(row['Total_Stores'])} store<br>
                <span style="font-size:0.8rem">{row['Forever_%']}%</span>
            </div>
            <div class="dm-alerts">{alerts_txt if alerts_txt != '✅ OK' else ''}</div>
        </div>"""

    cards_html += "</div>"
    st.markdown(cards_html, unsafe_allow_html=True)

# ── TAB 2: DETTAGLIO & MESSAGGI ─────────────────────────────────
with tab2:
    dm_selected = st.selectbox(
        "Seleziona District Manager",
        kpis["DISTRICT_MANAGER"].tolist(),
        format_func=lambda x: f"{'🔴' if kpis[kpis['DISTRICT_MANAGER']==x]['status'].values[0]=='critico' else ('⚠️' if kpis[kpis['DISTRICT_MANAGER']==x]['status'].values[0]=='attenzione' else '✅')} {x}"
    )

    row_sel = kpis[kpis["DISTRICT_MANAGER"] == dm_selected].iloc[0]

    col_msg, col_store = st.columns([1, 2])

    with col_msg:
        st.markdown('<div class="section-title">💬 Messaggio pronto</div>', unsafe_allow_html=True)
        msg = generate_message(row_sel, meta)
        st.text_area("", msg, height=320, label_visibility="collapsed")
        st.download_button(
            "⬇️ Scarica messaggio .txt",
            data=msg.encode("utf-8"),
            file_name=f"{dm_selected.replace(' ','_')}_{meta['month']:02d}_{meta['year']}.txt",
            mime="text/plain",
        )

    with col_store:
        st.markdown('<div class="section-title">🏪 Dettaglio Store</div>', unsafe_allow_html=True)
        sub = df[df["DISTRICT_MANAGER"] == dm_selected].copy()
        sub["Gross_%"] = (sub["SR_PLUS_GROSS_SALES"] / sub["TAM"] * 100).round(1)
        sub["Net_%"]   = (sub["SR_PLUS_NET_SALES"]   / sub["TAM"] * 100).round(1)
        display_cols = {
            "STORE": "Store", "CITY": "Città", "PROVINCE_CODE": "Prov",
            "TAM": "TAM", "SR_PLUS_GROSS_SALES": "Gross", "Gross_%": "Gross%",
            "SR_PLUS_NET_SALES": "Net", "Net_%": "Net%", "Activeforever": "Forever"
        }
        sub_display = sub[list(display_cols.keys())].rename(columns=display_cols).sort_values("Gross%")

        def color_gross(val):
            if val < soglia_gross: return "background-color:#FFCDD2; color:#B71C1C"
            if val >= soglia_gross + 20: return "background-color:#C8E6C9; color:#1B5E20"
            return "background-color:#FFE0B2; color:#E65100"

        def color_net(val):
            if val < soglia_net: return "background-color:#FFCDD2; color:#B71C1C"
            if val >= soglia_net + 20: return "background-color:#C8E6C9; color:#1B5E20"
            return "background-color:#FFE0B2; color:#E65100"

        styled = sub_display.style.applymap(color_gross, subset=["Gross%"]).applymap(color_net, subset=["Net%"])
        st.dataframe(styled, use_container_width=True, height=320)

        buf_dm = build_excel_dm(row_sel, df, meta)
        st.download_button(
            "⬇️ Scarica Excel DM",
            data=buf_dm,
            file_name=f"{dm_selected.replace(' ','_')}_{meta['month']:02d}_{meta['year']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ── TAB 3: DOWNLOAD ─────────────────────────────────────────────
with tab3:
    st.markdown('<div class="section-title">⬇️ Download pacchetto completo</div>', unsafe_allow_html=True)

    col_d1, col_d2 = st.columns(2)

    with col_d1:
        st.markdown("**📊 Riepilogo Generale Excel**")
        st.caption("Tabella completa tutti i DM con KPI, trend e alert — colorata per soglie")
        buf_rie = build_excel_riepilogo(kpis, meta)
        st.download_button(
            "⬇️ riepilogo_generale.xlsx",
            data=buf_rie,
            file_name=f"riepilogo_generale_{meta['month']:02d}_{meta['year']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with col_d2:
        st.markdown("**💬 Tutti i messaggi .txt**")
        st.caption("Un messaggio per ogni DM, pronti da copiare su WhatsApp")
        sep = "=" * 60 + "\n"
        all_msgs = sep.join([generate_message(row, meta) + "\n\n" for _, row in kpis.iterrows()])
        st.download_button(
            "⬇️ messaggi_whatsapp.txt",
            data=all_msgs.encode("utf-8"),
            file_name=f"messaggi_whatsapp_{meta['month']:02d}_{meta['year']}.txt",
            mime="text/plain",
            use_container_width=True,
        )

    st.markdown("---")
    st.markdown("**📦 Pacchetto ZIP — Excel singolo per ogni DM**")
    st.caption("Un file Excel per ciascun District Manager con dettaglio store")
    if st.button("🗜️ Genera ZIP", use_container_width=True):
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for _, row in kpis.iterrows():
                dm_buf = build_excel_dm(row, df, meta)
                safe = row["DISTRICT_MANAGER"].replace("/","_").replace("\\","_").replace(" ","_")[:40]
                zf.writestr(f"{safe}_{meta['month']:02d}_{meta['year']}.xlsx", dm_buf.read())
        zip_buf.seek(0)
        st.download_button(
            "⬇️ Scarica ZIP con tutti gli Excel",
            data=zip_buf,
            file_name=f"wind3_reload_DM_{meta['month']:02d}_{meta['year']}.zip",
            mime="application/zip",
            use_container_width=True,
        )

    st.markdown("---")
    st.markdown(f"""
    <div style="text-align:center; color:#7B97BF; font-size:0.75rem; padding:1rem;">
        Wind3 Reload Dashboard · Ambassador {zona_label} · {mese_str}<br>
        Soglie attive — Gross: {soglia_gross}% · Net: {soglia_net}%
    </div>
    """, unsafe_allow_html=True)
