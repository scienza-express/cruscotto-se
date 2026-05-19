"""
Scienza Express — Cruscotto Finanziario
Avvio:  streamlit run cruscotto_SE.py
Dipendenze: pip install streamlit pandas plotly openpyxl xlrd
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from pathlib import Path
import io

# ─────────────────────────────────────────────
# CONFIGURAZIONE PAGINA
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Cruscotto SE",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# MAPPATURA FORNITORI → CATEGORIA
# ─────────────────────────────────────────────
SUPPLIER_MAP = {
    "02283810501": "Prod. - Stampa b/n",
    "00954720678": "Prod. - Stampa colori",
    "01252290232": "Prod. - Stampa colori",
    "04061550275": "Prod. - Stampa colori",
    "01231830322": "Prod. - Grafica Bambini",
    "02489351201": "Prod. - Grafica Bambini",
    "13788560962": "Prod. - Grafica Bambini",
    "00771010576": "Prod. - Grafica Saggi",
    "01423560323": "Prod. - Grafica Saggi",
    "01677060442": "Prod. - Grafica Saggi",
    "03175010648": "Prod. - Grafica Saggi",
    "03266880800": "Prod. - Grafica Saggi",
    "07913550963": "Prod. - Grafica Saggi",
    "08194360015": "Prod. - Grafica Saggi",
    "08729160153": "Prod. - Grafica Saggi",
    "12454590964": "Prod. - Grafica Saggi",
    "03654040132": "Prod. - Grafica Saggi",
    "03931300044": "Prod. - Gestione/Redazione",
    "01274100310": "Prod. - Gestione/Redazione",
    "12736460010": "Prod. - Gestione/Redazione",
    "00996430328": "Prod. - Sito e app",
    "04275401208": "Prod. - Sito e app",
    "01502750977": "Prod. - Logistica",
    "06059191004": "Prod. - Logistica",
    "01417730320": "Comm. - Promozione",
    "01388260323": "Comm. - Comunicazione",
    "01635350935": "Comm. - Comunicazione",
    "02406690376": "Comm. - Comunicazione",
    "17022531002": "Comm. - Comunicazione",
    "04874660261": "Comm. - Comunicazione",
    "01196430407": "Comm. - Eventi",
    "02684631209": "Comm. - Eventi",
    "04933280481": "Comm. - Eventi",
    "02507120729": "Comm. - Distribuzione",
    "01541850382": "Comm. - Viaggi e Rapp.",
    "01691390692": "Comm. - Viaggi e Rapp.",
    "01729520302": "Comm. - Viaggi e Rapp.",
    "01986120382": "Comm. - Viaggi e Rapp.",
    "02028520449": "Comm. - Viaggi e Rapp.",
    "02468160227": "Comm. - Viaggi e Rapp.",
    "02538160033": "Comm. - Viaggi e Rapp.",
    "03017710348": "Comm. - Viaggi e Rapp.",
    "03044961203": "Comm. - Viaggi e Rapp.",
    "04367650969": "Comm. - Viaggi e Rapp.",
    "04530520404": "Comm. - Viaggi e Rapp.",
    "04581470285": "Comm. - Viaggi e Rapp.",
    "05006880966": "Comm. - Viaggi e Rapp.",
    "05403151003": "Comm. - Viaggi e Rapp.",
    "03553671201": "Comm. - Viaggi e Rapp.",
    "07492450965": "Comm. - Viaggi e Rapp.",
    "07516911000": "Comm. - Viaggi e Rapp.",
    "09771701001": "Comm. - Viaggi e Rapp.",
    "11805830012": "Comm. - Viaggi e Rapp.",
    "10158500966": "Comm. - Viaggi e Rapp.",
    "03939010165": "Amm. - Segreteria",
    "01322610328": "Amm. - Spese generali",
    "01367260328": "Amm. - Spese generali",
    "01472410933": "Amm. - Spese generali",
    "01807420938": "Amm. - Spese generali",
    "03763520966": "Amm. - Spese generali",
    "04570150278": "Amm. - Spese generali",
    "07501560150": "Amm. - Spese generali",
    "08106710158": "Amm. - Spese generali",
    "08539010010": "Amm. - Spese generali",
    "09381880013": "Amm. - Spese generali",
    "09732530150": "Amm. - Spese generali",
    "12581270969": "Amm. - Spese generali",
    "13378520152": "Amm. - Spese generali",
    "12878470157": "Amm. - Spese generali",
    "01573850516": "Amm. - Spese generali",
    "01691720468": "Amm. - Spese generali",
    "3336483DH":   "Amm. - Spese generali",
    "01097780322": "Nuovi progetti - RISE",
}

CATEGORIE_ORDINATE = [
    "Prod. - Stampa b/n",
    "Prod. - Stampa colori",
    "Prod. - Grafica Bambini",
    "Prod. - Grafica Saggi",
    "Prod. - Gestione/Redazione",
    "Prod. - Sito e app",
    "Prod. - Logistica",
    "Comm. - Promozione",
    "Comm. - Comunicazione",
    "Comm. - Distribuzione",
    "Comm. - Eventi",
    "Comm. - Viaggi e Rapp.",
    "Amm. - Segreteria",
    "Amm. - Spese generali",
    "Nuovi progetti - RISE",
    "Amm. - Non classificate",
]

BUDGET_ANNUO = {
    "Prod. - Stampa b/n":         0,
    "Prod. - Stampa colori":      0,
    "Prod. - Grafica Bambini":    15205,
    "Prod. - Grafica Saggi":      16494,
    "Prod. - Gestione/Redazione": 10790,
    "Prod. - Sito e app":         12630,
    "Prod. - Logistica":          16417,
    "Comm. - Promozione":         20000,
    "Comm. - Comunicazione":      10000,
    "Comm. - Distribuzione":       0,
    "Comm. - Eventi":             10000,
    "Comm. - Viaggi e Rapp.":      6000,
    "Amm. - Segreteria":           7000,
    "Amm. - Spese generali":      14000,
    "Nuovi progetti - RISE":           0,
    "Amm. - Non classificate":         0,
}

# ─────────────────────────────────────────────
# DATI CDA PRE-CALCOLATI DAI RENDICONTI PDF
# (lordo, IVA esente, sconto 60% già applicato)
# ─────────────────────────────────────────────
CDA_MENSILE = {
    "2026-01": 16012.60,
    "2026-02": 12001.40,
    "2026-03": 20803.70,
    "2026-04":  6067.56,
}

# ─────────────────────────────────────────────
# UTILITY
# ─────────────────────────────────────────────
def norm_piva(x):
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s.replace(".", "").isdigit():
        return str(int(float(s))).zfill(11)
    return s

def load_fatture_ricevute(uploaded_file) -> pd.DataFrame:
    engine = "xlrd" if str(uploaded_file.name).endswith(".xls") else "openpyxl"
    df = pd.read_excel(uploaded_file, engine=engine)
    df["Data emissione"] = pd.to_datetime(df["Data emissione"], errors="coerce")
    df["piva_norm"] = df["Partita IVA"].apply(norm_piva)
    df["Categoria"] = df["piva_norm"].map(SUPPLIER_MAP).fillna("Amm. - Non classificate")
    return df

def load_fatture_emesse(uploaded_file) -> pd.DataFrame:
    engine = "xlrd" if str(uploaded_file.name).endswith(".xls") else "openpyxl"
    df = pd.read_excel(uploaded_file, engine=engine)
    df["Data emissione"] = pd.to_datetime(df["Data emissione"], errors="coerce")
    return df

def classifica_ricavo(cliente):
    if pd.isna(cliente):
        return "Altro"
    cl = str(cliente).lower()
    if any(k in cl for k in ["consorzio", "cda"]):
        return "CDA"
    if any(k in cl for k in ["ministero", "ministeri", "e-shop", "eshop", "shop"]):
        return "Escludi"
    return "Diretta"

def filtra_trimestre(df, anno, trimestre):
    mesi = {1: (1,3), 2: (4,6), 3: (7,9), 4: (10,12)}
    m_start, m_end = mesi[trimestre]
    mask = (
        (df["Data emissione"].dt.year == anno) &
        (df["Data emissione"].dt.month >= m_start) &
        (df["Data emissione"].dt.month <= m_end)
    )
    return df[mask]

def cda_da_mensile(anno, trimestre):
    mesi = {1: [1,2,3], 2: [4,5,6], 3: [7,8,9], 4: [10,11,12]}
    totale = 0.0
    for m in mesi[trimestre]:
        key = f"{anno}-{m:02d}"
        totale += CDA_MENSILE.get(key, 0.0)
    return totale

def fmt_eur(v):
    return f"€ {v:,.0f}".replace(",", ".")

def delta_badge(eff, bgt):
    if bgt == 0:
        return ""
    diff = eff - bgt
    pct = diff / bgt * 100
    if diff >= 0:
        return f"▲ {fmt_eur(abs(diff))} (+{pct:.1f}%)"
    else:
        return f"▼ {fmt_eur(abs(diff))} ({pct:.1f}%)"

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.image("https://placehold.co/160x40/0070C0/FFFFFF?text=Scienza+Express",
             use_container_width=True)
    st.markdown("## Impostazioni")

    anno = st.selectbox("Anno", [2026, 2025], index=0)
    trimestre = st.selectbox("Trimestre", [1, 2, 3, 4],
                             format_func=lambda x: f"Q{x}", index=0)

    st.markdown("---")
    st.markdown("### Carica dati aggiornati")
    ft_emesse_file = st.file_uploader(
        "Fatture emesse (.xls / .xlsx)", type=["xls", "xlsx"], key="emesse")
    ft_ricevute_file = st.file_uploader(
        "Fatture ricevute (.xls / .xlsx)", type=["xls", "xlsx"], key="ricevute")

    st.markdown("---")
    st.markdown("### Ricavo sito e-commerce")
    sito_val = st.number_input(
        f"Sito Q{trimestre} {anno} (€)",
        min_value=0.0, value=10000.0, step=100.0,
        help="Inserisci il dato reale dalle statistiche WooCommerce")

    st.markdown("---")
    st.markdown("### Budget anno (modificabile)")
    b_cda      = st.number_input("Budget CDA lordo (€)",     value=87000, step=1000)
    b_diretta  = st.number_input("Budget Diretta (€)",        value=55000, step=1000)
    b_sito     = st.number_input("Budget Sito (€)",           value=40000, step=1000)

    st.markdown("---")
    st.markdown("### Stagionalità Q")
    st.caption("Coefficiente da applicare al budget annuo per stimare il budget di trimestre.")
    stag = {1: 0.35, 2: 0.20, 3: 0.10, 4: 0.35}
    coeff = st.slider(
        f"Quota Q{trimestre} sui costi (%)",
        min_value=5, max_value=60,
        value=int(stag.get(trimestre, 0.25) * 100),
        step=5, format="%d%%") / 100.0

# ─────────────────────────────────────────────
# CARICA / CALCOLA DATI
# ─────────────────────────────────────────────

# --- COSTI ---
if ft_ricevute_file:
    df_ric = load_fatture_ricevute(ft_ricevute_file)
    df_ric_q = filtra_trimestre(df_ric, anno, trimestre)
    costi_df = (df_ric_q.groupby("Categoria")["Imponibile"]
                .sum().reset_index()
                .rename(columns={"Imponibile": "Effettivo"}))
else:
    # Dati Q1 2026 pre-calcolati
    costi_precomp = {
        "Prod. - Stampa b/n":         21212.52,
        "Prod. - Stampa colori":           0.00,
        "Prod. - Grafica Bambini":      2100.00,
        "Prod. - Grafica Saggi":        3536.00,
        "Prod. - Gestione/Redazione":   6265.00,
        "Prod. - Sito e app":           4280.28 + 2424.02,
        "Prod. - Logistica":            5406.63,
        "Comm. - Promozione":              0.00,
        "Comm. - Comunicazione":        4635.97,
        "Comm. - Distribuzione":           0.00,
        "Comm. - Eventi":               3080.00,
        "Comm. - Viaggi e Rapp.":       1318.89,
        "Amm. - Segreteria":               0.00,
        "Amm. - Spese generali":         883.76,
        "Nuovi progetti - RISE":           0.00,
        "Amm. - Non classificate":         0.00,
    }
    costi_df = pd.DataFrame([
        {"Categoria": k, "Effettivo": v}
        for k, v in costi_precomp.items()
    ])

# Aggiungi budget
costi_df["Budget annuo"] = costi_df["Categoria"].map(BUDGET_ANNUO).fillna(0)
costi_df["Budget Q"]     = costi_df["Budget annuo"] * coeff
costi_df["Scostamento"]  = costi_df["Effettivo"] - costi_df["Budget Q"]

# Ordina per CATEGORIE_ORDINATE
costi_df["_ord"] = costi_df["Categoria"].apply(
    lambda x: CATEGORIE_ORDINATE.index(x) if x in CATEGORIE_ORDINATE else 99)
costi_df = costi_df.sort_values("_ord").drop(columns="_ord")

tot_costi = costi_df["Effettivo"].sum()

# --- RICAVI ---
cda_eff = cda_da_mensile(anno, trimestre)

if ft_emesse_file:
    df_emi = load_fatture_emesse(ft_emesse_file)
    df_emi_q = filtra_trimestre(df_emi, anno, trimestre)
    df_emi_q = df_emi_q.copy()
    df_emi_q["Canale"] = df_emi_q["Cliente"].apply(classifica_ricavo)
    diretta_eff = df_emi_q[df_emi_q["Canale"] == "Diretta"]["Imponibile"].sum()
else:
    diretta_eff = 14237.24  # Q1 2026 pre-calcolato

sito_eff = sito_val

tot_ricavi = cda_eff + diretta_eff + sito_eff
mol = tot_ricavi - tot_costi
mol_pct = mol / tot_ricavi if tot_ricavi else 0

# Budget Q ricavi (stagionalità uniforme ÷4 per ricavi, personalizzabile)
bgt_cda_q    = b_cda    / 4
bgt_dir_q    = b_diretta / 4
bgt_sito_q   = b_sito   / 4
bgt_ric_tot  = bgt_cda_q + bgt_dir_q + bgt_sito_q
bgt_cos_tot  = costi_df["Budget Q"].sum()
bgt_mol      = bgt_ric_tot - bgt_cos_tot

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
st.title(f"📚 Cruscotto Finanziario — Q{trimestre} {anno}")

if not ft_ricevute_file or not ft_emesse_file:
    st.info(
        "📂 Stai visualizzando i **dati pre-caricati Q1 2026**. "
        "Carica i file XLS nella barra laterale per aggiornare il trimestre.",
        icon="ℹ️")

# ─────────────────────────────────────────────
# KPI CARDS
# ─────────────────────────────────────────────
c1, c2, c3, c4 = st.columns(4)

def kpi(col, label, value, budget, positive_is_good=True):
    diff = value - budget
    pct  = diff / budget * 100 if budget else 0
    if positive_is_good:
        color = "normal" if diff >= 0 else "inverse"
    else:
        color = "inverse" if diff >= 0 else "normal"
    col.metric(label, fmt_eur(value),
               f"{'+' if diff>=0 else ''}{pct:.1f}% vs budget Q",
               delta_color=color)

kpi(c1, "💰 Ricavi totali",    tot_ricavi, bgt_ric_tot,  True)
kpi(c2, "💸 Costi totali",     tot_costi,  bgt_cos_tot,  False)
kpi(c3, "📊 MOL",              mol,        bgt_mol,       True)
c4.metric("📈 % MOL su ricavi", f"{mol_pct:.1%}",
          f"Budget: {bgt_mol/bgt_ric_tot:.1%}" if bgt_ric_tot else "—")

st.divider()

# ─────────────────────────────────────────────
# RICAVI — grafico e tabella
# ─────────────────────────────────────────────
col_ric, col_cos = st.columns(2)

with col_ric:
    st.subheader("Ricavi per canale")

    ricavi_df = pd.DataFrame({
        "Canale":    ["CDA", "Diretta", "Sito"],
        "Effettivo": [cda_eff, diretta_eff, sito_eff],
        "Budget Q":  [bgt_cda_q, bgt_dir_q, bgt_sito_q],
    })

    fig_ric = go.Figure()
    fig_ric.add_bar(
        name="Effettivo", x=ricavi_df["Canale"], y=ricavi_df["Effettivo"],
        marker_color="#0070C0",
        text=ricavi_df["Effettivo"].apply(lambda v: f"€{v:,.0f}"),
        textposition="outside")
    fig_ric.add_bar(
        name="Budget Q", x=ricavi_df["Canale"], y=ricavi_df["Budget Q"],
        marker_color="#D9E1F2",
        text=ricavi_df["Budget Q"].apply(lambda v: f"€{v:,.0f}"),
        textposition="outside")
    fig_ric.update_layout(
        barmode="group", height=320,
        margin=dict(t=20, b=20, l=10, r=10),
        legend=dict(orientation="h", y=-0.15),
        yaxis_tickprefix="€", yaxis_tickformat=",.0f")
    st.plotly_chart(fig_ric, use_container_width=True)

    # tabella ricavi
    ricavi_df["Scostamento"] = ricavi_df["Effettivo"] - ricavi_df["Budget Q"]
    st.dataframe(
        ricavi_df.style
            .format({"Effettivo": "€{:,.0f}", "Budget Q": "€{:,.0f}", "Scostamento": "€{:,.0f}"})
            .applymap(lambda v: "color: #C00000" if isinstance(v, (int, float)) and v < 0 else "",
                      subset=["Scostamento"]),
        use_container_width=True, hide_index=True)

# ─────────────────────────────────────────────
# COSTI — grafico orizzontale
# ─────────────────────────────────────────────
with col_cos:
    st.subheader("Costi per categoria")

    df_plot = costi_df[costi_df["Effettivo"] > 0].copy()
    df_plot = df_plot.sort_values("Effettivo")

    fig_cos = go.Figure()
    fig_cos.add_bar(
        name="Budget Q", y=df_plot["Categoria"], x=df_plot["Budget Q"],
        orientation="h", marker_color="#D9E1F2")
    fig_cos.add_bar(
        name="Effettivo", y=df_plot["Categoria"], x=df_plot["Effettivo"],
        orientation="h", marker_color="#0070C0",
        text=df_plot["Effettivo"].apply(lambda v: f"€{v:,.0f}"),
        textposition="outside")
    fig_cos.update_layout(
        barmode="overlay", height=320,
        margin=dict(t=20, b=20, l=10, r=10),
        legend=dict(orientation="h", y=-0.15),
        xaxis_tickprefix="€", xaxis_tickformat=",.0f")
    st.plotly_chart(fig_cos, use_container_width=True)

st.divider()

# ─────────────────────────────────────────────
# TABELLA COSTI DETTAGLIO
# ─────────────────────────────────────────────
st.subheader("Dettaglio costi")

costi_show = costi_df[["Categoria", "Effettivo", "Budget Q", "Scostamento"]].copy()

def color_scost(v):
    if isinstance(v, (int, float)):
        return "color: #C00000; font-weight:bold" if v > 0 else (
               "color: #375623; font-weight:bold" if v < 0 else "")
    return ""

st.dataframe(
    costi_show.style
        .format({
            "Effettivo":   "€ {:,.2f}",
            "Budget Q":    "€ {:,.2f}",
            "Scostamento": "€ {:,.2f}",
        })
        .applymap(color_scost, subset=["Scostamento"]),
    use_container_width=True, hide_index=True, height=420)

# Nota su scostamento costi
st.caption(
    "⚠️ Scostamento costi positivo = si è speso **più** del budget di trimestre (da analizzare). "
    "Negativo = si è speso meno (possibile anticipo di spese o risparmio).")

st.divider()

# ─────────────────────────────────────────────
# RIEPILOGO FINALE
# ─────────────────────────────────────────────
st.subheader("Riepilogo P&L semplificato")

col_a, col_b = st.columns(2)

with col_a:
    pl_data = {
        "Voce": [
            "CDA (sell-in lordo)",
            "Diretta (imponibile)",
            "Sito / e-commerce",
            "TOTALE RICAVI",
            "Totale costi operativi",
            "MOL",
        ],
        "Effettivo": [
            cda_eff, diretta_eff, sito_eff,
            tot_ricavi, tot_costi, mol,
        ],
        "Budget Q": [
            bgt_cda_q, bgt_dir_q, bgt_sito_q,
            bgt_ric_tot, bgt_cos_tot, bgt_mol,
        ],
    }
    pl_df = pd.DataFrame(pl_data)
    pl_df["Scostamento"] = pl_df["Effettivo"] - pl_df["Budget Q"]

    bold_rows = [3, 5]
    def style_pl(row):
        styles = [""] * len(row)
        if row.name in bold_rows:
            styles = ["font-weight: bold; background-color: #EBF0F8"] * len(row)
        if row.name == 5:
            styles = ["font-weight: bold; background-color: #D6E4BC"] * len(row)
        return styles

    st.dataframe(
        pl_df.style
            .format({
                "Effettivo":   "€ {:,.2f}",
                "Budget Q":    "€ {:,.2f}",
                "Scostamento": "€ {:,.2f}",
            })
            .apply(style_pl, axis=1),
        use_container_width=True, hide_index=True)

with col_b:
    # Waterfall MOL
    fig_wf = go.Figure(go.Waterfall(
        name="MOL",
        orientation="v",
        measure=["relative", "relative", "relative", "total", "relative", "total"],
        x=["CDA", "Diretta", "Sito", "Ricavi", "Costi", "MOL"],
        y=[cda_eff, diretta_eff, sito_eff, 0, -tot_costi, 0],
        connector={"line": {"color": "rgb(63, 63, 63)"}},
        increasing={"marker": {"color": "#0070C0"}},
        decreasing={"marker": {"color": "#C00000"}},
        totals={"marker": {"color": "#375623"}},
        text=[fmt_eur(cda_eff), fmt_eur(diretta_eff), fmt_eur(sito_eff),
              fmt_eur(tot_ricavi), fmt_eur(-tot_costi), fmt_eur(mol)],
        textposition="outside",
    ))
    fig_wf.update_layout(
        title="Waterfall ricavi → MOL",
        height=380,
        margin=dict(t=40, b=20, l=10, r=10),
        showlegend=False,
        yaxis_tickprefix="€", yaxis_tickformat=",.0f",
    )
    st.plotly_chart(fig_wf, use_container_width=True)

st.divider()

# ─────────────────────────────────────────────
# NOTA METODOLOGICA
# ─────────────────────────────────────────────
with st.expander("📌 Note metodologiche"):
    st.markdown(f"""
- **CDA** — sell-in lordo (prezzi di copertina, IVA esente) dai rendiconti mensili CDA.
  Dati Q1 2026 pre-caricati; per altri trimestri aggiorna il dizionario `CDA_MENSILE` nel codice
  o carica il file XLS delle fatture emesse (le righe CDA vengono sommate automaticamente).
- **Diretta** — imponibile delle fatture emesse, esclusi clienti Ministeri, e-shop e CDA.
- **Sito** — inserire manualmente dalla barra laterale (dato WooCommerce).
- **Costi soci** (Guido, Daniele, Massimo) — non inclusi; entreranno nei trimestri di competenza.
- **Budget Q** — budget annuo × {coeff:.0%} (coefficiente di stagionalità impostato nella barra laterale).
  Per i ricavi si usa sempre ÷ 4 (stagionalità uniforme); per i costi usa il cursore.
- **Stagionalità costi stampa** — il Q1 è tipicamente il trimestre più pesante (~35% del totale).
  La voce "Stampa b/n Q1 2026" include ordini Monteserra per il semestre.
- **Mappatura fornitori** — basata sul file `riclassificazione_fornitori_2025.xlsx`;
  i fornitori non riconosciuti finiscono in "Amm. - Non classificate".
""")

st.caption("Scienza Express Edizioni · Cruscotto interno · aggiornato automaticamente al caricamento dei file")
