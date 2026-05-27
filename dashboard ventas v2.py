"""
Dashboard Comercial - Catena Zapata
Santi Cattaneo - Gerencia Nacional de Ventas
"""

import pandas as pd
import numpy as np
from dash import Dash, dcc, html, Input, Output, State, no_update, callback_context
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import warnings, webbrowser, threading, os, sys, io, base64, unicodedata, hashlib, json
from datetime import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

warnings.filterwarnings('ignore')
print("[STARTUP] Módulo iniciando — importando dependencias", flush=True)

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                     Table, TableStyle, HRFlowable, KeepTogether)
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors as rl_colors
    from reportlab.lib.units import cm
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
    from reportlab.pdfgen import canvas as _rl_canvas
    PDF_AVAILABLE = True
    print("[STARTUP] ReportLab OK", flush=True)
except ImportError as _re:
    PDF_AVAILABLE = False
    print(f"[STARTUP] ReportLab no disponible: {_re}", flush=True)

EXCEL_PATH    = "/Volumes/santi 2T/dashboard/Dashboard.xlsx"
DRIVE_FOLDER  = "Dashboard"
DRIVE_FILENAME = "Dashboard.xlsx"
DRIVE_FILE_ID_FALLBACK = "1tPmdaJuR0pRMEjRkZ58JopLxdKztCXfn"
CREDS_PATH    = "/Users/santi/Downloads/catena-dashboard-fe7dc08d5408.json"
PORT = 8050

C = {
    'bg':      '#0D0D0D',
    'surf':    '#161616',
    'surf2':   '#1E1E1E',
    'border':  '#2A2A2A',
    'gold':    '#C9A84C',
    'red':     '#C0392B',
    'green':   '#27AE60',
    'text':    '#F0EDE8',
    'muted':   '#888888',
}
FONT = "Georgia, 'Times New Roman', serif"
MONO = "'Courier New', monospace"

# ── Carga ─────────────────────────────────────────────────────────────────────

_DRIVE_SERVICE_CACHE = None

def _get_drive_service():
    global _DRIVE_SERVICE_CACHE
    if _DRIVE_SERVICE_CACHE is None:
        google_creds_json = os.environ.get('GOOGLE_CREDENTIALS')
        if google_creds_json:
            creds = service_account.Credentials.from_service_account_info(
                json.loads(google_creds_json),
                scopes=['https://www.googleapis.com/auth/drive.readonly']
            )
        elif os.path.exists(CREDS_PATH):
            creds = service_account.Credentials.from_service_account_file(
                CREDS_PATH,
                scopes=['https://www.googleapis.com/auth/drive.readonly']
            )
        else:
            return None, None
        # cache_discovery=False evita guardar el documento de discovery en memoria
        _DRIVE_SERVICE_CACHE = build('drive', 'v3', credentials=creds, cache_discovery=False)
    service = _DRIVE_SERVICE_CACHE
    folder_res = service.files().list(
        q=f"name='{DRIVE_FOLDER}' and mimeType='application/vnd.google-apps.folder' and trashed=false",
        spaces='drive', fields='files(id)').execute()
    folders = folder_res.get('files', [])
    if folders:
        folder_id = folders[0]['id']
        q = f"name='{DRIVE_FILENAME}' and '{folder_id}' in parents and trashed=false"
    else:
        q = f"name='{DRIVE_FILENAME}' and trashed=false"
    file_res = service.files().list(
        q=q, spaces='drive', fields='files(id, modifiedTime)', orderBy='modifiedTime desc').execute()
    files = file_res.get('files', [])
    if files:
        return service, files[0]
    # fallback al ID hardcodeado si la búsqueda no encuentra nada
    try:
        meta = service.files().get(fileId=DRIVE_FILE_ID_FALLBACK,
                                   fields='id,modifiedTime').execute()
        return service, meta
    except Exception:
        return None, None

def get_drive_modified_time():
    """Retorna el modifiedTime del archivo en Drive sin descargarlo."""
    try:
        _, f = _get_drive_service()
        return f['modifiedTime'] if f else None
    except Exception:
        return None

def load_data():
    print("[STARTUP] load_data() iniciado", flush=True)
    service, f = _get_drive_service()
    if service and f:
        print(f"[STARTUP] Descargando desde Drive: {f.get('id','?')}", flush=True)
        request = service.files().get_media(fileId=f['id'])
        buf = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
            if status:
                print(f"[STARTUP] Descarga {int(status.progress()*100)}%", flush=True)
        buf.seek(0)
        print("[STARTUP] Descarga completa, parseando Excel", flush=True)
        xl = pd.ExcelFile(buf)
    else:
        if not os.path.exists(EXCEL_PATH):
            raise FileNotFoundError(f"Archivo no encontrado en Drive ni localmente: {EXCEL_PATH}")
        xl = pd.ExcelFile(EXCEL_PATH)
    dfs = {}
    for sheet in xl.sheet_names:
        df = xl.parse(sheet)
        df.columns = [str(c).strip() for c in df.columns]
        for col in df.select_dtypes(include='object').columns:
            try:
                df[col] = df[col].str.strip()
            except:
                pass
        dfs[sheet.strip().lower()] = df  # siempre minúscula

    # eliminar fila de totales de pend (Familia Producto == 'Total')
    if 'pend' in dfs and 'Familia Producto' in dfs['pend'].columns:
        dfs['pend'] = dfs['pend'][
            dfs['pend']['Familia Producto'].str.upper().str.strip() != 'TOTAL'
        ]

    return dfs

def month_cols(df):
    meses = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']
    return [c for c in df.columns if c in meses]

def get_ind(df, ind, groups, meses_sel=None):
    mc = month_cols(df)
    cols = groups + mc + ['Total']
    sub = df[df['Indicadores'] == ind][cols].copy()
    for c in mc + ['Total']:
        sub[c] = pd.to_numeric(sub[c], errors='coerce')
    if meses_sel:
        valid = [m for m in meses_sel if m in sub.columns]
        if valid:
            sub['Total'] = sub[valid].sum(axis=1)
    return sub

print("[STARTUP] Cargando datos...", flush=True)
try:
    DFS = load_data()
    DATA_OK = True
    print(f"[STARTUP] Datos cargados OK — hojas: {list(DFS.keys())}", flush=True)
except Exception as _e:
    print(f"[WARNING] No se pudo cargar el archivo al iniciar: {_e}", flush=True)
    DFS = {}
    DATA_OK = False
MC  = month_cols(DFS.get('x flia', pd.DataFrame()))
FAMILIAS       = sorted(DFS['x flia']['flia'].unique().tolist()) if DATA_OK else []
REPRESENTANTES = sorted(DFS['x repre']['Vendedor'].unique().tolist()) if DATA_OK else []
CANALES        = sorted([str(x) for x in DFS['x flia x canal']['Canal'].dropna().unique().tolist()]) if DATA_OK else []

# ── Credenciales ───────────────────────────────────────────────────────────────

def _norm(s):
    """Normaliza a ASCII lowercase para comparar usernames sin tildes."""
    return unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode('ascii').lower().strip()

def _make_pin(name, used):
    h = int(hashlib.md5(name.encode('utf-8')).hexdigest(), 16)
    pin = str(1000 + (h % 9000))
    while pin in used:
        h = (h * 6364136223846793005 + 1) & 0xFFFFFFFFFFFF
        pin = str(1000 + (h % 9000))
    return pin

_SKIP_REPS = {'directos casa interior', '69i'}

_used_pins = {'piso3', 'tio', 'sofi'}
_used_keys = {'jefe', 'florio', 'jorge'}
USERS = {
    'jefe':   {'password': 'piso3', 'role': 'admin',  'repre': None,
               'display_name': 'SANTIAGO CATTANEO', 'title': 'Jefe Nacional de Ventas'},
    'florio': {'password': 'tio',   'role': 'viewer', 'repre': None,
               'display_name': 'TOMAS FLORIO', 'title': 'Gerente Nacional de Ventas'},
    'jorge':  {'password': 'sofi',  'role': 'viewer', 'repre': None,
               'display_name': 'JORGE ESTEBAN', 'title': ''},
}
for _r in REPRESENTANTES:
    if _norm(_r) in _SKIP_REPS:
        continue
    _base = _norm(_r).split()[0].strip('.,;:-')  # primera palabra sin puntuación
    _key  = _base
    _suf  = 2
    while _key in _used_keys:             # evitar colisiones
        _key = f"{_base}{_suf}"; _suf += 1
    _used_keys.add(_key)
    _pin = _make_pin(_r, _used_pins)
    _used_pins.add(_pin)
    USERS[_key] = {'password': _pin, 'role': 'vendedor', 'repre': _r,
                   'display_name': _r.upper(), 'title': ''}

PL = dict(
    paper_bgcolor='rgba(0,0,0,0)',
    plot_bgcolor='rgba(0,0,0,0)',
    font=dict(family=FONT, color=C['text'], size=11),
    title_font=dict(family=FONT, color=C['gold'], size=13),
    xaxis=dict(gridcolor=C['border'], linecolor=C['border'], tickfont=dict(size=11)),
    yaxis=dict(gridcolor=C['border'], linecolor=C['border'], tickfont=dict(size=11)),
    margin=dict(l=40, r=20, t=40, b=40),
    legend=dict(bgcolor='rgba(0,0,0,0)', font=dict(size=9)),
    hoverlabel=dict(bgcolor=C['surf2'], bordercolor=C['gold'], font=dict(family=FONT)),
    bargap=0.28,
)

VAR_CAP = 500  # límite de visualización para variaciones extremas (líneas nuevas)

def _var(v, sin_ant=False, cajas=None):
    """Formatea variación %; '—' si NaN/inf; muestra cajas si no había ventas el año anterior."""
    if sin_ant:
        return f"{int(cajas):,} caj" if cajas is not None else '—'
    if v is None or (isinstance(v, float) and (np.isnan(v) or np.isinf(v))):
        return '—'
    return f"{v:+.0f}%"

def _cap_var(v, sin_ant=False):
    """Valor capeado para la barra; las líneas sin año anterior o extremas se capean a ±VAR_CAP."""
    if sin_ant or v is None or (isinstance(v, float) and (np.isnan(v) or np.isinf(v))):
        return VAR_CAP  # barra corta positiva para items sin año anterior
    return max(-VAR_CAP, min(VAR_CAP, v))

# ── Figuras ────────────────────────────────────────────────────────────────────

def fig_flia_ranking(flia_sel, canal_sel, meses_sel=None):
    """Dual-panel horizontal: familias ordenadas por volumen + variación % | participación %."""
    try:
        if canal_sel:
            df = DFS['x flia x canal']
            df = df[df['Canal'] == canal_sel]
            act = get_ind(df, 'Año Actual Cajas', ['flia'], meses_sel).groupby('flia')[['Total']].sum().reset_index()
            ant = get_ind(df, 'Año Anterior Cajas', ['flia'], meses_sel).groupby('flia')[['Total']].sum().reset_index()
        else:
            act = get_ind(DFS['x flia'], 'Año Actual Cajas', ['flia'], meses_sel)
            ant = get_ind(DFS['x flia'], 'Año Anterior Cajas', ['flia'], meses_sel)
        m = act.merge(ant, on='flia', how='outer', suffixes=('_a','_b'))
        m['Total_a'] = m['Total_a'].fillna(0)
        m['Total_b'] = m['Total_b'].fillna(0)
        m['var']    = (m['Total_a'] - m['Total_b']) / m['Total_b'].replace(0, np.nan) * 100
        m['sin_ant'] = m['Total_b'] == 0
        m = m[m['Total_a'] > 0]  # solo familias con ventas este año
        m['part']   = m['Total_a'] / m['Total_a'].sum() * 100
        if flia_sel:
            m = m[m['flia'] == flia_sel]
        m = m.sort_values('Total_a', ascending=True)

        col_var = [C['gold'] if row['sin_ant'] else (C['green'] if (pd.notna(row['var']) and row['var'] >= 0) else C['red'])
                   for _, row in m.iterrows()]
        col_vol = [C['gold'] if row['sin_ant'] else (C['green'] if (pd.notna(row['var']) and row['var'] >= 0) else C['red'])
                   for _, row in m.iterrows()]

        height = max(260, len(m) * 24 + 70)

        fig = make_subplots(
            rows=1, cols=2,
            column_widths=[0.38, 0.62],
            subplot_titles=['Variación % vs Año Anterior', 'Participación % y Cajas Año Actual'],
            horizontal_spacing=0.12,
        )

        # Panel izq: var% — capeado para no distorsionar escala
        var_display = [_cap_var(row['var'], row['sin_ant']) for _, row in m.iterrows()]
        var_labels  = [_var(row['var'], row['sin_ant'], row['Total_a']) for _, row in m.iterrows()]
        var_tpos    = ['inside' if row['sin_ant'] else 'outside' for _, row in m.iterrows()]
        var_tcol    = ['#FFFFFF' if row['sin_ant'] else C['text'] for _, row in m.iterrows()]
        hover_var   = [f"+{row['Total_a']:.0f} caj (sin año anterior)" if row['sin_ant'] else f"{row['var']:+.0f}%" for _, row in m.iterrows()]
        fig.add_trace(go.Bar(
            y=m['flia'], x=var_display, orientation='h',
            marker_color=col_var,
            text=var_labels,
            textposition=var_tpos,
            textfont=dict(size=12, color=var_tcol),
            customdata=hover_var,
            hovertemplate='<b>%{y}</b><br>Var: %{customdata}<extra></extra>',
        ), row=1, col=1)

        # Panel der: barras horizontales ordenadas, color por crecimiento/caída
        hover_der = [f"+{int(row['Total_a']):,} caj (sin año anterior)" if row['sin_ant']
                     else f"Anterior: {int(row['Total_b']):,}\nVar: {row['var']:+.0f}%"
                     for _, row in m.iterrows()]
        fig.add_trace(go.Bar(
            y=m['flia'], x=m['Total_a'], orientation='h',
            marker_color=col_vol,
            text=[f"{int(v):,} caj" for v in m['Total_a']],
            textposition='inside',
            insidetextanchor='middle',
            textfont=dict(size=10, color='#FFFFFF'),
            customdata=np.column_stack([m['part'].round(0), hover_der]),
            hovertemplate='<b>%{y}</b><br>Actual: %{x:,.0f} caj  (%{customdata[0]:.0f}%)<br>%{customdata[1]}<extra></extra>',
        ), row=1, col=2)
        # % participación afuera de las barras de cajas
        fig.add_trace(go.Scatter(
            y=m['flia'], x=m['Total_a'],
            text=[f"{p:.0f}%" for p in m['part']],
            mode='text', textposition='middle right',
            textfont=dict(size=14, color=C['text']),
            showlegend=False,
            hoverinfo='skip',
        ), row=1, col=2)

        fig.add_vline(x=0, line_width=1, line_color=C['muted'], line_dash='dot', col=1)

        pl_r = {k: v for k, v in PL.items() if k not in ('xaxis','yaxis','margin')}
        fig.update_layout(
            **pl_r,
            title='Familias — Variación y Participación',
            height=height,
            showlegend=False,
            margin=dict(l=10, r=90, t=46, b=24),
        )
        fig.update_xaxes(tickfont=dict(size=10), gridcolor=C['border'])
        fig.update_yaxes(tickfont=dict(size=10))
        fig.update_xaxes(ticksuffix='%', range=[-VAR_CAP * 1.3, VAR_CAP * 1.3], row=1, col=1)
        fig.update_xaxes(ticksuffix=' caj', row=1, col=2)
        return fig
    except Exception as e:
        return go.Figure().update_layout(**PL, title=f'Error flia ranking: {e}')

def fig_evolucion(flia_sel, repre_sel=None, canal_sel=None, meses_sel=None):
    try:
        if repre_sel:
            df = DFS['x repre']
            act = get_ind(df[df['Vendedor'] == repre_sel], 'Año Actual Cajas', ['flia'], meses_sel)
            ant = get_ind(df[df['Vendedor'] == repre_sel], 'Año Anterior Cajas', ['flia'], meses_sel)
        elif canal_sel:
            df = DFS['x flia x canal']
            act = get_ind(df[df['Canal'] == canal_sel], 'Año Actual Cajas', ['flia'], meses_sel)
            ant = get_ind(df[df['Canal'] == canal_sel], 'Año Anterior Cajas', ['flia'], meses_sel)
        else:
            df = DFS['x flia']
            act = get_ind(df, 'Año Actual Cajas', ['flia'], meses_sel)
            ant = get_ind(df, 'Año Anterior Cajas', ['flia'], meses_sel)
        if flia_sel:
            act = act[act['flia'] == flia_sel]
            ant = ant[ant['flia'] == flia_sel]
        mc_use = [m for m in MC if not meses_sel or m in meses_sel]
        palette = px.colors.qualitative.Set2
        fig = go.Figure()
        for i, (_, row) in enumerate(act.iterrows()):
            color = palette[i % len(palette)]
            vals = [row[m] for m in mc_use if pd.notna(row.get(m))]
            mok  = [m for m in mc_use if pd.notna(row.get(m))]
            fig.add_trace(go.Scatter(x=mok, y=vals, name=row['flia'], mode='lines+markers',
                                     line=dict(color=color, width=2),
                                     hovertemplate=f"{row['flia']}: %{{y:,.0f}}<extra></extra>"))
            row_ant = ant[ant['flia'] == row['flia']]
            if not row_ant.empty:
                vals_a = [row_ant.iloc[0][m] for m in mok if pd.notna(row_ant.iloc[0].get(m))]
                fig.add_trace(go.Scatter(x=mok[:len(vals_a)], y=vals_a, showlegend=False,
                                         mode='lines', line=dict(color=color, width=1, dash='dot'),
                                         opacity=0.35,
                                         hovertemplate=f"{row['flia']} ant: %{{y:,.0f}}<extra></extra>"))
        fig.update_layout(**PL, title='Evolución Mensual por Familia', height=300)
        return fig
    except Exception as e:
        return go.Figure().update_layout(**PL, title=f'Error: {e}')

def fig_ranking_ejecutivo(flia_sel=None, repre_sel=None, canal_sel=None, meses_sel=None):
    """Ranking de representantes y familias por var% — reemplaza el heatmap."""
    try:
        # ── Representantes ──
        df_repre = DFS['x repre']
        if repre_sel:
            df_repre = df_repre[df_repre['Vendedor'] == repre_sel]
        if flia_sel:
            df_repre = df_repre[df_repre['flia'] == flia_sel]
        if meses_sel:
            av = get_ind(df_repre, 'Año Actual Cajas', ['Vendedor','flia'], meses_sel)
            bv = get_ind(df_repre, 'Año Anterior Cajas', ['Vendedor','flia'], meses_sel)
            av_tot = av.groupby('Vendedor')['Total'].sum().reset_index()
            bv_tot = bv.groupby('Vendedor')['Total'].sum().reset_index()
            # outer para no perder reps que solo aparecen en un año
            rv_tot = av_tot.merge(bv_tot, on='Vendedor', how='outer', suffixes=('_a','_b'))
            rv_tot['Total_a'] = rv_tot['Total_a'].fillna(0)
            rv_tot['Total_b'] = rv_tot['Total_b'].fillna(0)
            rv_tot['sin_ant'] = rv_tot['Total_b'] == 0
            rv_tot['var'] = (rv_tot['Total_a'] - rv_tot['Total_b']) / rv_tot['Total_b'].replace(0, np.nan) * 100
            rv_tot = rv_tot[rv_tot['Total_a'] > 0]  # solo reps con ventas este año
            rv_tot = rv_tot.sort_values('var')
            rep_df = rv_tot.rename(columns={'var': 'Total_var', 'Total_a': 'Total_vol'}).reset_index(drop=True)
        else:
            rv = get_ind(df_repre, 'Var %', ['Vendedor','flia'])
            av = get_ind(df_repre, 'Año Actual Cajas', ['Vendedor','flia'])
            rv_tot = rv.groupby('Vendedor')['Total'].mean().mul(100).reset_index()
            av_tot = av.groupby('Vendedor')['Total'].sum().reset_index()
            # left desde av_tot para no perder reps sin Var % (nuevos)
            rep_df = av_tot.merge(rv_tot, on='Vendedor', how='left', suffixes=('_vol','_var'))
            rep_df = rep_df.rename(columns={'Total_vol': 'Total_vol', 'Total_var': 'Total_var'})
            rep_df['Total_var'] = rep_df['Total_var'].fillna(0)
            rep_df['sin_ant'] = rep_df['Total_var'] == 0
            rep_df = rep_df.sort_values('Total_var').reset_index(drop=True)

        sin_ant_rep = rep_df['sin_ant'] if 'sin_ant' in rep_df.columns else pd.Series([False] * len(rep_df))
        col_rep = [C['gold'] if sa else (C['green'] if (pd.notna(v) and v >= 0) else C['red'])
                   for v, sa in zip(rep_df['Total_var'], sin_ant_rep)]

        # ── Familias ──
        df_flia = DFS['x flia']
        if flia_sel:
            df_flia = df_flia[df_flia['flia'] == flia_sel]
        if meses_sel:
            fa_act = get_ind(df_flia, 'Año Actual Cajas', ['flia'], meses_sel)
            fa_ant = get_ind(df_flia, 'Año Anterior Cajas', ['flia'], meses_sel)
            fam_m = fa_act.merge(fa_ant, on='flia', how='outer', suffixes=('_a','_b'))
            fam_m['Total_a'] = fam_m['Total_a'].fillna(0)
            fam_m['Total_b'] = fam_m['Total_b'].fillna(0)
            fam_m['sin_ant']   = fam_m['Total_b'] == 0
            fam_m['Total_var'] = (fam_m['Total_a'] - fam_m['Total_b']) / fam_m['Total_b'].replace(0, np.nan) * 100
            fam_m = fam_m[fam_m['Total_a'] > 0]
            fam_m = fam_m.sort_values('Total_var').reset_index(drop=True)
            fam_df = fam_m.rename(columns={'Total_a': 'Total_vol'})
        else:
            fv = get_ind(df_flia, 'Var %', ['flia'])
            fa = get_ind(df_flia, 'Año Actual Cajas', ['flia'])
            # left desde fa para no perder familias sin Var % (nuevas)
            fam_df = fa[['flia','Total']].merge(fv[['flia','Total']], on='flia', how='left', suffixes=('_vol','_var'))
            fam_df['Total_var'] = fam_df['Total_var'].fillna(0) * 100
            fam_df['sin_ant']   = fam_df['Total_var'] == 0
            fam_df = fam_df.sort_values('Total_var').reset_index(drop=True)

        sin_ant_fam = fam_df['sin_ant'] if 'sin_ant' in fam_df.columns else pd.Series([False] * len(fam_df))
        col_fam = [C['gold'] if sa else (C['green'] if (pd.notna(v) and v >= 0) else C['red'])
                   for v, sa in zip(fam_df['Total_var'], sin_ant_fam)]

        fig = make_subplots(
            rows=1, cols=2,
            subplot_titles=['Representantes — Var % vs Año Anterior',
                            'Familias — Var % vs Año Anterior'],
            column_widths=[0.55, 0.45],
        )

        # Panel izq — representantes (sin variaciones extremas esperadas aquí)
        rep_sin_ant = rep_df.get('sin_ant', pd.Series([False] * len(rep_df)))
        rep_vol_list = list(rep_df['Total_vol'])
        rep_tpos = ['inside' if sa else 'outside' for sa in rep_sin_ant]
        rep_tcol = ['#FFFFFF' if sa else C['text'] for sa in rep_sin_ant]
        fig.add_trace(go.Bar(
            y=rep_df['Vendedor'].str[:22],
            x=[_cap_var(v, sa) for v, sa in zip(rep_df['Total_var'], rep_sin_ant)],
            orientation='h',
            marker_color=col_rep,
            textposition=rep_tpos,
            text=[_var(v, sa, cj) for v, sa, cj in zip(rep_df['Total_var'], rep_sin_ant, rep_vol_list)],
            textfont=dict(size=12, color=rep_tcol),
            customdata=rep_df['Total_vol'],
            hovertemplate='<b>%{y}</b><br>Cajas: %{customdata:,.0f}<extra></extra>',
        ), row=1, col=1)

        # Panel der — familias
        fam_vol_list = list(fam_df['Total_vol'])
        fam_tpos = ['inside' if sa else 'outside' for sa in sin_ant_fam]
        fam_tcol = ['#FFFFFF' if sa else C['text'] for sa in sin_ant_fam]
        fig.add_trace(go.Bar(
            y=fam_df['flia'].str[:18],
            x=[_cap_var(v, sa) for v, sa in zip(fam_df['Total_var'], sin_ant_fam)],
            orientation='h',
            marker_color=col_fam,
            textposition=fam_tpos,
            text=[_var(v, sa, cj) for v, sa, cj in zip(fam_df['Total_var'], sin_ant_fam, fam_vol_list)],
            textfont=dict(size=12, color=fam_tcol),
            customdata=fam_df['Total_vol'],
            hovertemplate='<b>%{y}</b><br>Cajas: %{customdata:,.0f}<extra></extra>',
        ), row=1, col=2)

        # Línea de cero en ambos paneles
        for col in [1, 2]:
            fig.add_vline(x=0, line_width=1, line_color=C['muted'], line_dash='dot', col=col)

        n_rep = len(rep_df)
        n_fam = len(fam_df)
        height = max(280, max(n_rep, n_fam) * 22 + 70)

        pl_r = {k: v for k, v in PL.items() if k not in ('xaxis','yaxis','margin')}
        fig.update_layout(
            **pl_r,
            title='Ranking Ejecutivo — Variación % vs Año Anterior',
            height=height,
            showlegend=False,
            margin=dict(l=10, r=80, t=46, b=24),
        )
        fig.update_xaxes(ticksuffix='%', tickfont=dict(size=10), gridcolor=C['border'],
                         range=[-VAR_CAP * 1.3, VAR_CAP * 1.3])
        fig.update_yaxes(tickfont=dict(size=10))
        return fig
    except Exception as e:
        return go.Figure().update_layout(**PL, title=f'Error ranking: {e}')

def fig_repre_ranking(flia_sel, canal_sel, repre_sel=None, meses_sel=None, solo=False):
    try:
        if canal_sel:
            df = DFS['x repre x canal']
            df = df[df['Canal'] == canal_sel]
            act = get_ind(df, 'Año Actual Cajas', ['Vendedor','flia'], meses_sel)
            ant = get_ind(df, 'Año Anterior Cajas', ['Vendedor','flia'], meses_sel)
        else:
            act = get_ind(DFS['x repre'], 'Año Actual Cajas', ['Vendedor','flia'], meses_sel)
            ant = get_ind(DFS['x repre'], 'Año Anterior Cajas', ['Vendedor','flia'], meses_sel)
        if flia_sel:
            act = act[act['flia'] == flia_sel]
            ant = ant[ant['flia'] == flia_sel]
        a = act.groupby('Vendedor')['Total'].sum().reset_index()
        b = ant.groupby('Vendedor')['Total'].sum().reset_index()
        m = a.merge(b, on='Vendedor', how='outer', suffixes=('_a','_b'))
        m['Total_a'] = m['Total_a'].fillna(0)
        m['Total_b'] = m['Total_b'].fillna(0)
        m = m[m['Total_a'] > 0]
        if solo and repre_sel:
            m = m[m['Vendedor'] == repre_sel]
        m['var']     = (m['Total_a'] - m['Total_b']) / m['Total_b'].replace(0, np.nan) * 100
        m['sin_ant'] = m['Total_b'] == 0
        m['part']    = m['Total_a'] / m['Total_a'].sum() * 100
        # Ordenar: mayor volumen arriba → invertir para que quede bien en horizontal (plotly dibuja de abajo a arriba)
        m = m.sort_values('Total_a', ascending=True)
        nombres = m['Vendedor'].str[:22]
        n = len(m)
        altura = max(320, n * 38 + 80)

        # Colores var%
        def _col_var_h(row):
            if repre_sel:
                if row['Vendedor'] != repre_sel:
                    return 'rgba(200,200,200,0.12)'
            if row['sin_ant']:      return C['gold']
            if pd.isna(row['var']): return C['muted']
            return C['green'] if row['var'] >= 0 else C['red']

        def _col_vol_h(v):
            if repre_sel:
                return C['gold'] if v == repre_sel else 'rgba(201,168,76,0.18)'
            return C['gold']

        col_var = [_col_var_h(r) for _, r in m.iterrows()]
        col_vol = [_col_vol_h(v) for v in m['Vendedor']]

        var_display = [_cap_var(r['var'], r['sin_ant']) for _, r in m.iterrows()]
        var_text    = [_var(r['var'], r['sin_ant'], r['Total_a']) for _, r in m.iterrows()]

        fig = make_subplots(
            rows=1, cols=2,
            column_widths=[0.38, 0.62],
            subplot_titles=['Variación % vs Año Anterior', 'Cajas Año Actual'],
            shared_yaxes=True,
        )

        # ── Panel izquierdo: Var% horizontal divergente ──
        fig.add_trace(go.Bar(
            y=nombres,
            x=var_display,
            orientation='h',
            marker_color=col_var,
            marker_line_width=0,
            text=var_text,
            textposition='outside',
            textfont=dict(size=11, color=C['text'], family=MONO),
            cliponaxis=False,
            hovertemplate='<b>%{y}</b><br>Var: %{text}<extra></extra>',
        ), row=1, col=1)
        # Línea de cero
        fig.add_vline(x=0, line_width=1, line_color=C['muted'], line_dash='dot', row=1, col=1)

        # ── Panel derecho: Cajas + participación ──
        fig.add_trace(go.Bar(
            y=nombres,
            x=m['Total_a'],
            orientation='h',
            marker_color=col_vol,
            marker_line_width=0,
            text=[f"{int(v):,}" for v in m['Total_a']],
            textposition='inside',
            insidetextanchor='middle',
            textfont=dict(size=10, color='#FFFFFF', family=MONO),
            hovertemplate='<b>%{y}</b><br>%{x:,.0f} cajas<extra></extra>',
        ), row=1, col=2)
        # % participación fuera de la barra, alineado a la derecha
        fig.add_trace(go.Scatter(
            y=nombres,
            x=m['Total_a'],
            text=[f"  {p:.0f}%" for p in m['part']],
            mode='text',
            textposition='middle right',
            textfont=dict(size=11, color=C['text']),
            showlegend=False,
            hoverinfo='skip',
        ), row=1, col=2)

        title = f'Representantes — Ranking | {repre_sel}' if repre_sel else 'Representantes — Ranking y Variación'
        _pl = {k: v for k, v in PL.items() if k != 'margin'}
        fig.update_layout(
            **_pl, title=title, height=altura, showlegend=False,
            margin=dict(l=10, r=60, t=46, b=20),
        )
        # Eje X izquierdo: rango centrado en 0
        cap = VAR_CAP * 1.25
        fig.update_xaxes(range=[-cap, cap], tickfont=dict(size=9), row=1, col=1)
        # Eje Y: nombres de representante
        fig.update_yaxes(tickfont=dict(size=11), automargin=True)
        # Eje X derecho: sin cero forzado, margen para el % afuera
        fig.update_xaxes(tickfont=dict(size=9), row=1, col=2)
        return fig
    except Exception as e:
        return go.Figure().update_layout(**PL, title=f'Error: {e}')

def fig_canal_mix(flia_sel, repre_sel, canal_sel=None, meses_sel=None):
    """Participación % por canal — gráfico de barras horizontales limpio."""
    try:
        if repre_sel:
            df = DFS['x repre x canal']
            df = df[df['Vendedor'] == repre_sel]
            act = get_ind(df, 'Año Actual Cajas', ['Canal','flia'], meses_sel)
            ant_df = DFS['x repre x canal'][DFS['x repre x canal']['Vendedor'] == repre_sel]
            ant = get_ind(ant_df, 'Año Anterior Cajas', ['Canal','flia'], meses_sel)
        else:
            df = DFS['x flia x canal']
            act = get_ind(df, 'Año Actual Cajas', ['Canal','flia'], meses_sel)
            ant = get_ind(df, 'Año Anterior Cajas', ['Canal','flia'], meses_sel)
        if flia_sel:
            act = act[act['flia'] == flia_sel]
            ant = ant[ant['flia'] == flia_sel]
        if canal_sel:
            act = act[act['Canal'] == canal_sel]
            ant = ant[ant['Canal'] == canal_sel]
        agg_a = act.groupby('Canal')['Total'].sum().reset_index()
        agg_b = ant.groupby('Canal')['Total'].sum().reset_index()
        m = agg_a.merge(agg_b, on='Canal', suffixes=('_a','_b'))
        m = m[m['Total_a'] > 0]
        m = m[~m['Canal'].str.upper().str.strip().isin(['TRAVEL RETAIL'])]
        m = m.sort_values('Total_a', ascending=True)
        total = m['Total_a'].sum()
        m['pct'] = m['Total_a'] / total * 100
        m['var'] = (m['Total_a'] - m['Total_b']) / m['Total_b'].replace(0, np.nan) * 100

        palette = ['#4E79A7','#F28E2B','#E15759','#76B7B2','#59A14F',
                   '#EDC948','#B07AA1','#FF9DA7','#9C755F','#BAB0AC']
        colors = [palette[i % len(palette)] for i in range(len(m))]

        total_ant    = m['Total_b'].sum()
        m['pct_ant'] = (m['Total_b'] / total_ant * 100).fillna(0) if total_ant > 0 else 0
        m['delta_pp'] = m['pct'] - m['pct_ant']

        x_rmax = m['pct'].max() + 16

        fig = go.Figure()

        # Barras año actual
        fig.add_trace(go.Bar(
            y=m['Canal'], x=m['pct'], orientation='h',
            marker_color=colors,
            marker_line_width=0,
            text=[f"  {p:.0f}%  ({int(v):,} caj)" for p, v in zip(m['pct'], m['Total_a'])],
            textposition='inside', insidetextanchor='start',
            textfont=dict(size=11, color='#FFFFFF', family='Helvetica'),
            customdata=list(zip(m['Total_a'], m['var'].fillna(float('nan')), m['pct_ant'])),
            hovertemplate='<b>%{y}</b><br>Act: %{x:.0f}%  ·  %{customdata[0]:,.0f} caj'
                          '<br>Ant: %{customdata[2]:.0f}%<extra></extra>',
        ))

        # Delta pp: tres trazas con color según signo
        _BLUE = '#4E79A7'
        _delta_cfg = [
            (m['delta_pp'] >  0.4, C['green'], '▲'),
            (m['delta_pp'] < -0.4, C['red'],   '▼'),
            (m['delta_pp'].abs() <= 0.4, _BLUE, '='),
        ]
        x_delta = m['pct'].max() + 2
        for mask, color, sym in _delta_cfg:
            sub = m[mask]
            if sub.empty:
                continue
            fig.add_trace(go.Scatter(
                y=sub['Canal'],
                x=[x_delta] * len(sub),
                mode='text',
                text=[f"{sym} {d:+.0f}pp" for d in sub['delta_pp']],
                textposition='middle right',
                textfont=dict(size=13, color=color, family='Helvetica'),
                hoverinfo='skip',
                showlegend=False,
            ))

        subtitulo = repre_sel or 'Región'
        pl_m = {k: v for k, v in PL.items() if k not in ('xaxis','yaxis','margin')}
        altura = max(200, len(m) * 36 + 60)
        fig.update_layout(
            **pl_m,
            title=f'Participación por Canal — {subtitulo}',
            height=altura,
            margin=dict(l=10, r=10, t=40, b=20),
            showlegend=False,
            xaxis=dict(showticklabels=False, showgrid=False, zeroline=False, range=[0, x_rmax]),
            yaxis=dict(tickfont=dict(size=11, color=C['text']), showgrid=False),
        )
        return fig
    except Exception as e:
        return go.Figure().update_layout(**PL, title=f'Error canal mix: {e}')

def fig_canal_barras(canal_sel, flia_sel=None, repre_sel=None, meses_sel=None):
    try:
        if repre_sel:
            df = DFS['x repre x canal']
            df = df[df['Vendedor'] == repre_sel]
        else:
            df = DFS['x flia x canal']
        if canal_sel:
            df = df[df['Canal'] == canal_sel]
        if flia_sel and 'flia' in df.columns:
            df = df[df['flia'] == flia_sel]
        act = get_ind(df, 'Año Actual Cajas', ['Canal','flia'], meses_sel)
        ant = get_ind(df, 'Año Anterior Cajas', ['Canal','flia'], meses_sel)
        a = act.groupby('Canal')['Total'].sum().reset_index()
        b = ant.groupby('Canal')['Total'].sum().reset_index()
        m = a.merge(b, on='Canal', suffixes=('_a','_b'))
        m = m[~m['Canal'].str.upper().str.strip().isin(['TRAVEL RETAIL'])]
        m['var'] = (m['Total_a'] - m['Total_b']) / m['Total_b'].replace(0, np.nan) * 100
        m = m.sort_values('Total_a', ascending=False)
        col_var = [C['green'] if (pd.notna(v) and v >= 0) else C['red'] for v in m['var']]
        fig = make_subplots(rows=1, cols=2, subplot_titles=['Cajas por Canal', 'Var % vs Año Anterior'])
        fig.add_trace(go.Bar(x=m['Canal'], y=m['Total_a'], marker_color=C['gold'],
                             text=[f"{int(v):,}" for v in m['Total_a']],
                             textposition='inside', insidetextanchor='middle',
                             textfont=dict(size=10, color='#FFFFFF'),
                             hovertemplate='%{x}<br>%{y:,.0f}<extra></extra>'), row=1, col=1)
        fig.add_trace(go.Bar(x=m['Canal'], y=m['var'].fillna(0).round(0), marker_color=col_var,
                             text=[_var(v) for v in m['var']],
                             textposition='outside',
                             textfont=dict(size=14, color=C['text']),
                             hovertemplate='%{x}<br>%{y:+.0f}%<extra></extra>'), row=1, col=2)
        _pl = {k:v for k,v in PL.items() if k != 'margin'}
        fig.update_layout(**_pl, title='Canales — Volumen y Variación', height=290, showlegend=False,
                          margin=dict(l=30, r=20, t=46, b=70))
        fig.update_xaxes(tickangle=-40, tickfont=dict(size=10))
        return fig
    except Exception as e:
        return go.Figure().update_layout(**PL, title=f'Error: {e}')

def analisis_clientes(repre_sel, flia_sel, meses_sel=None):
    """Devuelve dict con dfs preparados para la pestaña Clientes."""
    df = DFS['x cliente']
    act = get_ind(df, 'Año Actual Cajas', ['Vendedor','Cliente','flia'], meses_sel)
    ant = get_ind(df, 'Año Anterior Cajas', ['Vendedor','Cliente','flia'], meses_sel)
    if repre_sel:
        act = act[act['Vendedor'] == repre_sel]
        ant = ant[ant['Vendedor'] == repre_sel]
    if flia_sel:
        act = act[act['flia'] == flia_sel]
        ant = ant[ant['flia'] == flia_sel]
    act_agg = act[act['Total'] > 0].groupby(['Vendedor','Cliente'])['Total'].sum().reset_index().rename(columns={'Total':'act'})
    ant_agg = ant[ant['Total'] > 0].groupby(['Vendedor','Cliente'])['Total'].sum().reset_index().rename(columns={'Total':'ant'})
    claves  = ['Vendedor','Cliente']
    full = act_agg.merge(ant_agg, on=claves, how='outer')
    full['act'] = full['act'].fillna(0)
    full['ant'] = full['ant'].fillna(0)
    full['dif'] = full['act'] - full['ant']
    full['var'] = (full['dif'] / full['ant'].replace(0, np.nan)) * 100
    # categorías mutuamente excluyentes y lógicamente coherentes
    nuevos_cli   = full[(full['ant'] == 0) & (full['act'] > 0)].copy()   # no existían antes
    crecieron    = full[(full['ant'] > 0)  & (full['act'] > 0) & (full['dif'] > 0)].copy()
    cayeron      = full[(full['ant'] > 0)  & (full['act'] > 0) & (full['dif'] < 0)].copy()
    perdidos_cli = full[(full['ant'] > 0)  & (full['act'] == 0)].copy()  # no compraron este año
    activos      = full[(full['act'] > 0)  & (full['ant'] > 0)].copy()   # compraron ambos años
    con_crecimiento = full[full['dif'] > 0].copy()   # para el gráfico (incluye nuevos)
    con_caida       = full[full['dif'] < 0].copy()   # para el gráfico (incluye perdidos)
    return {
        'full': full, 'nuevos': con_crecimiento, 'perdidos': con_caida, 'activos': activos,
        'nuevos_cli': nuevos_cli, 'crecieron': crecieron,
        'cayeron': cayeron, 'perdidos_cli': perdidos_cli,
        'top_sube': full[full['dif'] > 0].nlargest(20, 'dif').sort_values('dif', ascending=False),
        'top_baja': full[full['dif'] < 0].nsmallest(20, 'dif').sort_values('dif'),
    }

def fig_clientes_variacion(datos):
    """Top 20 crecimiento y top 20 caída — gráfico principal."""
    try:
        top_sube = datos['top_sube']
        top_baja = datos['top_baja']
        h = max(380, max(len(top_sube), len(top_baja)) * 22 + 80)
        fig = make_subplots(
            rows=1, cols=2,
            subplot_titles=[f"Top {len(top_sube)} — Mayor Crecimiento (cajas)",
                            f"Top {len(top_baja)} — Mayor Caída (cajas)"],
            horizontal_spacing=0.1,
        )
        fig.add_trace(go.Bar(
            y=top_sube['Cliente'].str[:30], x=top_sube['dif'], orientation='h',
            marker_color=C['green'],
            text=[f"+{int(d):,} caj" for d in top_sube['dif']],
            textposition='inside', insidetextanchor='middle',
            textfont=dict(size=10, color='#FFFFFF'),
            customdata=top_sube[['Vendedor','act','ant']].values,
            hovertemplate='<b>%{y}</b><br>Rep: %{customdata[0]}<br>Δ: +%{x:,.0f} caj<br>Act: %{customdata[1]:,.0f}  Ant: %{customdata[2]:,.0f}<extra></extra>',
        ), row=1, col=1)
        fig.add_trace(go.Bar(
            y=top_baja['Cliente'].str[:30], x=top_baja['dif'], orientation='h',
            marker_color=C['red'],
            text=[f"{int(d):,} caj" for d in top_baja['dif']],
            textposition='inside', insidetextanchor='middle',
            textfont=dict(size=10, color='#FFFFFF'),
            customdata=top_baja[['Vendedor','act','ant']].values,
            hovertemplate='<b>%{y}</b><br>Rep: %{customdata[0]}<br>Δ: %{x:,.0f} caj<br>Act: %{customdata[1]:,.0f}  Ant: %{customdata[2]:,.0f}<extra></extra>',
        ), row=1, col=2)
        pl_c = {k: v for k, v in PL.items() if k not in ('margin',)}
        fig.update_layout(**pl_c, title='Clientes — Top 20 Crecimiento y Caída (cajas)',
                          height=h, showlegend=False, margin=dict(l=10, r=30, t=50, b=30))
        fig.update_xaxes(tickfont=dict(size=13), gridcolor=C['border'])
        fig.update_yaxes(tickfont=dict(size=13))
        return fig
    except Exception as e:
        return go.Figure().update_layout(**PL, title=f'Error clientes: {e}')

def fig_clientes_nuevos_perdidos(datos):
    """Clientes con crecimiento vs con caída — cajas totales dentro, % var afuera."""
    try:
        crec  = datos['nuevos'].nlargest(15, 'dif').sort_values('dif')
        caida = datos['perdidos'].nsmallest(15, 'dif').sort_values('dif', ascending=False)
        n_crec  = len(datos['nuevos'])
        n_caida = len(datos['perdidos'])
        h = max(300, max(len(crec), len(caida)) * 22 + 90)

        fig = make_subplots(
            rows=1, cols=2,
            subplot_titles=[f"Con Crecimiento ({n_crec} total)",
                            f"Con Caídas ({n_caida} total)"],
            horizontal_spacing=0.12,
        )

        if not crec.empty:
            y_crec = crec['Cliente'].str[:28]
            fig.add_trace(go.Bar(
                y=y_crec, x=crec['dif'], orientation='h',
                marker_color='rgba(39,174,96,0.85)',
                text=[f"+{int(v):,} caj" for v in crec['dif']],
                textposition='inside', insidetextanchor='middle',
                textfont=dict(size=10, color='#FFFFFF'),
                customdata=np.column_stack([crec['Vendedor'], crec['act'], crec['ant'], crec['var'].fillna(0).round(0)]),
                hovertemplate='<b>%{y}</b><br>Rep: %{customdata[0]}<br>Diferencia: +%{x:,.0f} caj<br>Actual: %{customdata[1]:,.0f}  Anterior: %{customdata[2]:,.0f}  %{customdata[3]:+.0f}%<extra></extra>',
            ), row=1, col=1)
        if not caida.empty:
            y_caida = caida['Cliente'].str[:28]
            fig.add_trace(go.Bar(
                y=y_caida, x=caida['dif'], orientation='h',
                marker_color='rgba(192,57,43,0.85)',
                text=[f"{int(v):,} caj" for v in caida['dif']],
                textposition='inside', insidetextanchor='middle',
                textfont=dict(size=10, color='#FFFFFF'),
                customdata=np.column_stack([caida['Vendedor'], caida['act'], caida['ant'], caida['var'].fillna(0).round(0)]),
                hovertemplate='<b>%{y}</b><br>Rep: %{customdata[0]}<br>Diferencia: %{x:,.0f} caj<br>Actual: %{customdata[1]:,.0f}  Anterior: %{customdata[2]:,.0f}  %{customdata[3]:+.0f}%<extra></extra>',
            ), row=1, col=2)

        pl_c = {k: v for k, v in PL.items() if k not in ('margin',)}
        fig.update_layout(**pl_c, title='Clientes — Con Crecimiento vs Con Caídas — Diferencia en cajas',
                          height=h, showlegend=False, margin=dict(l=10, r=20, t=50, b=30))
        fig.update_xaxes(tickfont=dict(size=11), gridcolor=C['border'], ticksuffix=' caj')
        fig.update_yaxes(tickfont=dict(size=11))
        return fig
    except Exception as e:
        return go.Figure().update_layout(**PL, title=f'Error crecimiento/caída: {e}')

def fig_pendientes(flia_sel=None, repre_sel=None):
    try:
        if 'pend' not in DFS:
            return go.Figure().update_layout(**PL, title='Sin datos de pendientes')
        df = DFS['pend'].copy()
        df.columns = [c.strip() for c in df.columns]
        df['Pedidos Pendientes'] = pd.to_numeric(df['Pedidos Pendientes'], errors='coerce')
        df = df[df['Pedidos Pendientes'] > 0]
        if repre_sel:
            df = df[df['Vendedor'] == repre_sel]
        if flia_sel:
            df = df[df['Familia Producto'] == flia_sel]
        agg = df.groupby('Vendedor')['Pedidos Pendientes'].sum().reset_index()
        agg = agg.sort_values('Pedidos Pendientes', ascending=False).head(15)
        titulo = 'Pedidos Pendientes por Vendedor'
        if repre_sel:
            titulo += f' — {repre_sel}'
        elif flia_sel:
            titulo += f' — {flia_sel}'
        fig = go.Figure(go.Bar(
            x=agg['Vendedor'].str[:28], y=agg['Pedidos Pendientes'],
            marker_color=C['gold'],
            hovertemplate='%{x}<br>Pendientes: %{y:,.0f}<extra></extra>'
        ))
        pl_p = {k: v for k, v in PL.items() if k != 'xaxis'}
        fig.update_layout(**pl_p, title=titulo, height=340)
        fig.update_xaxes(tickangle=-35, tickfont=dict(size=13), gridcolor=C['border'])
        return fig
    except Exception as e:
        return go.Figure().update_layout(**PL, title=f'Error: {e}')

# ── Análisis inteligente ───────────────────────────────────────────────────────

def generar_analisis(flia_sel=None, repre_sel=None, canal_sel=None, meses_sel=None):
    insights, alertas, oportunidades = [], [], []
    try:
        if repre_sel:
            df_repre_f = DFS['x repre'][DFS['x repre']['Vendedor'] == repre_sel]
            act = get_ind(df_repre_f, 'Año Actual Cajas', ['Vendedor','flia'], meses_sel)
            ant = get_ind(df_repre_f, 'Año Anterior Cajas', ['Vendedor','flia'], meses_sel)
            tot_a = act['Total'].sum(); tot_b = ant['Total'].sum()
            var_t = (tot_a - tot_b) / tot_b * 100 if tot_b else 0
            sign  = '+' if var_t >= 0 else ''
            insights.append(f"Representante {repre_sel} total: {int(tot_a):,} cajas | Var: {sign}{var_t:.0f}% vs año anterior")
            top3 = act.groupby('flia')['Total'].sum().nlargest(3)
            insights.append(f"Top familias: {', '.join([f'{f} ({int(v):,})' for f,v in top3.items()])}")
            var_f = act.merge(ant, on=['Vendedor','flia'], suffixes=('_a','_b'))
            var_f['var'] = (var_f['Total_a'] - var_f['Total_b']) / var_f['Total_b'].replace(0, np.nan)
            var_f_s = var_f.set_index('flia')['var'].dropna().sort_values()
            for f,v in var_f_s.head(3).items():
                alertas.append(f"Alerta {f}: {v*100:+.0f}% vs año anterior")
            for f,v in var_f_s.tail(3).items():
                oportunidades.append(f"Crecimiento {f}: {v*100:+.0f}% vs año anterior")
            foda = {
                'F': oportunidades[:],
                'D': alertas[:],
                'O': [f"Familia con mejor momentum para {repre_sel}",
                      "Oportunidad de recuperar clientes perdidos"],
                'A': [f"Familias con caída sostenida en cartera de {repre_sel}",
                      "Concentración de volumen en pocos clientes"]
            }
        elif flia_sel:
            act = get_ind(DFS['x flia'], 'Año Actual Cajas', ['flia'], meses_sel)
            act = act[act['flia'] == flia_sel]
            ant = get_ind(DFS['x flia'], 'Año Anterior Cajas', ['flia'], meses_sel)
            ant = ant[ant['flia'] == flia_sel]
            tot_a = act['Total'].sum(); tot_b = ant['Total'].sum()
            var_t = (tot_a - tot_b) / tot_b * 100 if tot_b else 0
            sign  = '+' if var_t >= 0 else ''
            insights.append(f"Familia {flia_sel}: {int(tot_a):,} cajas | Var: {sign}{var_t:.0f}% vs año anterior")
            var_f = get_ind(DFS['x flia'], 'Var %', ['flia'])
            var_f_s = var_f.set_index('flia')['Total'].dropna().sort_values()
            for f,v in var_f_s.head(3).items():
                alertas.append(f"Alerta {f}: {v*100:+.0f}% vs año anterior")
            for f,v in var_f_s.tail(3).items():
                oportunidades.append(f"Crecimiento {f}: {v*100:+.0f}% vs año anterior")
            foda = {
                'F': oportunidades[:],
                'D': alertas[:],
                'O': [f"Momentum positivo en familia {flia_sel}",
                      "Oportunidad de penetración en canales con baja participación"],
                'A': [f"Familia {flia_sel} bajo presión competitiva",
                      "Concentración de volumen en pocos representantes"]
            }
        else:
            act = get_ind(DFS['x flia'], 'Año Actual Cajas', ['flia'], meses_sel)
            ant = get_ind(DFS['x flia'], 'Año Anterior Cajas', ['flia'], meses_sel)
            var = get_ind(DFS['x flia'], 'Var %', ['flia'])
            tot_a = act['Total'].sum(); tot_b = ant['Total'].sum()
            var_t = (tot_a - tot_b) / tot_b * 100 if tot_b else 0
            sign  = '+' if var_t >= 0 else ''
            insights.append(f"Región total: {int(tot_a):,} cajas | Var: {sign}{var_t:.0f}% vs año anterior")
            top3 = act.set_index('flia')['Total'].nlargest(3)
            insights.append(f"Top familias: {', '.join([f'{f} ({int(v):,})' for f,v in top3.items()])}")
            var_f = var.set_index('flia')['Total'].dropna().sort_values()
            for f,v in var_f.head(3).items():
                alertas.append(f"Alerta {f}: {v*100:+.0f}% vs año anterior")
            for f,v in var_f.tail(3).items():
                oportunidades.append(f"Crecimiento {f}: {v*100:+.0f}% vs año anterior")
            repre_var = get_ind(DFS['x repre'], 'Var %', ['Vendedor','flia'])
            rv = repre_var.groupby('Vendedor')['Total'].mean().sort_values()
            alertas.append(f"Representante con mayor caída: {rv.index[0]} ({rv.iloc[0]*100:+.0f}%)")
            oportunidades.append(f"Representante con mayor crecimiento: {rv.index[-1]} ({rv.iloc[-1]*100:+.0f}%)")
            if len(MC) >= 2:
                u, p = MC[-1], MC[-2]
                tu = act[u].sum(); tp = act[p].sum()
                vt = (tu-tp)/tp*100 if tp else 0
                tend = "acelerando" if vt > 0 else "desacelerando"
                insights.append(f"Tendencia {p} a {u}: {tend} ({vt:+.0f}%)")
            foda = {
                'F': oportunidades[:],
                'D': alertas[:],
                'O': [f"Familia {var_f.index[-1]} con momentum positivo",
                      "Oportunidad de penetración en canales con baja participación"],
                'A': [f"Familia {var_f.index[0]} requiere plan de acción urgente",
                      "Concentración de volumen en pocos representantes"]
            }
    except Exception as e:
        foda = {'F':[], 'D':[], 'O':[], 'A':[]}
        insights.append(f"Error generando análisis: {e}")
    return insights, alertas, oportunidades, foda

def generar_red_flags(flia_sel=None, repre_sel=None, canal_sel=None, meses_sel=None):
    """Detecta señales de alerta críticas en los datos."""
    flags = []
    try:
        # Representantes con caída > 15%
        rv = get_ind(DFS['x repre'], 'Var %', ['Vendedor','flia'])
        if repre_sel:
            rv = rv[rv['Vendedor'] == repre_sel]
        rmed = rv.groupby('Vendedor')['Total'].mean() * 100
        for v, pct in rmed[rmed < -15].items():
            if repre_sel:
                flags.append(('CRITICO', f"Este representante: caída de {pct:.0f}% — requiere atención inmediata"))
            else:
                flags.append(('CRITICO', f"Rep. {v}: caída de {pct:.0f}% — requiere atención inmediata"))
        for v, pct in rmed[(rmed < -5) & (rmed >= -15)].items():
            if repre_sel:
                flags.append(('ALERTA', f"Este representante: caída de {pct:.0f}%"))
            else:
                flags.append(('ALERTA', f"Rep. {v}: caída de {pct:.0f}%"))

        # Familias con caída en todos los canales
        if repre_sel and 'x repre x canal' in DFS:
            src_canal = DFS['x repre x canal'][DFS['x repre x canal']['Vendedor'] == repre_sel]
            var_canal = get_ind(src_canal, 'Var %', ['flia','Canal'])
        else:
            var_canal = get_ind(DFS['x flia x canal'], 'Var %', ['flia','Canal'])
        familias_check = [flia_sel] if flia_sel else FAMILIAS
        for flia in familias_check:
            sub = var_canal[var_canal['flia'] == flia]['Total'] * 100
            if sub.dropna().empty:
                continue
            if (sub.dropna() < 0).all():
                flags.append(('ALERTA', f"Familia {flia}: caída en todos los canales ({sub.mean():.0f}% promedio)"))

        # Desaceleración mes a mes (últimos 2 meses disponibles)
        if len(MC) >= 2:
            if repre_sel and 'x repre' in DFS:
                src_decel = DFS['x repre'][DFS['x repre']['Vendedor'] == repre_sel]
                grps_decel = ['Vendedor', 'flia']
            else:
                src_decel = DFS['x flia']
                grps_decel = ['flia']
            act = get_ind(src_decel, 'Año Actual Cajas', grps_decel)
            u, p = MC[-1], MC[-2]
            tu = pd.to_numeric(act[u], errors='coerce').sum()
            tp = pd.to_numeric(act[p], errors='coerce').sum()
            if tp > 0 and (tu - tp) / tp * 100 < -10:
                vt = (tu - tp) / tp * 100
                label = "en tu zona" if repre_sel else "en total región"
                flags.append(('CRITICO', f"Desaceleración fuerte: {p}→{u} = {vt:+.0f}% {label}"))

        # Clientes perdidos (fueron a 0) — agregar por cliente antes de comparar
        # para no multiplicar por cantidad de familias ni de indicadores
        if 'x cliente' in DFS:
            cli = DFS['x cliente']
            act_c = get_ind(cli, 'Año Actual Cajas', ['Vendedor','Cliente','flia'])
            ant_c = get_ind(cli, 'Año Anterior Cajas', ['Vendedor','Cliente','flia'])
            if repre_sel:
                act_c = act_c[act_c['Vendedor'] == repre_sel]
                ant_c = ant_c[ant_c['Vendedor'] == repre_sel]
            act_agg = act_c.groupby('Cliente')['Total'].sum().reset_index()
            ant_agg = ant_c.groupby('Cliente')['Total'].sum().reset_index()
            merged = ant_agg[ant_agg['Total'] > 0].merge(act_agg, on='Cliente', suffixes=('_b','_a'), how='left')
            merged['Total_a'] = merged['Total_a'].fillna(0)
            perdidos = merged[merged['Total_a'] == 0]
            if not perdidos.empty:
                n = len(perdidos)
                flags.append(('ALERTA', f"{n} cliente{'s' if n>1 else ''} con 0 cajas este año (activos el año anterior)"))

    except Exception as e:
        flags.append(('INFO', f'No se pudieron calcular red flags: {e}'))

    if not flags:
        flags.append(('OK', 'Sin alertas críticas detectadas.'))
    return flags


def generar_analisis_quirurgico(flia_sel=None, repre_sel=None, canal_sel=None, meses_sel=None):
    """Genera datasets estructurados para el análisis quirúrgico de 4 niveles."""
    result = {}

    # 1. Variación por familia — nivel región
    try:
        src_f = DFS['x flia']
        if repre_sel:
            src_f = DFS['x repre'][DFS['x repre']['Vendedor'] == repre_sel]
        if flia_sel:
            src_f = src_f[src_f['flia'] == flia_sel]
        act_f = get_ind(src_f, 'Año Actual Cajas', ['flia'], meses_sel).groupby('flia')['Total'].sum()
        ant_f = get_ind(src_f, 'Año Anterior Cajas', ['flia'], meses_sel).groupby('flia')['Total'].sum()
        df_flia = pd.DataFrame({'Actual': act_f, 'Anterior': ant_f}).reset_index().rename(columns={'flia':'Familia'})
        df_flia['Var%'] = (df_flia['Actual'] - df_flia['Anterior']) / df_flia['Anterior'].replace(0, np.nan) * 100
        df_flia['Dif'] = df_flia['Actual'] - df_flia['Anterior']
        result['familias'] = df_flia.sort_values('Var%')
    except Exception:
        result['familias'] = pd.DataFrame()

    # 2. Variación por canal — nivel región (o filtrado por repre/flia)
    try:
        if repre_sel:
            src_c = DFS['x repre x canal'][DFS['x repre x canal']['Vendedor'] == repre_sel]
            if flia_sel:
                src_c = src_c[src_c['flia'] == flia_sel]
            act_c = get_ind(src_c, 'Año Actual Cajas', ['Canal'], meses_sel).groupby('Canal')['Total'].sum()
            ant_c = get_ind(src_c, 'Año Anterior Cajas', ['Canal'], meses_sel).groupby('Canal')['Total'].sum()
        elif flia_sel:
            src_c = DFS['x flia x canal'][DFS['x flia x canal']['flia'] == flia_sel]
            act_c = get_ind(src_c, 'Año Actual Cajas', ['Canal'], meses_sel).groupby('Canal')['Total'].sum()
            ant_c = get_ind(src_c, 'Año Anterior Cajas', ['Canal'], meses_sel).groupby('Canal')['Total'].sum()
        else:
            act_c = get_ind(DFS['x flia x canal'], 'Año Actual Cajas', ['Canal'], meses_sel).groupby('Canal')['Total'].sum()
            ant_c = get_ind(DFS['x flia x canal'], 'Año Anterior Cajas', ['Canal'], meses_sel).groupby('Canal')['Total'].sum()
        if canal_sel:
            act_c = act_c[[canal_sel]] if canal_sel in act_c.index else act_c
            ant_c = ant_c[[canal_sel]] if canal_sel in ant_c.index else ant_c
        df_canal = pd.DataFrame({'Actual': act_c, 'Anterior': ant_c}).reset_index()
        df_canal = df_canal[~df_canal['Canal'].str.upper().str.strip().isin(['TRAVEL RETAIL'])]
        df_canal['Var%'] = (df_canal['Actual'] - df_canal['Anterior']) / df_canal['Anterior'].replace(0, np.nan) * 100
        df_canal['Dif'] = df_canal['Actual'] - df_canal['Anterior']
        result['canales'] = df_canal.sort_values('Var%')
    except Exception:
        result['canales'] = pd.DataFrame()

    # 3. Matriz variación Representante × Familia
    try:
        src_rx = DFS['x repre'].copy()
        if repre_sel:
            src_rx = src_rx[src_rx['Vendedor'] == repre_sel]
        if flia_sel:
            src_rx = src_rx[src_rx['flia'] == flia_sel]
        if canal_sel:
            src_rxc = DFS['x repre x canal'][DFS['x repre x canal']['Canal'] == canal_sel].copy()
            if repre_sel:
                src_rxc = src_rxc[src_rxc['Vendedor'] == repre_sel]
            if flia_sel:
                src_rxc = src_rxc[src_rxc['flia'] == flia_sel]
            act_rx = get_ind(src_rxc, 'Año Actual Cajas', ['Vendedor','flia'], meses_sel)
            ant_rx = get_ind(src_rxc, 'Año Anterior Cajas', ['Vendedor','flia'], meses_sel)
        else:
            act_rx = get_ind(src_rx, 'Año Actual Cajas', ['Vendedor','flia'], meses_sel)
            ant_rx = get_ind(src_rx, 'Año Anterior Cajas', ['Vendedor','flia'], meses_sel)
        act_rx = act_rx.groupby(['Vendedor','flia'])['Total'].sum().reset_index()
        ant_rx = ant_rx.groupby(['Vendedor','flia'])['Total'].sum().reset_index()
        mx = act_rx.merge(ant_rx, on=['Vendedor','flia'], suffixes=('_a','_b'))
        mx['var'] = (mx['Total_a'] - mx['Total_b']) / mx['Total_b'].replace(0, np.nan) * 100
        pivot = mx.pivot_table(index='Vendedor', columns='flia', values='var', aggfunc='mean')
        flia_ord = pivot.mean().sort_values().index.tolist()
        rep_ord  = pivot.mean(axis=1).sort_values().index.tolist()
        pivot = pivot.loc[rep_ord, flia_ord]
        result['matriz'] = pivot
        result['tot_repre'] = act_rx.groupby('Vendedor')['Total'].sum()
    except Exception:
        result['matriz'] = pd.DataFrame()
        result['tot_repre'] = pd.Series(dtype=float)

    # 4. Fluctuación canal por representante
    try:
        src_rxc2 = DFS['x repre x canal'].copy()
        if repre_sel:
            src_rxc2 = src_rxc2[src_rxc2['Vendedor'] == repre_sel]
        if flia_sel:
            src_rxc2 = src_rxc2[src_rxc2['flia'] == flia_sel]
        if canal_sel:
            src_rxc2 = src_rxc2[src_rxc2['Canal'] == canal_sel]
        act2 = get_ind(src_rxc2, 'Año Actual Cajas', ['Vendedor','Canal'], meses_sel).groupby(['Vendedor','Canal'])['Total'].sum().reset_index()
        ant2 = get_ind(src_rxc2, 'Año Anterior Cajas', ['Vendedor','Canal'], meses_sel).groupby(['Vendedor','Canal'])['Total'].sum().reset_index()
        cr = act2.merge(ant2, on=['Vendedor','Canal'], suffixes=('_a','_b'))
        cr = cr[~cr['Canal'].str.upper().str.strip().isin(['TRAVEL RETAIL'])]
        cr['var'] = (cr['Total_a'] - cr['Total_b']) / cr['Total_b'].replace(0, np.nan) * 100
        cr['dif'] = cr['Total_a'] - cr['Total_b']
        result['canal_repre'] = cr.sort_values(['Vendedor','var'])
    except Exception:
        result['canal_repre'] = pd.DataFrame()

    # 5. Tendencia mensual — región total
    try:
        if len(MC) >= 2:
            act_m = get_ind(DFS['x flia'], 'Año Actual Cajas', ['flia'])
            ant_m = get_ind(DFS['x flia'], 'Año Anterior Cajas', ['flia'])
            tend_rows = []
            for m in MC:
                if m in act_m.columns and m in ant_m.columns:
                    a = pd.to_numeric(act_m[m], errors='coerce').sum()
                    b = pd.to_numeric(ant_m[m], errors='coerce').sum()
                    v = (a - b) / b * 100 if b else 0
                    tend_rows.append({'Mes': m, 'Actual': a, 'Anterior': b, 'Var%': v})
            result['tendencia'] = pd.DataFrame(tend_rows)
        else:
            result['tendencia'] = pd.DataFrame()
    except Exception:
        result['tendencia'] = pd.DataFrame()

    return result


# ── KPIs con filtros ───────────────────────────────────────────────────────────

def build_kpis(flia_sel=None, repre_sel=None, canal_sel=None, meses_sel=None):
    try:
        # Seleccionar hoja según filtros activos
        if canal_sel and repre_sel:
            df = DFS['x repre x canal'].copy()
            df = df[df['Canal'] == canal_sel]
            df = df[df['Vendedor'] == repre_sel]
            act = get_ind(df, 'Año Actual Cajas', ['Vendedor','flia'], meses_sel)
            ant = get_ind(df, 'Año Anterior Cajas', ['Vendedor','flia'], meses_sel)
        elif canal_sel:
            df = DFS['x flia x canal'].copy()
            df = df[df['Canal'] == canal_sel]
            act = get_ind(df, 'Año Actual Cajas', ['flia','Canal'], meses_sel)
            ant = get_ind(df, 'Año Anterior Cajas', ['flia','Canal'], meses_sel)
        elif repre_sel:
            df = DFS['x repre'].copy()
            df = df[df['Vendedor'] == repre_sel]
            act = get_ind(df, 'Año Actual Cajas', ['Vendedor','flia'], meses_sel)
            ant = get_ind(df, 'Año Anterior Cajas', ['Vendedor','flia'], meses_sel)
        else:
            act = get_ind(DFS['x flia'], 'Año Actual Cajas', ['flia'], meses_sel)
            ant = get_ind(DFS['x flia'], 'Año Anterior Cajas', ['flia'], meses_sel)

        if flia_sel:
            act = act[act['flia'] == flia_sel]
            ant = ant[ant['flia'] == flia_sel]

        tot_a = act['Total'].sum()
        tot_b = ant['Total'].sum()
        var_t = (tot_a - tot_b) / tot_b * 100 if tot_b else 0
    except:
        tot_a = tot_b = var_t = 0

    # Pendientes: filtrar por familia/representante si hay selección
    pend = 0
    try:
        if 'pend' in DFS:
            df_pend = DFS['pend'].copy()
            df_pend['Pedidos Pendientes'] = pd.to_numeric(df_pend['Pedidos Pendientes'], errors='coerce')
            df_pend = df_pend[df_pend['Pedidos Pendientes'] > 0]
            if repre_sel and 'Vendedor' in df_pend.columns:
                df_pend = df_pend[df_pend['Vendedor'] == repre_sel]
            if flia_sel and 'Familia Producto' in df_pend.columns:
                df_pend = df_pend[df_pend['Familia Producto'] == flia_sel]
            pend = df_pend['Pedidos Pendientes'].sum()
    except:
        pend = 0

    cv   = C['green'] if var_t >= 0 else C['red']
    sign = '+' if var_t >= 0 else ''

    # Número de representantes activos según filtros (incluyendo flia)
    try:
        if repre_sel:
            n_repre = 1
        elif flia_sel and canal_sel:
            n_repre = DFS['x repre x canal'][
                (DFS['x repre x canal']['Canal'] == canal_sel) &
                (DFS['x repre x canal']['flia'] == flia_sel)
            ]['Vendedor'].nunique()
        elif flia_sel:
            n_repre = DFS['x repre'][DFS['x repre']['flia'] == flia_sel]['Vendedor'].nunique()
        elif canal_sel:
            n_repre = DFS['x repre x canal'][DFS['x repre x canal']['Canal'] == canal_sel]['Vendedor'].nunique()
        else:
            n_repre = DFS['x repre']['Vendedor'].nunique()
    except:
        n_repre = '—'

    # Número de familias según filtros
    n_flia = 1 if flia_sel else len(FAMILIAS)

    filtro_label = " | ".join(filter(None, [
        flia_sel, repre_sel, canal_sel
    ])) or "Región completa"

    items = [
        ('CAJAS AÑO ACTUAL',   f"{int(tot_a):,}",     C['gold']),
        ('CAJAS AÑO ANTERIOR', f"{int(tot_b):,}",     C['muted']),
        ('VARIACION TOTAL',    f"{sign}{var_t:.0f}%",  cv),
        ('REPRESENTANTES',     str(n_repre),            C['gold']),
        ('FAMILIAS',           str(n_flia),             C['gold']),
        ('PENDIENTES',         f"{int(pend):,}",        C['red'] if pend > 0 else C['muted']),
    ]
    return html.Div([
        # Filtro activo
        html.Div(f"Filtro activo: {filtro_label}",
                 style={'color': C['muted'], 'fontSize': '9px', 'letterSpacing': '1px',
                        'marginBottom': '8px', 'textTransform': 'uppercase'}),
        html.Div([
            html.Div([
                html.Div(label, style={'color':C['muted'],'fontSize':'9px','letterSpacing':'2px',
                                       'textTransform':'uppercase','marginBottom':'4px'}),
                html.Div(val,   style={'color':color,'fontSize':'22px','fontWeight':'700','fontFamily':MONO}),
            ], style={'backgroundColor':C['surf'],'border':f"1px solid {C['border']}",
                      'borderRadius':'3px','padding':'14px','textAlign':'center'})
            for label, val, color in items
        ], style={'display':'grid','gridTemplateColumns':'repeat(6,1fr)','gap':'10px'}),
    ], style={'marginBottom':'16px'})


# ── Sistema de diseño PDF ejecutivo ───────────────────────────────────────────

def _pdf_ds():
    """Design system unificado para todos los PDFs ejecutivos."""
    G  = rl_colors.HexColor('#B8972A')
    R  = rl_colors.HexColor('#B03020')
    GR = rl_colors.HexColor('#217A45')
    BK = rl_colors.HexColor('#111111')
    MU = rl_colors.HexColor('#666666')
    LG = rl_colors.HexColor('#F7F7F4')
    WH = rl_colors.white
    return {
        'sec':   ParagraphStyle('_ps', fontSize=7.5, textColor=G, fontName='Helvetica-Bold',
                                spaceBefore=14, spaceAfter=4, leading=10, wordWrap='LTR'),
        'body':  ParagraphStyle('_pb', fontSize=8.5, textColor=BK, fontName='Helvetica',
                                leading=13, spaceAfter=3),
        'small': ParagraphStyle('_psm', fontSize=7, textColor=MU, fontName='Helvetica',
                                leading=10, spaceAfter=2),
        'alert': ParagraphStyle('_pa', fontSize=8.5, textColor=R, fontName='Helvetica',
                                leading=13, spaceAfter=3),
        'ok':    ParagraphStyle('_po', fontSize=8.5, textColor=GR, fontName='Helvetica',
                                leading=13, spaceAfter=3),
        'G': G, 'R': R, 'GR': GR, 'BK': BK, 'MU': MU, 'LG': LG, 'WH': WH,
    }

def _pdf_page_cb(titulo_doc, filtro_txt):
    """Callback para header y footer en cada página del PDF."""
    def _draw(c, doc):
        c.saveState()
        W, H = A4
        L, R_margin = 1.5*cm, 1.5*cm

        # ── Header band ──────────────────────────────────────────────────────
        c.setFillColor(rl_colors.HexColor('#0D0D0D'))
        c.rect(0, H - 2.1*cm, W, 2.1*cm, fill=1, stroke=0)
        # Línea dorada inferior del header
        c.setFillColor(rl_colors.HexColor('#B8972A'))
        c.rect(0, H - 2.1*cm - 1.5, W, 1.5, fill=1, stroke=0)
        # Nombre empresa (izquierda)
        c.setFont('Helvetica-Bold', 11)
        c.setFillColor(rl_colors.white)
        c.drawString(L, H - 1.3*cm, 'CATENA ZAPATA')
        # Título del reporte (derecha)
        c.setFont('Helvetica', 8.5)
        c.setFillColor(rl_colors.HexColor('#BBBBBB'))
        c.drawRightString(W - R_margin, H - 1.3*cm, titulo_doc)
        # Filtro + fecha (segunda línea, pequeño)
        c.setFont('Helvetica', 6.5)
        c.setFillColor(rl_colors.HexColor('#888888'))
        fecha_str = datetime.now().strftime('%d/%m/%Y')
        c.drawString(L, H - 1.85*cm, f"{filtro_txt}   ·   {fecha_str}")

        # ── Footer ───────────────────────────────────────────────────────────
        c.setStrokeColor(rl_colors.HexColor('#CCCCCC'))
        c.setLineWidth(0.4)
        c.line(L, 1.35*cm, W - R_margin, 1.35*cm)
        c.setFont('Helvetica', 6.5)
        c.setFillColor(rl_colors.HexColor('#999999'))
        c.drawString(L, 0.9*cm, 'CATENA ZAPATA  —  INFORMACIÓN CONFIDENCIAL')
        c.drawRightString(W - R_margin, 0.9*cm, f'Página {doc.page}')

        c.restoreState()
    return _draw

def _pdf_tbl(data, widths, var_cols=(), right_cols=(), center_cols=(), zebra=True):
    """Tabla ejecutiva: header oscuro, líneas horizontales, sin grilla vertical."""
    if not data or len(data) < 1:
        return Spacer(1, 0.1*cm)
    t = Table(data, colWidths=widths, repeatRows=1)
    BK  = rl_colors.HexColor('#1A1A1A')
    LG  = rl_colors.HexColor('#F7F7F4')
    SEP = rl_colors.HexColor('#E0E0DA')
    cmds = [
        # ── Encabezado ──
        ('BACKGROUND',    (0,0), (-1,0), BK),
        ('TEXTCOLOR',     (0,0), (-1,0), rl_colors.white),
        ('FONTNAME',      (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',      (0,0), (-1,0), 7.5),
        ('TOPPADDING',    (0,0), (-1,0), 7),
        ('BOTTOMPADDING', (0,0), (-1,0), 7),
        ('LEFTPADDING',   (0,0), (-1,0), 7),
        ('RIGHTPADDING',  (0,0), (-1,0), 7),
        # ── Filas de datos ──
        ('FONTNAME',      (0,1), (-1,-1), 'Helvetica'),
        ('FONTSIZE',      (0,1), (-1,-1), 8),
        ('TEXTCOLOR',     (0,1), (-1,-1), rl_colors.HexColor('#1A1A1A')),
        ('TOPPADDING',    (0,1), (-1,-1), 5),
        ('BOTTOMPADDING', (0,1), (-1,-1), 5),
        ('LEFTPADDING',   (0,1), (-1,-1), 7),
        ('RIGHTPADDING',  (0,1), (-1,-1), 7),
        # Solo líneas horizontales
        ('LINEBELOW',     (0,0), (-1,0),  0.3, SEP),
        ('LINEBELOW',     (0,1), (-1,-1), 0.3, SEP),
        ('VALIGN',        (0,0), (-1,-1), 'MIDDLE'),
    ]
    # Zebra
    if zebra:
        for ri in range(1, len(data)):
            bg = LG if ri % 2 == 0 else rl_colors.white
            cmds.append(('BACKGROUND', (0,ri), (-1,ri), bg))
    # Alineación numérica
    for col in right_cols:
        cmds.append(('ALIGN', (col,0), (col,-1), 'RIGHT'))
    for col in center_cols:
        cmds.append(('ALIGN', (col,0), (col,-1), 'CENTER'))
    # Color variaciones
    R  = rl_colors.HexColor('#B03020')
    GR = rl_colors.HexColor('#217A45')
    for col in var_cols:
        for ri, row in enumerate(data[1:], 1):
            val = row[col] if col < len(row) else ''
            if isinstance(val, str) and val not in ('—', '', 'NUEVO'):
                raw = val.replace('+','').replace('%','').strip()
                try:
                    pct = float(raw)
                    color = GR if pct >= 0 else R
                    cmds.append(('TEXTCOLOR', (col,ri), (col,ri), color))
                    cmds.append(('FONTNAME',  (col,ri), (col,ri), 'Helvetica-Bold'))
                except ValueError:
                    pass
    t.setStyle(TableStyle(cmds))
    return t

def _pdf_kpi_tiles(tiles):
    """tiles = [(label, valor, color_hex), ...]  — fila de métricas destacadas."""
    ds = _pdf_ds()
    cells = []
    for label, valor, color_hex in tiles:
        lbl_st = ParagraphStyle('_kl', fontSize=6.5, textColor=rl_colors.HexColor('#888888'),
                                fontName='Helvetica', leading=9, alignment=TA_CENTER)
        val_st = ParagraphStyle('_kv', fontSize=17, textColor=rl_colors.HexColor(color_hex),
                                fontName='Helvetica-Bold', leading=21, alignment=TA_CENTER)
        cells.append([Paragraph(label.upper(), lbl_st), Paragraph(str(valor), val_st)])

    # Cada tile es una columna de 2 filas (label + valor) — armar como tabla interna
    tile_tbls = []
    for label, valor, color_hex in tiles:
        lbl_st = ParagraphStyle(f'_kl{id(label)}', fontSize=6.5,
                                textColor=rl_colors.HexColor('#999999'),
                                fontName='Helvetica', leading=9, alignment=TA_CENTER)
        val_st = ParagraphStyle(f'_kv{id(label)}', fontSize=18,
                                textColor=rl_colors.HexColor(color_hex),
                                fontName='Helvetica-Bold', leading=22, alignment=TA_CENTER)
        inner = Table(
            [[Paragraph(valor, val_st)], [Paragraph(label.upper(), lbl_st)]],
            colWidths=[None]
        )
        inner.setStyle(TableStyle([
            ('ALIGN',        (0,0), (-1,-1), 'CENTER'),
            ('TOPPADDING',   (0,0), (-1,-1), 6),
            ('BOTTOMPADDING',(0,0), (-1,-1), 6),
        ]))
        tile_tbls.append(inner)

    n = len(tile_tbls)
    W_page = 18*cm
    tile_w = W_page / n
    outer = Table([tile_tbls], colWidths=[tile_w]*n)
    outer.setStyle(TableStyle([
        ('BACKGROUND',    (0,0), (-1,-1), rl_colors.HexColor('#F7F7F4')),
        ('LINEAFTER',     (0,0), (-2,-1), 0.5, rl_colors.HexColor('#DDDDDA')),
        ('TOPPADDING',    (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ('LEFTPADDING',   (0,0), (-1,-1), 0),
        ('RIGHTPADDING',  (0,0), (-1,-1), 0),
        ('BOX',           (0,0), (-1,-1), 0.5, rl_colors.HexColor('#DDDDDA')),
    ]))
    return outer

def _pdf_section(txt, ds):
    """Título de sección dorado con línea fina."""
    return KeepTogether([
        HRFlowable(width='100%', thickness=0.4,
                   color=rl_colors.HexColor('#DDDDDA'), spaceAfter=2),
        Paragraph(txt, ds['sec']),
    ])

def _pdf_alert_row(nivel, msg, ds):
    badges = {'CRITICO': ('#B03020','CRÍTICO'), 'ALERTA': ('#B36A00','ALERTA'),
              'OK': ('#217A45','OK'), 'INFO': ('#555555','INFO')}
    bg, lbl = badges.get(nivel, ('#555555', nivel))
    badge_st = ParagraphStyle(f'_ba{nivel}', fontSize=6.5, fontName='Helvetica-Bold',
                              textColor=rl_colors.white, alignment=TA_CENTER)
    msg_st = ParagraphStyle(f'_bm{nivel}', fontSize=8, fontName='Helvetica',
                            textColor=rl_colors.HexColor('#1A1A1A'), leading=11)
    badge_tbl = Table([[Paragraph(lbl, badge_st)]], colWidths=[1.4*cm])
    badge_tbl.setStyle(TableStyle([
        ('BACKGROUND',    (0,0), (-1,-1), rl_colors.HexColor(bg)),
        ('TOPPADDING',    (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING',   (0,0), (-1,-1), 4),
        ('RIGHTPADDING',  (0,0), (-1,-1), 4),
    ]))
    row_tbl = Table([[badge_tbl, Paragraph(msg, msg_st)]], colWidths=[1.6*cm, 16.4*cm])
    row_tbl.setStyle(TableStyle([
        ('VALIGN',        (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING',    (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LINEBELOW',     (0,0), (-1,-1), 0.3, rl_colors.HexColor('#E8E8E4')),
    ]))
    return row_tbl


# ── PDF por representante ──────────────────────────────────────────────────────

def generar_pdf_repre(repre_sel):
    if not PDF_AVAILABLE:
        return None
    ds = _pdf_ds()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=1.5*cm, rightMargin=1.5*cm,
                            topMargin=2.5*cm, bottomMargin=1.8*cm)
    filtro_txt = repre_sel or "Región completa"
    cb = _pdf_page_cb(f"Informe de Representante", filtro_txt)
    story = []

    # ── KPIs destacados ───────────────────────────────────────────────────────
    try:
        df_r = DFS['x repre'][DFS['x repre']['Vendedor'] == repre_sel]
        act = get_ind(df_r, 'Año Actual Cajas', ['Vendedor','flia'])
        ant = get_ind(df_r, 'Año Anterior Cajas', ['Vendedor','flia'])
        tot_a = act['Total'].sum(); tot_b = ant['Total'].sum()
        var_t = (tot_a - tot_b) / tot_b * 100 if tot_b else 0
        all_act = get_ind(DFS['x repre'], 'Año Actual Cajas', ['Vendedor','flia'])
        rk_df = all_act.groupby('Vendedor')['Total'].sum().rank(ascending=False, method='min')
        rank = int(rk_df.get(repre_sel, 0))
        sign = '+' if var_t >= 0 else ''
        var_col = '#217A45' if var_t >= 0 else '#B03020'
        story.append(_pdf_kpi_tiles([
            ('Cajas Año Actual',   f"{int(tot_a):,}",           '#111111'),
            ('Cajas Año Anterior', f"{int(tot_b):,}",           '#555555'),
            ('Variación %',        f"{sign}{var_t:.0f}%",       var_col),
            ('Ranking Nacional',   f"#{rank} / {len(REPRESENTANTES)}", '#B8972A'),
        ]))
        story.append(Spacer(1, 0.4*cm))
    except Exception as e:
        story.append(Paragraph(f"Error KPIs: {e}", ds['alert']))

    # ── Evolución mensual ─────────────────────────────────────────────────────
    story.append(_pdf_section("Evolución Mensual", ds))
    try:
        df_r2 = DFS['x repre'][DFS['x repre']['Vendedor'] == repre_sel]
        act2  = get_ind(df_r2, 'Año Actual Cajas', ['Vendedor','flia'])
        ant2  = get_ind(df_r2, 'Año Anterior Cajas', ['Vendedor','flia'])
        rows = [['Mes', 'Año Actual', 'Año Anterior', 'Variación %']]
        for m in reversed(MC):
            if m in act2.columns:
                va = pd.to_numeric(act2[m], errors='coerce').sum()
                vb = pd.to_numeric(ant2[m], errors='coerce').sum() if m in ant2.columns else 0
                vp = (va - vb) / vb * 100 if vb else 0
                s = '+' if vp >= 0 else ''
                rows.append([m, f"{int(va):,}", f"{int(vb):,}", f"{s}{vp:.0f}%"])
        story.append(_pdf_tbl(rows, [3*cm, 5*cm, 5*cm, 5*cm],
                              var_cols=(3,), right_cols=(1,2,3)))
        story.append(Spacer(1, 0.3*cm))
    except Exception as e:
        story.append(Paragraph(f"Error evolución: {e}", ds['alert']))

    # ── Variación por familia ─────────────────────────────────────────────────
    story.append(_pdf_section("Variación por Familia vs Año Anterior", ds))
    try:
        rv = get_ind(DFS['x repre'], 'Var %', ['Vendedor','flia'])
        rv_r = rv[rv['Vendedor'] == repre_sel].copy()
        rv_r['pct'] = rv_r['Total'] * 100
        rv_r = rv_r.dropna(subset=['pct']).sort_values('pct')
        act_f = get_ind(df_r, 'Año Actual Cajas', ['Vendedor','flia']).groupby('flia')['Total'].sum()
        ant_f = get_ind(df_r, 'Año Anterior Cajas', ['Vendedor','flia']).groupby('flia')['Total'].sum()
        rows = [['Familia', 'Cajas Actual', 'Cajas Anterior', 'Variación %']]
        for _, row in rv_r.iterrows():
            s = '+' if row['pct'] >= 0 else ''
            a_val = int(act_f.get(row['flia'], 0))
            b_val = int(ant_f.get(row['flia'], 0))
            rows.append([row['flia'], f"{a_val:,}", f"{b_val:,}", f"{s}{row['pct']:.0f}%"])
        story.append(_pdf_tbl(rows, [6*cm, 4*cm, 4*cm, 4*cm],
                              var_cols=(3,), right_cols=(1,2,3)))
        story.append(Spacer(1, 0.3*cm))
    except Exception as e:
        story.append(Paragraph(f"Error familias: {e}", ds['alert']))

    # ── Mix por canal ─────────────────────────────────────────────────────────
    story.append(_pdf_section("Mix por Canal de Venta", ds))
    try:
        df_c = DFS['x repre x canal'][DFS['x repre x canal']['Vendedor'] == repre_sel]
        ac = get_ind(df_c, 'Año Actual Cajas', ['Canal','flia']).groupby('Canal')['Total'].sum().reset_index()
        bc = get_ind(df_c, 'Año Anterior Cajas', ['Canal','flia']).groupby('Canal')['Total'].sum().reset_index()
        mc2 = ac.merge(bc, on='Canal', suffixes=('_a','_b'))
        mc2['pct'] = mc2['Total_a'] / mc2['Total_a'].sum() * 100
        mc2['var'] = (mc2['Total_a'] - mc2['Total_b']) / mc2['Total_b'].replace(0, np.nan) * 100
        mc2 = mc2[mc2['Total_a'] > 0].sort_values('Total_a', ascending=False)
        rows = [['Canal', 'Cajas Actual', 'Participación %', 'Variación %']]
        for _, r in mc2.iterrows():
            s = '+' if pd.notna(r['var']) and r['var'] >= 0 else ''
            rows.append([r['Canal'], f"{int(r['Total_a']):,}",
                         f"{r['pct']:.0f}%", f"{s}{r['var']:.0f}%" if pd.notna(r['var']) else '—'])
        story.append(_pdf_tbl(rows, [6*cm, 4*cm, 4*cm, 4*cm],
                              var_cols=(3,), right_cols=(1,2,3)))
        story.append(Spacer(1, 0.3*cm))
    except Exception as e:
        story.append(Paragraph(f"Error canales: {e}", ds['alert']))

    # ── Top clientes ──────────────────────────────────────────────────────────
    if 'x cliente' in DFS:
        story.append(_pdf_section("Clientes — Mayor Variación", ds))
        try:
            cli = DFS['x cliente'][DFS['x cliente']['Vendedor'] == repre_sel]
            ac2 = get_ind(cli, 'Año Actual Cajas', ['Vendedor','Cliente','flia'])
            an2 = get_ind(cli, 'Año Anterior Cajas', ['Vendedor','Cliente','flia'])
            mc3 = ac2.merge(an2, on=['Vendedor','Cliente','flia'], suffixes=('_a','_b'))
            mc3 = mc3[mc3['Total_b'] > 0]
            mc3['dif'] = mc3['Total_a'] - mc3['Total_b']
            mc3['vp']  = mc3['dif'] / mc3['Total_b'] * 100

            rows_up = [['Cliente', 'Familia', 'Cajas Act.', 'Δ Cajas', 'Var %']]
            for _, r in mc3.nlargest(8, 'dif').iterrows():
                s = '+' if r['vp'] >= 0 else ''
                rows_up.append([str(r['Cliente'])[:28], str(r['flia'])[:16],
                                 f"{int(r['Total_a']):,}", f"+{int(r['dif']):,}",
                                 f"{s}{r['vp']:.0f}%"])
            story.append(Paragraph("Mayor crecimiento", ds['small']))
            story.append(_pdf_tbl(rows_up, [5.5*cm, 3.5*cm, 3*cm, 3*cm, 3*cm],
                                  var_cols=(4,), right_cols=(2,3,4)))

            story.append(Spacer(1, 0.3*cm))
            rows_dn = [['Cliente', 'Familia', 'Cajas Act.', 'Δ Cajas', 'Var %']]
            for _, r in mc3.nsmallest(8, 'dif').iterrows():
                s = '+' if r['vp'] >= 0 else ''
                rows_dn.append([str(r['Cliente'])[:28], str(r['flia'])[:16],
                                 f"{int(r['Total_a']):,}", f"{int(r['dif']):,}",
                                 f"{s}{r['vp']:.0f}%"])
            story.append(Paragraph("Mayor caída", ds['small']))
            story.append(_pdf_tbl(rows_dn, [5.5*cm, 3.5*cm, 3*cm, 3*cm, 3*cm],
                                  var_cols=(4,), right_cols=(2,3,4)))
        except Exception as e:
            story.append(Paragraph(f"Error clientes: {e}", ds['alert']))

    doc.build(story, onFirstPage=cb, onLaterPages=cb)
    buf.seek(0)
    return buf.read()


def generar_pdf_resumen(flia_sel=None, repre_sel=None, canal_sel=None):
    """PDF ejecutivo A4 — resumen completo de la selección activa."""
    if not PDF_AVAILABLE:
        return None
    ds = _pdf_ds()
    filtro_txt = " | ".join(filter(None, [flia_sel, repre_sel, canal_sel])) or "Región completa"
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=1.5*cm, rightMargin=1.5*cm,
                            topMargin=2.5*cm, bottomMargin=1.8*cm)
    cb = _pdf_page_cb("Informe de Ventas — Jefatura Nacional", filtro_txt)
    story = []

    # ── KPIs destacados ───────────────────────────────────────────────────────
    try:
        if canal_sel and repre_sel:
            df_k = DFS['x repre x canal']
            df_k = df_k[(df_k['Canal']==canal_sel) & (df_k['Vendedor']==repre_sel)]
            act_k = get_ind(df_k, 'Año Actual Cajas', ['Vendedor','flia'])
            ant_k = get_ind(df_k, 'Año Anterior Cajas', ['Vendedor','flia'])
        elif canal_sel:
            df_k = DFS['x flia x canal'][DFS['x flia x canal']['Canal']==canal_sel]
            act_k = get_ind(df_k, 'Año Actual Cajas', ['flia','Canal'])
            ant_k = get_ind(df_k, 'Año Anterior Cajas', ['flia','Canal'])
        elif repre_sel:
            df_k = DFS['x repre'][DFS['x repre']['Vendedor']==repre_sel]
            act_k = get_ind(df_k, 'Año Actual Cajas', ['Vendedor','flia'])
            ant_k = get_ind(df_k, 'Año Anterior Cajas', ['Vendedor','flia'])
        else:
            act_k = get_ind(DFS['x flia'], 'Año Actual Cajas', ['flia'])
            ant_k = get_ind(DFS['x flia'], 'Año Anterior Cajas', ['flia'])
        if flia_sel:
            act_k = act_k[act_k['flia']==flia_sel]
            ant_k = ant_k[ant_k['flia']==flia_sel]
        tot_a = act_k['Total'].sum(); tot_b = ant_k['Total'].sum()
        var_t = (tot_a - tot_b) / tot_b * 100 if tot_b else 0
        sign = '+' if var_t >= 0 else ''
        var_col = '#217A45' if var_t >= 0 else '#B03020'
        pend_tot = 0
        if 'pend' in DFS:
            dp = DFS['pend'].copy(); dp.columns = [c.strip() for c in dp.columns]
            dp['Pedidos Pendientes'] = pd.to_numeric(dp['Pedidos Pendientes'], errors='coerce')
            if repre_sel: dp = dp[dp['Vendedor'].str.strip()==repre_sel]
            pend_tot = int(dp['Pedidos Pendientes'].sum())
        n_reps = DFS['x repre']['Vendedor'].nunique() if 'x repre' in DFS else 0
        n_flias = DFS['x flia']['flia'].nunique() if 'x flia' in DFS else 0
        story.append(_pdf_kpi_tiles([
            ('Cajas Año Actual',   f"{int(tot_a):,}",     '#111111'),
            ('Cajas Año Anterior', f"{int(tot_b):,}",     '#555555'),
            ('Variación %',        f"{sign}{var_t:.0f}%", var_col),
            ('Pendientes',         f"{pend_tot:,}",       '#B03020' if pend_tot > 0 else '#555555'),
        ]))
        story.append(Spacer(1, 0.4*cm))
    except Exception as e:
        story.append(Paragraph(f"Error KPIs: {e}", ds['alert']))

    # ── Familias ─────────────────────────────────────────────────────────────
    story.append(_pdf_section("Familias — Variación vs Año Anterior", ds))
    try:
        if repre_sel:
            src = DFS['x repre'][DFS['x repre']['Vendedor']==repre_sel]
            af = get_ind(src, 'Año Actual Cajas', ['Vendedor','flia']).groupby('flia')['Total'].sum()
            bf = get_ind(src, 'Año Anterior Cajas', ['Vendedor','flia']).groupby('flia')['Total'].sum()
        else:
            af = get_ind(DFS['x flia'], 'Año Actual Cajas', ['flia']).groupby('flia')['Total'].sum()
            bf = get_ind(DFS['x flia'], 'Año Anterior Cajas', ['flia']).groupby('flia')['Total'].sum()
        mf = pd.DataFrame({'a':af,'b':bf}).reset_index().rename(columns={'index':'flia','flia':'flia'})
        if 'flia' not in mf.columns:
            mf.columns = ['flia','a','b']
        if flia_sel: mf = mf[mf['flia']==flia_sel]
        mf['var']  = (mf['a']-mf['b'])/mf['b'].replace(0,np.nan)*100
        mf['part'] = mf['a']/mf['a'].sum()*100
        mf = mf.sort_values('a', ascending=False)
        rows = [['Familia','Cajas Actual','Cajas Anterior','Variación %','Part. %']]
        for _, r in mf.iterrows():
            s = '+' if pd.notna(r['var']) and r['var']>=0 else ''
            vstr = f"{s}{r['var']:.0f}%" if pd.notna(r['var']) else '—'
            rows.append([r['flia'], f"{int(r['a']):,}", f"{int(r['b']):,}", vstr, f"{r['part']:.0f}%"])
        story.append(_pdf_tbl(rows, [5.5*cm,3.5*cm,3.5*cm,2.8*cm,2.7*cm],
                              var_cols=(3,), right_cols=(1,2,3,4)))
        story.append(Spacer(1,0.3*cm))
    except Exception as e:
        story.append(Paragraph(f"Error familias: {e}", ds['alert']))

    # ── Representantes ────────────────────────────────────────────────────────
    story.append(_pdf_section("Representantes — Ranking y Variación", ds))
    try:
        ar = get_ind(DFS['x repre'], 'Año Actual Cajas', ['Vendedor','flia'])
        br = get_ind(DFS['x repre'], 'Año Anterior Cajas', ['Vendedor','flia'])
        if flia_sel: ar=ar[ar['flia']==flia_sel]; br=br[br['flia']==flia_sel]
        if repre_sel: ar=ar[ar['Vendedor']==repre_sel]; br=br[br['Vendedor']==repre_sel]
        a2=ar.groupby('Vendedor')['Total'].sum().reset_index()
        b2=br.groupby('Vendedor')['Total'].sum().reset_index()
        mr=a2.merge(b2,on='Vendedor',suffixes=('_a','_b'))
        mr['var'] =(mr['Total_a']-mr['Total_b'])/mr['Total_b'].replace(0,np.nan)*100
        mr['part']=mr['Total_a']/mr['Total_a'].sum()*100
        mr=mr.sort_values('Total_a',ascending=False)
        rows=[['#','Representante','Cajas Actual','Variación %','Participación %']]
        for i,(_,r) in enumerate(mr.iterrows(),1):
            s='+' if pd.notna(r['var']) and r['var']>=0 else ''
            rows.append([str(i),r['Vendedor'][:32],f"{int(r['Total_a']):,}",
                         f"{s}{r['var']:.0f}%" if pd.notna(r['var']) else '—',f"{r['part']:.0f}%"])
        story.append(_pdf_tbl(rows,[1*cm,6*cm,3.5*cm,3*cm,4.5*cm],
                              var_cols=(3,),right_cols=(2,3,4),center_cols=(0,)))
        story.append(Spacer(1,0.3*cm))
    except Exception as e:
        story.append(Paragraph(f"Error representantes: {e}", ds['alert']))

    # ── Mix por canal ─────────────────────────────────────────────────────────
    story.append(_pdf_section("Mix por Canal de Venta", ds))
    try:
        if repre_sel:
            df_c=DFS['x repre x canal'][DFS['x repre x canal']['Vendedor']==repre_sel]
            act_c=get_ind(df_c,'Año Actual Cajas',['Canal','flia'])
            ant_c=get_ind(df_c,'Año Anterior Cajas',['Canal','flia'])
        elif canal_sel:
            df_c=DFS['x flia x canal'][DFS['x flia x canal']['Canal']==canal_sel]
            act_c=get_ind(df_c,'Año Actual Cajas',['Canal','flia'])
            ant_c=get_ind(df_c,'Año Anterior Cajas',['Canal','flia'])
        else:
            act_c=get_ind(DFS['x flia x canal'],'Año Actual Cajas',['Canal','flia'])
            ant_c=get_ind(DFS['x flia x canal'],'Año Anterior Cajas',['Canal','flia'])
        if flia_sel: act_c=act_c[act_c['flia']==flia_sel]; ant_c=ant_c[ant_c['flia']==flia_sel]
        ac2=act_c.groupby('Canal')['Total'].sum().reset_index()
        bc2=ant_c.groupby('Canal')['Total'].sum().reset_index()
        mc2=ac2.merge(bc2,on='Canal',suffixes=('_a','_b'))
        mc2['pct']=mc2['Total_a']/mc2['Total_a'].sum()*100
        mc2['var']=(mc2['Total_a']-mc2['Total_b'])/mc2['Total_b'].replace(0,np.nan)*100
        mc2=mc2[mc2['Total_a']>0].sort_values('Total_a',ascending=False)
        rows=[['Canal','Cajas Actual','Participación %','Variación %']]
        for _,r in mc2.iterrows():
            s='+' if pd.notna(r['var']) and r['var']>=0 else ''
            rows.append([r['Canal'],f"{int(r['Total_a']):,}",
                         f"{r['pct']:.0f}%",f"{s}{r['var']:.0f}%" if pd.notna(r['var']) else '—'])
        story.append(_pdf_tbl(rows,[6*cm,4*cm,4*cm,4*cm],
                              var_cols=(3,),right_cols=(1,2,3)))
        story.append(Spacer(1,0.3*cm))
    except Exception as e:
        story.append(Paragraph(f"Error canal: {e}", ds['alert']))

    # ── Alertas ───────────────────────────────────────────────────────────────
    story.append(_pdf_section("Alertas Automáticas", ds))
    try:
        flags = generar_red_flags(flia_sel=flia_sel, repre_sel=repre_sel, canal_sel=canal_sel)
        for nivel, msg in flags:
            story.append(_pdf_alert_row(nivel, msg, ds))
        story.append(Spacer(1,0.2*cm))
    except Exception as e:
        story.append(Paragraph(f"Error alertas: {e}", ds['alert']))

    # ── Pendientes ────────────────────────────────────────────────────────────
    if 'pend' in DFS:
        story.append(_pdf_section("Pedidos Pendientes — Top 10", ds))
        try:
            df_p=DFS['pend'].copy(); df_p.columns=[c.strip() for c in df_p.columns]
            df_p['Pedidos Pendientes']=pd.to_numeric(df_p['Pedidos Pendientes'],errors='coerce')
            df_p=df_p[df_p['Pedidos Pendientes']>0]
            if repre_sel: df_p=df_p[df_p['Vendedor'].str.strip()==repre_sel]
            agg_p=df_p.groupby('Vendedor')['Pedidos Pendientes'].sum().reset_index()
            agg_p=agg_p.sort_values('Pedidos Pendientes',ascending=False).head(10)
            total_p=agg_p['Pedidos Pendientes'].sum()
            rows=[['Representante','Pendientes','% del Total']]
            for _,r in agg_p.iterrows():
                rows.append([r['Vendedor'][:35],f"{int(r['Pedidos Pendientes']):,}",
                              f"{r['Pedidos Pendientes']/total_p*100:.0f}%"])
            story.append(_pdf_tbl(rows,[9*cm,4.5*cm,4.5*cm],right_cols=(1,2)))
        except Exception as e:
            story.append(Paragraph(f"Error pendientes: {e}", ds['alert']))

    doc.build(story, onFirstPage=cb, onLaterPages=cb)
    buf.seek(0)
    return buf.read()


def generar_pdf_tab(tab, flia_sel=None, repre_sel=None, canal_sel=None):
    """PDF ejecutivo por pestaña — diseño profesional para directorio."""
    if not PDF_AVAILABLE:
        return None
    ds = _pdf_ds()
    filtro_txt = " | ".join(filter(None, [flia_sel, repre_sel, canal_sel])) or "Región completa"
    TAB_LABELS = {
        'region':'Región','repre':'Representantes','clientes':'Clientes',
        'canales':'Canales','analisis':'Análisis','pendientes':'Pendientes',
    }
    titulo_doc = f"Informe de {TAB_LABELS.get(tab, tab)}"
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=1.5*cm, rightMargin=1.5*cm,
                            topMargin=2.5*cm, bottomMargin=1.8*cm)
    cb = _pdf_page_cb(titulo_doc, filtro_txt)
    story = []

    # ── REGIÓN ────────────────────────────────────────────────────────────────
    if tab == 'region':
        # KPIs
        try:
            act = get_ind(DFS['x flia'], 'Año Actual Cajas', ['flia'])
            ant = get_ind(DFS['x flia'], 'Año Anterior Cajas', ['flia'])
            if flia_sel:
                act = act[act['flia']==flia_sel]; ant = ant[ant['flia']==flia_sel]
            ta, tb = act['Total'].sum(), ant['Total'].sum()
            vt = (ta-tb)/tb*100 if tb else 0
            vcol = '#27AE60' if vt >= 0 else '#C0392B'
            story.append(_pdf_kpi_tiles([
                ('Cajas Año Actual', f"{int(ta):,}", '#2C3E50'),
                ('Cajas Año Anterior', f"{int(tb):,}", '#2C3E50'),
                ('Variación %', f"{'+'if vt>=0 else ''}{vt:.0f}%", vcol),
            ]))
            story.append(Spacer(1, 0.3*cm))
        except Exception as e:
            story.append(Paragraph(f"Error KPIs: {e}", ds['alert']))

        # Familias
        story.append(_pdf_section("RANKING FAMILIAS", ds))
        try:
            act_f = get_ind(DFS['x flia'], 'Año Actual Cajas', ['flia'])
            ant_f = get_ind(DFS['x flia'], 'Año Anterior Cajas', ['flia'])
            if flia_sel:
                act_f = act_f[act_f['flia']==flia_sel]; ant_f = ant_f[ant_f['flia']==flia_sel]
            mf = act_f.merge(ant_f, on='flia', suffixes=('_a','_b'))
            mf['var']  = (mf['Total_a']-mf['Total_b'])/mf['Total_b'].replace(0,np.nan)*100
            mf['part'] = mf['Total_a']/mf['Total_a'].sum()*100
            mf = mf.sort_values('Total_a', ascending=False).head(12)
            rows = [['Familia','Cajas Actual','Cajas Anterior','Var %','Part %']]
            for _, r in mf.iterrows():
                s = '+' if r['var']>=0 else ''
                rows.append([r['flia'], f"{int(r['Total_a']):,}", f"{int(r['Total_b']):,}",
                              f"{s}{r['var']:.0f}%", f"{r['part']:.0f}%"])
            story.append(_pdf_tbl(rows, [5*cm, 3*cm, 3*cm, 2.2*cm, 2.3*cm], var_cols=(3,), right_cols=(1,2,3,4)))
            story.append(Spacer(1, 0.3*cm))
        except Exception as e:
            story.append(Paragraph(f"Error familias: {e}", ds['alert']))

        # Representantes
        story.append(_pdf_section("RANKING REPRESENTANTES", ds))
        try:
            ar = get_ind(DFS['x repre'], 'Año Actual Cajas', ['Vendedor','flia'])
            br = get_ind(DFS['x repre'], 'Año Anterior Cajas', ['Vendedor','flia'])
            if flia_sel:
                ar = ar[ar['flia']==flia_sel]; br = br[br['flia']==flia_sel]
            a2 = ar.groupby('Vendedor')['Total'].sum().reset_index()
            b2 = br.groupby('Vendedor')['Total'].sum().reset_index()
            mr = a2.merge(b2, on='Vendedor', suffixes=('_a','_b'))
            mr['var']  = (mr['Total_a']-mr['Total_b'])/mr['Total_b'].replace(0,np.nan)*100
            mr['part'] = mr['Total_a']/mr['Total_a'].sum()*100
            mr = mr.sort_values('Total_a', ascending=False)
            rows = [['#','Representante','Cajas Actual','Var %','Part %']]
            for i, (_, r) in enumerate(mr.iterrows(), 1):
                s = '+' if r['var']>=0 else ''
                rows.append([str(i), r['Vendedor'][:30], f"{int(r['Total_a']):,}",
                              f"{s}{r['var']:.0f}%", f"{r['part']:.0f}%"])
            story.append(_pdf_tbl(rows, [0.8*cm, 6.2*cm, 3*cm, 2.2*cm, 2.3*cm], var_cols=(3,), right_cols=(2,3,4)))
        except Exception as e:
            story.append(Paragraph(f"Error representantes: {e}", ds['alert']))

    # ── REPRESENTANTES ────────────────────────────────────────────────────────
    elif tab == 'repre':
        if repre_sel:
            try:
                df_r = DFS['x repre'][DFS['x repre']['Vendedor']==repre_sel]
                act = get_ind(df_r, 'Año Actual Cajas', ['Vendedor','flia'])
                ant = get_ind(df_r, 'Año Anterior Cajas', ['Vendedor','flia'])
                ta, tb = act['Total'].sum(), ant['Total'].sum()
                vt = (ta-tb)/tb*100 if tb else 0
                all_act = get_ind(DFS['x repre'], 'Año Actual Cajas', ['Vendedor','flia'])
                rk = all_act.groupby('Vendedor')['Total'].sum().rank(ascending=False, method='min')
                rank = int(rk.get(repre_sel, 0))
                vcol = '#27AE60' if vt >= 0 else '#C0392B'
                story.append(_pdf_kpi_tiles([
                    ('Cajas Año Actual', f"{int(ta):,}", '#2C3E50'),
                    ('Cajas Año Anterior', f"{int(tb):,}", '#2C3E50'),
                    ('Variación %', f"{'+'if vt>=0 else ''}{vt:.0f}%", vcol),
                    ('Ranking Nacional', f"#{rank} de {len(REPRESENTANTES)}", '#2C3E50'),
                ]))
                story.append(Spacer(1, 0.3*cm))
            except Exception as e:
                story.append(Paragraph(f"Error KPIs: {e}", ds['alert']))

            # Evolución mensual
            story.append(_pdf_section("EVOLUCIÓN MENSUAL (cajas año actual)", ds))
            try:
                df_r2 = DFS['x repre'][DFS['x repre']['Vendedor']==repre_sel]
                act2  = get_ind(df_r2, 'Año Actual Cajas', ['Vendedor','flia'])
                ant2  = get_ind(df_r2, 'Año Anterior Cajas', ['Vendedor','flia'])
                mes_rows = [['Mes','Año Actual','Año Anterior','Var %']]
                for m in reversed(MC):
                    if m in act2.columns:
                        va = pd.to_numeric(act2[m], errors='coerce').sum()
                        vb = pd.to_numeric(ant2[m], errors='coerce').sum() if m in ant2.columns else 0
                        vp = (va-vb)/vb*100 if vb else 0
                        s2 = '+' if vp>=0 else ''
                        mes_rows.append([m, f"{int(va):,}", f"{int(vb):,}", f"{s2}{vp:.0f}%"])
                story.append(_pdf_tbl(mes_rows, [2.5*cm, 4*cm, 4*cm, 3*cm], var_cols=(3,), right_cols=(1,2,3)))
                story.append(Spacer(1, 0.3*cm))
            except Exception as e:
                story.append(Paragraph(f"Error evolución: {e}", ds['alert']))

            # Mix por canal
            story.append(_pdf_section("MIX POR CANAL", ds))
            try:
                df_c = DFS['x repre x canal'][DFS['x repre x canal']['Vendedor']==repre_sel]
                act_c = get_ind(df_c, 'Año Actual Cajas', ['Canal','flia'])
                ant_c = get_ind(df_c, 'Año Anterior Cajas', ['Canal','flia'])
                ac = act_c.groupby('Canal')['Total'].sum().reset_index()
                bc = ant_c.groupby('Canal')['Total'].sum().reset_index()
                mc2 = ac.merge(bc, on='Canal', suffixes=('_a','_b'))
                mc2['pct'] = mc2['Total_a']/mc2['Total_a'].sum()*100
                mc2['var'] = (mc2['Total_a']-mc2['Total_b'])/mc2['Total_b'].replace(0,np.nan)*100
                mc2 = mc2[mc2['Total_a']>0].sort_values('Total_a', ascending=False)
                rows = [['Canal','Cajas','Part %','Var %']]
                for _, r in mc2.iterrows():
                    s = '+' if r['var']>=0 else ''
                    rows.append([r['Canal'], f"{int(r['Total_a']):,}",
                                  f"{r['pct']:.0f}%", f"{s}{r['var']:.0f}%"])
                story.append(_pdf_tbl(rows, [5*cm, 3.5*cm, 3*cm, 3*cm], var_cols=(3,), right_cols=(1,2,3)))
                story.append(Spacer(1, 0.3*cm))
            except Exception as e:
                story.append(Paragraph(f"Error canales: {e}", ds['alert']))

            # Variación por familia
            story.append(_pdf_section("VARIACIÓN POR FAMILIA", ds))
            try:
                rv = get_ind(DFS['x repre'], 'Var %', ['Vendedor','flia'])
                rv_r = rv[rv['Vendedor']==repre_sel].copy()
                rv_r['pct'] = rv_r['Total']*100
                rv_r = rv_r.dropna(subset=['pct']).sort_values('pct')
                fam_rows = [['Familia', 'Var %']]
                for _, row in rv_r.iterrows():
                    s2 = '+' if row['pct']>=0 else ''
                    fam_rows.append([row['flia'], f"{s2}{row['pct']:.0f}%"])
                story.append(_pdf_tbl(fam_rows, [12*cm, 4*cm], var_cols=(1,), right_cols=(1,)))
                story.append(Spacer(1, 0.3*cm))
            except Exception as e:
                story.append(Paragraph(f"Error familias: {e}", ds['alert']))

            # Top clientes
            story.append(_pdf_section("TOP CLIENTES — VARIACIÓN", ds))
            try:
                if 'x cliente' in DFS:
                    cli = DFS['x cliente'][DFS['x cliente']['Vendedor']==repre_sel]
                    ac2 = get_ind(cli, 'Año Actual Cajas', ['Vendedor','Cliente','flia'])
                    an2 = get_ind(cli, 'Año Anterior Cajas', ['Vendedor','Cliente','flia'])
                    mc3 = ac2.merge(an2, on=['Vendedor','Cliente','flia'], suffixes=('_a','_b'))
                    mc3 = mc3[mc3['Total_b']>0]
                    mc3['dif'] = mc3['Total_a']-mc3['Total_b']
                    mc3['vp']  = mc3['dif']/mc3['Total_b']*100
                    rows = [['Cliente','Familia','Cajas Act','Var %']]
                    for _, r in mc3.nlargest(8,'dif').iterrows():
                        s2 = '+' if r['vp']>=0 else ''
                        rows.append([str(r['Cliente'])[:28], str(r['flia'])[:14],
                                      f"{int(r['Total_a']):,}", f"{s2}{r['vp']:.0f}%"])
                    story.append(_pdf_section("Mayor crecimiento", ds))
                    story.append(_pdf_tbl(rows, [7*cm, 3.5*cm, 2.5*cm, 2.5*cm], var_cols=(3,)))
                    story.append(Spacer(1, 0.2*cm))
                    rows2 = [['Cliente','Familia','Cajas Act','Var %']]
                    for _, r in mc3.nsmallest(8,'dif').iterrows():
                        s2 = '+' if r['vp']>=0 else ''
                        rows2.append([str(r['Cliente'])[:28], str(r['flia'])[:14],
                                       f"{int(r['Total_a']):,}", f"{s2}{r['vp']:.0f}%"])
                    story.append(_pdf_section("Mayor caída", ds))
                    story.append(_pdf_tbl(rows2, [7*cm, 3.5*cm, 2.5*cm, 2.5*cm], var_cols=(3,)))
            except Exception as e:
                story.append(Paragraph(f"Error clientes: {e}", ds['alert']))
        else:
            # Sin rep seleccionado: ranking general
            story.append(_pdf_section("RANKING GENERAL DE REPRESENTANTES", ds))
            try:
                ar = get_ind(DFS['x repre'], 'Año Actual Cajas', ['Vendedor','flia'])
                br = get_ind(DFS['x repre'], 'Año Anterior Cajas', ['Vendedor','flia'])
                if flia_sel:
                    ar = ar[ar['flia']==flia_sel]; br = br[br['flia']==flia_sel]
                a2 = ar.groupby('Vendedor')['Total'].sum().reset_index()
                b2 = br.groupby('Vendedor')['Total'].sum().reset_index()
                mr = a2.merge(b2, on='Vendedor', suffixes=('_a','_b'))
                mr['var']  = (mr['Total_a']-mr['Total_b'])/mr['Total_b'].replace(0,np.nan)*100
                mr['part'] = mr['Total_a']/mr['Total_a'].sum()*100
                mr = mr.sort_values('Total_a', ascending=False)
                rows = [['#','Representante','Cajas Actual','Var %','Part %']]
                for i, (_, r) in enumerate(mr.iterrows(), 1):
                    s = '+' if r['var']>=0 else ''
                    rows.append([str(i), r['Vendedor'][:30], f"{int(r['Total_a']):,}",
                                  f"{s}{r['var']:.0f}%", f"{r['part']:.0f}%"])
                story.append(_pdf_tbl(rows, [0.8*cm, 6.2*cm, 3*cm, 2.2*cm, 2.3*cm], var_cols=(3,), right_cols=(2,3,4)))
            except Exception as e:
                story.append(Paragraph(f"Error ranking: {e}", ds['alert']))

    # ── CLIENTES ──────────────────────────────────────────────────────────────
    elif tab == 'clientes':
        try:
            datos = analisis_clientes(repre_sel, flia_sel)
            n, p, a = len(datos['nuevos']), len(datos['perdidos']), len(datos['activos'])
            total_crec  = int(datos['nuevos']['dif'].sum())
            total_caida = int(datos['perdidos']['dif'].abs().sum())
            story.append(_pdf_kpi_tiles([
                ('Activos ambos años', str(a), '#2C3E50'),
                ('Con Crecimiento', str(n), '#217A45'),
                ('Con Caídas', str(p), '#8B2222'),
                ('Cajas ganadas', f"+{total_crec:,}", '#217A45'),
                ('Cajas perdidas', f"-{total_caida:,}", '#8B2222'),
            ]))
            story.append(Spacer(1, 0.3*cm))

            # Top crecimiento
            story.append(_pdf_section("TOP 10 — MAYOR CRECIMIENTO (cajas)", ds))
            rows = [['Cliente','Representante','Cajas Act','Δ Cajas','Var %']]
            for _, r in datos['top_sube'].head(10).iterrows():
                s = '+' if r['dif'] >= 0 else ''
                vstr = f"{r['var']:+.0f}%" if pd.notna(r['var']) else '—'
                rows.append([str(r['Cliente'])[:26], str(r['Vendedor'])[:20],
                              f"{int(r['act']):,}", f"{s}{int(r['dif']):,}", vstr])
            story.append(_pdf_tbl(rows, [5*cm, 4*cm, 2.5*cm, 2.5*cm, 2.5*cm], var_cols=(4,), right_cols=(2,3,4)))
            story.append(Spacer(1, 0.3*cm))

            # Top caída
            story.append(_pdf_section("TOP 10 — MAYOR CAÍDA (cajas)", ds))
            rows2 = [['Cliente','Representante','Cajas Act','Δ Cajas','Var %']]
            for _, r in datos['top_baja'].head(10).iterrows():
                vstr = f"{r['var']:+.0f}%" if pd.notna(r['var']) else '—'
                rows2.append([str(r['Cliente'])[:26], str(r['Vendedor'])[:20],
                               f"{int(r['act']):,}", f"{int(r['dif']):,}", vstr])
            story.append(_pdf_tbl(rows2, [5*cm, 4*cm, 2.5*cm, 2.5*cm, 2.5*cm], var_cols=(4,), right_cols=(2,3,4)))
            story.append(Spacer(1, 0.3*cm))

            # Con crecimiento / con caídas
            if not datos['nuevos'].empty:
                story.append(_pdf_section(f"CON CRECIMIENTO — TOP 10 (de {n} total)", ds))
                rowsN = [['Cliente','Representante','Cajas Act','Δ Cajas']]
                for _, r in datos['nuevos'].nlargest(10,'dif').iterrows():
                    rowsN.append([str(r['Cliente'])[:30], str(r['Vendedor'])[:22],
                                  f"{int(r['act']):,}", f"+{int(r['dif']):,}"])
                story.append(_pdf_tbl(rowsN, [6*cm, 4.5*cm, 2.5*cm, 2.5*cm]))
                story.append(Spacer(1, 0.3*cm))

            if not datos['perdidos'].empty:
                story.append(_pdf_section(f"CON CAÍDAS — TOP 10 (de {p} total)", ds))
                rowsP = [['Cliente','Representante','Cajas Act','Δ Cajas']]
                for _, r in datos['perdidos'].nsmallest(10,'dif').iterrows():
                    rowsP.append([str(r['Cliente'])[:30], str(r['Vendedor'])[:22],
                                  f"{int(r['act']):,}", f"{int(r['dif']):,}"])
                story.append(_pdf_tbl(rowsP, [6*cm, 4.5*cm, 2.5*cm, 2.5*cm]))

        except Exception as e:
            story.append(Paragraph(f"Error clientes: {e}", ds['alert']))

    # ── CANALES ───────────────────────────────────────────────────────────────
    elif tab == 'canales':
        story.append(_pdf_section("MIX POR CANAL", ds))
        try:
            if repre_sel:
                df_c = DFS['x repre x canal'][DFS['x repre x canal']['Vendedor']==repre_sel]
                act_c = get_ind(df_c, 'Año Actual Cajas', ['Canal','flia'])
                ant_c = get_ind(df_c, 'Año Anterior Cajas', ['Canal','flia'])
            else:
                act_c = get_ind(DFS['x flia x canal'], 'Año Actual Cajas', ['Canal','flia'])
                ant_c = get_ind(DFS['x flia x canal'], 'Año Anterior Cajas', ['Canal','flia'])
            if flia_sel:
                act_c = act_c[act_c['flia']==flia_sel]; ant_c = ant_c[ant_c['flia']==flia_sel]
            ac = act_c.groupby('Canal')['Total'].sum().reset_index()
            bc = ant_c.groupby('Canal')['Total'].sum().reset_index()
            mc2 = ac.merge(bc, on='Canal', suffixes=('_a','_b'))
            mc2['pct'] = mc2['Total_a']/mc2['Total_a'].sum()*100
            mc2['var'] = (mc2['Total_a']-mc2['Total_b'])/mc2['Total_b'].replace(0,np.nan)*100
            mc2 = mc2[mc2['Total_a']>0].sort_values('Total_a', ascending=False)
            rows = [['Canal','Cajas Año Actual','Cajas Año Anterior','Participación %','Var %']]
            for _, r in mc2.iterrows():
                s = '+' if r['var']>=0 else ''
                rows.append([r['Canal'], f"{int(r['Total_a']):,}", f"{int(r['Total_b']):,}",
                              f"{r['pct']:.0f}%", f"{s}{r['var']:.0f}%"])
            story.append(_pdf_tbl(rows, [4.5*cm, 3.5*cm, 3.5*cm, 3*cm, 2.5*cm], var_cols=(4,), right_cols=(1,2,3,4)))
        except Exception as e:
            story.append(Paragraph(f"Error canales: {e}", ds['alert']))

    # ── ANÁLISIS ──────────────────────────────────────────────────────────────
    elif tab == 'analisis':
        story.append(_pdf_section("RED FLAGS — ALERTAS AUTOMÁTICAS", ds))
        try:
            flags = generar_red_flags(flia_sel=flia_sel, repre_sel=repre_sel, canal_sel=canal_sel)
            for nivel, msg in flags:
                story.append(_pdf_alert_row(nivel, msg, ds))
        except Exception as e:
            story.append(Paragraph(f"Error red flags: {e}", ds['alert']))

        # Análisis quirúrgico para PDF
        try:
            aq_pdf = generar_analisis_quirurgico(flia_sel=flia_sel, repre_sel=repre_sel, canal_sel=canal_sel)
        except Exception as _eq:
            aq_pdf = {}

        def _vt_pdf(v):
            if pd.isna(v): return '—'
            return f"{'+'if v>=0 else ''}{v:.0f}%"

        cell_st    = ParagraphStyle('_cs',  fontSize=7, fontName='Helvetica', leading=9)
        cell_wh_st = ParagraphStyle('_csw', fontSize=7, fontName='Helvetica-Bold', leading=9,
                                    textColor=rl_colors.white)
        hdr_st  = ParagraphStyle('_hs', fontSize=7, fontName='Helvetica-Bold', leading=9,
                                 textColor=rl_colors.white)

        RED   = rl_colors.HexColor('#C0392B')
        GREEN = rl_colors.HexColor('#27AE60')
        GOLD  = rl_colors.HexColor('#C9A84C')
        GREY  = rl_colors.HexColor('#555555')

        def _pdf_var_color(v):
            if pd.isna(v): return None
            if v <= -10:   return rl_colors.HexColor('#7B1F1F')
            if v <= -5:    return rl_colors.HexColor('#3D1A00')
            if v <  0:     return rl_colors.HexColor('#251800')
            if v < 10:     return rl_colors.HexColor('#0D2E0D')
            return rl_colors.HexColor('#0A3A12')

        # Tabla familias
        story.append(Spacer(1, 0.3*cm))
        story.append(_pdf_section("FAMILIAS — VARIACIÓN YTD", ds))
        df_fq = aq_pdf.get('familias', pd.DataFrame())
        if not df_fq.empty:
            hdr_row = [Paragraph(x, hdr_st) for x in ['Familia','Actual','Anterior','Dif','Var%']]
            rows_fq = [hdr_row]
            ts_fq = [('BACKGROUND',(0,0),(-1,0),rl_colors.HexColor('#1A1A1A')),
                     ('GRID',(0,0),(-1,-1),0.3,rl_colors.HexColor('#CCCCCC')),
                     ('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3),
                     ('LEFTPADDING',(0,0),(-1,-1),4),('RIGHTPADDING',(0,0),(-1,-1),4),
                     ('ALIGN',(1,0),(-1,-1),'RIGHT')]
            for i, (_, r) in enumerate(df_fq.iterrows(), 1):
                v = r.get('Var%', np.nan)
                dif = r.get('Dif', r['Actual'] - r['Anterior'])
                bg = _pdf_var_color(v)
                rows_fq.append([
                    Paragraph(str(r['Familia'])[:30], cell_st),
                    Paragraph(f"{int(r['Actual']):,}", cell_st),
                    Paragraph(f"{int(r['Anterior']):,}", cell_st),
                    Paragraph(f"{int(dif):+,}", cell_st),
                    Paragraph(_vt_pdf(v), cell_wh_st if bg else cell_st),
                ])
                if bg:
                    ts_fq.append(('BACKGROUND',(4,i),(4,i),bg))
            tf_fq = Table(rows_fq, colWidths=[6*cm,3*cm,3*cm,2.5*cm,2.5*cm])
            tf_fq.setStyle(TableStyle(ts_fq))
            story.append(tf_fq)

        # Tabla canales
        story.append(Spacer(1, 0.3*cm))
        story.append(_pdf_section("CANALES — VARIACIÓN YTD", ds))
        df_cq = aq_pdf.get('canales', pd.DataFrame())
        if not df_cq.empty:
            hdr_row2 = [Paragraph(x, hdr_st) for x in ['Canal','Actual','Anterior','Dif','Var%']]
            rows_cq = [hdr_row2]
            ts_cq = [('BACKGROUND',(0,0),(-1,0),rl_colors.HexColor('#1A1A1A')),
                     ('GRID',(0,0),(-1,-1),0.3,rl_colors.HexColor('#CCCCCC')),
                     ('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3),
                     ('LEFTPADDING',(0,0),(-1,-1),4),('RIGHTPADDING',(0,0),(-1,-1),4),
                     ('ALIGN',(1,0),(-1,-1),'RIGHT')]
            for i, (_, r) in enumerate(df_cq.iterrows(), 1):
                v = r.get('Var%', np.nan)
                dif = r.get('Dif', r['Actual'] - r['Anterior'])
                bg = _pdf_var_color(v)
                rows_cq.append([
                    Paragraph(str(r['Canal'])[:30], cell_st),
                    Paragraph(f"{int(r['Actual']):,}", cell_st),
                    Paragraph(f"{int(r['Anterior']):,}", cell_st),
                    Paragraph(f"{int(dif):+,}", cell_st),
                    Paragraph(_vt_pdf(v), cell_wh_st if bg else cell_st),
                ])
                if bg:
                    ts_cq.append(('BACKGROUND',(4,i),(4,i),bg))
            tf_cq = Table(rows_cq, colWidths=[6*cm,3*cm,3*cm,2.5*cm,2.5*cm])
            tf_cq.setStyle(TableStyle(ts_cq))
            story.append(tf_cq)

        # Matriz Representante × Familia
        story.append(Spacer(1, 0.3*cm))
        story.append(_pdf_section("MAPA DE TENDENCIAS — REPRESENTANTE × FAMILIA", ds))
        df_mx = aq_pdf.get('matriz', pd.DataFrame())
        if not df_mx.empty:
            flias_mx = list(df_mx.columns)
            hdr_mx = [Paragraph('Rep.', hdr_st)] + [Paragraph(f[:9], hdr_st) for f in flias_mx]
            rows_mx = [hdr_mx]
            ts_mx   = [('BACKGROUND',(0,0),(-1,0),rl_colors.HexColor('#1A1A1A')),
                       ('GRID',(0,0),(-1,-1),0.3,rl_colors.HexColor('#333333')),
                       ('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3),
                       ('LEFTPADDING',(0,0),(-1,-1),3),('RIGHTPADDING',(0,0),(-1,-1),3),
                       ('ALIGN',(1,0),(-1,-1),'CENTER'),('FONTSIZE',(0,0),(-1,-1),7)]
            for ri, rep in enumerate(df_mx.index, 1):
                row_d = df_mx.loc[rep]
                cells_mx = [Paragraph(rep[:18], cell_st)]
                for ci, flia in enumerate(flias_mx, 1):
                    v = row_d.get(flia, np.nan)
                    bg = _pdf_var_color(v)
                    cells_mx.append(Paragraph(_vt_pdf(v), cell_wh_st if bg else cell_st))
                    if bg:
                        ts_mx.append(('BACKGROUND',(ci,ri),(ci,ri),bg))
                rows_mx.append(cells_mx)
            n_flias = len(flias_mx)
            rep_w = 3.5*cm
            flia_w = (17*cm - rep_w) / max(n_flias, 1)
            tf_mx = Table(rows_mx, colWidths=[rep_w] + [flia_w]*n_flias)
            tf_mx.setStyle(TableStyle(ts_mx))
            story.append(tf_mx)

    # ── PENDIENTES ────────────────────────────────────────────────────────────
    elif tab == 'pendientes':
        story.append(_pdf_section("PEDIDOS PENDIENTES POR VENDEDOR", ds))
        try:
            if 'pend' in DFS:
                df_p = DFS['pend'].copy()
                df_p.columns = [c.strip() for c in df_p.columns]
                df_p['Pedidos Pendientes'] = pd.to_numeric(df_p['Pedidos Pendientes'], errors='coerce')
                df_p = df_p[df_p['Pedidos Pendientes']>0]
                if repre_sel:
                    df_p = df_p[df_p['Vendedor'].str.strip()==repre_sel]
                agg_p = df_p.groupby('Vendedor')['Pedidos Pendientes'].sum().reset_index()
                agg_p = agg_p.sort_values('Pedidos Pendientes', ascending=False)
                total_p = agg_p['Pedidos Pendientes'].sum()
                rows = [['Vendedor','Pendientes','% del Total']]
                for _, r in agg_p.iterrows():
                    rows.append([r['Vendedor'][:35], f"{int(r['Pedidos Pendientes']):,}",
                                  f"{r['Pedidos Pendientes']/total_p*100:.0f}%"])
                rows.append(['TOTAL', f"{int(total_p):,}", '100%'])
                story.append(_pdf_tbl(rows, [8*cm, 4*cm, 3.5*cm]))
                story.append(Spacer(1, 0.4*cm))

                # Detalle por familia
                story.append(_pdf_section("DETALLE POR FAMILIA", ds))
                df_det = df_p.sort_values('Pedidos Pendientes', ascending=False)
                rows2 = [['Familia Producto','Vendedor','Pendientes']]
                for _, r in df_det.iterrows():
                    rows2.append([str(r.get('Familia Producto',''))[:30],
                                   str(r.get('Vendedor',''))[:25],
                                   f"{int(r.get('Pedidos Pendientes',0)):,}"])
                story.append(_pdf_tbl(rows2, [7*cm, 5*cm, 3.5*cm]))
        except Exception as e:
            story.append(Paragraph(f"Error pendientes: {e}", ds['alert']))

    doc.build(story, onFirstPage=cb, onLaterPages=cb)
    buf.seek(0)
    return buf.read()


INSIGHTS, ALERTAS, OPORTUNIDADES, FODA = generar_analisis()  # global defaults for PDF

# ── Estilos ────────────────────────────────────────────────────────────────────

CARD  = {'backgroundColor':C['surf'],'border':f"1px solid {C['border']}",
         'borderRadius':'3px','padding':'16px','marginBottom':'14px'}
CARD_G= {**CARD,'borderColor':C['gold']}
LABEL = {'color':C['muted'],'fontSize':'9px','letterSpacing':'2px',
         'textTransform':'uppercase','marginBottom':'6px'}
SEC   = {'color':C['gold'],'fontSize':'10px','letterSpacing':'3px',
         'textTransform':'uppercase','borderBottom':f"1px solid {C['border']}",
         'paddingBottom':'6px','marginBottom':'12px'}
G2    = {'display':'grid','gridTemplateColumns':'1fr 1fr','gap':'14px'}
TS    = {'backgroundColor':'transparent','color':C['muted'],'border':'none',
         'borderBottom':'2px solid transparent','padding':'10px 18px',
         'fontSize':'10px','letterSpacing':'2px','textTransform':'uppercase'}
TSS   = {**TS,'color':C['gold'],'borderBottom':f"2px solid {C['gold']}"}
DD    = {'backgroundColor':C['surf2'],'color':C['text'],'border':f"1px solid {C['border']}",
         'borderRadius':'2px','fontSize':'12px'}

# ── App ────────────────────────────────────────────────────────────────────────
print("[STARTUP] Creando app Dash...", flush=True)
app = Dash(__name__, suppress_callback_exceptions=True)
app.title = "Dashboard Ventas — Catena Zapata"
server = app.server
print("[STARTUP] App Dash creada OK", flush=True)

# CSS para dropdowns (texto visible sobre fondo oscuro)
app.index_string = '''
<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>{%title%}</title>
        {%favicon%}
        {%css%}
        <style>
            body { background-color: #0D0D0D; margin: 0; }

            /* ── Tabs ── */
            .dash-tab, .tab {
                background-color: transparent !important;
                color: #888888 !important;
                border: none !important;
                border-bottom: 2px solid transparent !important;
                padding: 10px 18px !important;
                font-size: 10px !important;
                letter-spacing: 2px !important;
                text-transform: uppercase !important;
                cursor: pointer !important;
            }
            .dash-tab--selected, .tab--selected { color: #C9A84C !important; border-bottom: 2px solid #C9A84C !important; background-color: transparent !important; }
            .dash-tabs { border-bottom: none !important; }

            /* ── Botones ── */
            button { -webkit-appearance: none !important; appearance: none !important; }
            button:disabled { opacity: 1 !important; cursor: not-allowed !important; }

            /* ── Login inputs (fallback CSS) ── */
            input[id="login-user"], input[id="login-pass"],
            #login-user > input, #login-pass > input,
            #login-user input, #login-pass input {
                background-color: #FFFFFF !important;
                color: #111111 !important;
                -webkit-text-fill-color: #111111 !important;
                color-scheme: light !important;
                -webkit-appearance: none !important;
                appearance: none !important;
                width: 100% !important;
                box-sizing: border-box !important;
                padding: 14px 16px !important;
                font-size: 16px !important;
                border: 1px solid #BBBBBB !important;
                border-radius: 4px !important;
                outline: none !important;
                display: block !important;
            }
            input[id="login-user"]:-webkit-autofill,
            input[id="login-pass"]:-webkit-autofill,
            #login-user input:-webkit-autofill,
            #login-pass input:-webkit-autofill {
                -webkit-box-shadow: 0 0 0px 1000px #FFFFFF inset !important;
                -webkit-text-fill-color: #111111 !important;
            }

            /* ── Scrollbar ── */
            ::-webkit-scrollbar { width: 6px; height: 6px; }
            ::-webkit-scrollbar-track { background: #161616; }
            ::-webkit-scrollbar-thumb { background: #2A2A2A; border-radius: 3px; }

            /* ── Impresión A4 ── */
            @media print {
                @page { size: A4 portrait; margin: 1.2cm; }
                .no-print { display: none !important; }
                body { background: white !important; color: #111 !important; font-size: 10pt; }
                .js-plotly-plot .plotly { background: white !important; }
                /* Fuerza salto de página entre cards */
                .print-break { page-break-before: always; }
            }
        </style>
    </head>
    <body>
        {%app_entry%}
        <footer>
            {%config%}
            {%scripts%}
            {%renderer%}
        </footer>
        <script>
        /* ── Dropdown dark theme via MutationObserver ── */
        (function() {
            var BG   = '#1E1E1E';
            var BG2  = '#2A2A2A';
            var TEXT = '#F0EDE8';
            var MUTE = '#888888';
            var GOLD = '#C9A84C';

            function paint(el) {
                var cls = (el.className || '').toString();
                if (!cls) return;
                /* control */
                if (cls.match(/control/i) && !cls.match(/indicator|container/i)) {
                    el.style.setProperty('background-color', BG,  'important');
                    el.style.setProperty('border-color',     BG2, 'important');
                    el.style.setProperty('box-shadow',       'none', 'important');
                    el.style.setProperty('min-height',       '36px', 'important');
                }
                /* input value / placeholder */
                if (cls.match(/singleValue|placeholder/i)) {
                    el.style.setProperty('color', cls.match(/placeholder/i) ? MUTE : TEXT, 'important');
                }
                /* menu / options */
                if (cls.match(/menu$/i) || cls.match(/MenuList/i)) {
                    el.style.setProperty('background-color', BG,  'important');
                    el.style.setProperty('border',     '1px solid '+BG2, 'important');
                    el.style.setProperty('z-index',    '9999', 'important');
                }
                if (cls.match(/option/i)) {
                    el.style.setProperty('background-color', BG,   'important');
                    el.style.setProperty('color',            TEXT,  'important');
                    el.style.setProperty('cursor',           'pointer', 'important');
                    el.addEventListener('mouseenter', function(){ this.style.setProperty('background-color', BG2, 'important'); });
                    el.addEventListener('mouseleave', function(){ this.style.setProperty('background-color', BG,  'important'); });
                }
                /* indicator separator */
                if (cls.match(/indicatorSeparator/i)) {
                    el.style.setProperty('background-color', BG2, 'important');
                }
                /* value container */
                if (cls.match(/ValueContainer|container/i)) {
                    el.style.setProperty('background-color', BG, 'important');
                }
            }

            function scanAll() {
                /* paint all elements inside .dash-dropdown */
                document.querySelectorAll('.dash-dropdown *').forEach(paint);
                /* also paint portal-rendered menus (outside .dash-dropdown in DOM) */
                document.querySelectorAll('[class*="menu"], [class*="option"]').forEach(paint);
            }

            var observer = new MutationObserver(function(muts) {
                muts.forEach(function(m) { m.addedNodes.forEach(function(n) { if (n.querySelectorAll) { paint(n); n.querySelectorAll('*').forEach(paint); } }); });
            });

            function init() {
                scanAll();
                observer.observe(document.body, { childList: true, subtree: true });
            }

            /* ── Forzar estilos en inputs de login — busca el <input> real ── */
            (function(){
                var PROPS = [
                    ['background-color',       '#ffffff'],
                    ['color',                  '#111111'],
                    ['-webkit-text-fill-color','#111111'],
                    ['color-scheme',           'light'],
                    ['-webkit-appearance',     'none'],
                    ['appearance',             'none'],
                    ['width',                  '100%'],
                    ['box-sizing',             'border-box'],
                    ['padding',                '14px 16px'],
                    ['font-size',              '16px'],
                    ['font-family',            'inherit'],
                    ['border',                 '1px solid #BBBBBB'],
                    ['border-radius',          '4px'],
                    ['outline',                'none'],
                    ['display',                'block'],
                    ['margin',                 '0'],
                ];
                function fix() {
                    ['login-user', 'login-pass'].forEach(function(wid) {
                        var wrap = document.getElementById(wid);
                        if (!wrap) return;
                        var inp = (wrap.tagName === 'INPUT') ? wrap : wrap.querySelector('input');
                        if (!inp) return;
                        PROPS.forEach(function(p){ inp.style.setProperty(p[0], p[1], 'important'); });
                    });
                }
                /* Disparar en múltiples momentos para cubrir el ciclo de render de Dash */
                fix();
                [50, 200, 500, 1000, 2000].forEach(function(ms){ setTimeout(fix, ms); });
                /* Re-aplicar en cada cambio del DOM (Dash puede re-renderizar) */
                new MutationObserver(fix).observe(document.body, {childList:true, subtree:true});
                /* Re-aplicar al hacer foco o click */
                document.addEventListener('focus', fix, true);
                document.addEventListener('click', fix, true);
            })();

            if (document.readyState === 'loading') {
                document.addEventListener('DOMContentLoaded', init);
            } else {
                init();
            }
        })();
        </script>
    </body>
</html>
'''

_LI = {
    'display': 'block', 'width': '100%', 'marginBottom': '14px',
    'padding': '0', 'border': 'none', 'backgroundColor': 'transparent',
}
_BTN_LOGOUT = {
    'backgroundColor': 'transparent', 'color': C['muted'],
    'border': f"1px solid {C['border']}", 'padding': '6px 14px',
    'fontSize': '9px', 'letterSpacing': '1.5px', 'textTransform': 'uppercase',
    'cursor': 'pointer', 'borderRadius': '2px', 'fontFamily': FONT,
    'WebkitAppearance': 'none', 'appearance': 'none',
}

app.layout = html.Div([

    dcc.Store(id='auth-store', storage_type='session'),

    # ── LOGIN PAGE ────────────────────────────────────────────────────────────
    html.Div(id='login-page', children=[
        html.Div([
            html.H1("CATENA ZAPATA",
                    style={'color': C['gold'], 'fontSize': '22px', 'letterSpacing': '5px',
                           'textTransform': 'uppercase', 'margin': '0 0 4px 0',
                           'fontWeight': '400', 'fontFamily': FONT, 'textAlign': 'center'}),
            html.P("Gerencia Nacional de Ventas",
                   style={'color': '#888888', 'fontSize': '9px', 'letterSpacing': '3px',
                          'textTransform': 'uppercase', 'textAlign': 'center', 'margin': '0 0 28px 0'}),
            html.Hr(style={'border': 'none', 'borderTop': '1px solid #333333', 'marginBottom': '24px'}),
            dcc.Input(id='login-user', type='text', placeholder='Usuario', debounce=False, style=_LI),
            dcc.Input(id='login-pass', type='password', placeholder='Contraseña', debounce=False, style=_LI),
            html.Div(id='login-error',
                     style={'color': '#CC0000', 'fontSize': '11px', 'marginBottom': '10px',
                            'textAlign': 'center', 'minHeight': '16px'}),
            html.Button('INGRESAR', id='btn-login', n_clicks=0, style={
                'backgroundColor': C['gold'], 'color': '#111', 'border': 'none',
                'padding': '13px 0', 'fontSize': '11px', 'letterSpacing': '2.5px',
                'textTransform': 'uppercase', 'cursor': 'pointer', 'borderRadius': '3px',
                'fontFamily': FONT, 'fontWeight': '700', 'width': '100%',
                'WebkitAppearance': 'none', 'appearance': 'none',
            }),
            html.Div("Vendedores: primera palabra del nombre / su clave",
                     style={'marginTop': '18px', 'textAlign': 'center',
                            'color': '#888888', 'fontSize': '9px', 'fontFamily': MONO}),
        ], style={
            'backgroundColor': '#161616',
            'border': '1px solid #2A2A2A',
            'borderRadius': '6px',
            'padding': '44px 40px',
            'width': '420px',
            'boxSizing': 'border-box',
            'colorScheme': 'light',
        }),
    ], style={'display': 'flex', 'justifyContent': 'center', 'alignItems': 'center',
              'minHeight': '100vh', 'backgroundColor': C['bg']}),

    # ── DASHBOARD PAGE ────────────────────────────────────────────────────────
    html.Div(id='dashboard-page', style={'display': 'none'}, children=[

        # Header
        html.Div([
            html.Div([
                html.H1("CATENA ZAPATA", style={'color':C['gold'],'fontSize':'20px',
                                                'letterSpacing':'4px','textTransform':'uppercase',
                                                'margin':0,'fontWeight':'400','fontFamily':FONT}),
                html.P("Gerencia Nacional de Ventas",
                       style={'color':C['muted'],'fontSize':'10px','letterSpacing':'2px',
                              'textTransform':'uppercase','margin':0}),
            ]),
            html.Div([
                html.Div(id='user-badge', style={'color':C['gold'],'fontSize':'10px',
                                                  'letterSpacing':'1px','marginBottom':'2px',
                                                  'textTransform':'uppercase'}),
                html.Div(id='timestamp', style={'color':C['muted'],'fontSize':'10px','marginBottom':'6px'}),
                html.Div([
                    html.Button("↓ Resumen PDF / A4", id='btn-print', style={
                        'backgroundColor': C['gold'], 'color': '#111', 'border': 'none',
                        'padding': '6px 14px', 'fontSize': '9px', 'letterSpacing': '1.5px',
                        'textTransform': 'uppercase', 'cursor': 'pointer', 'borderRadius': '2px',
                        'fontFamily': FONT, 'fontWeight': '600',
                        'WebkitAppearance': 'none', 'appearance': 'none', 'marginRight': '8px',
                    }),
                    html.Button("↻ Actualizar datos", id='btn-refresh', n_clicks=0, style={
                        'backgroundColor': 'transparent', 'color': C['gold'], 'border': f"1px solid {C['gold']}",
                        'padding': '6px 14px', 'fontSize': '9px', 'letterSpacing': '1.5px',
                        'textTransform': 'uppercase', 'cursor': 'pointer', 'borderRadius': '2px',
                        'fontFamily': FONT, 'fontWeight': '600',
                        'WebkitAppearance': 'none', 'appearance': 'none', 'marginRight': '8px',
                    }),
                    html.Button("SALIR", id='btn-logout', n_clicks=0, style=_BTN_LOGOUT),
                ]),
            ], style={'textAlign':'right'}),
            html.Div(id='_print_dummy', style={'display':'none'}),
        ], style={'backgroundColor':C['surf'],'borderBottom':f"2px solid {C['gold']}",
                  'padding':'14px 28px','display':'flex','justifyContent':'space-between','alignItems':'center'}),

        # ── Tabs (sticky) ────────────────────────────────────────────────────
        html.Div([
            dcc.Tabs(id='tabs', value='region', children=[
                dcc.Tab(label='REGION',         value='region',    style=TS, selected_style=TSS),
                dcc.Tab(label='REPRESENTANTES', value='repre',     style=TS, selected_style=TSS),
                dcc.Tab(label='CLIENTES',       value='clientes',  style=TS, selected_style=TSS),
                dcc.Tab(label='CANALES',        value='canales',   style=TS, selected_style=TSS),
                dcc.Tab(label='ANALISIS',       value='analisis',  style=TS, selected_style=TSS),
                dcc.Tab(label='PENDIENTES',     value='pendientes',style=TS, selected_style=TSS),
            ], style={'border':'none'}),
        ], style={
            'backgroundColor': C['surf'],
            'borderBottom': f"1px solid {C['border']}",
            'padding': '0 28px',
            'position': 'sticky', 'top': 0, 'zIndex': 200,
        }),

        # ── KPIs + filtros (sticky bajo los tabs) ────────────────────────────
        html.Div([
            html.Div(id='kpis'),
            html.Div([
                html.Div([
                    html.Div('Familia', style=LABEL),
                    dcc.Dropdown(id='dd-flia',
                        options=[{'label':'Todas','value':''}]+[{'label':f,'value':f} for f in FAMILIAS],
                        value='', clearable=False, style=DD),
                ], style={'flex':1}),
                html.Div(id='repre-wrap', style={'flex':1}, children=[
                    html.Div('Representante', style=LABEL),
                    dcc.Dropdown(id='dd-repre',
                        options=[{'label':'Todos','value':''}]+[{'label':r,'value':r} for r in REPRESENTANTES],
                        value='', clearable=False, style=DD),
                ]),
                html.Div([
                    html.Div('Canal', style=LABEL),
                    dcc.Dropdown(id='dd-canal',
                        options=[{'label':'Todos','value':''}]+[{'label':c,'value':c} for c in CANALES],
                        value='', clearable=False, style=DD),
                ], style={'flex':1}),
                html.Div([
                    html.Div('Meses', style=LABEL),
                    dcc.Dropdown(
                        id='dd-meses',
                        options=[{'label': m, 'value': m} for m in reversed(MC)],
                        value=None,
                        multi=True,
                        placeholder='Todos',
                        clearable=True,
                        style=DD,
                    ),
                ], style={'flex': 1.2}),
            ], style={'display':'flex','gap':'14px','marginBottom':'0','alignItems':'flex-end'}),
        ], style={
            'position': 'sticky', 'top': '38px', 'zIndex': 100,
            'backgroundColor': C['bg'],
            'padding': '16px 28px 12px',
            'borderBottom': f"1px solid {C['border']}",
        }),

        # ── Contenido scrolleable ─────────────────────────────────────────────
        html.Div(id='content', style={'padding': '16px 28px 24px'}),

        dcc.Download(id='download-pdf'),
        dcc.Download(id='download-resumen'),
        dcc.Download(id='download-tab-pdf'),
        dcc.Interval(id='interval', interval=86400000, n_intervals=0),
        html.Div(id='_refresh_dummy', style={'display': 'none'}),
        dcc.Store(id='data-version', data=0),
        dcc.Store(id='drive-modified', data=''),

    ]),

], style={'backgroundColor':C['bg'],'minHeight':'100vh','fontFamily':FONT,'color':C['text']})


# ── Callbacks ──────────────────────────────────────────────────────────────────

@app.callback(
    Output('kpis','children'),
    Input('interval','n_intervals'),
    Input('data-version','data'),
    Input('dd-flia','value'),
    Input('dd-repre','value'),
    Input('dd-canal','value'),
    Input('dd-meses','value'),
    State('auth-store','data'),
)
def cb_kpis(n, _ver, flia, repre, canal, meses, auth):
    if auth and auth.get('role') == 'vendedor':
        repre = auth.get('repre', '')
    return build_kpis(flia or None, repre or None, canal or None, meses_sel=meses or None)

@app.callback(
    Output('data-version', 'data'),
    Output('drive-modified', 'data'),
    Output('timestamp', 'children'),
    Input('interval', 'n_intervals'),
    State('drive-modified', 'data'),
    State('data-version', 'data'),
    prevent_initial_call=False,
)
def cb_auto_sync(n, last_modified, version):
    global DFS, MC, FAMILIAS, REPRESENTANTES, CANALES
    current_modified = get_drive_modified_time()
    now = datetime.now().strftime('%d/%m/%Y %H:%M')
    if current_modified and current_modified != last_modified:
        DFS = load_data()
        MC = month_cols(DFS['x flia'])
        FAMILIAS = sorted(DFS['x flia']['flia'].unique().tolist())
        REPRESENTANTES = sorted(DFS['x repre']['Vendedor'].unique().tolist())
        CANALES = sorted([str(x) for x in DFS['x flia x canal']['Canal'].dropna().unique().tolist()])
        return (version or 0) + 1, current_modified, f"Actualizado automáticamente: {now}"
    return no_update, last_modified or '', f"Verificado: {now}"

@app.callback(
    Output('content','children'),
    Input('tabs','value'),
    Input('data-version','data'),
    Input('dd-flia','value'),
    Input('dd-repre','value'),
    Input('dd-canal','value'),
    Input('dd-meses','value'),
    State('auth-store','data'),
)
def cb_content(tab, _ver, flia, repre, canal, meses, auth):
    if auth and auth.get('role') == 'vendedor':
        repre = auth.get('repre', '')
    flia  = flia  or None
    repre = repre or None
    canal = canal or None
    meses = meses or None

    _pbt = html.Div(
        html.Button("↓ Imprimir resumen de esta pestaña", id='btn-tab-pdf', n_clicks=0, style={
            'backgroundColor': 'transparent', 'color': C['gold'],
            'border': f"1px solid {C['gold']}", 'padding': '6px 16px',
            'fontSize': '9px', 'letterSpacing': '1.5px', 'textTransform': 'uppercase',
            'cursor': 'pointer', 'borderRadius': '2px', 'fontFamily': FONT,
            'WebkitAppearance': 'none', 'appearance': 'none',
        }),
        style={'textAlign': 'right', 'marginBottom': '10px'},
    )

    if tab == 'region':
        if auth and auth.get('role') == 'vendedor':
            return html.Div("Acceso restringido.", style={
                'color': C['muted'], 'textAlign': 'center',
                'padding': '60px', 'fontSize': '13px', 'letterSpacing': '2px'
            })
        return html.Div([
            _pbt,
            html.Div([dcc.Graph(figure=fig_flia_ranking(flia, canal, meses), config={'displayModeBar':False})], style=CARD),
            html.Div([dcc.Graph(figure=fig_evolucion(flia, repre, canal, meses), config={'displayModeBar':False})], style=CARD),
            html.Div([dcc.Graph(figure=fig_ranking_ejecutivo(flia, repre, canal, meses), config={'displayModeBar':False})], style=CARD),
        ])

    elif tab == 'repre':
        habilitado = bool(repre and PDF_AVAILABLE)
        btn_style = {
            'backgroundColor': C['gold'] if habilitado else C['surf2'],
            'color': C['bg'] if habilitado else C['muted'],
            'border': f"1px solid {C['gold'] if habilitado else C['border']}",
            'padding': '9px 20px',
            'fontSize': '10px',
            'letterSpacing': '2px',
            'textTransform': 'uppercase',
            'cursor': 'pointer' if habilitado else 'not-allowed',
            'borderRadius': '2px',
            'fontFamily': FONT,
            'marginBottom': '14px',
            'WebkitAppearance': 'none',
            'MozAppearance': 'none',
            'appearance': 'none',
            'outline': 'none',
            'opacity': '1',
        }
        pdf_label = (
            f"↓ Generar PDF — {repre}" if habilitado
            else ("← Seleccioná un representante para generar PDF" if PDF_AVAILABLE
                  else "PDF no disponible (pip install reportlab)")
        )
        return html.Div([
            html.Div([
                html.Button(pdf_label, id='btn-pdf', style=btn_style),
                _pbt,
            ], style={'display': 'flex', 'justifyContent': 'space-between', 'alignItems': 'center', 'marginBottom': '4px'}),
            html.Div([dcc.Graph(figure=fig_repre_ranking(flia, canal, repre, meses, solo=(auth and auth.get('role')=='vendedor')), config={'displayModeBar':False})], style=CARD),
            html.Div([dcc.Graph(figure=fig_canal_mix(flia, repre, canal, meses), config={'displayModeBar':False})], style=CARD),
        ])

    elif tab == 'clientes':
        try:
            datos = analisis_clientes(repre, flia, meses)
            n_nuevos   = len(datos['nuevos_cli'])
            n_crec     = len(datos['crecieron'])
            n_caida    = len(datos['cayeron'])
            n_perdidos = len(datos['perdidos_cli'])
            dif_crec   = int(datos['crecieron']['dif'].sum())
            dif_caida  = int(datos['cayeron']['dif'].abs().sum())
            stats_items = [
                ('CLIENTES NUEVOS',        str(n_nuevos),       C['gold']),
                ('CRECIERON',              str(n_crec),         C['green']),
                ('CAYERON',                str(n_caida),        C['red']),
                ('CAJAS CRECIMIENTO',      f"{dif_crec:,}",     C['green']),
                ('CAJAS CAÍDAS',           f"{dif_caida:,}",    C['red']),
            ]
            stats_bar = html.Div([
                html.Div([
                    html.Div(lbl, style={'color':C['muted'],'fontSize':'8px','letterSpacing':'1.5px',
                                         'textTransform':'uppercase','marginBottom':'3px'}),
                    html.Div(val, style={'color':col,'fontSize':'18px','fontWeight':'700','fontFamily':MONO}),
                ], style={'backgroundColor':C['surf'],'border':f"1px solid {C['border']}",
                          'borderRadius':'3px','padding':'12px','textAlign':'center'})
                for lbl, val, col in stats_items
            ], style={'display':'grid','gridTemplateColumns':'repeat(5,1fr)','gap':'10px','marginBottom':'14px'})
        except Exception as e:
            datos = {'top_sube': pd.DataFrame(), 'top_baja': pd.DataFrame(),
                     'nuevos': pd.DataFrame(), 'perdidos': pd.DataFrame()}
            stats_bar = html.Div(f"Error stats: {e}", style={'color': C['red']})
        return html.Div([
            _pbt,
            stats_bar,
            html.Div([dcc.Graph(figure=fig_clientes_nuevos_perdidos(datos), config={'displayModeBar':False})], style=CARD),
        ])

    elif tab == 'canales':
        return html.Div([
            _pbt,
            html.Div([dcc.Graph(figure=fig_canal_barras(canal, flia, repre, meses), config={'displayModeBar':False})], style=CARD),
            html.Div([dcc.Graph(figure=fig_canal_mix(flia, repre, canal, meses), config={'displayModeBar':False})], style=CARD),
        ])

    elif tab == 'analisis':
        flags_d = generar_red_flags(flia_sel=flia, repre_sel=repre, canal_sel=canal, meses_sel=meses)
        aq = generar_analisis_quirurgico(flia_sel=flia, repre_sel=repre, canal_sel=canal, meses_sel=meses)

        _flag_colors = {'CRITICO': C['red'], 'ALERTA': '#E67E22', 'OK': C['green'], 'INFO': C['muted']}

        def _flag_item(nivel, msg):
            bc = _flag_colors.get(nivel, C['muted'])
            return html.Div([
                html.Span(nivel, style={'backgroundColor': bc, 'color': 'white', 'fontSize': '8px',
                    'letterSpacing': '1px', 'padding': '2px 6px', 'borderRadius': '2px',
                    'marginRight': '8px', 'fontWeight': '700', 'flexShrink': '0'}),
                html.Span(msg, style={'fontSize': '12px', 'lineHeight': '1.5'}),
            ], style={'padding': '6px 10px', 'borderLeft': f"2px solid {bc}", 'marginBottom': '4px',
                      'display': 'flex', 'alignItems': 'flex-start'})

        def _var_color(v):
            if pd.isna(v): return '#222222', C['muted']
            if v <= -20:   return '#5C0000', '#FFB3B3'
            if v <= -10:   return '#7B1F1F', '#FFAAAA'
            if v <= -5:    return '#3D1A00', '#FF9057'
            if v <  0:     return '#251800', C['gold']
            if v <  5:     return '#0A1A0A', '#7FCF7F'
            if v < 15:     return '#0D2E0D', C['green']
            return '#0A3A12', '#52FF7A'

        def _vt(v):
            if pd.isna(v): return '—'
            return f"{'+'if v>=0 else ''}{v:.0f}%"

        def _th(txt, align='right'):
            return html.Th(txt, style={'padding':'5px 8px','fontSize':'8px','letterSpacing':'1px',
                'color':C['muted'],'textTransform':'uppercase','borderBottom':f"1px solid {C['border']}",
                'textAlign':align,'fontWeight':'400','whiteSpace':'nowrap'})

        def _td(txt, color=None, align='right', bold=False, bg=None):
            st = {'padding':'5px 8px','fontSize':'11px','textAlign':align,
                  'color': color or C['text'], 'borderBottom':f"1px solid {C['border']}",
                  'fontFamily':MONO}
            if bold: st['fontWeight'] = '700'
            if bg:   st['backgroundColor'] = bg
            return html.Td(txt, style=st)

        def _tabla_diag(df, col_name, titulo):
            if df is None or df.empty:
                return html.Div([html.Div(titulo, style=SEC),
                    html.Div("Sin datos", style={'color':C['muted'],'fontSize':'12px'})], style=CARD)
            tot_a = df['Actual'].sum(); tot_b = df['Anterior'].sum()
            tot_v = (tot_a - tot_b) / tot_b * 100 if tot_b else 0
            filas = []
            for _, row in df.iterrows():
                v = row.get('Var%', np.nan)
                bg_c, txt_c = _var_color(v)
                dif = row.get('Dif', row['Actual'] - row['Anterior'])
                filas.append(html.Tr([
                    html.Td(str(row[col_name])[:26], style={'padding':'5px 8px','fontSize':'11px',
                        'borderBottom':f"1px solid {C['border']}"}),
                    _td(f"{int(row['Actual']):,}"),
                    _td(f"{int(row['Anterior']):,}", color=C['muted']),
                    _td(f"{int(dif):+,}", color=C['green'] if dif>=0 else C['red']),
                    html.Td(_vt(v), style={'padding':'5px 8px','fontSize':'11px','textAlign':'right',
                        'backgroundColor':bg_c,'color':txt_c,'fontWeight':'700',
                        'fontFamily':MONO,'borderBottom':f"1px solid {C['border']}"}),
                ]))
            bg_t, txt_t = _var_color(tot_v)
            filas.append(html.Tr([
                html.Td("TOTAL", style={'padding':'6px 8px','fontSize':'8px','letterSpacing':'2px',
                    'color':C['muted'],'textTransform':'uppercase','fontWeight':'700',
                    'borderTop':f"1px solid {C['border']}"}),
                _td(f"{int(tot_a):,}", bold=True),
                _td(f"{int(tot_b):,}", color=C['muted'], bold=True),
                _td(f"{int(tot_a-tot_b):+,}", color=C['green'] if tot_a>=tot_b else C['red'], bold=True),
                html.Td(_vt(tot_v), style={'padding':'6px 8px','fontSize':'11px','textAlign':'right',
                    'backgroundColor':bg_t,'color':txt_t,'fontWeight':'700',
                    'fontFamily':MONO,'borderTop':f"1px solid {C['border']}"}),
            ], style={'borderTop':f"2px solid {C['border']}"}))
            return html.Div([
                html.Div(titulo, style=SEC),
                html.Table([
                    html.Thead(html.Tr([_th(col_name,'left'),_th('Actual'),_th('Anterior'),_th('Dif'),_th('Var%')])),
                    html.Tbody(filas)
                ], style={'width':'100%','borderCollapse':'collapse'})
            ], style=CARD)

        def _tabla_matriz(pivot, tot_repre):
            if pivot is None or pivot.empty:
                return html.Div("Sin datos para la matriz.", style={'color':C['muted'],'fontSize':'12px'})
            flias = list(pivot.columns)
            reps  = list(pivot.index)
            header = html.Thead(html.Tr([
                html.Th("Representante", style={'padding':'5px 8px','fontSize':'8px','color':C['muted'],
                    'textTransform':'uppercase','letterSpacing':'1px',
                    'borderBottom':f"1px solid {C['border']}",
                    'textAlign':'left','fontWeight':'400','minWidth':'120px'}),
                html.Th("Total", style={'padding':'5px 8px','fontSize':'8px','color':C['muted'],
                    'textTransform':'uppercase','letterSpacing':'1px',
                    'borderBottom':f"1px solid {C['border']}",
                    'textAlign':'right','fontWeight':'400','whiteSpace':'nowrap'}),
                *[html.Th(f[:13], style={'padding':'4px 5px','fontSize':'7px','color':C['muted'],
                    'textTransform':'uppercase','letterSpacing':'0.5px',
                    'borderBottom':f"1px solid {C['border']}",
                    'textAlign':'center','fontWeight':'400','whiteSpace':'nowrap'}) for f in flias],
                html.Th("Peor flia", style={'padding':'5px 8px','fontSize':'8px','color':C['red'],
                    'textTransform':'uppercase','letterSpacing':'1px',
                    'borderBottom':f"1px solid {C['border']}",
                    'textAlign':'left','fontWeight':'400','whiteSpace':'nowrap'}),
            ]))
            rows = []
            for rep in reps:
                row_data = pivot.loc[rep]
                rep_mean = row_data.dropna().mean() if not row_data.isna().all() else np.nan
                bg_mean, txt_mean = _var_color(rep_mean)
                tot = tot_repre.get(rep, 0) if tot_repre is not None else 0
                valid = row_data.dropna()
                worst_flia = valid.idxmin() if not valid.empty else None
                worst_val  = valid.min()    if not valid.empty else np.nan
                cells = []
                for flia in flias:
                    v = row_data.get(flia, np.nan)
                    bg_c, txt_c = _var_color(v)
                    cells.append(html.Td(_vt(v), style={
                        'padding':'5px 5px','fontSize':'10px','textAlign':'center',
                        'backgroundColor':bg_c,'color':txt_c,'fontFamily':MONO,
                        'fontWeight':'600','borderRight':'1px solid #111',
                        'borderBottom':'1px solid #111','whiteSpace':'nowrap',
                    }))
                rows.append(html.Tr([
                    html.Td([
                        html.Div(rep[:22], style={'fontSize':'10px','fontWeight':'600'}),
                        html.Div(_vt(rep_mean), style={'fontSize':'8px','color':txt_mean,
                            'fontFamily':MONO,'marginTop':'2px',
                            'backgroundColor':bg_mean,'padding':'1px 4px',
                            'borderRadius':'2px','display':'inline-block'}),
                    ], style={'padding':'5px 8px','borderBottom':'1px solid #111','verticalAlign':'top'}),
                    html.Td(f"{int(tot):,}" if not pd.isna(tot) else '—',
                        style={'padding':'5px 8px','fontSize':'10px','textAlign':'right',
                            'color':C['gold'],'fontFamily':MONO,'fontWeight':'700',
                            'borderBottom':'1px solid #111','whiteSpace':'nowrap'}),
                    *cells,
                    html.Td([
                        html.Div(str(worst_flia)[:14] if worst_flia else '—',
                            style={'fontSize':'9px','color':C['red']}),
                        html.Div(_vt(worst_val),
                            style={'fontSize':'8px','color':C['red'],'fontFamily':MONO}),
                    ], style={'padding':'5px 8px','borderBottom':'1px solid #111','verticalAlign':'top'}),
                ]))
            return html.Table([header, html.Tbody(rows)],
                style={'width':'100%','borderCollapse':'collapse'})

        def _sec_canal_repre(df):
            if df is None or df.empty:
                return html.Div("Sin datos.", style={'color':C['muted'],'fontSize':'12px'})
            reps = sorted(df['Vendedor'].unique())
            cards = []
            for rep in reps:
                sub = df[df['Vendedor'] == rep].sort_values('var')
                rep_a = sub['Total_a'].sum(); rep_b = sub['Total_b'].sum()
                rep_v = (rep_a - rep_b) / rep_b * 100 if rep_b else 0
                bg_rv, txt_rv = _var_color(rep_v)
                filas_r = []
                for _, row in sub.iterrows():
                    v = row['var']
                    bg_c, txt_c = _var_color(v)
                    filas_r.append(html.Tr([
                        html.Td(str(row['Canal'])[:20], style={'padding':'4px 8px','fontSize':'10px',
                            'borderBottom':f"1px solid {C['border']}"}),
                        html.Td(f"{int(row['Total_a']):,}", style={'padding':'4px 8px','fontSize':'10px',
                            'textAlign':'right','fontFamily':MONO,
                            'borderBottom':f"1px solid {C['border']}"}),
                        html.Td(f"{int(row['dif']):+,}", style={'padding':'4px 8px','fontSize':'10px',
                            'textAlign':'right','fontFamily':MONO,
                            'color':C['green'] if row['dif']>=0 else C['red'],
                            'borderBottom':f"1px solid {C['border']}"}),
                        html.Td(_vt(v), style={'padding':'4px 8px','fontSize':'10px',
                            'textAlign':'right','fontFamily':MONO,'fontWeight':'700',
                            'backgroundColor':bg_c,'color':txt_c,
                            'borderBottom':f"1px solid {C['border']}"}),
                    ]))
                cards.append(html.Div([
                    html.Div([
                        html.Span(rep[:22], style={'fontSize':'10px','fontWeight':'700','letterSpacing':'0.5px'}),
                        html.Span(_vt(rep_v), style={'fontSize':'10px','fontWeight':'700',
                            'fontFamily':MONO,'marginLeft':'8px',
                            'color':txt_rv,'backgroundColor':bg_rv,
                            'padding':'1px 6px','borderRadius':'2px'}),
                    ], style={'marginBottom':'8px','display':'flex','alignItems':'center','justifyContent':'space-between'}),
                    html.Table([
                        html.Thead(html.Tr([_th('Canal','left'),_th('Actual'),_th('Dif'),_th('Var%')])),
                        html.Tbody(filas_r)
                    ], style={'width':'100%','borderCollapse':'collapse'})
                ], style={**CARD,'marginBottom':'0'}))
            cols = '1fr 1fr' if len(reps) > 1 else '1fr'
            return html.Div(cards, style={'display':'grid','gridTemplateColumns':cols,'gap':'14px'})

        def _sec_tendencia(df):
            if df is None or df.empty:
                return None
            filas_t = []
            for _, row in df.iterrows():
                v = row['Var%']
                bg_c, txt_c = _var_color(v)
                filas_t.append(html.Tr([
                    html.Td(row['Mes'], style={'padding':'4px 10px','fontSize':'11px',
                        'borderBottom':f"1px solid {C['border']}"}),
                    html.Td(f"{int(row['Actual']):,}", style={'padding':'4px 10px','fontSize':'11px',
                        'textAlign':'right','fontFamily':MONO,
                        'borderBottom':f"1px solid {C['border']}"}),
                    html.Td(f"{int(row['Anterior']):,}", style={'padding':'4px 10px','fontSize':'11px',
                        'textAlign':'right','fontFamily':MONO,'color':C['muted'],
                        'borderBottom':f"1px solid {C['border']}"}),
                    html.Td(_vt(v), style={'padding':'4px 10px','fontSize':'11px',
                        'textAlign':'right','fontFamily':MONO,'fontWeight':'700',
                        'backgroundColor':bg_c,'color':txt_c,
                        'borderBottom':f"1px solid {C['border']}"}),
                ]))
            return html.Div([
                html.Div("Tendencia Mensual — Región Total", style=SEC),
                html.Table([
                    html.Thead(html.Tr([_th('Mes','left'),_th('Actual'),_th('Anterior'),_th('Var%')])),
                    html.Tbody(filas_t)
                ], style={'width':'100%','borderCollapse':'collapse'})
            ], style=CARD)

        # ── Ensamblar secciones
        _df_empty = pd.DataFrame()
        sec_flags = html.Div([
            html.Div("Alertas Automáticas", style={**SEC, 'color': C['red']}),
            *[_flag_item(nivel, msg) for nivel, msg in flags_d]
        ], style={**CARD, 'borderColor': C['red']})

        sec_region = html.Div([
            _tabla_diag(aq.get('familias', _df_empty), 'Familia', 'Familias — Variación YTD'),
            _tabla_diag(aq.get('canales',  _df_empty), 'Canal',   'Canales — Variación YTD'),
        ], style=G2)

        _pivot  = aq.get('matriz',     _df_empty)
        _tot_r  = aq.get('tot_repre',  pd.Series(dtype=float))
        sec_matriz = html.Div([
            html.Div("Mapa de Tendencias — Representante × Familia", style=SEC),
            html.Div("Variación % vs año anterior. Orden: de mayor caída a mayor crecimiento — familias (columnas) y representantes (filas).",
                style={'color':C['muted'],'fontSize':'10px','marginBottom':'10px'}),
            html.Div(_tabla_matriz(_pivot, _tot_r), style={'overflowX':'auto'}),
        ], style=CARD)

        _cr = aq.get('canal_repre', _df_empty)
        sec_canales_rep = html.Div([
            html.Div("Fluctuación de Canales por Representante", style=SEC),
            html.Div("Canal | Cajas actuales | Diferencia vs año anterior | Var%. Ordenado de mayor caída a mayor crecimiento.",
                style={'color':C['muted'],'fontSize':'10px','marginBottom':'10px'}),
            _sec_canal_repre(_cr),
        ], style=CARD)

        _tend_el = _sec_tendencia(aq.get('tendencia', _df_empty))
        _children = [_pbt, sec_flags, sec_region, sec_matriz, sec_canales_rep]
        if _tend_el:
            _children.append(_tend_el)

        return html.Div(_children)

    elif tab == 'pendientes':
        if 'pend' not in DFS:
            return html.Div("Sin datos de pendientes.", style={'color': C['muted']})
        df = DFS['pend'].copy()
        df.columns = [c.strip() for c in df.columns]
        df['Pedidos Pendientes'] = pd.to_numeric(df['Pedidos Pendientes'], errors='coerce')
        df = df[df['Pedidos Pendientes'] > 0]
        if repre:
            df = df[df['Vendedor'] == repre]
        if flia:
            df = df[df['Familia Producto'] == flia]
        df = df.sort_values('Pedidos Pendientes', ascending=False)

        # Subtotal por vendedor para la tabla
        df_agg = df.groupby('Vendedor')['Pedidos Pendientes'].sum().reset_index()
        df_agg = df_agg.sort_values('Pedidos Pendientes', ascending=False)
        total_pend = df_agg['Pedidos Pendientes'].sum()

        return html.Div([
            _pbt,
            html.Div([dcc.Graph(figure=fig_pendientes(flia_sel=flia, repre_sel=repre), config={'displayModeBar':False})], style=CARD),
            html.Div([
                html.Div([
                    html.Div("Detalle por Vendedor", style=SEC),
                    html.Div(f"Total región: {int(total_pend):,} pedidos pendientes",
                             style={'color': C['red'], 'fontSize': '12px', 'marginBottom': '12px'}),
                    html.Table([
                        html.Thead(html.Tr([
                            html.Th("Vendedor",   style={'textAlign':'left','color':C['muted'],'fontSize':'9px','padding':'6px 10px'}),
                            html.Th("Pendientes", style={'textAlign':'right','color':C['muted'],'fontSize':'9px','padding':'6px 10px'}),
                            html.Th("% del Total", style={'textAlign':'right','color':C['muted'],'fontSize':'9px','padding':'6px 10px'}),
                        ])),
                        html.Tbody([
                            html.Tr([
                                html.Td(str(row['Vendedor']).strip()[:38],
                                        style={'padding':'5px 10px','fontSize':'11px','borderBottom':f"1px solid {C['border']}"}),
                                html.Td(f"{int(row['Pedidos Pendientes']):,}",
                                        style={'padding':'5px 10px','fontSize':'11px','textAlign':'right',
                                               'color':C['gold'],'borderBottom':f"1px solid {C['border']}"}),
                                html.Td(f"{row['Pedidos Pendientes']/total_pend*100:.0f}%",
                                        style={'padding':'5px 10px','fontSize':'11px','textAlign':'right',
                                               'color':C['muted'],'borderBottom':f"1px solid {C['border']}"}),
                            ]) for _, row in df_agg.iterrows()
                        ])
                    ], style={'width':'100%','borderCollapse':'collapse'})
                ]),
                html.Div([
                    html.Div("Detalle por Familia", style={**SEC, 'marginTop': '20px'}),
                    html.Table([
                        html.Thead(html.Tr([
                            html.Th("Familia",    style={'textAlign':'left','color':C['muted'],'fontSize':'9px','padding':'6px 10px'}),
                            html.Th("Vendedor",   style={'textAlign':'left','color':C['muted'],'fontSize':'9px','padding':'6px 10px'}),
                            html.Th("Pendientes", style={'textAlign':'right','color':C['muted'],'fontSize':'9px','padding':'6px 10px'}),
                        ])),
                        html.Tbody([
                            html.Tr([
                                html.Td(str(row.get('Familia Producto','')),
                                        style={'padding':'5px 10px','fontSize':'11px','borderBottom':f"1px solid {C['border']}"}),
                                html.Td(str(row.get('Vendedor','')).strip()[:38],
                                        style={'padding':'5px 10px','fontSize':'11px','borderBottom':f"1px solid {C['border']}"}),
                                html.Td(f"{int(row.get('Pedidos Pendientes',0)):,}",
                                        style={'padding':'5px 10px','fontSize':'11px','textAlign':'right',
                                               'color':C['gold'],'borderBottom':f"1px solid {C['border']}"}),
                            ]) for _, row in df.iterrows()
                        ])
                    ], style={'width':'100%','borderCollapse':'collapse'})
                ]),
            ], style=CARD),
        ])

    return html.Div()


@app.callback(
    Output('auth-store', 'data'),
    Output('login-error', 'children'),
    Input('btn-login', 'n_clicks'),
    Input('btn-logout', 'n_clicks'),
    State('login-user', 'value'),
    State('login-pass', 'value'),
    prevent_initial_call=True,
)
def cb_auth(n_login, n_logout, usuario, password):
    trigger = (callback_context.triggered or [{}])[0].get('prop_id', '').split('.')[0]
    if trigger == 'btn-logout':
        return None, ""
    if not usuario or not password:
        return no_update, "Completá usuario y contraseña."
    key = _norm(usuario)
    if key in USERS and USERS[key]['password'] == password.strip():
        return USERS[key], ""
    return no_update, "Usuario o contraseña incorrectos."


@app.callback(
    Output('login-page', 'style'),
    Output('dashboard-page', 'style'),
    Output('user-badge', 'children'),
    Output('btn-refresh', 'style'),
    Input('auth-store', 'data'),
)
def cb_toggle_pages(auth):
    _login_show = {'display': 'flex', 'justifyContent': 'center', 'alignItems': 'center',
                   'minHeight': '100vh', 'backgroundColor': C['bg']}
    _login_hide = {'display': 'none'}
    _dash_show  = {}
    _dash_hide  = {'display': 'none'}
    _btn_refresh_show = {
        'backgroundColor': 'transparent', 'color': C['gold'], 'border': f"1px solid {C['gold']}",
        'padding': '6px 14px', 'fontSize': '9px', 'letterSpacing': '1.5px',
        'textTransform': 'uppercase', 'cursor': 'pointer', 'borderRadius': '2px',
        'fontFamily': FONT, 'fontWeight': '600',
        'WebkitAppearance': 'none', 'appearance': 'none', 'marginRight': '8px',
    }
    _btn_refresh_hide = {'display': 'none'}
    if auth:
        role = auth.get('role')
        btn_st = _btn_refresh_show if role == 'admin' else _btn_refresh_hide
        name  = auth.get('display_name') or (auth.get('repre') or '').upper()
        title = auth.get('title', '')
        badge = html.Div([
            html.Div(name,  style={'color': C['gold'], 'fontSize': '11px',
                                   'letterSpacing': '1px', 'fontWeight': '600'}),
            html.Div(title, style={'color': C['muted'], 'fontSize': '9px',
                                   'letterSpacing': '1px'}) if title else None,
        ])
        return _login_hide, _dash_show, badge, btn_st
    return _login_show, _dash_hide, "", _btn_refresh_hide


@app.callback(
    Output('dd-repre', 'value'),
    Output('dd-repre', 'disabled'),
    Output('repre-wrap', 'style'),
    Input('auth-store', 'data'),
)
def cb_repre_lock(auth):
    if auth and auth.get('role') == 'vendedor':
        return auth.get('repre', ''), True, {'flex': 1, 'opacity': '0.55', 'pointerEvents': 'none'}
    return '', False, {'flex': 1}


@app.callback(
    Output('dd-flia',  'options'),
    Output('dd-repre', 'options'),
    Output('dd-canal', 'options'),
    Output('dd-meses', 'options'),
    Input('data-version', 'data'),
)
def cb_update_dropdown_options(_ver):
    flia_opts  = [{'label':'Todas','value':''}] + [{'label':f,'value':f} for f in FAMILIAS]
    repre_opts = [{'label':'Todos','value':''}] + [{'label':r,'value':r} for r in REPRESENTANTES]
    canal_opts = [{'label':'Todos','value':''}] + [{'label':c,'value':c} for c in CANALES]
    mes_opts   = [{'label':m,'value':m} for m in reversed(MC)]
    return flia_opts, repre_opts, canal_opts, mes_opts


@app.callback(
    Output('download-pdf', 'data'),
    Input('btn-pdf', 'n_clicks'),
    State('dd-repre', 'value'),
    prevent_initial_call=True,
)
def cb_pdf(n_clicks, repre):
    if not n_clicks or not repre or not PDF_AVAILABLE:
        return None
    pdf_bytes = generar_pdf_repre(repre)
    if not pdf_bytes:
        return None
    nombre = repre.replace(' ', '_').replace('/', '-')
    return dcc.send_bytes(pdf_bytes, filename=f"reporte_{nombre}.pdf")


@app.callback(
    Output('download-resumen', 'data'),
    Input('btn-print', 'n_clicks'),
    State('dd-flia',  'value'),
    State('dd-repre', 'value'),
    State('dd-canal', 'value'),
    prevent_initial_call=True,
)
def cb_resumen_pdf(n_clicks, flia, repre, canal):
    if not n_clicks or not PDF_AVAILABLE:
        return None
    pdf_bytes = generar_pdf_resumen(flia or None, repre or None, canal or None)
    if not pdf_bytes:
        return None
    filtro = "_".join(filter(None, [flia, repre, canal])).replace(' ','_') or "region"
    return dcc.send_bytes(pdf_bytes, filename=f"resumen_{filtro}_{datetime.now().strftime('%Y%m%d')}.pdf")


@app.callback(
    Output('download-tab-pdf', 'data'),
    Input('btn-tab-pdf', 'n_clicks'),
    State('tabs', 'value'),
    State('dd-flia', 'value'),
    State('dd-repre', 'value'),
    State('dd-canal', 'value'),
    State('auth-store', 'data'),
    prevent_initial_call=True,
)
def cb_tab_pdf(n_clicks, tab, flia, repre, canal, auth):
    if not n_clicks or not PDF_AVAILABLE:
        return None
    if auth and auth.get('role') == 'vendedor':
        repre = auth.get('repre', '')
    pdf_bytes = generar_pdf_tab(
        tab,
        flia_sel=flia or None,
        repre_sel=repre or None,
        canal_sel=canal or None,
    )
    if not pdf_bytes:
        return None
    TAB_LABELS = {
        'region':'region','repre':'representantes','clientes':'clientes',
        'canales':'canales','analisis':'analisis','pendientes':'pendientes',
    }
    filtro = "_".join(filter(None, [flia, repre, canal])).replace(' ','_') or "todos"
    nombre = f"{TAB_LABELS.get(tab, tab)}_{filtro}_{datetime.now().strftime('%Y%m%d')}.pdf"
    return dcc.send_bytes(pdf_bytes, filename=nombre)


@app.callback(
    Output('data-version', 'data', allow_duplicate=True),
    Output('timestamp', 'children', allow_duplicate=True),
    Input('btn-refresh', 'n_clicks'),
    State('data-version', 'data'),
    prevent_initial_call=True,
)
def cb_refresh_data(n, version):
    global DFS, MC, FAMILIAS, REPRESENTANTES, CANALES
    DFS = load_data()
    MC = month_cols(DFS['x flia'])
    FAMILIAS = sorted(DFS['x flia']['flia'].unique().tolist())
    REPRESENTANTES = sorted(DFS['x repre']['Vendedor'].unique().tolist())
    CANALES = sorted([str(x) for x in DFS['x flia x canal']['Canal'].dropna().unique().tolist()])
    ts = f"Recargado desde Drive: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    return (version or 0) + 1, ts

# ── Run ────────────────────────────────────────────────────────────────────────

@app.callback(
    Output('tabs', 'value'),
    Input('auth-store', 'data'),
    prevent_initial_call=True,
)
def cb_vendor_tab_redirect(auth):
    if auth and auth.get('role') == 'vendedor':
        return 'repre'
    return no_update


def open_browser():
    webbrowser.open(f"http://localhost:{PORT}")

if __name__ == '__main__':
    print("\n" + "="*55)
    print("  DASHBOARD VENTAS — CATENA ZAPATA")
    print(f"  Archivo: {EXCEL_PATH}")
    print(f"  URL: http://localhost:{PORT}")
    print(f"  PDF: {'disponible (reportlab)' if PDF_AVAILABLE else 'instalar: pip install reportlab'}")
    print("  Detener: Ctrl+C")
    print("="*55)
    print("\n  CREDENCIALES DE ACCESO")
    print(f"  {'USUARIO':<22} {'CLAVE':<8} {'REPRESENTANTE'}")
    print("  " + "-"*60)
    for _u, _data in USERS.items():
        _rep = _data.get('repre') or '(admin)'
        print(f"  {_u:<22} {_data['password']:<8} {_rep}")
    print("="*55 + "\n")
    threading.Timer(1.5, open_browser).start()
    app.run(debug=False, port=PORT, host='localhost')
