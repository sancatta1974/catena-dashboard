"""
Dashboard Comercial - Catena Zapata
Santi Cattaneo - Jefatura Nacional de Ventas
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

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                     Table, TableStyle, HRFlowable)
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors as rl_colors
    from reportlab.lib.units import cm
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

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
    service, f = _get_drive_service()
    if service and f:
        request = service.files().get_media(fileId=f['id'])
        buf = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        buf.seek(0)
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

try:
    DFS = load_data()
    DATA_OK = True
except Exception as _e:
    print(f"[WARNING] No se pudo cargar el archivo al iniciar: {_e}")
    DFS = {}
    DATA_OK = False
MC  = month_cols(DFS.get('x flia', pd.DataFrame()))
FAMILIAS       = sorted(DFS['x flia']['flia'].unique().tolist()) if DATA_OK else []
REPRESENTANTES = sorted(DFS['x repre']['Vendedor'].unique().tolist()) if DATA_OK else []
CANALES        = sorted(DFS['x flia x canal']['Canal'].unique().tolist()) if DATA_OK else []

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

_SKIP_REPS = {'directos casa interior'}

_used_pins = {'piso3'}
_used_keys = {'jefe'}
USERS = {'jefe': {'password': 'piso3', 'role': 'admin', 'repre': None}}
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
    USERS[_key] = {'password': _pin, 'role': 'vendedor', 'repre': _r}

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

def _var(v):
    """Formatea variación % con 1 decimal; muestra '—' si es NaN/inf."""
    if v is None or (isinstance(v, float) and (np.isnan(v) or np.isinf(v))):
        return '—'
    return f"{v:+.0f}%"

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
        m = act.merge(ant, on='flia', suffixes=('_a','_b'))
        m['var']  = (m['Total_a'] - m['Total_b']) / m['Total_b'].replace(0, np.nan) * 100
        m['part'] = m['Total_a'] / m['Total_a'].sum() * 100
        if flia_sel:
            m = m[m['flia'] == flia_sel]
        m = m.dropna(subset=['Total_a']).sort_values('Total_a', ascending=True)

        col_var = [C['green'] if (pd.notna(v) and v >= 0) else C['red'] for v in m['var']]
        col_vol = col_var

        height = max(260, len(m) * 24 + 70)

        fig = make_subplots(
            rows=1, cols=2,
            column_widths=[0.38, 0.62],
            subplot_titles=['Variación % vs Año Anterior', 'Participación % y Cajas Año Actual'],
            horizontal_spacing=0.12,
        )

        # Panel izq: var%
        fig.add_trace(go.Bar(
            y=m['flia'], x=m['var'].fillna(0).round(0), orientation='h',
            marker_color=col_var,
            text=[_var(v) for v in m['var']],
            textposition='outside',
            textfont=dict(size=14, color=C['text']),
            hovertemplate='<b>%{y}</b><br>Var: %{x:+.0f}%<extra></extra>',
        ), row=1, col=1)

        # Panel der: barras horizontales ordenadas, color por crecimiento/caída
        fig.add_trace(go.Bar(
            y=m['flia'], x=m['Total_a'], orientation='h',
            marker_color=col_vol,
            text=[f"{int(v):,} caj" for v in m['Total_a']],
            textposition='inside',
            insidetextanchor='middle',
            textfont=dict(size=10, color='#FFFFFF'),
            customdata=np.column_stack([m['part'].round(0), m['Total_b'], m['var'].fillna(0).round(0)]),
            hovertemplate='<b>%{y}</b><br>Actual: %{x:,.0f} caj  (%{customdata[0]:.0f}%)<br>Anterior: %{customdata[1]:,.0f}<br>Var: %{customdata[2]:+.0f}%<extra></extra>',
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
        fig.update_xaxes(ticksuffix='%',    row=1, col=1)
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
            rv_tot = av.groupby('Vendedor')['Total'].sum().reset_index().merge(
                bv.groupby('Vendedor')['Total'].sum().reset_index(), on='Vendedor', suffixes=('_a','_b'))
            rv_tot['var'] = (rv_tot['Total_a'] - rv_tot['Total_b']) / rv_tot['Total_b'].replace(0, np.nan) * 100
            rv_tot = rv_tot.sort_values('var')
            rep_df = rv_tot.rename(columns={'var': 'Total_var', 'Total_a': 'Total_vol'}).reset_index(drop=True)
        else:
            rv = get_ind(df_repre, 'Var %', ['Vendedor','flia'])
            av = get_ind(df_repre, 'Año Actual Cajas', ['Vendedor','flia'])
            rv_tot = rv.groupby('Vendedor')['Total'].mean().mul(100).reset_index()
            av_tot = av.groupby('Vendedor')['Total'].sum().reset_index()
            rep_df = rv_tot.merge(av_tot, on='Vendedor', suffixes=('_var','_vol'))
            rep_df = rep_df.sort_values('Total_var').reset_index(drop=True)
        col_rep = [C['green'] if (pd.notna(v) and v >= 0) else C['red'] for v in rep_df['Total_var']]

        # ── Familias ──
        df_flia = DFS['x flia']
        if flia_sel:
            df_flia = df_flia[df_flia['flia'] == flia_sel]
        if meses_sel:
            fa_act = get_ind(df_flia, 'Año Actual Cajas', ['flia'], meses_sel)
            fa_ant = get_ind(df_flia, 'Año Anterior Cajas', ['flia'], meses_sel)
            fam_m = fa_act.merge(fa_ant, on='flia', suffixes=('_a','_b'))
            fam_m['Total_var'] = (fam_m['Total_a'] - fam_m['Total_b']) / fam_m['Total_b'].replace(0, np.nan) * 100
            fam_m = fam_m.sort_values('Total_var').reset_index(drop=True)
            fam_df = fam_m.rename(columns={'Total_a': 'Total_vol'})
        else:
            fv = get_ind(df_flia, 'Var %', ['flia'])
            fa = get_ind(df_flia, 'Año Actual Cajas', ['flia'])
            fam_df = fv.merge(fa[['flia','Total']], on='flia', suffixes=('_var','_vol'))
            fam_df['Total_var'] = fam_df['Total_var'] * 100
            fam_df = fam_df.sort_values('Total_var').reset_index(drop=True)
        col_fam = [C['green'] if (pd.notna(v) and v >= 0) else C['red'] for v in fam_df['Total_var']]

        fig = make_subplots(
            rows=1, cols=2,
            subplot_titles=['Representantes — Var % vs Año Anterior',
                            'Familias — Var % vs Año Anterior'],
            column_widths=[0.55, 0.45],
        )

        # Panel izq — representantes
        fig.add_trace(go.Bar(
            y=rep_df['Vendedor'].str[:22],
            x=rep_df['Total_var'].fillna(0).round(0),
            orientation='h',
            marker_color=col_rep,
            textposition='outside',
            text=[_var(v) for v in rep_df['Total_var']],
            textfont=dict(size=14, color=C['text']),
            customdata=rep_df['Total_vol'],
            hovertemplate='<b>%{y}</b><br>Var: %{x:+.0f}%<br>Cajas: %{customdata:,.0f}<extra></extra>',
        ), row=1, col=1)

        # Panel der — familias
        fig.add_trace(go.Bar(
            y=fam_df['flia'].str[:18],
            x=fam_df['Total_var'].fillna(0).round(0),
            orientation='h',
            marker_color=col_fam,
            textposition='outside',
            text=[_var(v) for v in fam_df['Total_var']],
            textfont=dict(size=14, color=C['text']),
            customdata=fam_df['Total_vol'],
            hovertemplate='<b>%{y}</b><br>Var: %{x:+.0f}%<br>Cajas: %{customdata:,.0f}<extra></extra>',
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
        fig.update_xaxes(ticksuffix='%', tickfont=dict(size=10), gridcolor=C['border'])
        fig.update_yaxes(tickfont=dict(size=10))
        return fig
    except Exception as e:
        return go.Figure().update_layout(**PL, title=f'Error ranking: {e}')

def fig_repre_ranking(flia_sel, canal_sel, repre_sel=None, meses_sel=None):
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
        m = a.merge(b, on='Vendedor', suffixes=('_a','_b'))
        m['var']  = (m['Total_a'] - m['Total_b']) / m['Total_b'].replace(0, np.nan) * 100
        m['part'] = m['Total_a'] / m['Total_a'].sum() * 100
        m = m.sort_values('Total_a', ascending=False)
        nombres = m['Vendedor'].str[:18]

        # Colores: resaltar el representante seleccionado
        if repre_sel:
            col_vol = [C['gold'] if v == repre_sel else 'rgba(201,168,76,0.25)' for v in m['Vendedor']]
            col_var = []
            for i, (v, pct) in enumerate(zip(m['Vendedor'], m['var'])):
                if v == repre_sel:
                    col_var.append(C['green'] if pct >= 0 else C['red'])
                else:
                    col_var.append('rgba(200,200,200,0.15)')
        else:
            col_vol = [C['gold']] * len(m)
            col_var = [C['green'] if (pd.notna(v) and v >= 0) else C['red'] for v in m['var']]

        fig = make_subplots(rows=1, cols=2,
                            subplot_titles=['Var % vs Año Anterior', 'Cajas Año Actual (% participación)'])
        fig.add_trace(go.Bar(x=nombres, y=m['var'].fillna(0).round(0), marker_color=col_var,
                             text=[_var(v) for v in m['var']],
                             textposition='outside',
                             textfont=dict(size=14, color=C['text']),
                             hovertemplate='%{x}<br>%{y:+.0f}%<extra></extra>'), row=1, col=1)
        fig.add_trace(go.Bar(x=nombres, y=m['Total_a'], marker_color=col_vol,
                             text=[f"{int(v):,}" for v in m['Total_a']],
                             textposition='inside', insidetextanchor='middle',
                             textfont=dict(size=10, color='#FFFFFF'),
                             hovertemplate='%{x}<br>%{y:,.0f} cajas<extra></extra>'), row=1, col=2)
        # % participación afuera de la barra de cajas
        fig.add_trace(go.Scatter(
            x=nombres, y=m['Total_a'],
            text=[f"{p:.0f}%" for p in m['part']],
            mode='text', textposition='top center',
            textfont=dict(size=15, color=C['text']),
            showlegend=False, hoverinfo='skip',
        ), row=1, col=2)
        fig.add_vline(x=0, line_width=1, line_color=C['muted'], line_dash='dot', col=1)
        title = f'Representantes — Ranking | Filtro: {repre_sel}' if repre_sel else 'Representantes — Ranking y Variación'
        _pl = {k:v for k,v in PL.items() if k != 'margin'}
        fig.update_layout(**_pl, title=title, height=340, showlegend=False,
                          margin=dict(l=30, r=20, t=46, b=70))
        fig.update_xaxes(tickangle=-40, tickfont=dict(size=10))
        fig.update_yaxes(tickfont=dict(size=10))
        return fig
    except Exception as e:
        return go.Figure().update_layout(**PL, title=f'Error: {e}')

def fig_canal_mix(flia_sel, repre_sel, canal_sel=None, meses_sel=None):
    """Barras horizontales de participación % por canal — reemplaza el pie chart."""
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
        m = m[m['Total_a'] > 0].sort_values('Total_a', ascending=True)
        total = m['Total_a'].sum()
        m['pct']  = m['Total_a'] / total * 100
        m['var']  = (m['Total_a'] - m['Total_b']) / m['Total_b'].replace(0, np.nan) * 100
        palette = px.colors.qualitative.Set2
        colors  = [palette[i % len(palette)] for i in range(len(m))]
        col_var = [C['green'] if (pd.notna(v) and v >= 0) else C['red'] for v in m['var']]

        fig = make_subplots(
            rows=1, cols=2,
            column_widths=[0.35, 0.65],
            subplot_titles=['Var % vs Año Anterior', 'Participación % Año Actual'],
        )
        # Panel izq: var%
        fig.add_trace(go.Bar(
            y=m['Canal'], x=m['var'].fillna(0).round(0), orientation='h',
            marker_color=col_var,
            text=[_var(v) for v in m['var']],
            textposition='outside',
            textfont=dict(size=14, color=C['text']),
            hovertemplate='<b>%{y}</b><br>%{x:+.0f}%<extra></extra>',
        ), row=1, col=1)
        # Panel der: barras horizontales de % con etiqueta de volumen
        fig.add_trace(go.Bar(
            y=m['Canal'], x=m['pct'], orientation='h',
            marker_color=colors,
            text=[f"{p:.0f}%  ({int(v):,} caj)" for p, v in zip(m['pct'], m['Total_a'])],
            textposition='inside', insidetextanchor='middle',
            textfont=dict(size=10, color='#FFFFFF'),
            hovertemplate='<b>%{y}</b><br>%{x:.0f}%  —  %{text}<extra></extra>',
        ), row=1, col=2)

        subtitulo = repre_sel or 'Región'
        pl_m = {k: v for k, v in PL.items() if k not in ('xaxis','yaxis','margin')}
        fig.update_layout(
            **pl_m,
            title=f'Mix por Canal — {subtitulo}',
            height=max(220, len(m) * 26 + 70),
            margin=dict(l=10, r=80, t=46, b=24),
            showlegend=False,
        )
        fig.update_xaxes(tickfont=dict(size=10), gridcolor=C['border'])
        fig.update_yaxes(tickfont=dict(size=10))
        fig.update_xaxes(ticksuffix='%', row=1, col=1)
        fig.update_xaxes(ticksuffix='%', row=1, col=2)
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
    activos  = full[(full['act'] > 0) & (full['ant'] > 0)].copy()
    con_crecimiento = full[full['dif'] > 0].copy()
    con_caida       = full[full['dif'] < 0].copy()
    return {
        'full': full, 'nuevos': con_crecimiento, 'perdidos': con_caida, 'activos': activos,
        'top_sube': full[full['dif'] > 0].nlargest(20, 'dif').sort_values('dif'),
        'top_baja': full[full['dif'] < 0].nsmallest(20, 'dif').sort_values('dif', ascending=False),
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
        crec  = datos['nuevos'].nlargest(15, 'act').sort_values('act')
        caida = datos['perdidos'].nlargest(15, 'act').sort_values('act')
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
                y=y_crec, x=crec['act'], orientation='h',
                marker_color='rgba(39,174,96,0.85)',
                text=[f"{int(v):,} caj" for v in crec['act']],
                textposition='inside', insidetextanchor='middle',
                textfont=dict(size=10, color='#FFFFFF'),
                customdata=np.column_stack([crec['Vendedor'], crec['dif'], crec['var'].fillna(0).round(0)]),
                hovertemplate='<b>%{y}</b><br>Rep: %{customdata[0]}<br>Cajas: %{x:,.0f}<br>Δ: +%{customdata[1]:,.0f}  %{customdata[2]:+.0f}%<extra></extra>',
            ), row=1, col=1)
        if not caida.empty:
            y_caida = caida['Cliente'].str[:28]
            fig.add_trace(go.Bar(
                y=y_caida, x=caida['act'], orientation='h',
                marker_color='rgba(192,57,43,0.85)',
                text=[f"{int(v):,} caj" for v in caida['act']],
                textposition='inside', insidetextanchor='middle',
                textfont=dict(size=10, color='#FFFFFF'),
                customdata=np.column_stack([caida['Vendedor'], caida['dif'], caida['var'].fillna(0).round(0)]),
                hovertemplate='<b>%{y}</b><br>Rep: %{customdata[0]}<br>Cajas: %{x:,.0f}<br>Δ: %{customdata[1]:,.0f}  %{customdata[2]:+.0f}%<extra></extra>',
            ), row=1, col=2)

        pl_c = {k: v for k, v in PL.items() if k not in ('margin',)}
        fig.update_layout(**pl_c, title='Clientes — Con Crecimiento vs Con Caídas (top 15)',
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
            act = get_ind(DFS['x flia'], 'Año Actual Cajas', ['flia'])
            u, p = MC[-1], MC[-2]
            tu = pd.to_numeric(act[u], errors='coerce').sum()
            tp = pd.to_numeric(act[p], errors='coerce').sum()
            if tp > 0 and (tu - tp) / tp * 100 < -10:
                vt = (tu - tp) / tp * 100
                flags.append(('CRITICO', f"Desaceleración fuerte: {p}→{u} = {vt:+.0f}% en total región"))

        # Clientes perdidos (fueron a 0)
        if 'x cliente' in DFS:
            cli = DFS['x cliente']
            act_c = get_ind(cli, 'Año Actual Cajas', ['Vendedor','Cliente','flia'])
            ant_c = get_ind(cli, 'Año Anterior Cajas', ['Vendedor','Cliente','flia'])
            if repre_sel:
                act_c = act_c[act_c['Vendedor'] == repre_sel]
                ant_c = ant_c[ant_c['Vendedor'] == repre_sel]
            merged = ant_c[ant_c['Total'] > 0].merge(act_c, on=['Vendedor','Cliente','flia'], suffixes=('_b','_a'))
            perdidos = merged[merged['Total_a'] == 0]
            if not perdidos.empty:
                n = len(perdidos)
                flags.append(('ALERTA', f"{n} cliente{'s' if n>1 else ''} con 0 cajas este año (activos el año anterior)"))

    except Exception as e:
        flags.append(('INFO', f'No se pudieron calcular red flags: {e}'))

    if not flags:
        flags.append(('OK', 'Sin alertas críticas detectadas.'))
    return flags


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

    # Número de representantes activos según filtros
    n_repre = 1 if repre_sel else (
        DFS['x repre x canal'][DFS['x repre x canal']['Canal']==canal_sel]['Vendedor'].nunique()
        if canal_sel else DFS['x repre']['Vendedor'].nunique()
    )

    filtro_label = " | ".join(filter(None, [
        flia_sel, repre_sel, canal_sel
    ])) or "Región completa"

    items = [
        ('CAJAS AÑO ACTUAL',   f"{int(tot_a):,}",     C['gold']),
        ('CAJAS AÑO ANTERIOR', f"{int(tot_b):,}",     C['muted']),
        ('VARIACION TOTAL',    f"{sign}{var_t:.0f}%",  cv),
        ('REPRESENTANTES',     str(n_repre),            C['gold']),
        ('FAMILIAS',           str(len(FAMILIAS)),      C['gold']),
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


# ── PDF por representante ──────────────────────────────────────────────────────

def generar_pdf_repre(repre_sel):
    if not PDF_AVAILABLE:
        return None

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=2*cm, rightMargin=2*cm,
                            topMargin=2*cm, bottomMargin=2*cm)

    styles = getSampleStyleSheet()
    titulo   = ParagraphStyle('titulo',   parent=styles['Heading1'],
                               textColor=rl_colors.HexColor('#8B6914'), fontSize=18, spaceAfter=6)
    subtitulo= ParagraphStyle('subtitulo',parent=styles['Heading2'],
                               textColor=rl_colors.HexColor('#444444'), fontSize=10, spaceAfter=4)
    cuerpo   = ParagraphStyle('cuerpo',   parent=styles['Normal'],
                               textColor=rl_colors.HexColor('#1A1A1A'), fontSize=9, spaceAfter=3)
    alerta_s = ParagraphStyle('alerta',   parent=styles['Normal'],
                               textColor=rl_colors.HexColor('#C0392B'), fontSize=9, spaceAfter=3)
    ok_s     = ParagraphStyle('ok',       parent=styles['Normal'],
                               textColor=rl_colors.HexColor('#1E7A40'), fontSize=9, spaceAfter=3)

    story = []

    # Encabezado
    story.append(Paragraph("CATENA ZAPATA", titulo))
    story.append(Paragraph(f"Reporte de Representante — {repre_sel}", subtitulo))
    story.append(Paragraph(f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}", cuerpo))
    story.append(HRFlowable(width="100%", thickness=1, color=rl_colors.HexColor('#C9A84C')))
    story.append(Spacer(1, 0.4*cm))

    # KPIs del representante
    try:
        df_r = DFS['x repre']
        df_r = df_r[df_r['Vendedor'] == repre_sel]
        act = get_ind(df_r, 'Año Actual Cajas', ['Vendedor','flia'])
        ant = get_ind(df_r, 'Año Anterior Cajas', ['Vendedor','flia'])
        tot_a = act['Total'].sum()
        tot_b = ant['Total'].sum()
        var_t = (tot_a - tot_b) / tot_b * 100 if tot_b else 0

        # Ranking
        all_act = get_ind(DFS['x repre'], 'Año Actual Cajas', ['Vendedor','flia'])
        ranking_df = all_act.groupby('Vendedor')['Total'].sum().sort_values(ascending=False).reset_index()
        ranking_df['rank'] = range(1, len(ranking_df)+1)
        rank = ranking_df[ranking_df['Vendedor']==repre_sel]['rank'].iloc[0] if repre_sel in ranking_df['Vendedor'].values else '—'

        sign = '+' if var_t >= 0 else ''
        story.append(Paragraph("RESUMEN EJECUTIVO", subtitulo))
        kpi_data = [
            ['Métrica', 'Valor'],
            ['Cajas Año Actual', f"{int(tot_a):,}"],
            ['Cajas Año Anterior', f"{int(tot_b):,}"],
            ['Variación %', f"{sign}{var_t:.0f}%"],
            ['Ranking Nacional', f"#{rank} de {len(REPRESENTANTES)}"],
        ]
        t = Table(kpi_data, colWidths=[8*cm, 6*cm])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), rl_colors.HexColor('#C9A84C')),
            ('TEXTCOLOR',  (0,0), (-1,0), rl_colors.black),
            ('FONTSIZE',   (0,0), (-1,-1), 9),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [rl_colors.white, rl_colors.HexColor('#F5F5F5')]),
            ('TEXTCOLOR',  (0,1), (-1,-1), rl_colors.HexColor('#1A1A1A')),
            ('GRID',       (0,0), (-1,-1), 0.5, rl_colors.HexColor('#CCCCCC')),
            ('PADDING',    (0,0), (-1,-1), 6),
            ('ALIGN',      (1,0), (1,-1), 'RIGHT'),
        ]))
        story.append(t)
        story.append(Spacer(1, 0.5*cm))
    except Exception as e:
        story.append(Paragraph(f"Error generando KPIs: {e}", alerta_s))

    # Evolución mensual
    try:
        story.append(Paragraph("EVOLUCIÓN MENSUAL", subtitulo))
        df_r2 = DFS['x repre'][DFS['x repre']['Vendedor'] == repre_sel]
        act2  = get_ind(df_r2, 'Año Actual Cajas', ['Vendedor','flia'])
        # Agrupar todos los meses
        mes_data = [['Mes', 'Cajas Año Actual']]
        for m in MC:
            if m in act2.columns:
                val = pd.to_numeric(act2[m], errors='coerce').sum()
                mes_data.append([m, f"{int(val):,}" if pd.notna(val) else '—'])
        t2 = Table(mes_data, colWidths=[4*cm, 6*cm])
        t2.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), rl_colors.HexColor('#C9A84C')),
            ('TEXTCOLOR',  (0,0), (-1,0), rl_colors.black),
            ('FONTSIZE',   (0,0), (-1,-1), 8),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [rl_colors.white, rl_colors.HexColor('#F5F5F5')]),
            ('TEXTCOLOR',  (0,1), (-1,-1), rl_colors.HexColor('#1A1A1A')),
            ('GRID',       (0,0), (-1,-1), 0.5, rl_colors.HexColor('#CCCCCC')),
            ('PADDING',    (0,0), (-1,-1), 5),
            ('ALIGN',      (1,0), (1,-1), 'RIGHT'),
        ]))
        story.append(t2)
        story.append(Spacer(1, 0.5*cm))
    except Exception as e:
        story.append(Paragraph(f"Error evolución: {e}", alerta_s))

    # Red flags del representante
    try:
        story.append(Paragraph("ALERTAS DEL REPRESENTANTE", subtitulo))
        rv = get_ind(DFS['x repre'], 'Var %', ['Vendedor','flia'])
        rv_r = rv[rv['Vendedor'] == repre_sel].copy()
        rv_r['pct'] = rv_r['Total'] * 100
        rv_r = rv_r.dropna(subset=['pct'])

        if rv_r.empty:
            story.append(Paragraph("Sin datos de variación disponibles.", cuerpo))
        else:
            for _, row in rv_r.sort_values('pct').iterrows():
                sign2 = '+' if row['pct'] >= 0 else ''
                style_use = ok_s if row['pct'] >= 0 else alerta_s
                story.append(Paragraph(f"• {row['flia']}: {sign2}{row['pct']:.0f}%", style_use))
        story.append(Spacer(1, 0.5*cm))
    except Exception as e:
        story.append(Paragraph(f"Error alertas: {e}", alerta_s))

    # Top/bottom clientes
    try:
        story.append(Paragraph("TOP CLIENTES — VARIACIÓN", subtitulo))
        if 'x cliente' in DFS:
            cli = DFS['x cliente']
            cli = cli[cli['Vendedor'] == repre_sel]
            act_c = get_ind(cli, 'Año Actual Cajas', ['Vendedor','Cliente','flia'])
            ant_c = get_ind(cli, 'Año Anterior Cajas', ['Vendedor','Cliente','flia'])
            m = act_c.merge(ant_c, on=['Vendedor','Cliente','flia'], suffixes=('_a','_b'))
            m = m[m['Total_b'] > 0]
            m['dif'] = m['Total_a'] - m['Total_b']
            m['var_pct'] = (m['dif'] / m['Total_b']) * 100

            cli_data = [['Cliente', 'Familia', 'Cajas Act', 'Var%']]
            for _, row in m.nlargest(5, 'dif').iterrows():
                sign3 = '+' if row['var_pct'] >= 0 else ''
                cli_data.append([str(row['Cliente'])[:30], str(row['flia'])[:15],
                                  f"{int(row['Total_a']):,}", f"{sign3}{row['var_pct']:.0f}%"])
            for _, row in m.nsmallest(5, 'dif').iterrows():
                sign3 = '+' if row['var_pct'] >= 0 else ''
                cli_data.append([str(row['Cliente'])[:30], str(row['flia'])[:15],
                                  f"{int(row['Total_a']):,}", f"{sign3}{row['var_pct']:.0f}%"])

            t3 = Table(cli_data, colWidths=[7*cm, 3.5*cm, 3*cm, 2.5*cm])
            t3.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), rl_colors.HexColor('#C9A84C')),
                ('TEXTCOLOR',  (0,0), (-1,0), rl_colors.black),
                ('FONTSIZE',   (0,0), (-1,-1), 7.5),
                ('ROWBACKGROUNDS', (0,1), (-1,-1), [rl_colors.white, rl_colors.HexColor('#F5F5F5')]),
                ('TEXTCOLOR',  (0,1), (-1,-1), rl_colors.HexColor('#1A1A1A')),
                ('GRID',       (0,0), (-1,-1), 0.5, rl_colors.HexColor('#CCCCCC')),
                ('PADDING',    (0,0), (-1,-1), 4),
                ('ALIGN',      (2,0), (3,-1), 'RIGHT'),
            ]))
            story.append(t3)
    except Exception as e:
        story.append(Paragraph(f"Error clientes: {e}", alerta_s))

    doc.build(story)
    buf.seek(0)
    return buf.read()


def generar_pdf_resumen(flia_sel=None, repre_sel=None, canal_sel=None):
    """PDF ejecutivo A4 — resumen de ventas de la selección actual."""
    if not PDF_AVAILABLE:
        return None

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=1.8*cm, rightMargin=1.8*cm,
                            topMargin=1.8*cm, bottomMargin=1.8*cm)

    W, _ = A4

    def make_styles():
        s = getSampleStyleSheet()
        gold  = rl_colors.HexColor('#8B6914')
        red   = rl_colors.HexColor('#C0392B')
        green = rl_colors.HexColor('#1E7A40')
        dark  = rl_colors.HexColor('#1A1A1A')
        light = rl_colors.HexColor('#1A1A1A')
        muted = rl_colors.HexColor('#555555')
        return {
            'h1':    ParagraphStyle('h1',    fontSize=20, textColor=gold,   spaceAfter=2,  fontName='Helvetica-Bold'),
            'h2':    ParagraphStyle('h2',    fontSize=11, textColor=muted,  spaceAfter=8,  fontName='Helvetica'),
            'sec':   ParagraphStyle('sec',   fontSize=8,  textColor=gold,   spaceAfter=4,  fontName='Helvetica-Bold',
                                    textTransform='uppercase', spaceBefore=10),
            'body':  ParagraphStyle('body',  fontSize=8,  textColor=dark,   spaceAfter=2,  fontName='Helvetica',
                                    leading=12),
            'alert': ParagraphStyle('alert', fontSize=8,  textColor=red,    spaceAfter=2,  fontName='Helvetica'),
            'ok':    ParagraphStyle('ok',    fontSize=8,  textColor=green,  spaceAfter=2,  fontName='Helvetica'),
            'dark':  dark, 'light': light, 'gold': gold, 'red': red, 'green': green, 'muted': muted,
        }

    def tbl(data, col_widths, header_row=True):
        t = Table(data, colWidths=col_widths, repeatRows=1 if header_row else 0)
        style_cmds = [
            ('FONTSIZE',  (0,0), (-1,-1), 8),
            ('FONTNAME',  (0,0), (-1,-1), 'Helvetica'),
            ('GRID',      (0,0), (-1,-1), 0.4, rl_colors.HexColor('#CCCCCC')),
            ('PADDING',   (0,0), (-1,-1), 5),
            ('VALIGN',    (0,0), (-1,-1), 'MIDDLE'),
            ('ROWBACKGROUNDS', (0, 1 if header_row else 0), (-1,-1),
             [rl_colors.white, rl_colors.HexColor('#F5F5F5')]),
            ('TEXTCOLOR', (0, 1 if header_row else 0), (-1,-1), rl_colors.HexColor('#1A1A1A')),
        ]
        if header_row:
            style_cmds += [
                ('BACKGROUND', (0,0), (-1,0), rl_colors.HexColor('#C9A84C')),
                ('TEXTCOLOR',  (0,0), (-1,0), rl_colors.black),
                ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
            ]
        t.setStyle(TableStyle(style_cmds))
        return t

    st = make_styles()
    story = []

    # ── Encabezado ──
    filtro_txt = " | ".join(filter(None, [flia_sel, repre_sel, canal_sel])) or "Región completa"
    story.append(Paragraph("CATENA ZAPATA", st['h1']))
    story.append(Paragraph("Reporte de Ventas — Jefatura Nacional de Ventas", st['h2']))
    story.append(Paragraph(f"Filtro activo: {filtro_txt}   •   {datetime.now().strftime('%d/%m/%Y %H:%M')}", st['h2']))
    story.append(HRFlowable(width="100%", thickness=1.5, color=st['gold'], spaceAfter=8))

    # ── KPIs ──
    story.append(Paragraph("RESUMEN EJECUTIVO", st['sec']))
    try:
        kpi_vals = {}
        if canal_sel and repre_sel:
            df_k = DFS['x repre x canal']
            df_k = df_k[df_k['Canal'] == canal_sel]
            df_k = df_k[df_k['Vendedor'] == repre_sel]
            act_k = get_ind(df_k, 'Año Actual Cajas', ['Vendedor','flia'])
            ant_k = get_ind(df_k, 'Año Anterior Cajas', ['Vendedor','flia'])
        elif canal_sel:
            df_k = DFS['x flia x canal'][DFS['x flia x canal']['Canal'] == canal_sel]
            act_k = get_ind(df_k, 'Año Actual Cajas', ['flia','Canal'])
            ant_k = get_ind(df_k, 'Año Anterior Cajas', ['flia','Canal'])
        elif repre_sel:
            df_k = DFS['x repre'][DFS['x repre']['Vendedor'] == repre_sel]
            act_k = get_ind(df_k, 'Año Actual Cajas', ['Vendedor','flia'])
            ant_k = get_ind(df_k, 'Año Anterior Cajas', ['Vendedor','flia'])
        else:
            act_k = get_ind(DFS['x flia'], 'Año Actual Cajas', ['flia'])
            ant_k = get_ind(DFS['x flia'], 'Año Anterior Cajas', ['flia'])
        if flia_sel:
            act_k = act_k[act_k['flia'] == flia_sel]
            ant_k = ant_k[ant_k['flia'] == flia_sel]
        tot_a = act_k['Total'].sum(); tot_b = ant_k['Total'].sum()
        var_t = (tot_a - tot_b) / tot_b * 100 if tot_b else 0
        sign  = '+' if var_t >= 0 else ''

        kpi_data = [
            ['Métrica', 'Año Actual', 'Año Anterior', 'Variación'],
            ['Cajas Totales', f"{int(tot_a):,}", f"{int(tot_b):,}", f"{sign}{var_t:.0f}%"],
        ]
        story.append(tbl(kpi_data, [5*cm, 3.5*cm, 3.5*cm, 3*cm]))
        story.append(Spacer(1, 0.3*cm))
    except Exception as e:
        story.append(Paragraph(f"Error KPIs: {e}", st['alert']))

    # ── Top familias ──
    story.append(Paragraph("TOP FAMILIAS POR VOLUMEN", st['sec']))
    try:
        if repre_sel:
            df_fam = DFS['x repre'][DFS['x repre']['Vendedor'] == repre_sel]
            act_f = get_ind(df_fam, 'Año Actual Cajas', ['Vendedor','flia'])
            ant_f = get_ind(df_fam, 'Año Anterior Cajas', ['Vendedor','flia'])
            act_f = act_f.groupby('flia')['Total'].sum().reset_index().rename(columns={'Total':'Total_a'})
            ant_f = ant_f.groupby('flia')['Total'].sum().reset_index().rename(columns={'Total':'Total_b'})
            mf = act_f.merge(ant_f, on='flia', how='outer').fillna(0)
        else:
            act_f = get_ind(DFS['x flia'], 'Año Actual Cajas', ['flia'])
            ant_f = get_ind(DFS['x flia'], 'Año Anterior Cajas', ['flia'])
            mf = act_f.merge(ant_f, on='flia', suffixes=('_a','_b'))
        if flia_sel:
            mf = mf[mf['flia'] == flia_sel]
        mf['var'] = (mf['Total_a'] - mf['Total_b']) / mf['Total_b'].replace(0, np.nan) * 100
        mf['part'] = mf['Total_a'] / mf['Total_a'].sum() * 100
        mf = mf.sort_values('Total_a', ascending=False).head(10)

        fam_hdr = [['Familia', 'Cajas Actual', 'Cajas Anterior', 'Var %', 'Part %']]
        for _, r in mf.iterrows():
            s = '+' if pd.notna(r['var']) and r['var'] >= 0 else ''
            vstr = f"{s}{r['var']:.0f}%" if pd.notna(r['var']) else '—'
            fam_hdr.append([r['flia'], f"{int(r['Total_a']):,}", f"{int(r['Total_b']):,}",
                             vstr, f"{r['part']:.0f}%"])
        story.append(tbl(fam_hdr, [5*cm, 3*cm, 3*cm, 2.2*cm, 2.8*cm]))
        story.append(Spacer(1, 0.3*cm))
    except Exception as e:
        story.append(Paragraph(f"Error familias: {e}", st['alert']))

    # ── Top representantes ──
    story.append(Paragraph("REPRESENTANTES — RANKING Y VARIACIÓN", st['sec']))
    try:
        act_r = get_ind(DFS['x repre'], 'Año Actual Cajas', ['Vendedor','flia'])
        ant_r = get_ind(DFS['x repre'], 'Año Anterior Cajas', ['Vendedor','flia'])
        if flia_sel:
            act_r = act_r[act_r['flia'] == flia_sel]
            ant_r = ant_r[ant_r['flia'] == flia_sel]
        if repre_sel:
            act_r = act_r[act_r['Vendedor'] == repre_sel]
            ant_r = ant_r[ant_r['Vendedor'] == repre_sel]
        ar = act_r.groupby('Vendedor')['Total'].sum().reset_index()
        br = ant_r.groupby('Vendedor')['Total'].sum().reset_index()
        mr = ar.merge(br, on='Vendedor', suffixes=('_a','_b'))
        mr['var']  = (mr['Total_a'] - mr['Total_b']) / mr['Total_b'].replace(0, np.nan) * 100
        mr['part'] = mr['Total_a'] / mr['Total_a'].sum() * 100
        mr = mr.sort_values('Total_a', ascending=False)

        rep_hdr = [['#', 'Representante', 'Cajas Actual', 'Var %', 'Participación']]
        for i, (_, r) in enumerate(mr.iterrows(), 1):
            s = '+' if r['var'] >= 0 else ''
            rep_hdr.append([str(i), r['Vendedor'][:30], f"{int(r['Total_a']):,}",
                             f"{s}{r['var']:.0f}%", f"{r['part']:.0f}%"])
        story.append(tbl(rep_hdr, [0.8*cm, 6*cm, 3*cm, 2.2*cm, 2.5*cm]))
        story.append(Spacer(1, 0.3*cm))
    except Exception as e:
        story.append(Paragraph(f"Error representantes: {e}", st['alert']))

    # ── Mix por canal ──
    story.append(Paragraph("MIX POR CANAL", st['sec']))
    try:
        if repre_sel:
            df_c = DFS['x repre x canal'][DFS['x repre x canal']['Vendedor'] == repre_sel]
            act_c = get_ind(df_c, 'Año Actual Cajas', ['Canal','flia'])
        elif canal_sel:
            df_c = DFS['x flia x canal'][DFS['x flia x canal']['Canal'] == canal_sel]
            act_c = get_ind(df_c, 'Año Actual Cajas', ['Canal','flia'])
        else:
            act_c = get_ind(DFS['x flia x canal'], 'Año Actual Cajas', ['Canal','flia'])
        if flia_sel:
            act_c = act_c[act_c['flia'] == flia_sel]
        agg_c = act_c.groupby('Canal')['Total'].sum().reset_index()
        agg_c['pct'] = agg_c['Total'] / agg_c['Total'].sum() * 100
        agg_c = agg_c[agg_c['Total'] > 0].sort_values('Total', ascending=False)

        can_hdr = [['Canal', 'Cajas', 'Participación %']]
        for _, r in agg_c.iterrows():
            can_hdr.append([r['Canal'], f"{int(r['Total']):,}", f"{r['pct']:.0f}%"])
        story.append(tbl(can_hdr, [7*cm, 4*cm, 4*cm]))
        story.append(Spacer(1, 0.3*cm))
    except Exception as e:
        story.append(Paragraph(f"Error canal: {e}", st['alert']))

    # ── Alertas ──
    story.append(Paragraph("ALERTAS AUTOMÁTICAS", st['sec']))
    try:
        flags = generar_red_flags(flia_sel=flia_sel, repre_sel=repre_sel, canal_sel=canal_sel)
        icon = {'CRITICO': '⚠ ', 'ALERTA': '! ', 'OK': '✓ ', 'INFO': '→ '}
        for nivel, msg in flags:
            sty = st['alert'] if nivel in ('CRITICO','ALERTA') else (st['ok'] if nivel == 'OK' else st['body'])
            story.append(Paragraph(f"{icon.get(nivel,'')}{msg}", sty))
    except Exception as e:
        story.append(Paragraph(f"Error alertas: {e}", st['alert']))

    # ── Pendientes ──
    if 'pend' in DFS:
        story.append(Paragraph("PEDIDOS PENDIENTES (TOP 10)", st['sec']))
        try:
            df_p = DFS['pend'].copy()
            df_p.columns = [c.strip() for c in df_p.columns]
            df_p['Pedidos Pendientes'] = pd.to_numeric(df_p['Pedidos Pendientes'], errors='coerce')
            df_p = df_p[df_p['Pedidos Pendientes'] > 0]
            if repre_sel:
                df_p = df_p[df_p['Vendedor'].str.strip() == repre_sel]
            agg_p = df_p.groupby('Vendedor')['Pedidos Pendientes'].sum().reset_index()
            agg_p = agg_p.sort_values('Pedidos Pendientes', ascending=False).head(10)
            total_p = agg_p['Pedidos Pendientes'].sum()
            pend_hdr = [['Vendedor', 'Pendientes', '% del Total']]
            for _, r in agg_p.iterrows():
                pend_hdr.append([r['Vendedor'][:35],
                                  f"{int(r['Pedidos Pendientes']):,}",
                                  f"{r['Pedidos Pendientes']/total_p*100:.0f}%"])
            story.append(tbl(pend_hdr, [8*cm, 3.5*cm, 3.5*cm]))
        except Exception as e:
            story.append(Paragraph(f"Error pendientes: {e}", st['alert']))

    doc.build(story)
    buf.seek(0)
    return buf.read()


def generar_pdf_tab(tab, flia_sel=None, repre_sel=None, canal_sel=None):
    """PDF de resumen específico por pestaña."""
    if not PDF_AVAILABLE:
        return None

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=1.8*cm, rightMargin=1.8*cm,
                            topMargin=1.8*cm, bottomMargin=1.8*cm)

    gold_hex  = '#8B6914'
    red_hex   = '#C0392B'
    green_hex = '#1E7A40'

    def _st():
        s = getSampleStyleSheet()
        return {
            'h1':  ParagraphStyle('_h1',  fontSize=18, textColor=rl_colors.HexColor(gold_hex),
                                  fontName='Helvetica-Bold', spaceAfter=2),
            'h2':  ParagraphStyle('_h2',  fontSize=10, textColor=rl_colors.HexColor('#555555'),
                                  fontName='Helvetica', spaceAfter=8),
            'sec': ParagraphStyle('_sec', fontSize=8,  textColor=rl_colors.HexColor(gold_hex),
                                  fontName='Helvetica-Bold', textTransform='uppercase',
                                  spaceBefore=10, spaceAfter=4),
            'body':ParagraphStyle('_body',fontSize=8,  textColor=rl_colors.HexColor('#1A1A1A'),
                                  fontName='Helvetica', leading=12, spaceAfter=2),
            'alrt':ParagraphStyle('_alrt',fontSize=8,  textColor=rl_colors.HexColor(red_hex),
                                  fontName='Helvetica', spaceAfter=2),
            'ok':  ParagraphStyle('_ok',  fontSize=8,  textColor=rl_colors.HexColor(green_hex),
                                  fontName='Helvetica', spaceAfter=2),
        }

    def _tbl(data, widths):
        t = Table(data, colWidths=widths, repeatRows=1)
        t.setStyle(TableStyle([
            ('BACKGROUND',     (0,0), (-1,0),  rl_colors.HexColor(gold_hex)),
            ('TEXTCOLOR',      (0,0), (-1,0),  rl_colors.white),
            ('FONTNAME',       (0,0), (-1,0),  'Helvetica-Bold'),
            ('FONTSIZE',       (0,0), (-1,-1), 8),
            ('FONTNAME',       (0,1), (-1,-1), 'Helvetica'),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [rl_colors.white, rl_colors.HexColor('#F5F5F5')]),
            ('TEXTCOLOR',      (0,1), (-1,-1), rl_colors.HexColor('#1A1A1A')),
            ('GRID',           (0,0), (-1,-1), 0.4, rl_colors.HexColor('#CCCCCC')),
            ('PADDING',        (0,0), (-1,-1), 5),
            ('VALIGN',         (0,0), (-1,-1), 'MIDDLE'),
        ]))
        return t

    st   = _st()
    filtro_txt = " | ".join(filter(None, [flia_sel, repre_sel, canal_sel])) or "Región completa"
    TAB_LABELS = {
        'region':'Región','repre':'Representantes','clientes':'Clientes',
        'canales':'Canales','analisis':'Análisis','pendientes':'Pendientes',
    }
    story = []
    story.append(Paragraph("CATENA ZAPATA", st['h1']))
    story.append(Paragraph(f"Resumen — {TAB_LABELS.get(tab, tab)}", st['h2']))
    story.append(Paragraph(f"Filtro: {filtro_txt}   •   {datetime.now().strftime('%d/%m/%Y %H:%M')}", st['h2']))
    story.append(HRFlowable(width="100%", thickness=1.5,
                             color=rl_colors.HexColor(gold_hex), spaceAfter=8))

    # ── REGIÓN ────────────────────────────────────────────────────────────────
    if tab == 'region':
        # KPIs
        story.append(Paragraph("KPIs GENERALES", st['sec']))
        try:
            act = get_ind(DFS['x flia'], 'Año Actual Cajas', ['flia'])
            ant = get_ind(DFS['x flia'], 'Año Anterior Cajas', ['flia'])
            if flia_sel:
                act = act[act['flia']==flia_sel]; ant = ant[ant['flia']==flia_sel]
            ta, tb = act['Total'].sum(), ant['Total'].sum()
            vt = (ta-tb)/tb*100 if tb else 0
            story.append(_tbl([['Cajas Año Actual','Cajas Año Anterior','Variación %'],
                                [f"{int(ta):,}", f"{int(tb):,}", f"{'+' if vt>=0 else ''}{vt:.0f}%"]],
                               [5.5*cm, 5.5*cm, 4*cm]))
            story.append(Spacer(1, 0.3*cm))
        except Exception as e:
            story.append(Paragraph(f"Error KPIs: {e}", st['alrt']))

        # Familias
        story.append(Paragraph("RANKING FAMILIAS", st['sec']))
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
            story.append(_tbl(rows, [5*cm, 3*cm, 3*cm, 2.2*cm, 2.3*cm]))
            story.append(Spacer(1, 0.3*cm))
        except Exception as e:
            story.append(Paragraph(f"Error familias: {e}", st['alrt']))

        # Representantes
        story.append(Paragraph("RANKING REPRESENTANTES", st['sec']))
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
            story.append(_tbl(rows, [0.8*cm, 6.2*cm, 3*cm, 2.2*cm, 2.3*cm]))
        except Exception as e:
            story.append(Paragraph(f"Error representantes: {e}", st['alrt']))

    # ── REPRESENTANTES ────────────────────────────────────────────────────────
    elif tab == 'repre':
        if repre_sel:
            # KPIs del rep + ranking
            story.append(Paragraph("RESUMEN DEL REPRESENTANTE", st['sec']))
            try:
                df_r = DFS['x repre'][DFS['x repre']['Vendedor']==repre_sel]
                act = get_ind(df_r, 'Año Actual Cajas', ['Vendedor','flia'])
                ant = get_ind(df_r, 'Año Anterior Cajas', ['Vendedor','flia'])
                ta, tb = act['Total'].sum(), ant['Total'].sum()
                vt = (ta-tb)/tb*100 if tb else 0
                all_act = get_ind(DFS['x repre'], 'Año Actual Cajas', ['Vendedor','flia'])
                rk = all_act.groupby('Vendedor')['Total'].sum().rank(ascending=False, method='min')
                rank = int(rk.get(repre_sel, 0))
                s = '+' if vt>=0 else ''
                story.append(_tbl(
                    [['Cajas Año Actual','Cajas Año Anterior','Variación %','Ranking Nacional'],
                     [f"{int(ta):,}", f"{int(tb):,}", f"{s}{vt:.0f}%", f"#{rank} de {len(REPRESENTANTES)}"]],
                    [4*cm, 4*cm, 3.5*cm, 4*cm]))
                story.append(Spacer(1, 0.3*cm))
            except Exception as e:
                story.append(Paragraph(f"Error KPIs: {e}", st['alrt']))

            # Evolución mensual
            story.append(Paragraph("EVOLUCIÓN MENSUAL (cajas año actual)", st['sec']))
            try:
                df_r2 = DFS['x repre'][DFS['x repre']['Vendedor']==repre_sel]
                act2  = get_ind(df_r2, 'Año Actual Cajas', ['Vendedor','flia'])
                ant2  = get_ind(df_r2, 'Año Anterior Cajas', ['Vendedor','flia'])
                mes_rows = [['Mes','Año Actual','Año Anterior','Var %']]
                for m in MC:
                    if m in act2.columns:
                        va = pd.to_numeric(act2[m], errors='coerce').sum()
                        vb = pd.to_numeric(ant2[m], errors='coerce').sum() if m in ant2.columns else 0
                        vp = (va-vb)/vb*100 if vb else 0
                        s2 = '+' if vp>=0 else ''
                        mes_rows.append([m, f"{int(va):,}", f"{int(vb):,}", f"{s2}{vp:.0f}%"])
                story.append(_tbl(mes_rows, [2.5*cm, 4*cm, 4*cm, 3*cm]))
                story.append(Spacer(1, 0.3*cm))
            except Exception as e:
                story.append(Paragraph(f"Error evolución: {e}", st['alrt']))

            # Mix por canal
            story.append(Paragraph("MIX POR CANAL", st['sec']))
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
                story.append(_tbl(rows, [5*cm, 3.5*cm, 3*cm, 3*cm]))
                story.append(Spacer(1, 0.3*cm))
            except Exception as e:
                story.append(Paragraph(f"Error canales: {e}", st['alrt']))

            # Alertas por familia
            story.append(Paragraph("VARIACIÓN POR FAMILIA", st['sec']))
            try:
                rv = get_ind(DFS['x repre'], 'Var %', ['Vendedor','flia'])
                rv_r = rv[rv['Vendedor']==repre_sel].copy()
                rv_r['pct'] = rv_r['Total']*100
                rv_r = rv_r.dropna(subset=['pct']).sort_values('pct')
                for _, row in rv_r.iterrows():
                    s2 = '+' if row['pct']>=0 else ''
                    sty = st['ok'] if row['pct']>=0 else st['alrt']
                    story.append(Paragraph(f"• {row['flia']}: {s2}{row['pct']:.0f}%", sty))
                story.append(Spacer(1, 0.3*cm))
            except Exception as e:
                story.append(Paragraph(f"Error familias: {e}", st['alrt']))

            # Top clientes
            story.append(Paragraph("TOP CLIENTES — VARIACIÓN", st['sec']))
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
                    story.append(Paragraph("Mayor crecimiento:", st['body']))
                    story.append(_tbl(rows, [7*cm, 3.5*cm, 2.5*cm, 2.5*cm]))
                    story.append(Spacer(1, 0.2*cm))
                    rows2 = [['Cliente','Familia','Cajas Act','Var %']]
                    for _, r in mc3.nsmallest(8,'dif').iterrows():
                        s2 = '+' if r['vp']>=0 else ''
                        rows2.append([str(r['Cliente'])[:28], str(r['flia'])[:14],
                                       f"{int(r['Total_a']):,}", f"{s2}{r['vp']:.0f}%"])
                    story.append(Paragraph("Mayor caída:", st['body']))
                    story.append(_tbl(rows2, [7*cm, 3.5*cm, 2.5*cm, 2.5*cm]))
            except Exception as e:
                story.append(Paragraph(f"Error clientes: {e}", st['alrt']))
        else:
            # Sin rep seleccionado: ranking general
            story.append(Paragraph("RANKING GENERAL DE REPRESENTANTES", st['sec']))
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
                story.append(_tbl(rows, [0.8*cm, 6.2*cm, 3*cm, 2.2*cm, 2.3*cm]))
            except Exception as e:
                story.append(Paragraph(f"Error ranking: {e}", st['alrt']))

    # ── CLIENTES ──────────────────────────────────────────────────────────────
    elif tab == 'clientes':
        try:
            datos = analisis_clientes(repre_sel, flia_sel)
            n, p, a = len(datos['nuevos']), len(datos['perdidos']), len(datos['activos'])
            total_crec  = int(datos['nuevos']['dif'].sum())
            total_caida = int(datos['perdidos']['dif'].abs().sum())
            story.append(Paragraph("RESUMEN DE CLIENTES", st['sec']))
            story.append(_tbl(
                [['Activos ambos años','Con Crecimiento','Con Caídas',
                  'Cajas ganadas','Cajas perdidas'],
                 [str(a), str(n), str(p), f"{total_crec:,}", f"{total_caida:,}"]],
                [3.5*cm, 3*cm, 3*cm, 3.5*cm, 3.5*cm]))
            story.append(Spacer(1, 0.3*cm))

            # Top crecimiento
            story.append(Paragraph("TOP 10 — MAYOR CRECIMIENTO (cajas)", st['sec']))
            rows = [['Cliente','Representante','Cajas Act','Δ Cajas','Var %']]
            for _, r in datos['top_sube'].head(10).iterrows():
                s = '+' if r['dif'] >= 0 else ''
                vstr = f"{r['var']:+.0f}%" if pd.notna(r['var']) else '—'
                rows.append([str(r['Cliente'])[:26], str(r['Vendedor'])[:20],
                              f"{int(r['act']):,}", f"{s}{int(r['dif']):,}", vstr])
            story.append(_tbl(rows, [5*cm, 4*cm, 2.5*cm, 2.5*cm, 2.5*cm]))
            story.append(Spacer(1, 0.3*cm))

            # Top caída
            story.append(Paragraph("TOP 10 — MAYOR CAÍDA (cajas)", st['sec']))
            rows2 = [['Cliente','Representante','Cajas Act','Δ Cajas','Var %']]
            for _, r in datos['top_baja'].head(10).iterrows():
                vstr = f"{r['var']:+.0f}%" if pd.notna(r['var']) else '—'
                rows2.append([str(r['Cliente'])[:26], str(r['Vendedor'])[:20],
                               f"{int(r['act']):,}", f"{int(r['dif']):,}", vstr])
            story.append(_tbl(rows2, [5*cm, 4*cm, 2.5*cm, 2.5*cm, 2.5*cm]))
            story.append(Spacer(1, 0.3*cm))

            # Con crecimiento / con caídas
            if not datos['nuevos'].empty:
                story.append(Paragraph(f"CON CRECIMIENTO — TOP 10 (de {n} total)", st['sec']))
                rowsN = [['Cliente','Representante','Cajas Act','Δ Cajas']]
                for _, r in datos['nuevos'].nlargest(10,'dif').iterrows():
                    rowsN.append([str(r['Cliente'])[:30], str(r['Vendedor'])[:22],
                                  f"{int(r['act']):,}", f"+{int(r['dif']):,}"])
                story.append(_tbl(rowsN, [6*cm, 4.5*cm, 2.5*cm, 2.5*cm]))
                story.append(Spacer(1, 0.3*cm))

            if not datos['perdidos'].empty:
                story.append(Paragraph(f"CON CAÍDAS — TOP 10 (de {p} total)", st['sec']))
                rowsP = [['Cliente','Representante','Cajas Act','Δ Cajas']]
                for _, r in datos['perdidos'].nsmallest(10,'dif').iterrows():
                    rowsP.append([str(r['Cliente'])[:30], str(r['Vendedor'])[:22],
                                  f"{int(r['act']):,}", f"{int(r['dif']):,}"])
                story.append(_tbl(rowsP, [6*cm, 4.5*cm, 2.5*cm, 2.5*cm]))

        except Exception as e:
            story.append(Paragraph(f"Error clientes: {e}", st['alrt']))

    # ── CANALES ───────────────────────────────────────────────────────────────
    elif tab == 'canales':
        story.append(Paragraph("MIX POR CANAL", st['sec']))
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
            story.append(_tbl(rows, [4.5*cm, 3.5*cm, 3.5*cm, 3*cm, 2.5*cm]))
        except Exception as e:
            story.append(Paragraph(f"Error canales: {e}", st['alrt']))

    # ── ANÁLISIS ──────────────────────────────────────────────────────────────
    elif tab == 'analisis':
        story.append(Paragraph("RED FLAGS — ALERTAS AUTOMÁTICAS", st['sec']))
        try:
            flags = generar_red_flags(flia_sel=flia_sel, repre_sel=repre_sel, canal_sel=canal_sel)
            icon  = {'CRITICO':'⚠ ','ALERTA':'! ','OK':'✓ ','INFO':'→ '}
            for nivel, msg in flags:
                sty = st['alrt'] if nivel in ('CRITICO','ALERTA') else (st['ok'] if nivel=='OK' else st['body'])
                story.append(Paragraph(f"{icon.get(nivel,'')}{msg}", sty))
        except Exception as e:
            story.append(Paragraph(f"Error red flags: {e}", st['alrt']))

        # Generar análisis con los filtros activos (no usar globals)
        insights_l, _, _, foda_l = generar_analisis(flia_sel=flia_sel, repre_sel=repre_sel, canal_sel=canal_sel)

        story.append(Spacer(1, 0.4*cm))
        story.append(Paragraph("INSIGHTS CLAVE", st['sec']))
        for i in insights_l:
            story.append(Paragraph(f"• {i}", st['body']))

        story.append(Spacer(1, 0.4*cm))
        story.append(Paragraph("FODA AUTOMÁTICO", st['sec']))

        foda_cell_st = ParagraphStyle('_fc', fontSize=8, textColor=rl_colors.HexColor('#1A1A1A'),
                                      fontName='Helvetica', leading=11, spaceAfter=2)
        foda_hdr_st  = ParagraphStyle('_fh', fontSize=8, textColor=rl_colors.white,
                                      fontName='Helvetica-Bold', leading=10)

        def _foda_cell(items):
            return [Paragraph(f"• {x}", foda_cell_st) for x in items] or [Paragraph('—', foda_cell_st)]

        foda_data = [
            [Paragraph('FORTALEZAS', foda_hdr_st), Paragraph('DEBILIDADES', foda_hdr_st)],
            [_foda_cell(foda_l['F']), _foda_cell(foda_l['D'])],
            [Paragraph('OPORTUNIDADES', foda_hdr_st), Paragraph('AMENAZAS', foda_hdr_st)],
            [_foda_cell(foda_l['O']), _foda_cell(foda_l['A'])],
        ]
        tf = Table(foda_data, colWidths=[8.5*cm, 8.5*cm])
        tf.setStyle(TableStyle([
            ('BACKGROUND',  (0,0), (-1,0),  rl_colors.HexColor(green_hex)),
            ('BACKGROUND',  (0,2), (-1,2),  rl_colors.HexColor('#8B6914')),
            ('GRID',        (0,0), (-1,-1), 0.4, rl_colors.HexColor('#CCCCCC')),
            ('TOPPADDING',  (0,0), (-1,-1), 6),
            ('BOTTOMPADDING',(0,0),(-1,-1), 6),
            ('LEFTPADDING', (0,0), (-1,-1), 6),
            ('RIGHTPADDING',(0,0), (-1,-1), 6),
            ('VALIGN',      (0,0), (-1,-1), 'TOP'),
            ('ROWBACKGROUNDS', (0,1), (-1,1), [rl_colors.HexColor('#F0FFF4')]),
            ('ROWBACKGROUNDS', (0,3), (-1,3), [rl_colors.HexColor('#FFF8E6')]),
        ]))
        story.append(tf)

    # ── PENDIENTES ────────────────────────────────────────────────────────────
    elif tab == 'pendientes':
        story.append(Paragraph("PEDIDOS PENDIENTES POR VENDEDOR", st['sec']))
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
                story.append(_tbl(rows, [8*cm, 4*cm, 3.5*cm]))
                story.append(Spacer(1, 0.4*cm))

                # Detalle por familia
                story.append(Paragraph("DETALLE POR FAMILIA", st['sec']))
                df_det = df_p.sort_values('Pedidos Pendientes', ascending=False)
                rows2 = [['Familia Producto','Vendedor','Pendientes']]
                for _, r in df_det.iterrows():
                    rows2.append([str(r.get('Familia Producto',''))[:30],
                                   str(r.get('Vendedor',''))[:25],
                                   f"{int(r.get('Pedidos Pendientes',0)):,}"])
                story.append(_tbl(rows2, [7*cm, 5*cm, 3.5*cm]))
        except Exception as e:
            story.append(Paragraph(f"Error pendientes: {e}", st['alrt']))

    doc.build(story)
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

app = Dash(__name__, suppress_callback_exceptions=True)
app.title = "Dashboard Ventas — Catena Zapata"
server = app.server

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

_LI = {   # login input style
    'backgroundColor': C['surf2'], 'color': C['text'],
    'border': f"1px solid {C['border']}", 'borderRadius': '2px',
    'padding': '10px 12px', 'fontSize': '13px', 'fontFamily': FONT,
    'width': '100%', 'boxSizing': 'border-box', 'outline': 'none', 'marginBottom': '10px',
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
            html.P("Dashboard Comercial",
                   style={'color': C['muted'], 'fontSize': '9px', 'letterSpacing': '3px',
                          'textTransform': 'uppercase', 'textAlign': 'center', 'margin': '0 0 28px 0'}),
            html.Hr(style={'border': 'none', 'borderTop': f"1px solid {C['border']}", 'marginBottom': '24px'}),
            dcc.Input(id='login-user', type='text', placeholder='Usuario', style=_LI, debounce=False),
            dcc.Input(id='login-pass', type='password', placeholder='Contraseña', style=_LI, debounce=False),
            html.Div(id='login-error',
                     style={'color': C['red'], 'fontSize': '11px', 'marginBottom': '10px',
                            'textAlign': 'center', 'minHeight': '16px'}),
            html.Button('INGRESAR', id='btn-login', n_clicks=0, style={
                'backgroundColor': C['gold'], 'color': '#111', 'border': 'none',
                'padding': '11px 0', 'fontSize': '10px', 'letterSpacing': '2.5px',
                'textTransform': 'uppercase', 'cursor': 'pointer', 'borderRadius': '2px',
                'fontFamily': FONT, 'fontWeight': '700', 'width': '100%',
                'WebkitAppearance': 'none', 'appearance': 'none',
            }),
            html.Div([
                html.Span("Admin: jefe / piso3   |   Vendedores: primera palabra del nombre / su clave",
                          style={'color': C['muted'], 'fontSize': '9px', 'fontFamily': MONO}),
            ], style={'marginTop': '18px', 'textAlign': 'center'}),
        ], style={
            'backgroundColor': C['surf'], 'border': f"1px solid {C['border']}",
            'borderRadius': '4px', 'padding': '40px 36px', 'width': '380px', 'boxSizing': 'border-box',
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
                html.P("Dashboard Comercial — Jefatura Nacional de Ventas",
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
                        options=[{'label': m, 'value': m} for m in MC],
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
        CANALES = sorted(DFS['x flia x canal']['Canal'].unique().tolist())
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
            html.Div([dcc.Graph(figure=fig_repre_ranking(flia, canal, repre, meses), config={'displayModeBar':False})], style=CARD),
            html.Div([dcc.Graph(figure=fig_canal_mix(flia, repre, canal, meses), config={'displayModeBar':False})], style=CARD),
        ])

    elif tab == 'clientes':
        try:
            datos = analisis_clientes(repre, flia, meses)
            n, p, a = len(datos['nuevos']), len(datos['perdidos']), len(datos['activos'])
            total_crec  = int(datos['nuevos']['dif'].sum())
            total_caida = int(datos['perdidos']['dif'].abs().sum())
            stats_items = [
                ('ACTIVOS AMBOS AÑOS',    str(a),              C['gold']),
                ('CON CRECIMIENTO',        str(n),              C['green']),
                ('CON CAÍDAS',             str(p),              C['red']),
                ('CAJAS CRECIMIENTO',      f"{total_crec:,}",   C['green']),
                ('CAJAS CAÍDAS',           f"{total_caida:,}",  C['red']),
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
        insights_d, alertas_d, oportunidades_d, foda_d = generar_analisis(flia_sel=flia, repre_sel=repre, canal_sel=canal, meses_sel=meses)

        def item(txt, bc):
            return html.Div(txt, style={'padding':'8px 12px','backgroundColor':C['surf2'],
                                        'borderLeft':f"3px solid {bc}",
                                        'marginBottom':'6px','fontSize':'12px','lineHeight':'1.5'})

        flag_colors = {'CRITICO': C['red'], 'ALERTA': '#E67E22', 'OK': C['green'], 'INFO': C['muted']}

        return html.Div([
            _pbt,
            # Red Flags
            html.Div([
                html.Div("Red Flags — Alertas Automáticas", style={**SEC, 'color': C['red']}),
                *[item(f"[{nivel}] {msg}", flag_colors.get(nivel, C['muted'])) for nivel, msg in flags_d]
            ], style={**CARD, 'borderColor': C['red']}),

            html.Div([
                html.Div("Insights Clave", style=SEC),
                *[item(i, C['gold']) for i in insights_d]
            ], style=CARD_G),

            html.Div([
                html.Div([
                    html.Div("Alertas", style={**SEC,'color':C['red']}),
                    *[item(a, C['red']) for a in alertas_d]
                ], style=CARD),
                html.Div([
                    html.Div("Oportunidades", style={**SEC,'color':C['green']}),
                    *[item(o, C['green']) for o in oportunidades_d]
                ], style=CARD),
            ], style=G2),

            html.Div([
                html.Div("Análisis FODA Automático", style=SEC),
                html.Div([
                    html.Div([
                        html.Div("FORTALEZAS", style={**LABEL,'color':C['green']}),
                        *[html.Div(f"• {x}", style={'fontSize':'12px','marginBottom':'5px'}) for x in foda_d['F']]
                    ], style={**CARD,'borderColor':C['green']}),
                    html.Div([
                        html.Div("DEBILIDADES", style={**LABEL,'color':C['red']}),
                        *[html.Div(f"• {x}", style={'fontSize':'12px','marginBottom':'5px'}) for x in foda_d['D']]
                    ], style={**CARD,'borderColor':C['red']}),
                    html.Div([
                        html.Div("OPORTUNIDADES", style={**LABEL,'color':C['gold']}),
                        *[html.Div(f"• {x}", style={'fontSize':'12px','marginBottom':'5px'}) for x in foda_d['O']]
                    ], style={**CARD,'borderColor':C['gold']}),
                    html.Div([
                        html.Div("AMENAZAS", style={**LABEL,'color':C['muted']}),
                        *[html.Div(f"• {x}", style={'fontSize':'12px','marginBottom':'5px'}) for x in foda_d['A']]
                    ], style=CARD),
                ], style=G2),
            ], style=CARD),
        ])

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
    Input('auth-store', 'data'),
)
def cb_toggle_pages(auth):
    _login_show = {'display': 'flex', 'justifyContent': 'center', 'alignItems': 'center',
                   'minHeight': '100vh', 'backgroundColor': C['bg']}
    _login_hide = {'display': 'none'}
    _dash_show  = {}
    _dash_hide  = {'display': 'none'}
    if auth:
        if auth.get('role') == 'admin':
            badge = "Admin — acceso completo"
        else:
            badge = f"{auth.get('repre', '')} — vista personal"
        return _login_hide, _dash_show, badge
    return _login_show, _dash_hide, ""


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
    mes_opts   = [{'label':m,'value':m} for m in MC]
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
    CANALES = sorted(DFS['x flia x canal']['Canal'].unique().tolist())
    ts = f"Recargado desde Drive: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    return (version or 0) + 1, ts

# ── Run ────────────────────────────────────────────────────────────────────────

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
