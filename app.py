from flask import Flask, request, send_file, render_template_string
import pandas as pd
import io, os, gc
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB

# ── Estilos ───────────────────────────────────────────────────────────────────
BLK='000000'; WHT='FFFFFF'; ORG='E97132'; GRY='F2F2F2'; GRN='C6EFCE'; RED='FFC7CE'
FMT_EUR='_-* #,##0.00\\ "€"_-;\\-* #,##0.00\\ "€"_-;_-* "-"??\\ "€"_-;_-@_-'
FMT_INT='#,##0'; FMT_PCT='0.0%'
fill  = lambda c: PatternFill('solid', fgColor=c)
font  = lambda c=BLK, b=False, sz=11: Font(name='Calibri', size=sz, bold=b, color=c)
align = lambda h='center': Alignment(horizontal=h, vertical='center')
r2    = lambda n: round(float(n), 2) if n is not None else None

# ── Carga optimizada del maestro ──────────────────────────────────────────────
COLS_NEEDED = ['tienda','COMPARABLE','temporada','Vender_en','seccion','gama',
               'articulo','codart','color','neto','qty','fecha']

def load_maestro(file_obj):
    # Leer cabecera para detectar nombres exactos (case-insensitive)
    header_df = pd.read_excel(file_obj, nrows=0)
    file_obj.seek(0)
    col_map = {}
    for col in header_df.columns:
        cl = col.strip().lower()
        for needed in COLS_NEEDED:
            if cl == needed.lower():
                col_map[col] = needed

    missing = [c for c in COLS_NEEDED if c not in col_map.values()]
    if missing:
        raise ValueError(f'Columnas no encontradas: {missing}. Cabecera: {list(header_df.columns)}')

    # Cargar SOLO las columnas necesarias — ahorra ~90% de memoria
    df = pd.read_excel(file_obj, usecols=list(col_map.keys()),
                       dtype={'codart': str, 'color': str, 'tienda': str,
                              'articulo': str, 'gama': str, 'seccion': str,
                              'temporada': str, 'Vender_en': str})
    df.rename(columns=col_map, inplace=True)

    # Tipos optimizados
    df['neto']  = pd.to_numeric(df['neto'],  errors='coerce').fillna(0).astype('float32')
    df['qty']   = pd.to_numeric(df['qty'],   errors='coerce').fillna(0).astype('float32')
    df['fecha'] = pd.to_numeric(df['fecha'], errors='coerce').fillna(0).astype('int16')

    # Strings a categoría (enorme ahorro en columnas repetitivas)
    for col in ['tienda','COMPARABLE','temporada','Vender_en','seccion','gama','articulo','color']:
        df[col] = df[col].fillna('').astype(str).str.strip().astype('category')

    # Normalizar COMPARABLE: 0 → ''
    # FIX: replace instead of rename_categories to avoid duplicate error
    df['COMPARABLE'] = df['COMPARABLE'].astype(str).replace({'0':'','nan':'','None':''}).astype('category')

    # Filtrar solo años válidos
    df = df[df['fecha'].isin([2025, 2026])].copy()
    return df

# ── Helpers de cálculo ────────────────────────────────────────────────────────
def map_vend(v):
    if v == 'SS26': return 'SS26'
    if v == 'NOS CONTINUATIVO': return 'NOS'
    if v == 'FIN EXISTENCIAS':  return 'FE'
    if v == 'CEREMONIA':        return 'CEREMONIA'
    return None

def pneto(df, idx):
    g26 = df[df['fecha']==2026].groupby(idx)['neto'].sum().rename(2026)
    g25 = df[df['fecha']==2025].groupby(idx)['neto'].sum().rename(2025)
    out = pd.concat([g26, g25], axis=1).fillna(0).reset_index()
    out['DIF'] = out[2026] - out[2025]
    return out.sort_values(2026, ascending=False).reset_index(drop=True)

# ── Cálculos de cada pestaña ──────────────────────────────────────────────────
def calc_resumen(df):
    def build_block(data):
        d26 = data[data['fecha']==2026]; d25 = data[data['fecha']==2025]
        by26 = d26.groupby('seccion')['neto'].sum()
        by25 = d25.groupby('seccion')['neto'].sum()
        tot26 = float(by26.sum())
        d26v = d26.copy(); d26v['tg'] = d26v['Vender_en'].astype(str).map(map_vend)
        tn = d26v[d26v['tg'].notna()].groupby(['seccion','tg'])['neto'].sum().unstack(fill_value=0)
        for c in ['SS26','NOS','FE','CEREMONIA']:
            if c not in tn.columns: tn[c] = 0.0
        rows = []
        for sec in by26.sort_values(ascending=False).index:
            n26=r2(by26.get(sec,0)); n25=r2(by25.get(sec,0))
            pct = round(n26/tot26,4) if tot26 else 0
            ts = [r2(tn.loc[sec,c]) if sec in tn.index else 0 for c in ['SS26','NOS','FE','CEREMONIA']]
            tp = [round(t/n26,4) if n26 else 0 for t in ts]
            rows.append([sec,n26,n25,r2(n26-n25),pct,None,sec]+ts+[r2(sum(ts)),None,sec]+tp)
        tn26=r2(by26.sum()); tn25=r2(by25.sum())
        tt=[r2(tn[c].sum()) if c in tn.columns else 0 for c in ['SS26','NOS','FE','CEREMONIA']]
        ttp=[round(t/tn26,4) if tn26 else 0 for t in tt]
        rows.append(['Total general',tn26,tn25,r2(tn26-tn25),1.0,None,'Total general']+tt+[r2(sum(tt)),None,'Total general']+ttp)
        return rows
    hdr=['SECCIÓN',2026,2025,'DIF NETO','%',None,'SECCIÓN','SS26','NOS','FE','CEREMONIA','Total general',None,'SECCIÓN','SS26%','NOS%','FE%','CER%']
    comp = df[df['COMPARABLE'].isin(['SI','SI - Reformada'])]
    return [['TOTALES'],[],[],hdr]+build_block(df)+[[],['COMPARABLES'],[],[],hdr]+build_block(comp)

def calc_fac_tiendas(df):
    t=pneto(df,'tienda'); t.columns=['tienda',2026,2025,'dif neto']; return t

def calc_fac_seccion(df):
    t=pneto(df,['tienda','seccion']); t.columns=['tienda','seccion',2026,2025,'DIF NETO']; return t

def calc_gamas_comp(df):
    d=df[df['COMPARABLE'].isin(['SI','SI - Reformada'])]
    g26=d[d['fecha']==2026].groupby(['seccion','gama']).agg(N26=('neto','sum'),Q26=('qty','sum'))
    g25=d[d['fecha']==2025].groupby(['seccion','gama']).agg(N25=('neto','sum'),Q25=('qty','sum'))
    r=g26.join(g25,how='outer').fillna(0).reset_index()
    r['DIF NETO']=r['N26']-r['N25']; r['DIF QTY']=r['Q26']-r['Q25']
    r['PM 26']=r.apply(lambda x:round(x['N26']/x['Q26'],2) if x['Q26'] else None,axis=1)
    r['PM 25']=r.apply(lambda x:round(x['N25']/x['Q25'],2) if x['Q25'] else None,axis=1)
    r['DIF PM']=r.apply(lambda x:round(x['PM 26']-x['PM 25'],2) if pd.notna(x.get('PM 26')) and pd.notna(x.get('PM 25')) else None,axis=1)
    return r[['seccion','gama','N26','N25','DIF NETO','Q26','Q25','DIF QTY','PM 26','PM 25','DIF PM']].sort_values('N26',ascending=False).reset_index(drop=True)

def calc_top_cia_cod(df):
    r=df.groupby(['fecha','seccion','temporada','Vender_en','gama','codart','articulo']).agg(neto=('neto','sum'),qty=('qty','sum')).reset_index()
    r.columns=['fecha','seccion','temporada','Vender_en','gama','codart','articulo','Suma de neto','Suma de qty']
    return r.sort_values('Suma de neto',ascending=False).reset_index(drop=True)

def calc_gamas_tot(df):
    g26=df[df['fecha']==2026].groupby(['seccion','gama']).agg(N26=('neto','sum'),Q26=('qty','sum'))
    g25=df[df['fecha']==2025].groupby(['seccion','gama']).agg(N25=('neto','sum'),Q25=('qty','sum'))
    r=g26.join(g25,how='outer').fillna(0).reset_index()
    r['DIF NETO']=r['N26']-r['N25']; r['DIF QTY']=r['Q26']-r['Q25']
    return r[['seccion','gama','N26','N25','DIF NETO','Q26','Q25','DIF QTY']].sort_values('N26',ascending=False).reset_index(drop=True)

def calc_gamas_tienda(df):
    g26=df[df['fecha']==2026].groupby(['tienda','seccion','gama']).agg(N26=('neto','sum'),Q26=('qty','sum'))
    g25=df[df['fecha']==2025].groupby(['tienda','seccion','gama']).agg(N25=('neto','sum'),Q25=('qty','sum'))
    r=g26.join(g25,how='outer').fillna(0).reset_index()
    r['DIF NETO']=r['N26']-r['N25']; r['DIF QTY']=r['Q26']-r['Q25']
    return r[['tienda','seccion','gama','N26','N25','DIF NETO','Q26','Q25','DIF QTY']].sort_values('N26',ascending=False).reset_index(drop=True)

def calc_top_ventas_cia(df):
    r=df.groupby(['fecha','seccion','temporada','Vender_en','gama','codart','articulo','color']).agg(neto=('neto','sum'),qty=('qty','sum')).reset_index()
    r.columns=['fecha','seccion','temporada','Vender_en','gama','codart','articulo','color','Suma de neto','Suma de qty']
    return r.sort_values('Suma de neto',ascending=False).reset_index(drop=True)

def calc_top_venta_tienda(df):
    r=df.groupby(['tienda','fecha','seccion','temporada','Vender_en','gama','codart','articulo','color']).agg(neto=('neto','sum'),qty=('qty','sum')).reset_index()
    r.columns=['tienda','fecha','seccion','temporada','Vender_en','gama','codart','articulo','color','Suma de neto','Suma de qty']
    return r.sort_values('Suma de neto',ascending=False).reset_index(drop=True)

def calc_top_tienda_cod(df):
    r=df.groupby(['tienda','fecha','seccion','temporada','Vender_en','gama','codart','articulo']).agg(neto=('neto','sum'),qty=('qty','sum')).reset_index()
    r.columns=['tienda','fecha','seccion','temporada','Vender_en','gama','codart','articulo','Suma de neto','Suma de qty']
    return r.sort_values('Suma de neto',ascending=False).reset_index(drop=True)

# ── Escritura Excel con formato ───────────────────────────────────────────────
SHEET_COLS = {
    'FACTURACIÓN TIENDAS':    [('tienda',47,None,False),(2026,13,FMT_EUR,False),(2025,13,FMT_EUR,False),('dif neto',13,FMT_EUR,True)],
    'FACTURACIÓN SECCIÓN':    [('tienda',47,None,False),('seccion',16,None,False),(2026,13,FMT_EUR,False),(2025,13,FMT_EUR,False),('DIF NETO',13,FMT_EUR,True)],
    'GAMAS CIA COMPARABLES':  [('seccion',14,None,False),('gama',17,None,False),('N26',13,FMT_EUR,False),('N25',13,FMT_EUR,False),('DIF NETO',12,FMT_EUR,True),('Q26',10,FMT_INT,False),('Q25',11,FMT_INT,False),('DIF QTY',11,FMT_INT,True),('PM 26',11,FMT_EUR,False),('PM 25',11,FMT_EUR,False),('DIF PM',11,FMT_EUR,True)],
    'TOP CIA COD':            [('fecha',10,None,False),('seccion',14,None,False),('temporada',17,None,False),('Vender_en',17,None,False),('gama',15,None,False),('codart',12,None,False),('articulo',57,None,False),('Suma de neto',13,FMT_EUR,False),('Suma de qty',11,FMT_INT,False)],
    'GAMAS CIA TOTALES':      [('seccion',14,None,False),('gama',17,None,False),('N26',13,FMT_EUR,False),('N25',13,FMT_EUR,False),('DIF NETO',12,FMT_EUR,True),('Q26',10,FMT_INT,False),('Q25',11,FMT_INT,False),('DIF QTY',11,FMT_INT,True)],
    'GAMAS TIENDA':           [('tienda',47,None,False),('seccion',14,None,False),('gama',17,None,False),('N26',12,FMT_EUR,False),('N25',12,FMT_EUR,False),('DIF NETO',12,FMT_EUR,True),('Q26',10,FMT_INT,False),('Q25',11,FMT_INT,False),('DIF QTY',11,FMT_INT,True)],
    'TOP VENTAS CIA':         [('fecha',10,None,False),('seccion',14,None,False),('temporada',17,None,False),('Vender_en',17,None,False),('gama',15,None,False),('codart',12,None,False),('articulo',57,None,False),('color',22,None,False),('Suma de neto',13,FMT_EUR,False),('Suma de qty',11,FMT_INT,False)],
    'TOP VENTA TIENDA':       [('tienda',47,None,False),('fecha',10,None,False),('seccion',14,None,False),('temporada',17,None,False),('Vender_en',17,None,False),('gama',15,None,False),('codart',12,None,False),('articulo',57,None,False),('color',22,None,False),('Suma de neto',13,FMT_EUR,False),('Suma de qty',11,FMT_INT,False)],
    'TOP TIENDA COD':         [('tienda',47,None,False),('fecha',10,None,False),('seccion',14,None,False),('temporada',17,None,False),('Vender_en',17,None,False),('gama',15,None,False),('codart',12,None,False),('articulo',57,None,False),('Suma de neto',13,FMT_EUR,False),('Suma de qty',11,FMT_INT,False)],
}

def write_data_sheet(wb, name, df):
    ws = wb.create_sheet(name)
    col_defs = SHEET_COLS[name]
    for ci,(lbl,width,fmt,orange) in enumerate(col_defs,1):
        c=ws.cell(1,ci,value=lbl)
        c.fill=fill(ORG if orange else BLK); c.font=font(WHT,True); c.alignment=align()
        ws.column_dimensions[get_column_letter(ci)].width=width
    ws.row_dimensions[1].height=20; ws.freeze_panes='A2'
    ws.auto_filter.ref=f'A1:{get_column_letter(len(col_defs))}1'
    for ri,row in enumerate(df.itertuples(index=False),2):
        alt=ri%2==0
        for ci,(_,_,fmt,_) in enumerate(col_defs,1):
            val=row[ci-1] if ci-1<len(row) else None
            if val is pd.NA or (isinstance(val,float) and pd.isna(val)): val=None
            c=ws.cell(ri,ci,value=val); c.font=font(); c.alignment=align()
            if alt: c.fill=fill(GRY)
            if fmt and val is not None: c.number_format=fmt
            cn=str(col_defs[ci-1][0])
            if 'DIF' in cn.upper() and isinstance(val,(int,float)) and val is not None:
                c.fill=fill(GRN if val>0 else RED if val<0 else (GRY if alt else 'FFFFFF'))

def write_resumen_sheet(wb, rows):
    ws=wb.create_sheet('RESUMEN')
    widths={'A':16,'B':14,'C':14,'D':12,'E':11,'G':16,'H':14,'I':13,'J':12,'K':12,'L':14,'N':13,'O':7,'P':7,'Q':7,'R':12}
    for col,w in widths.items(): ws.column_dimensions[col].width=w
    for ri,row in enumerate(rows,1):
        if not row: continue
        if row[0] in('TOTALES','COMPARABLES') and all(v is None for v in row[1:]):
            ws.cell(ri,1,value=row[0]).font=font(BLK,True,13); continue
        is_hdr=row[0]=='SECCIÓN'; is_tot=row[0]=='Total general'
        for ci,val in enumerate(row,1):
            if val is None: continue
            c=ws.cell(ri,ci,value=val)
            if is_hdr:
                c.fill=fill(BLK); c.font=font(WHT,True); c.alignment=align()
            elif is_tot:
                c.fill=fill(ORG); c.font=font(WHT,True); c.alignment=align()
                if ci in(2,3,4,8,9,10,11,12): c.number_format=FMT_EUR
                if ci==5 or ci in(15,16,17,18): c.number_format=FMT_PCT
            else:
                c.font=font(); c.alignment=align()
                if ci in(2,3,4,8,9,10,11,12): c.number_format=FMT_EUR
                if ci==5 or ci in(15,16,17,18): c.number_format=FMT_PCT
                if ci==4 and isinstance(val,(int,float)):
                    c.fill=fill(GRN if val>=0 else RED)

def generar_excel(df, tipo, semana, anio):
    wb=Workbook(); wb.remove(wb.active)
    write_resumen_sheet(wb, calc_resumen(df))

    tabs=[
        ('FACTURACIÓN TIENDAS',   calc_fac_tiendas),
        ('FACTURACIÓN SECCIÓN',   calc_fac_seccion),
        ('GAMAS CIA COMPARABLES', calc_gamas_comp),
        ('TOP CIA COD',           calc_top_cia_cod),
        ('GAMAS CIA TOTALES',     calc_gamas_tot),
        ('GAMAS TIENDA',          calc_gamas_tienda),
        ('TOP VENTAS CIA',        calc_top_ventas_cia),
        ('TOP VENTA TIENDA',      calc_top_venta_tienda),
        ('TOP TIENDA COD',        calc_top_tienda_cod),
    ]
    for name, fn in tabs:
        result = fn(df)
        write_data_sheet(wb, name, result)
        del result; gc.collect()  # liberar memoria tras cada pestaña

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    del wb; gc.collect()
    return buf, f'PESOS_{tipo}_W{semana:02d}_{anio}.xlsx'

# ── HTML ──────────────────────────────────────────────────────────────────────
HTML = '''<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Generador PESOS</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@300;400;500;600&display=swap');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{--ink:#0f0f14;--paper:#f5f4ef;--accent:#c8f035;--mid:#7a7a8a;--border:#d8d7cf;--card:#fff;--orange:#e97132;--green:#c6efce;--green-dk:#276221;--red:#ffc7ce}
body{font-family:'DM Sans',sans-serif;background:var(--paper);color:var(--ink);min-height:100vh}
header{background:var(--ink);padding:18px 36px;display:flex;align-items:center;justify-content:space-between}
.logo{font-family:'DM Mono',monospace;font-size:12px;font-weight:500;color:var(--accent);letter-spacing:.14em;text-transform:uppercase}
.pill{font-family:'DM Mono',monospace;font-size:11px;color:#555;background:#1a1a22;padding:5px 12px;border-radius:20px}
main{max-width:760px;margin:0 auto;padding:36px 20px 80px}
.week-row{display:flex;gap:20px;align-items:center;margin-bottom:28px;padding:18px 22px;background:var(--card);border-radius:12px;border:1px solid var(--border)}
.wlbl{display:flex;flex-direction:column;gap:4px;align-items:center}
.wlbl label{font-size:10px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--mid)}
.wlbl input{font-family:'DM Mono',monospace;font-size:22px;font-weight:500;width:68px;border:none;border-bottom:2px solid var(--ink);background:transparent;color:var(--ink);padding:2px 0;text-align:center;outline:none;-moz-appearance:textfield}
.wlbl input::-webkit-inner-spin-button{-webkit-appearance:none}
.wlbl input:focus{border-color:var(--accent)}
.wsep{font-family:'DM Mono',monospace;font-size:20px;color:var(--border)}
.drop-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:28px}
.dz{border:1.5px dashed var(--border);border-radius:12px;padding:26px 18px;text-align:center;cursor:pointer;transition:all .2s;background:var(--card);position:relative;min-height:145px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:10px}
.dz:hover,.dz.over{border-color:var(--ink);background:#fafaf6}
.dz.loaded{border-color:var(--accent);border-style:solid;background:#f6ffe8}
.dz input[type=file]{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
.dz-icon{font-size:26px}.dz-lbl{font-size:10px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--mid)}
.dz-sub{font-size:11px;color:#bbb}.dz-opt{font-size:10px;color:#ccc;font-style:italic}
.dz-name{font-family:'DM Mono',monospace;font-size:10px;background:var(--accent);color:var(--ink);padding:2px 8px;border-radius:3px;font-weight:600;max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.dz-meta{font-family:'DM Mono',monospace;font-size:9px;color:var(--mid);line-height:1.8}
.dz-clr{position:absolute;top:8px;right:10px;background:none;border:none;font-size:14px;color:#bbb;cursor:pointer;padding:3px 6px;border-radius:4px;z-index:2}
.dz-clr:hover{color:var(--ink);background:#f0ede4}
.fmt-bar{display:flex;gap:7px;flex-wrap:wrap;margin-bottom:24px;padding:10px 14px;background:var(--card);border-radius:10px;border:1px solid var(--border);align-items:center}
.fmt-bar>span:first-child{font-size:10px;font-weight:600;color:var(--mid)}
.chip{font-size:10px;font-weight:600;padding:3px 8px;border-radius:4px}
.c-blk{background:var(--ink);color:var(--accent)}.c-org{background:var(--orange);color:#fff}
.c-grn{background:var(--green);color:var(--green-dk)}.c-red{background:var(--red);color:#9c0006}
#prog{display:none;padding:20px 22px;background:var(--card);border-radius:12px;border:1px solid var(--border);margin-bottom:24px;text-align:center}
.spinner{width:36px;height:36px;border:3px solid var(--border);border-top-color:var(--ink);border-radius:50%;animation:spin .8s linear infinite;margin:0 auto 12px}
@keyframes spin{to{transform:rotate(360deg)}}
.prog-lbl{font-family:'DM Mono',monospace;font-size:12px;color:var(--mid)}
.prog-bar-bg{width:100%;height:3px;background:var(--border);border-radius:2px;overflow:hidden;margin-top:12px}
.prog-bar-fill{height:100%;background:var(--ink);border-radius:2px;width:0%;transition:width .4s ease}
#err{display:none;background:#fff5f5;border:1px solid #ffd0d0;border-radius:10px;padding:14px 18px;font-size:12px;color:#cc2222;margin-bottom:20px;font-family:'DM Mono',monospace;white-space:pre-wrap}
#btn{width:100%;padding:16px;background:var(--ink);color:var(--accent);border:none;border-radius:12px;font-family:'DM Mono',monospace;font-size:13px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;cursor:pointer;transition:all .2s;display:flex;align-items:center;justify-content:center;gap:10px}
#btn:hover:not(:disabled){background:#1e1e2e;transform:translateY(-1px)}
#btn:disabled{opacity:.3;cursor:not-allowed;transform:none}
.results{display:flex;flex-direction:column;gap:12px;margin-top:20px}
.result-card{background:var(--ink);border-radius:14px;padding:22px 26px;color:#fff;display:flex;align-items:center;justify-content:space-between;gap:20px}
.result-info{display:flex;align-items:center;gap:14px}
.result-icon{font-size:26px}
.result-name{font-family:'DM Mono',monospace;font-size:13px;color:var(--accent);font-weight:500}
.result-sub{font-size:11px;color:#666;margin-top:2px}
.btn-dl{padding:10px 20px;background:var(--accent);color:var(--ink);border:none;border-radius:8px;font-family:'DM Mono',monospace;font-size:12px;font-weight:700;text-transform:uppercase;cursor:pointer;white-space:nowrap;text-decoration:none;display:inline-block}
.btn-dl:hover{background:#d4ff40}
</style>
</head>
<body>
<header>
  <span class="logo">⬡ Generador PESOS</span>
  <span class="pill" id="wpill">W__ · ____</span>
</header>
<main>
  <div class="week-row">
    <div class="wlbl"><label>Semana</label><input type="number" id="isem" min="1" max="53" value="{{ sem }}"/></div>
    <span class="wsep">/</span>
    <div class="wlbl"><label>Año</label><input type="number" id="ianio" min="2020" max="2099" value="{{ anio }}"/></div>
  </div>
  <div class="drop-grid">
    <div class="dz" id="dz-fp">
      <input type="file" id="f-fp" accept=".xlsx,.xls"/>
      <div class="dz-icon">📋</div><div class="dz-lbl">Maestro Full Price</div>
      <div class="dz-sub">Arrastra aquí o haz clic</div>
    </div>
    <div class="dz" id="dz-out">
      <input type="file" id="f-out" accept=".xlsx,.xls"/>
      <div class="dz-icon">🏪</div><div class="dz-lbl">Maestro Outlet</div>
      <div class="dz-sub">Arrastra aquí o haz clic</div>
      <div class="dz-opt">opcional</div>
    </div>
  </div>
  <div class="fmt-bar">
    <span>FORMATO:</span>
    <span class="chip c-blk">Cabeceras negras</span>
    <span class="chip c-org">DIF naranja</span>
    <span class="chip c-grn">▲ Positivo</span>
    <span class="chip c-red">▼ Negativo</span>
  </div>
  <div id="err"></div>
  <div id="prog">
    <div class="spinner"></div>
    <div class="prog-lbl" id="prog-lbl">Procesando... puede tardar 1-2 minutos</div>
    <div class="prog-bar-bg"><div class="prog-bar-fill" id="prog-fill"></div></div>
  </div>
  <button id="btn" disabled>→ &nbsp;Generar PESOS</button>
  <div class="results" id="results"></div>
</main>
<script>
const files={};
const $=id=>document.getElementById(id);
function upPill(){$('wpill').textContent=`W${String($('isem').value).padStart(2,'0')} · ${$('ianio').value}`;}
function chkReady(){$('btn').disabled=!Object.keys(files).length;}
function showErr(m){const e=$('err');e.textContent=m;e.style.display='block';}
function hideErr(){$('err').style.display='none';}

function setupDz(key){
  const dz=$(`dz-${key}`);
  dz.addEventListener('dragover',e=>{e.preventDefault();dz.classList.add('over');});
  dz.addEventListener('dragleave',()=>dz.classList.remove('over'));
  dz.addEventListener('drop',e=>{e.preventDefault();dz.classList.remove('over');if(e.dataTransfer.files[0])setFile(e.dataTransfer.files[0],key);});
  $(`f-${key}`).addEventListener('change',e=>{if(e.target.files[0])setFile(e.target.files[0],key);});
}
function setFile(file,key){
  files[key]=file;
  const dz=$(`dz-${key}`);dz.classList.add('loaded');
  dz.innerHTML=`<button class="dz-clr" onclick="clrFile('${key}')">✕</button>
    <input type="file" id="f-${key}" accept=".xlsx,.xls"/>
    <div class="dz-icon">${key==='fp'?'📋':'🏪'}</div>
    <div class="dz-lbl">${key==='fp'?'Full Price':'Outlet'}</div>
    <div class="dz-name">${file.name}</div>
    <div class="dz-meta">${(file.size/1024/1024).toFixed(1)} MB</div>`;
  $(`f-${key}`).addEventListener('change',e=>{if(e.target.files[0])setFile(e.target.files[0],key);});
  hideErr();chkReady();
}
function clrFile(key){
  delete files[key];
  const dz=$(`dz-${key}`);dz.classList.remove('loaded');
  dz.innerHTML=`<input type="file" id="f-${key}" accept=".xlsx,.xls"/>
    <div class="dz-icon">${key==='fp'?'📋':'🏪'}</div>
    <div class="dz-lbl">${key==='fp'?'Maestro Full Price':'Maestro Outlet'}</div>
    <div class="dz-sub">Arrastra aquí o haz clic</div>
    ${key==='out'?'<div class="dz-opt">opcional</div>':''}`;
  $(`f-${key}`).addEventListener('change',e=>{if(e.target.files[0])setFile(e.target.files[0],key);});
  chkReady();
}
setupDz('fp');setupDz('out');
$('isem').addEventListener('input',upPill);$('ianio').addEventListener('input',upPill);upPill();

let progInterval=null;
function startFakeProgress(){
  let p=5;
  $('prog-fill').style.width='5%';
  progInterval=setInterval(()=>{
    if(p<90){p+=Math.random()*3;$('prog-fill').style.width=Math.min(p,90)+'%';}
  },800);
}
function stopProgress(){
  clearInterval(progInterval);
  $('prog-fill').style.width='100%';
}

$('btn').addEventListener('click',async()=>{
  hideErr();$('results').innerHTML='';
  $('btn').disabled=true;$('prog').style.display='block';
  startFakeProgress();
  const sem=$('isem').value,anio=$('ianio').value;
  const resultCards=[];
  try{
    for(const[key,file]of Object.entries(files)){
      const tipo=key==='fp'?'FULL_PRICE':'OUTLET';
      $('prog-lbl').textContent=`Generando ${tipo}... puede tardar 1-2 minutos`;
      const fd=new FormData();
      fd.append('file',file);fd.append('tipo',tipo);
      fd.append('semana',sem);fd.append('anio',anio);
      const resp=await fetch('/generar',{method:'POST',body:fd});
      if(!resp.ok){const t=await resp.text();throw new Error(t);}
      const blob=await resp.blob();
      const fname=`PESOS_${tipo}_W${String(sem).padStart(2,'0')}_${anio}.xlsx`;
      resultCards.push({key,tipo,fname,url:URL.createObjectURL(blob)});
    }
    stopProgress();
    setTimeout(()=>{$('prog').style.display='none';},400);
    $('results').innerHTML=resultCards.map(({key,tipo,fname,url})=>`
      <div class="result-card">
        <div class="result-info">
          <span class="result-icon">${key==='fp'?'📋':'🏪'}</span>
          <div><div class="result-name">${fname}</div>
          <div class="result-sub">${tipo==='FULL_PRICE'?'Full Price':'Outlet'} · 10 pestañas formateadas</div></div>
        </div>
        <a class="btn-dl" href="${url}" download="${fname}">↓ Descargar</a>
      </div>`).join('');
  }catch(err){
    stopProgress();$('prog').style.display='none';
    showErr('Error: '+err.message);
  }
  $('btn').disabled=false;
});
</script>
</body>
</html>'''

# ── Rutas ─────────────────────────────────────────────────────────────────────
@app.route('/')
def index():
    now=datetime.now()
    return render_template_string(HTML, sem=now.isocalendar()[1], anio=now.year)

@app.route('/generar', methods=['POST'])
def generar():
    try:
        file  = request.files.get('file')
        tipo  = request.form.get('tipo','FULL_PRICE')
        sem   = int(request.form.get('semana',1))
        anio  = int(request.form.get('anio',datetime.now().year))
        if not file: return 'No se recibió ningún archivo',400
        df = load_maestro(file)
        if len(df)==0: return 'El archivo no tiene datos válidos',400
        buf, fname = generar_excel(df, tipo, sem, anio)
        del df; gc.collect()
        return send_file(buf,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True, download_name=fname)
    except Exception as e:
        return str(e), 500

if __name__=='__main__':
    port=int(os.environ.get('PORT',5000))
    app.run(host='0.0.0.0',port=port,debug=False)
