#!/usr/bin/env python3
"""
Reads the xlsx source files and regenerates index.html with fresh data.
Run directly: uv run --with openpyxl python3 -I rebuild.py
"""
import os, sys
from datetime import datetime
from collections import Counter

BASE     = os.path.dirname(os.path.abspath(__file__))
DELTA    = os.path.join(BASE, 'TERNIUM - Delta Consolidado.xlsx')
PL_FILE  = os.path.join(BASE, 'TERNIUM - Parking Lots Consolidado.xlsx')
MC_FILE  = os.path.join(BASE, 'Situación Actual Material Cluster.xlsx')
OUT      = os.path.join(BASE, 'index.html')

try:
    import openpyxl
except ImportError:
    sys.exit('openpyxl not found. Run: pip install openpyxl')

# ── safe encoding for JS single-quoted strings ────────────────────────────────
def s(v, maxlen=None):
    if v is None:
        return ''
    t = str(v).strip().replace('\n', ' ').replace('\r', '')
    t = t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    t = t.replace("'", '&#39;').replace('"', '&quot;')
    if maxlen:
        t = t[:maxlen]
    return t

def pct(n, d):
    return f'{n/d*100:.1f}' if d else '0'

def bpct(n, d):
    return f'{min(n/d*100, 100):.1f}' if d else '0'

# ── DELTAS ────────────────────────────────────────────────────────────────────
wb_d  = openpyxl.load_workbook(DELTA, data_only=True)
ws_d  = wb_d['GAP List']
rows_d = [r for r in ws_d.iter_rows(min_row=5, values_only=True) if any(c is not None for c in r)]

total_d   = len(rows_d)
stat_d    = Counter(r[9]  for r in rows_d if r[9])
prio_d    = Counter(r[7]  for r in rows_d if r[7])
type_d    = Counter(r[12] for r in rows_d if r[12])
mod_d     = Counter(r[1]  for r in rows_d if r[1])

open_d    = stat_d.get('Creado', 0)
closed_d  = stat_d.get('Cerrado', 0)
aprov_d   = stat_d.get('Aprobado', 0)
cancel_d  = sum(v for k,v in stat_d.items() if 'Cancelado' in str(k))
reject_d  = stat_d.get('Rechazado', 0)
mand_d    = prio_d.get('1 - Obligatorio', 0)
iface_d   = sum(v for k,v in type_d.items() if 'Interfaz' in str(k))
enh_d     = sum(v for k,v in type_d.items() if 'nhace' in str(k))
newcfg_d  = sum(v for k,v in type_d.items() if 'Nueva' in str(k))
rep_d     = sum(v for k,v in type_d.items() if 'Reporte' in str(k))
mig_d     = type_d.get('Conversion/Migración', 0)

open_rows_d = [r for r in rows_d if str(r[9] or '').strip() == 'Creado']

# ── PARKING LOTS ──────────────────────────────────────────────────────────────
wb_p  = openpyxl.load_workbook(PL_FILE, data_only=True)
ws_p  = wb_p['Parking Lots']
rows_p = [r for r in ws_p.iter_rows(min_row=5, values_only=True) if any(c is not None for c in r)]

total_p      = len(rows_p)
stat_p       = Counter(r[10] for r in rows_p if r[10])
mod_p        = Counter(r[1]  for r in rows_p if r[1])
pl_closed    = stat_p.get('Cerrado', 0) + stat_p.get('Cerrado-Duplicidad', 0)
pl_pc        = stat_p.get('Pendiente Cliente', 0)
pl_ps        = stat_p.get('Pendiente SAP', 0)
pl_pending   = pl_pc + pl_ps
pl_clos_pct  = pct(pl_closed, total_p)

def ws_rows(mods):
    return [r for r in rows_p if str(r[1] or '').strip() in mods]

def js_pl_rows(rows):
    lines = []
    for r in rows:
        lines.append(
            "      ['" + s(r[0]) + "','" + s(r[1]) + "','" + s(r[2], 45) +
            "','" + s(r[5], 75) + "','" + s(r[7]) + "','" + s(r[8]) + "','" + s(r[10]) + "']"
        )
    return ',\n'.join(lines)

pp_rows    = ws_rows({'PP', 'AR-PP'})
ba_rows    = ws_rows({'B&A'})
mmim_rows  = ws_rows({'MM-IM', 'AR-MM-IM'})
co_rows    = ws_rows({'CO', 'co', 'AR-CO'})
mmpur_rows = ws_rows({'MM-PUR', 'AR-MM-PUR'})
sd_rows    = ws_rows({'SD', 'AR-SD'})

# ── MATERIAL CLUSTER ──────────────────────────────────────────────────────────
wb_mc  = openpyxl.load_workbook(MC_FILE, data_only=True)
ws_mc  = wb_mc['Sheet1']
mc_rows = [r for r in ws_mc.iter_rows(min_row=2, values_only=True) if r[0] and str(r[0]).strip()]

def js_mc_rows():
    lines = []
    for r in mc_rows:
        cells = [s(r[i] if i < len(r) else '') for i in range(5)]
        lines.append("      ['" + "','".join(cells) + "']")
    return ',\n'.join(lines)

# ── pre-compute CARD_DATA JS blocks ──────────────────────────────────────────
today = datetime.now().strftime('%B %-d, %Y')

# total-deltas rows
td_lines = []
for m, c in sorted(mod_d.items(), key=lambda x: -x[1]):
    op = sum(1 for r in rows_d if str(r[1] or '').strip() == m and str(r[9] or '').strip() == 'Creado')
    top_types = Counter(str(r[12] or '') for r in rows_d if str(r[1] or '').strip() == m)
    top = top_types.most_common(1)[0][0] if top_types else '-'
    td_lines.append(f"      ['{s(m)}','{c}','{op}','{s(top)}']")
total_deltas_js = ',\n'.join(td_lines)

# open-deltas rows
om_lines = []
for r in open_rows_d[:12]:
    om_lines.append(f"      ['{s(r[0])}','{s(r[1])}','{s(r[3],60)}','{s(r[7])}','{s(r[12])}']")
open_modal_js = ',\n'.join(om_lines)

# closed-deltas rows
cd_lines = []
for m in sorted(mod_d.keys()):
    cnt = sum(1 for r in rows_d if str(r[1] or '').strip() == m and str(r[9] or '').strip() == 'Cerrado')
    if cnt > 0:
        cd_lines.append(f"      ['{s(m)}','{cnt}','Cerrado','Agreed SAP solution']")
closed_modal_js = ',\n'.join(cd_lines)

# mandatory-deltas rows
md_lines = []
for r in rows_d:
    if str(r[7] or '').startswith('1'):
        md_lines.append(f"      ['{s(r[0])}','{s(r[1])}','{s(r[3],60)}','{s(r[9])}','{s(r[12])}']")
mand_modal_js = ',\n'.join(md_lines)

# total-pl rows
pl_mod_lines = []
for m, c in sorted(mod_p.items(), key=lambda x: -x[1]):
    cl = sum(1 for r in rows_p if str(r[1] or '').strip() == m and 'Cerrado' in str(r[10] or ''))
    pc = sum(1 for r in rows_p if str(r[1] or '').strip() == m and str(r[10] or '') == 'Pendiente Cliente')
    ps = sum(1 for r in rows_p if str(r[1] or '').strip() == m and str(r[10] or '') == 'Pendiente SAP')
    pl_mod_lines.append(f"      ['{s(m)}','{c}','{cl}','{pc}','{ps}']")
total_pl_js = ',\n'.join(pl_mod_lines)

# pl-closed rows
plc_lines = []
for m, c in sorted(mod_p.items(), key=lambda x: -x[1]):
    cl = sum(1 for r in rows_p if str(r[1] or '').strip() == m and 'Cerrado' in str(r[10] or ''))
    if cl > 0:
        plc_lines.append(f"      ['{s(m)}','{cl}','{pct(cl,c)}%','{c}']")
pl_closed_js = ',\n'.join(plc_lines)

# pl-pending rows
pending_by_mod = {}
for r in rows_p:
    m  = str(r[1] or '').strip()
    st = str(r[10] or '').strip()
    if 'Pendiente' in st:
        if m not in pending_by_mod:
            pending_by_mod[m] = {'client': 0, 'sap': 0}
        if st == 'Pendiente Cliente':
            pending_by_mod[m]['client'] += 1
        else:
            pending_by_mod[m]['sap'] += 1

plp_lines = []
for m, v in sorted(pending_by_mod.items(), key=lambda x: -(x[1]['client']+x[1]['sap'])):
    tot = v['client'] + v['sap']
    risk = 'High' if tot >= 10 else ('Medium' if tot >= 3 else 'Low')
    vc = v['client']
    vs = v['sap']
    plp_lines.append(f"      ['{s(m)}','{vc}','{vs}','{tot}','{risk}']")
pl_pending_js = ',\n'.join(plp_lines)

# interfaces rows
iface_lines = []
for r in rows_d:
    t = str(r[12] or '').strip()
    if t in ('Interfaz', 'Enhacement', 'Enhancement'):
        iface_lines.append(f"      ['{s(r[0])}','{s(r[1])}','{s(r[3],60)}','{s(r[12])}','{s(r[7])}']")
iface_modal_js = ',\n'.join(iface_lines)

# ── bar-drill CARD_DATA: deltas by status ─────────────────────────────────────
def js_delta_rows_for(filter_fn):
    lines = []
    for r in rows_d:
        if filter_fn(r):
            lines.append(f"      ['{s(r[0])}','{s(r[1])}','{s(r[3],60)}','{s(r[7])}','{s(r[9])}','{s(r[12])}']")
    return ',\n'.join(lines)

BAR_STATUS_KEYS = {
    'Creado':    ('creado',   'Open (Creado)'),
    'Cerrado':   ('cerrado',  'Closed (Cerrado)'),
    'Aprobado':  ('aprobado', 'Approved (Aprobado)'),
    'Rechazado': ('rechazado','Rejected (Rechazado)'),
}
for raw_st, (slug, label) in BAR_STATUS_KEYS.items():
    pass  # built inline in CARD_DATA template

def cancelado_rows_js():
    lines = []
    for r in rows_d:
        if 'Cancelado' in str(r[9] or ''):
            lines.append(f"      ['{s(r[0])}','{s(r[1])}','{s(r[3],60)}','{s(r[7])}','{s(r[9])}','{s(r[12])}']")
    return ',\n'.join(lines)

def status_rows_js(status_value):
    return js_delta_rows_for(lambda r: str(r[9] or '').strip() == status_value)

# bar-drill: deltas by type
TYPE_GROUPS = [
    ('enh',    'Enhancement',        lambda r: 'nhace' in str(r[12] or '')),
    ('iface',  'Interface',          lambda r: 'Interfaz' in str(r[12] or '')),
    ('newcfg', 'New Configuration',  lambda r: 'Nueva' in str(r[12] or '')),
    ('rep',    'Report',             lambda r: 'Reporte' in str(r[12] or '')),
    ('mig',    'Migration / Other',  lambda r: 'Migra' in str(r[12] or '') or 'Conver' in str(r[12] or '')),
]

def type_rows_js(filter_fn):
    lines = []
    for r in rows_d:
        if filter_fn(r):
            lines.append(f"      ['{s(r[0])}','{s(r[1])}','{s(r[3],60)}','{s(r[9])}','{s(r[12])}']")
    return ',\n'.join(lines)

# bar-drill: deltas by module
def mod_rows_js(module):
    lines = []
    for r in rows_d:
        if str(r[1] or '').strip() == module:
            lines.append(f"      ['{s(r[0])}','{s(r[1])}','{s(r[3],60)}','{s(r[7])}','{s(r[9])}','{s(r[12])}']")
    return ',\n'.join(lines)

# bar-drill: PLs by status
def pl_stat_rows_js(status_value):
    lines = []
    for r in rows_p:
        st = str(r[10] or '').strip()
        match = ('Cerrado' in st) if status_value == '__cerrado__' else (st == status_value)
        if match:
            lines.append(
                "      ['" + s(r[0]) + "','" + s(r[1]) + "','" + s(r[2], 45) +
                "','" + s(r[5], 75) + "','" + s(r[7]) + "','" + s(r[8]) + "','" + s(r[10]) + "']"
            )
    return ',\n'.join(lines)

# bar chart — deltas by module (clickable)
mod_bars_html = ''
for m, c in sorted(mod_d.items(), key=lambda x: -x[1]):
    w   = bpct(c, total_d)
    key = s(m).replace(' ', '_').replace('&amp;', 'AND').replace('&', 'AND')
    mod_bars_html += (
        f'        <div class="bar-row bar-clickable" onclick="openModal(\'bar-mod-{key}\')" title="Click to drill into {s(m)}">'
        f'<span class="bar-label">{s(m)}</span>'
        f'<div class="bar-track"><div class="bar-fill" style="width:{w}%;background:var(--sap-blue);">{c}</div></div>'
        f'<span class="bar-count">{c}</span></div>\n'
    )

# bar chart — PLs by status (clickable)
pl_bars_html = ''
PL_STAT_COLORS = {}
for st, c in sorted(stat_p.items(), key=lambda x: -x[1]):
    w   = bpct(c, total_p)
    col = 'var(--sap-green)' if 'Cerrado' in st else ('var(--sap-red)' if st == 'Pendiente Cliente' else ('var(--sap-orange)' if st == 'Pendiente SAP' else 'var(--sap-grey)'))
    key = s(st).replace(' ', '_').replace('-', '_')
    pl_bars_html += (
        f'        <div class="bar-row bar-clickable" onclick="openModal(\'bar-plstat-{key}\')" title="Click to view {s(st)} items">'
        f'<span class="bar-label">{s(st)}</span>'
        f'<div class="bar-track"><div class="bar-fill" style="width:{w}%;background:{col};">{c}</div></div>'
        f'<span class="bar-count">{c}</span></div>\n'
    )

# open priority table
# open priority table (clickable rows)
open_p12 = [r for r in open_rows_d if str(r[7] or '').startswith(('1', '2'))][:8]
open_rows_html = ''
for r in open_p12:
    ptag = 'tag-red' if str(r[7] or '').startswith('1') else 'tag-gold'
    did  = s(r[0])
    open_rows_html += f'''        <tr class="row-clickable" onclick="openModal('delta-{did}')" title="Click to view full details for {did}">
          <td><strong style="color:var(--sap-blue)">{did}</strong></td>
          <td>{s(r[1])}</td>
          <td>{s(r[3], 60)}</td>
          <td><span class="tag {ptag}">{s(r[7])}</span></td>
          <td>{s(r[12])}</td>
          <td><span class="tag tag-orange">Creado</span></td>
        </tr>\n'''

# pending PLs table (clickable rows)
pending_rows_html = ''
for m, v in sorted(pending_by_mod.items(), key=lambda x: -(x[1]['client']+x[1]['sap'])):
    tot = v['client'] + v['sap']
    tag = 'tag-red' if tot >= 10 else ('tag-orange' if tot >= 3 else 'tag-green')
    imp  = 'Design Impact' if v['client'] >= 5 else 'Config / Other'
    mkey = s(m).replace(' ', '_').replace('&amp;', 'AND').replace('&', 'AND').replace('-', '_')
    pending_rows_html += f'''        <tr class="row-clickable" onclick="openModal('pl-pending-mod-{mkey}')" title="Click to view pending PLs for {m}">
          <td><strong style="color:var(--sap-blue)">{m}</strong></td>
          <td>{v["client"]}</td>
          <td>{v["sap"]}</td>
          <td><span class="tag {tag}">{tot}</span></td>
          <td>{imp}</td>
        </tr>\n'''

# ── delta detail cards (one per open-priority row) ────────────────────────────
def delta_detail_card_js(delta_id):
    """Return JS object literal for a single delta's full detail view."""
    row = next((r for r in rows_d if str(r[0] or '').strip() == delta_id), None)
    if not row:
        return f"{{ title: '{delta_id} — Detail', tableLabel: 'Delta not found', headers: ['Field','Value'], rows: [] }}"
    ptag_js = 'tag-red' if str(row[7] or '').startswith('1') else ('tag-gold' if str(row[7] or '').startswith('2') else 'tag-grey')
    stat_js = str(row[9] or '')
    stat_color = '#BB0000' if stat_js == 'Creado' else ('#107E3E' if stat_js == 'Cerrado' else '#E8A000')
    # Build detail rows as key-value pairs
    fields = [
        ('Delta ID',           s(row[0])),
        ('SAP Module',         s(row[1])),
        ('Priority',           s(row[7])),
        ('Status',             s(row[9])),
        ('Type',               s(row[12])),
        ('Macro Process',      s(row[3])),
        ('Process',            s(row[4])),
        ('Short Description',  s(row[5])),
        ('Detailed Description', s(row[6])),
        ('SAP Alternative',    s(row[15])),
        ('Complexity',         s(row[14])),
        ('Permanent / Temp',   s(row[21])),
        ('Client Owner',       s(row[16])),
        ('SAP Owner',          s(row[17])),
        ('Business Owner',     s(row[18])),
        ('Created Date',       s(row[19])),
        ('Close Date',         s(row[20])),
        ('In SOW (Interfaz)',   s(row[13])),
    ]
    rows_js = ',\n'.join(f"      ['{f}','{v}']" for f, v in fields if v)
    return (
        f"{{ title: 'Delta {s(row[0])} — {s(row[1])} · {s(row[7])}',"
        f" tableLabel: '{s(row[5][:80]) if row[5] else s(row[3])}', "
        f"headers: ['Field','Value'], rows: [\n{rows_js}\n    ] }}"
    )

# ── pending-PL-by-module cards ─────────────────────────────────────────────────
def js_pl_pending_for_mod_key(mkey):
    """Return the set of module strings that map to this mkey."""
    # reverse-map: find the module(s) whose sanitised key matches mkey
    found = []
    for m in pending_by_mod.keys():
        k = s(m).replace(' ', '_').replace('&amp;', 'AND').replace('&', 'AND').replace('-', '_')
        if k == mkey:
            found.append(m)
    return found

def js_pl_pending_rows_for_mod(mods):
    """Full PL rows for a list of module strings, filtered to Pendiente statuses."""
    lines = []
    for r in rows_p:
        m  = str(r[1] or '').strip()
        st = str(r[10] or '').strip()
        if m in mods and 'Pendiente' in st:
            lines.append(
                "      ['" + s(r[0]) + "','" + s(r[1]) + "','" + s(r[2], 45) +
                "','" + s(r[5], 75) + "','" + s(r[6], 80) + "','" + s(r[7]) +
                "','" + s(r[8]) + "','" + s(r[11]) + "','" + s(r[12]) + "','" + s(r[10]) + "']"
            )
    return ',\n'.join(lines)

# pre-build the pending-mod CARD_DATA fragment (avoids chr(39) hacks in f-string)
_pmc_parts = []
for _m, _v in sorted(pending_by_mod.items(), key=lambda x: -(x[1]['client']+x[1]['sap'])):
    _tot = _v['client'] + _v['sap']
    _mkey = s(_m).replace(' ', '_').replace('&amp;', 'AND').replace('&', 'AND').replace('-', '_')
    _mods = js_pl_pending_for_mod_key(_mkey)
    _rows = js_pl_pending_rows_for_mod(_mods)
    _pmc_parts.append(
        f"  'pl-pending-mod-{_mkey}': {{"
        f" title: 'Pending PLs — {s(_m)} ({_tot} open)',"
        f" tableLabel: 'Parking Lots pending resolution for module {s(_m)}',"
        f" headers: ['PL ID','Module','Workstream','Short Desc','Full Description','Priority','Type','SAP Owner','Client Owner','Status'],"
        f" rows: [\n{_rows}\n    ] }}"
    )
_pending_mod_cards_js = ',\n'.join(_pmc_parts)

# ── HTML template ─────────────────────────────────────────────────────────────
html = f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Ternium — SAP S/4HANA Implementation Status Report</title>
<style>
  :root {{
    --sap-blue:#0070F2; --sap-dark:#1D2D3E; --sap-mid:#354A5E; --sap-light:#EBF8FF;
    --sap-gold:#E8A000; --sap-green:#107E3E; --sap-red:#BB0000; --sap-orange:#E76500;
    --sap-grey:#6A6D70; --sap-bg:#F5F6F7; --sap-border:#D9D9D9; --white:#FFFFFF;
    --radius:6px; --shadow:0 2px 8px rgba(0,0,0,0.08);
  }}
  * {{ box-sizing:border-box; margin:0; padding:0; }}
  body {{ font-family:'72','SAP72',-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif; background:var(--sap-bg); color:var(--sap-dark); font-size:14px; line-height:1.5; }}
  .header {{ background:var(--sap-dark); color:var(--white); padding:0 32px; display:flex; align-items:center; justify-content:space-between; height:56px; position:sticky; top:0; z-index:100; box-shadow:0 2px 6px rgba(0,0,0,0.3); }}
  .header-brand {{ display:flex; align-items:center; gap:16px; }}
  .sap-logo {{ background:var(--sap-blue); color:white; font-weight:700; font-size:18px; padding:4px 10px; border-radius:3px; letter-spacing:1px; }}
  .header-title {{ font-size:16px; font-weight:600; opacity:0.95; }}
  .report-date {{ font-size:12px; opacity:0.7; }}
  .hero {{ background:linear-gradient(135deg,var(--sap-blue) 0%,var(--sap-mid) 100%); color:white; padding:32px 32px 28px; }}
  .hero h1 {{ font-size:24px; font-weight:700; margin-bottom:6px; }}
  .hero p {{ font-size:14px; opacity:0.85; max-width:700px; }}
  .hero-badges {{ display:flex; gap:12px; margin-top:16px; flex-wrap:wrap; }}
  .badge {{ background:rgba(255,255,255,0.15); border:1px solid rgba(255,255,255,0.25); border-radius:20px; padding:4px 14px; font-size:12px; font-weight:600; }}
  main {{ padding:24px 32px; max-width:1400px; margin:0 auto; }}
  .section-title {{ font-size:18px; font-weight:700; color:var(--sap-dark); margin:32px 0 16px; display:flex; align-items:center; gap:10px; }}
  .section-title::after {{ content:''; flex:1; height:1px; background:var(--sap-border); }}
  .section-icon {{ width:28px; height:28px; background:var(--sap-blue); border-radius:50%; display:flex; align-items:center; justify-content:center; color:white; font-size:13px; font-weight:700; flex-shrink:0; }}
  .kpi-grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(180px,1fr)); gap:16px; margin-bottom:8px; }}
  .kpi-card {{ background:var(--white); border-radius:var(--radius); box-shadow:var(--shadow); padding:20px 20px 16px; border-top:4px solid var(--sap-blue); transition:transform 0.15s; cursor:pointer; }}
  .kpi-card:hover {{ transform:translateY(-2px); box-shadow:0 4px 16px rgba(0,112,242,0.15); }}
  .kpi-card.green {{ border-top-color:var(--sap-green); }} .kpi-card.red {{ border-top-color:var(--sap-red); }}
  .kpi-card.gold {{ border-top-color:var(--sap-gold); }} .kpi-card.orange {{ border-top-color:var(--sap-orange); }}
  .kpi-value {{ font-size:36px; font-weight:700; line-height:1; color:var(--sap-dark); }}
  .kpi-label {{ font-size:12px; color:var(--sap-grey); margin-top:6px; font-weight:600; text-transform:uppercase; letter-spacing:0.5px; }}
  .kpi-sub {{ font-size:11px; color:var(--sap-grey); margin-top:4px; }}
  .kpi-hint {{ font-size:10px; color:var(--sap-blue); margin-top:8px; opacity:0.7; }}
  .modal-backdrop {{ display:none; position:fixed; inset:0; background:rgba(29,45,62,0.55); z-index:200; align-items:center; justify-content:center; padding:24px; }}
  .modal-backdrop.open {{ display:flex; }}
  .modal-box {{ background:var(--white); border-radius:8px; box-shadow:0 8px 40px rgba(0,0,0,0.22); width:100%; max-width:780px; max-height:88vh; display:flex; flex-direction:column; overflow:hidden; animation:modalIn 0.18s ease; }}
  @keyframes modalIn {{ from {{ opacity:0; transform:translateY(12px) scale(0.98); }} to {{ opacity:1; transform:translateY(0) scale(1); }} }}
  .modal-header {{ padding:16px 20px; display:flex; align-items:center; justify-content:space-between; border-bottom:1px solid var(--sap-border); background:var(--sap-dark); color:var(--white); flex-shrink:0; }}
  .modal-header-title {{ font-size:15px; font-weight:700; }}
  .modal-close {{ background:none; border:none; color:rgba(255,255,255,0.7); font-size:22px; cursor:pointer; line-height:1; padding:0 4px; }}
  .modal-close:hover {{ color:white; }}
  .modal-body {{ padding:20px; overflow-y:auto; flex:1; }}
  .modal-subtitle {{ font-size:12px; font-weight:700; text-transform:uppercase; letter-spacing:0.5px; color:var(--sap-grey); margin-bottom:10px; }}
  .modal-table-wrap {{ overflow-x:auto; border:1px solid var(--sap-border); border-radius:6px; margin-bottom:20px; }}
  .modal-table-wrap table {{ width:100%; }} .modal-table-wrap thead th {{ font-size:11px; }} .modal-table-wrap tbody td {{ font-size:12px; }}
  .notes-label {{ font-size:12px; font-weight:700; text-transform:uppercase; letter-spacing:0.5px; color:var(--sap-grey); margin-bottom:6px; }}
  .notes-area {{ width:100%; height:96px; border:1px solid var(--sap-border); border-radius:6px; padding:10px 12px; font-size:13px; font-family:inherit; resize:vertical; outline:none; transition:border-color 0.15s; color:var(--sap-dark); }}
  .notes-area:focus {{ border-color:var(--sap-blue); }}
  .notes-footer {{ display:flex; align-items:center; justify-content:space-between; margin-top:8px; }}
  .char-counter {{ font-size:11px; color:var(--sap-grey); }} .char-counter.warn {{ color:var(--sap-orange); font-weight:700; }}
  .save-btn {{ background:var(--sap-blue); color:white; border:none; border-radius:4px; padding:7px 20px; font-size:13px; font-weight:600; cursor:pointer; transition:background 0.15s; }}
  .save-btn:hover {{ background:#0058c1; }}
  .save-flash {{ font-size:12px; color:var(--sap-green); font-weight:700; opacity:0; transition:opacity 0.3s; margin-right:12px; }}
  .save-flash.show {{ opacity:1; }}
  .filter-bar {{ display:flex; align-items:center; flex-wrap:wrap; gap:6px; margin-bottom:14px; padding:10px 12px; background:var(--sap-bg); border-radius:6px; border:1px solid var(--sap-border); }}
  .filter-bar-label {{ font-size:11px; font-weight:700; color:var(--sap-grey); text-transform:uppercase; letter-spacing:0.5px; margin-right:4px; }}
  .filter-pill {{ padding:4px 12px; border-radius:12px; border:1px solid var(--sap-border); background:var(--white); font-size:11px; font-weight:600; cursor:pointer; color:var(--sap-dark); transition:all 0.12s; }}
  .filter-pill:hover {{ border-color:var(--sap-blue); color:var(--sap-blue); }} .filter-pill.active {{ background:var(--sap-blue); color:white; border-color:var(--sap-blue); }}
  .pill-count {{ opacity:0.75; font-size:10px; }} .modal-count {{ font-size:11px; color:var(--sap-grey); margin-left:auto; white-space:nowrap; }}
  .phase-timeline {{ display:grid; grid-template-columns:repeat(5,1fr); gap:0; background:var(--white); border-radius:var(--radius); box-shadow:var(--shadow); overflow:hidden; }}
  .phase {{ padding:20px 16px; border-right:1px solid var(--sap-border); }} .phase:last-child {{ border-right:none; }}
  .phase-num {{ width:32px; height:32px; border-radius:50%; display:flex; align-items:center; justify-content:center; font-weight:700; font-size:14px; margin-bottom:10px; }}
  .phase-name {{ font-size:13px; font-weight:700; color:var(--sap-dark); margin-bottom:4px; }}
  .phase-desc {{ font-size:11px; color:var(--sap-grey); margin-bottom:12px; line-height:1.4; }}
  .phase-status-pill {{ display:inline-block; padding:3px 10px; border-radius:12px; font-size:11px; font-weight:700; text-transform:uppercase; }}
  .pill-done {{ background:#E6F4EE; color:var(--sap-green); }} .pill-active {{ background:#FEF3E2; color:var(--sap-orange); }} .pill-pending {{ background:#F5F6F7; color:var(--sap-grey); }}
  .phase.done {{ border-top:4px solid var(--sap-green); }} .phase.active {{ border-top:4px solid var(--sap-gold); }} .phase.pending {{ border-top:4px solid var(--sap-border); }}
  .phase.done .phase-num {{ background:#E6F4EE; color:var(--sap-green); }} .phase.active .phase-num {{ background:#FEF3E2; color:var(--sap-gold); }} .phase.pending .phase-num {{ background:#F5F6F7; color:var(--sap-grey); }}
  .phase-progress {{ margin-top:12px; }}
  .progress-bar-bg {{ height:6px; background:var(--sap-border); border-radius:3px; overflow:hidden; margin-top:4px; }}
  .progress-bar-fill {{ height:100%; border-radius:3px; transition:width 0.6s ease; }}
  .fill-green {{ background:var(--sap-green); }} .fill-gold {{ background:var(--sap-gold); }} .fill-grey {{ background:var(--sap-border); }}
  .phase-pct {{ font-size:11px; color:var(--sap-grey); font-weight:600; }}
  .ws-grid {{ display:grid; grid-template-columns:repeat(auto-fill,minmax(260px,1fr)); gap:16px; }}
  .ws-card {{ background:var(--white); border-radius:var(--radius); cursor:pointer; transition:transform 0.15s,box-shadow 0.15s; box-shadow:var(--shadow); overflow:hidden; }}
  .ws-card:hover {{ transform:translateY(-2px); box-shadow:0 4px 16px rgba(0,112,242,0.13); }}
  .ws-card-header {{ padding:12px 16px; display:flex; align-items:center; justify-content:space-between; border-bottom:1px solid var(--sap-border); }}
  .ws-module {{ font-weight:700; font-size:14px; }} .ws-badge {{ padding:2px 9px; border-radius:10px; font-size:11px; font-weight:700; }}
  .ws-card-body {{ padding:14px 16px; }} .ws-stat-row {{ display:flex; justify-content:space-between; margin-bottom:8px; font-size:12px; }}
  .ws-stat-label {{ color:var(--sap-grey); }} .ws-stat-value {{ font-weight:600; }}
  .ws-mini-bar {{ display:flex; height:8px; border-radius:4px; overflow:hidden; margin-top:10px; gap:1px; }} .ws-mini-seg {{ flex:1; }}
  .charts-row {{ display:grid; grid-template-columns:1fr 1fr; gap:16px; }}
  .chart-card {{ background:var(--white); border-radius:var(--radius); box-shadow:var(--shadow); padding:20px; }}
  .chart-title {{ font-size:14px; font-weight:700; margin-bottom:16px; color:var(--sap-dark); }}
  .bar-chart {{ display:flex; flex-direction:column; gap:10px; }}
  .bar-row {{ display:flex; align-items:center; gap:10px; }}
  .bar-label {{ font-size:12px; width:160px; flex-shrink:0; color:var(--sap-dark); white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
  .bar-track {{ flex:1; height:20px; background:var(--sap-bg); border-radius:3px; overflow:hidden; }}
  .bar-fill {{ height:100%; border-radius:3px; display:flex; align-items:center; padding-left:6px; color:white; font-size:11px; font-weight:600; transition:width 0.6s ease; }}
  .bar-count {{ font-size:12px; font-weight:700; width:30px; text-align:right; flex-shrink:0; }}
  .bar-clickable {{ cursor:pointer; border-radius:4px; padding:2px 4px; margin:-2px -4px; transition:background 0.12s; }}
  .bar-clickable:hover {{ background:var(--sap-light); }}
  .bar-clickable:hover .bar-label {{ color:var(--sap-blue); }}
  .row-clickable {{ cursor:pointer; transition:background 0.12s; }}
  .row-clickable:hover {{ background:var(--sap-light) !important; }}
  .row-clickable td:first-child::after {{ content:' ▸'; font-size:10px; color:var(--sap-blue); opacity:0; transition:opacity 0.12s; }}
  .row-clickable:hover td:first-child::after {{ opacity:1; }}
  .table-card {{ background:var(--white); border-radius:var(--radius); box-shadow:var(--shadow); overflow:hidden; }}
  .table-header {{ padding:16px 20px; display:flex; align-items:center; justify-content:space-between; border-bottom:1px solid var(--sap-border); }}
  .table-title {{ font-size:14px; font-weight:700; }}
  table {{ width:100%; border-collapse:collapse; }}
  thead th {{ background:var(--sap-bg); padding:10px 14px; text-align:left; font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:0.5px; color:var(--sap-grey); border-bottom:1px solid var(--sap-border); }}
  tbody td {{ padding:10px 14px; font-size:13px; border-bottom:1px solid var(--sap-bg); vertical-align:top; }}
  tbody tr:hover {{ background:var(--sap-light); }}
  .tag {{ display:inline-block; padding:2px 8px; border-radius:10px; font-size:11px; font-weight:600; }}
  .tag-green {{ background:#E6F4EE; color:var(--sap-green); }} .tag-red {{ background:#FCEAEA; color:var(--sap-red); }}
  .tag-orange {{ background:#FEF3E2; color:var(--sap-orange); }} .tag-blue {{ background:#E8F3FF; color:var(--sap-blue); }}
  .tag-grey {{ background:#F5F6F7; color:var(--sap-grey); }} .tag-gold {{ background:#FEF7E0; color:var(--sap-gold); }}
  footer {{ margin-top:48px; padding:20px 32px; background:var(--sap-dark); color:rgba(255,255,255,0.5); font-size:12px; text-align:center; }}
  @media print {{ .header {{ position:static; }} .kpi-card:hover {{ transform:none; }} main {{ padding:16px; }} }}
  @media (max-width:900px) {{ .phase-timeline {{ grid-template-columns:1fr; }} .charts-row {{ grid-template-columns:1fr; }} main {{ padding:16px; }} }}
</style>
</head>
<body>

<header class="header">
  <div class="header-brand">
    <span class="sap-logo">SAP</span>
    <span class="header-title">Ternium — S/4HANA Implementation Status Report</span>
  </div>
  <span class="report-date">Report Date: {today}</span>
</header>

<div class="hero">
  <h1>Ternium S/4HANA Implementation</h1>
  <p>FIN · CO · COGM · B&amp;A workstreams — Brazil, Americas, Spain. SAP Activate Methodology — 5-Phase Governance View.</p>
  <div class="hero-badges">
    <span class="badge">SAP Activate</span>
    <span class="badge">FIN · CO · COGM · B&amp;A</span>
    <span class="badge">BR · AMI · ES</span>
    <span class="badge">{total_d} Deltas</span>
    <span class="badge">{total_p} Parking Lots</span>
  </div>
</div>

<main>

  <div class="section-title"><span class="section-icon">◆</span> Executive Summary</div>
  <div class="kpi-grid">
    <div class="kpi-card" data-card="total-deltas" onclick="openModal('total-deltas')">
      <div class="kpi-value">{total_d}</div>
      <div class="kpi-label">Total Deltas / GAPs</div>
      <div class="kpi-sub">Across all modules</div>
      <div class="kpi-hint">▸ Click to view details</div>
    </div>
    <div class="kpi-card red" data-card="open-deltas" onclick="openModal('open-deltas')">
      <div class="kpi-value">{open_d}</div>
      <div class="kpi-label">Open (Creado)</div>
      <div class="kpi-sub">{pct(open_d, total_d)}% of total</div>
      <div class="kpi-hint">▸ Click to view details</div>
    </div>
    <div class="kpi-card green" data-card="closed-deltas" onclick="openModal('closed-deltas')">
      <div class="kpi-value">{closed_d}</div>
      <div class="kpi-label">Closed Deltas</div>
      <div class="kpi-sub">{pct(closed_d, total_d)}% resolved</div>
      <div class="kpi-hint">▸ Click to view details</div>
    </div>
    <div class="kpi-card gold" data-card="mandatory-deltas" onclick="openModal('mandatory-deltas')">
      <div class="kpi-value">{mand_d}</div>
      <div class="kpi-label">Mandatory Deltas</div>
      <div class="kpi-sub">Priority 1 – GoLive blockers</div>
      <div class="kpi-hint">▸ Click to view details</div>
    </div>
    <div class="kpi-card" data-card="total-pl" onclick="openModal('total-pl')">
      <div class="kpi-value">{total_p}</div>
      <div class="kpi-label">Parking Lots</div>
      <div class="kpi-sub">Total raised</div>
      <div class="kpi-hint">▸ Click to view details</div>
    </div>
    <div class="kpi-card green" data-card="pl-closed" onclick="openModal('pl-closed')">
      <div class="kpi-value">{pl_closed}</div>
      <div class="kpi-label">PLs Closed</div>
      <div class="kpi-sub">{pl_clos_pct}% closure rate</div>
      <div class="kpi-hint">▸ Click to view details</div>
    </div>
    <div class="kpi-card orange" data-card="pl-pending" onclick="openModal('pl-pending')">
      <div class="kpi-value">{pl_pending}</div>
      <div class="kpi-label">PLs Pending</div>
      <div class="kpi-sub">{pl_pc} client · {pl_ps} SAP</div>
      <div class="kpi-hint">▸ Click to view details</div>
    </div>
    <div class="kpi-card gold" data-card="interfaces" onclick="openModal('interfaces')">
      <div class="kpi-value">{iface_d + enh_d}</div>
      <div class="kpi-label">Interfaces &amp; Enhancements</div>
      <div class="kpi-sub">{iface_d} interfaces + {enh_d} enhancements</div>
      <div class="kpi-hint">▸ Click to view details</div>
    </div>
  </div>

  <div class="section-title"><span class="section-icon">▶</span> SAP Activate — 5 Phases</div>
  <div class="phase-timeline">
    <div class="phase done">
      <div class="phase-num">1</div><div class="phase-name">Discover</div>
      <div class="phase-desc">Fit-to-standard workshops, initial scoping, stakeholder alignment, system demo.</div>
      <span class="phase-status-pill pill-done">✓ Complete</span>
      <div class="phase-progress"><div class="phase-pct">100%</div><div class="progress-bar-bg"><div class="progress-bar-fill fill-green" style="width:100%"></div></div></div>
    </div>
    <div class="phase done">
      <div class="phase-num">2</div><div class="phase-name">Prepare</div>
      <div class="phase-desc">Project governance, infrastructure setup, team onboarding, sandbox provisioning.</div>
      <span class="phase-status-pill pill-done">✓ Complete</span>
      <div class="phase-progress"><div class="phase-pct">100%</div><div class="progress-bar-bg"><div class="progress-bar-fill fill-green" style="width:100%"></div></div></div>
    </div>
    <div class="phase active">
      <div class="phase-num">3</div><div class="phase-name">Explore</div>
      <div class="phase-desc">Business process workshops, delta/GAP identification, parking lots resolution, backlog build.</div>
      <span class="phase-status-pill pill-active">⚡ In Progress</span>
      <div class="phase-progress"><div class="phase-pct">~72%</div><div class="progress-bar-bg"><div class="progress-bar-fill fill-gold" style="width:72%"></div></div></div>
    </div>
    <div class="phase pending">
      <div class="phase-num">4</div><div class="phase-name">Realize</div>
      <div class="phase-desc">Configuration, development of enhancements &amp; interfaces, unit testing, UAT execution.</div>
      <span class="phase-status-pill pill-pending">Upcoming</span>
      <div class="phase-progress"><div class="phase-pct">0%</div><div class="progress-bar-bg"><div class="progress-bar-fill fill-grey" style="width:0%"></div></div></div>
    </div>
    <div class="phase pending">
      <div class="phase-num">5</div><div class="phase-name">Deploy</div>
      <div class="phase-desc">Cutover execution, data migration, GoLive, hypercare &amp; stabilization.</div>
      <span class="phase-status-pill pill-pending">Upcoming</span>
      <div class="phase-progress"><div class="phase-pct">0%</div><div class="progress-bar-bg"><div class="progress-bar-fill fill-grey" style="width:0%"></div></div></div>
    </div>
  </div>

  <div class="section-title"><span class="section-icon">▤</span> Workstream Status</div>
  <div class="ws-grid">

    <div class="ws-card" onclick="openModal('ws-pp')">
      <div class="ws-card-header" style="border-left:4px solid #0070F2;">
        <span class="ws-module">PP — Production Planning</span>
        <span class="ws-badge tag-blue">{mod_d.get('PP',0)} Deltas</span>
      </div>
      <div class="ws-card-body">
        <div class="ws-stat-row"><span class="ws-stat-label">Interfaces</span><span class="ws-stat-value">High volume</span></div>
        <div class="ws-stat-row"><span class="ws-stat-label">Parking Lots</span><span class="ws-stat-value">{len(pp_rows)} total</span></div>
        <div class="ws-stat-row"><span class="ws-stat-label">Key risk</span><span class="ws-stat-value">NCA MX integrations</span></div>
        <div class="ws-mini-bar">
          <div class="ws-mini-seg" style="background:var(--sap-red);flex:7"></div>
          <div class="ws-mini-seg" style="background:var(--sap-green);flex:2"></div>
          <div class="ws-mini-seg" style="background:var(--sap-grey);flex:2"></div>
        </div>
        <div class="kpi-hint" style="margin-top:10px">▸ Click to view PL details</div>
      </div>
    </div>

    <div class="ws-card" onclick="openModal('ws-ba')">
      <div class="ws-card-header" style="border-left:4px solid #E8A000;">
        <span class="ws-module">B&amp;A — Budget &amp; Analytics</span>
        <span class="ws-badge tag-gold">{mod_d.get('B&A',0)} Deltas</span>
      </div>
      <div class="ws-card-body">
        <div class="ws-stat-row"><span class="ws-stat-label">Enhancements</span><span class="ws-stat-value">Largest workstream</span></div>
        <div class="ws-stat-row"><span class="ws-stat-label">Parking Lots</span><span class="ws-stat-value">{len(ba_rows)} total</span></div>
        <div class="ws-stat-row"><span class="ws-stat-label">Key risk</span><span class="ws-stat-value">SAC + DSP scope</span></div>
        <div class="ws-mini-bar">
          <div class="ws-mini-seg" style="background:var(--sap-red);flex:8"></div>
          <div class="ws-mini-seg" style="background:var(--sap-green);flex:1"></div>
          <div class="ws-mini-seg" style="background:var(--sap-grey);flex:1"></div>
        </div>
        <div class="kpi-hint" style="margin-top:10px">▸ Click to view PL details</div>
      </div>
    </div>

    <div class="ws-card" onclick="openModal('ws-mmim')">
      <div class="ws-card-header" style="border-left:4px solid #107E3E;">
        <span class="ws-module">MM-IM — Inventory Mgmt</span>
        <span class="ws-badge tag-green">{mod_d.get('MM-IM',0)} Deltas</span>
      </div>
      <div class="ws-card-body">
        <div class="ws-stat-row"><span class="ws-stat-label">Parking Lots</span><span class="ws-stat-value">{len(mmim_rows)} total</span></div>
        <div class="ws-stat-row"><span class="ws-stat-label">PLs AR</span><span class="ws-stat-value">{mod_p.get("AR-MM-IM",0)} Argentina scope</span></div>
        <div class="ws-stat-row"><span class="ws-stat-label">Key risk</span><span class="ws-stat-value">Granel NCA movements</span></div>
        <div class="ws-mini-bar">
          <div class="ws-mini-seg" style="background:var(--sap-red);flex:6"></div>
          <div class="ws-mini-seg" style="background:var(--sap-green);flex:5"></div>
          <div class="ws-mini-seg" style="background:var(--sap-grey);flex:4"></div>
        </div>
        <div class="kpi-hint" style="margin-top:10px">▸ Click to view PL details</div>
      </div>
    </div>

    <div class="ws-card" onclick="openModal('ws-co')">
      <div class="ws-card-header" style="border-left:4px solid #BB0000;">
        <span class="ws-module">CO — Controlling</span>
        <span class="ws-badge tag-red">{mod_d.get('CO',0)} Deltas + {mod_d.get('B&A + CO',0)} Cross</span>
      </div>
      <div class="ws-card-body">
        <div class="ws-stat-row"><span class="ws-stat-label">Parking Lots</span><span class="ws-stat-value">{len(co_rows)} total (highest)</span></div>
        <div class="ws-stat-row"><span class="ws-stat-label">PLs AR</span><span class="ws-stat-value">{mod_p.get("AR-CO",0)} Argentina scope</span></div>
        <div class="ws-stat-row"><span class="ws-stat-label">Key risk</span><span class="ws-stat-value">Standard cost + COGM</span></div>
        <div class="ws-mini-bar">
          <div class="ws-mini-seg" style="background:var(--sap-red);flex:6"></div>
          <div class="ws-mini-seg" style="background:var(--sap-green);flex:3"></div>
          <div class="ws-mini-seg" style="background:var(--sap-grey);flex:3"></div>
        </div>
        <div class="kpi-hint" style="margin-top:10px">▸ Click to view PL details</div>
      </div>
    </div>

    <div class="ws-card" onclick="openModal('ws-mmpur')">
      <div class="ws-card-header" style="border-left:4px solid #6A6D70;">
        <span class="ws-module">MM-PUR — Purchasing</span>
        <span class="ws-badge tag-grey">{mod_d.get('MM-PUR',0)} Deltas</span>
      </div>
      <div class="ws-card-body">
        <div class="ws-stat-row"><span class="ws-stat-label">Parking Lots</span><span class="ws-stat-value">{len(mmpur_rows)} total</span></div>
        <div class="ws-stat-row"><span class="ws-stat-label">PLs AR</span><span class="ws-stat-value">{mod_p.get("AR-MM-PUR",0)} Argentina scope</span></div>
        <div class="ws-stat-row"><span class="ws-stat-label">Status</span><span class="ws-stat-value">Progressing</span></div>
        <div class="ws-mini-bar">
          <div class="ws-mini-seg" style="background:var(--sap-red);flex:4"></div>
          <div class="ws-mini-seg" style="background:var(--sap-green);flex:4"></div>
          <div class="ws-mini-seg" style="background:var(--sap-grey);flex:2"></div>
        </div>
        <div class="kpi-hint" style="margin-top:10px">▸ Click to view PL details</div>
      </div>
    </div>

    <div class="ws-card" onclick="openModal('ws-sd')">
      <div class="ws-card-header" style="border-left:4px solid #E76500;">
        <span class="ws-module">SD — Sales &amp; Distribution</span>
        <span class="ws-badge tag-orange">{mod_d.get('SD',0)} Deltas</span>
      </div>
      <div class="ws-card-body">
        <div class="ws-stat-row"><span class="ws-stat-label">Parking Lots</span><span class="ws-stat-value">{len(sd_rows)} total</span></div>
        <div class="ws-stat-row"><span class="ws-stat-label">PLs AR</span><span class="ws-stat-value">{mod_p.get("AR-SD",0)} Argentina scope</span></div>
        <div class="ws-stat-row"><span class="ws-stat-label">Status</span><span class="ws-stat-value">Low delta volume</span></div>
        <div class="ws-mini-bar">
          <div class="ws-mini-seg" style="background:var(--sap-green);flex:8"></div>
          <div class="ws-mini-seg" style="background:var(--sap-red);flex:2"></div>
        </div>
        <div class="kpi-hint" style="margin-top:10px">▸ Click to view PL details</div>
      </div>
    </div>

    <div class="ws-card" onclick="openModal('ws-matcluster')">
      <div class="ws-card-header" style="border-left:4px solid #5C2D91;">
        <span class="ws-module">Material Cluster — Situación Actual</span>
        <span class="ws-badge" style="background:#EDE7F6;color:#5C2D91;">5 Countries</span>
      </div>
      <div class="ws-card-body">
        <div class="ws-stat-row"><span class="ws-stat-label">Template Global</span><span class="ws-stat-value">MTS — Usiminas model</span></div>
        <div class="ws-stat-row"><span class="ws-stat-label">Exceptions</span><span class="ws-stat-value">Centro América · Colombia</span></div>
        <div class="ws-stat-row"><span class="ws-stat-label">Key issue</span><span class="ws-stat-value">Prod. vendido ≠ costeado</span></div>
        <div class="ws-mini-bar">
          <div class="ws-mini-seg" style="background:#5C2D91;flex:3"></div>
          <div class="ws-mini-seg" style="background:var(--sap-orange);flex:2"></div>
          <div class="ws-mini-seg" style="background:var(--sap-green);flex:2"></div>
        </div>
        <div class="kpi-hint" style="margin-top:10px">▸ Click to view country details</div>
      </div>
    </div>

  </div>

  <div class="section-title"><span class="section-icon">▣</span> Delta / GAP Analysis</div>
  <div class="charts-row">
    <div class="chart-card">
      <div class="chart-title">Deltas by Status <span style="font-size:10px;color:var(--sap-blue);font-weight:400;margin-left:6px;">▸ click a bar to drill in</span></div>
      <div class="bar-chart">
        <div class="bar-row bar-clickable" onclick="openModal('bar-status-Creado')" title="Click to view open deltas"><span class="bar-label">Creado (Open)</span><div class="bar-track"><div class="bar-fill" style="width:{bpct(open_d,total_d)}%;background:var(--sap-red);">{open_d}</div></div><span class="bar-count">{open_d}</span></div>
        <div class="bar-row bar-clickable" onclick="openModal('bar-status-Cerrado')" title="Click to view closed deltas"><span class="bar-label">Cerrado (Closed)</span><div class="bar-track"><div class="bar-fill" style="width:{bpct(closed_d,total_d)}%;background:var(--sap-green);">{closed_d}</div></div><span class="bar-count">{closed_d}</span></div>
        <div class="bar-row bar-clickable" onclick="openModal('bar-status-Aprobado')" title="Click to view approved deltas"><span class="bar-label">Aprobado</span><div class="bar-track"><div class="bar-fill" style="width:{bpct(aprov_d,total_d)}%;background:var(--sap-blue);">{aprov_d}</div></div><span class="bar-count">{aprov_d}</span></div>
        <div class="bar-row bar-clickable" onclick="openModal('bar-status-Cancelado')" title="Click to view cancelled deltas"><span class="bar-label">Cancelado</span><div class="bar-track"><div class="bar-fill" style="width:{bpct(cancel_d,total_d)}%;background:var(--sap-grey);">{cancel_d}</div></div><span class="bar-count">{cancel_d}</span></div>
        <div class="bar-row bar-clickable" onclick="openModal('bar-status-Rechazado')" title="Click to view rejected deltas"><span class="bar-label">Rechazado</span><div class="bar-track"><div class="bar-fill" style="width:{bpct(reject_d,total_d)}%;background:var(--sap-orange);">{reject_d}</div></div><span class="bar-count">{reject_d}</span></div>
      </div>
    </div>
    <div class="chart-card">
      <div class="chart-title">Deltas by Type <span style="font-size:10px;color:var(--sap-blue);font-weight:400;margin-left:6px;">▸ click a bar to drill in</span></div>
      <div class="bar-chart">
        <div class="bar-row bar-clickable" onclick="openModal('bar-type-enh')" title="Click to view Enhancement deltas"><span class="bar-label">Enhancement</span><div class="bar-track"><div class="bar-fill" style="width:{bpct(enh_d,total_d)}%;background:var(--sap-blue);">{enh_d}</div></div><span class="bar-count">{enh_d}</span></div>
        <div class="bar-row bar-clickable" onclick="openModal('bar-type-iface')" title="Click to view Interface deltas"><span class="bar-label">Interface</span><div class="bar-track"><div class="bar-fill" style="width:{bpct(iface_d,total_d)}%;background:var(--sap-orange);">{iface_d}</div></div><span class="bar-count">{iface_d}</span></div>
        <div class="bar-row bar-clickable" onclick="openModal('bar-type-newcfg')" title="Click to view New Configuration deltas"><span class="bar-label">New Configuration</span><div class="bar-track"><div class="bar-fill" style="width:{bpct(newcfg_d,total_d)}%;background:var(--sap-gold);">{newcfg_d}</div></div><span class="bar-count">{newcfg_d}</span></div>
        <div class="bar-row bar-clickable" onclick="openModal('bar-type-rep')" title="Click to view Report deltas"><span class="bar-label">Report</span><div class="bar-track"><div class="bar-fill" style="width:{bpct(rep_d,total_d)}%;background:var(--sap-mid);">{rep_d}</div></div><span class="bar-count">{rep_d}</span></div>
        <div class="bar-row bar-clickable" onclick="openModal('bar-type-mig')" title="Click to view Migration / Other deltas"><span class="bar-label">Migration / Other</span><div class="bar-track"><div class="bar-fill" style="width:{bpct(mig_d,total_d)}%;background:var(--sap-grey);">{mig_d}</div></div><span class="bar-count">{mig_d}</span></div>
      </div>
    </div>
    <div class="chart-card">
      <div class="chart-title">Deltas by Module <span style="font-size:10px;color:var(--sap-blue);font-weight:400;margin-left:6px;">▸ click a bar to drill in</span></div>
      <div class="bar-chart">
{mod_bars_html}      </div>
    </div>
    <div class="chart-card">
      <div class="chart-title">Parking Lots by Status <span style="font-size:10px;color:var(--sap-blue);font-weight:400;margin-left:6px;">▸ click a bar to drill in</span></div>
      <div class="bar-chart">
{pl_bars_html}      </div>
    </div>
  </div>

  <div class="section-title"><span class="section-icon">⚠</span> Open Priority Items — GoLive Risk</div>
  <div class="table-card">
    <div class="table-header">
      <span class="table-title">Priority 1 &amp; 2 Open Deltas</span>
      <span style="display:flex;align-items:center;gap:12px">
        <span style="font-size:10px;color:var(--sap-blue);">▸ click a row to view full detail</span>
        <span class="tag tag-red">{len(open_p12)} items shown — resolve before Realize phase</span>
      </span>
    </div>
    <table>
      <thead><tr><th>Delta ID</th><th>Module</th><th>Description</th><th>Priority</th><th>Type</th><th>Status</th></tr></thead>
      <tbody>
{open_rows_html}      </tbody>
    </table>
  </div>

  <div class="section-title" style="margin-top:32px"><span class="section-icon">⏳</span> Parking Lots — Pending Resolution</div>
  <div class="table-card">
    <div class="table-header">
      <span class="table-title">Items Awaiting Client or SAP Response</span>
      <span style="display:flex;align-items:center;gap:12px">
        <span style="font-size:10px;color:var(--sap-blue);">▸ click a row to view pending PL list</span>
        <span class="tag tag-orange">{pl_pending} open items</span>
      </span>
    </div>
    <table>
      <thead><tr><th>Module</th><th>Pending Client</th><th>Pending SAP</th><th>Total Open</th><th>Priority Type</th></tr></thead>
      <tbody>
{pending_rows_html}      </tbody>
    </table>
  </div>

  <div class="section-title" style="margin-top:32px"><span class="section-icon">!</span> Key Risks &amp; Recommendations</div>
  <div class="charts-row">
    <div class="chart-card">
      <div class="chart-title">Top Risks</div>
      <table style="font-size:12px">
        <thead><tr><th>Risk</th><th>Impact</th><th>Owner</th></tr></thead>
        <tbody>
          <tr><td>{open_d} open deltas entering Realize with no approved solution</td><td><span class="tag tag-red">High</span></td><td>PMO</td></tr>
          <tr><td>NCA MX interfaces (mandatory PP deltas) undefined</td><td><span class="tag tag-red">High</span></td><td>PP WS Lead</td></tr>
          <tr><td>{pl_pc} PLs pending client response — blocking configuration</td><td><span class="tag tag-orange">Medium</span></td><td>Client SMEs</td></tr>
          <tr><td>B&amp;A SAC scope creep (enhancements out of SOW)</td><td><span class="tag tag-orange">Medium</span></td><td>B&amp;A Lead</td></tr>
          <tr><td>CO PLs highest volume ({len(co_rows)}) — standard cost design risk</td><td><span class="tag tag-orange">Medium</span></td><td>CO WS Lead</td></tr>
        </tbody>
      </table>
    </div>
    <div class="chart-card">
      <div class="chart-title">Recommendations for Next Sprint</div>
      <table style="font-size:12px">
        <thead><tr><th>#</th><th>Action</th><th>Phase</th></tr></thead>
        <tbody>
          <tr><td>1</td><td>Prioritize resolution of {mand_d} Priority-1 deltas before Realize entry</td><td>Explore → Realize</td></tr>
          <tr><td>2</td><td>Schedule interface design sessions for all {iface_d} Interface-type deltas</td><td>Explore</td></tr>
          <tr><td>3</td><td>Escalate {pl_pc} client-pending PLs in next Steering Committee</td><td>Governance</td></tr>
          <tr><td>4</td><td>Conduct GoLive strategy session for Permanent vs Temporary GAPs</td><td>Explore</td></tr>
          <tr><td>5</td><td>Baseline project plan in MS Project with Realize milestones</td><td>Prepare → Realize</td></tr>
        </tbody>
      </table>
    </div>
  </div>

</main>

<footer>
  Ternium S/4HANA Implementation Status Report · Auto-generated from source xlsx · {today} · SAP Global Delivery Hub Americas
</footer>

<!-- MODAL -->
<div class="modal-backdrop" id="modal-backdrop" onclick="handleBackdropClick(event)">
  <div class="modal-box" id="modal-box">
    <div class="modal-header">
      <span class="modal-header-title" id="modal-title">Details</span>
      <button class="modal-close" onclick="closeModal()">&#10005;</button>
    </div>
    <div class="modal-body">
      <div class="modal-subtitle" id="modal-table-label">Related Items</div>
      <div id="modal-filter-bar" class="filter-bar" style="display:none">
        <span class="filter-bar-label">Filter by Status:</span>
        <span id="modal-filter-pills"></span>
        <span class="modal-count" id="modal-count"></span>
      </div>
      <div class="modal-table-wrap">
        <table><thead id="modal-thead"></thead><tbody id="modal-tbody"></tbody></table>
      </div>
      <div class="notes-label">Notes</div>
      <textarea class="notes-area" id="modal-notes" maxlength="500" placeholder="Add your notes here (max 500 characters)&#8230;"></textarea>
      <div class="notes-footer">
        <span class="char-counter" id="char-counter">500 characters remaining</span>
        <div style="display:flex;align-items:center">
          <span class="save-flash" id="save-flash">Saved &#10003;</span>
          <button class="save-btn" onclick="saveNote()">Save Note</button>
        </div>
      </div>
    </div>
  </div>
</div>

<script>
const CARD_DATA = {{
  'total-deltas': {{
    title: 'Total Deltas / GAPs — {total_d} Items',
    tableLabel: 'All Deltas by Module',
    headers: ['Module', 'Total Deltas', 'Open', 'Top Type'],
    rows: [
{total_deltas_js}
    ]
  }},
  'open-deltas': {{
    title: 'Open Deltas (Creado) — {open_d} Items',
    tableLabel: 'Open Deltas requiring resolution',
    headers: ['Delta ID', 'Module', 'Description', 'Priority', 'Type'],
    rows: [
{open_modal_js}
    ]
  }},
  'closed-deltas': {{
    title: 'Closed Deltas — {closed_d} Items',
    tableLabel: 'Resolved and closed deltas',
    headers: ['Module', 'Count Closed', 'Status', 'Resolution Type'],
    rows: [
{closed_modal_js}
    ]
  }},
  'mandatory-deltas': {{
    title: 'Mandatory Deltas (Priority 1) — {mand_d} GoLive Blockers',
    tableLabel: 'Priority 1 – must be resolved before GoLive',
    headers: ['Delta ID', 'Module', 'Description', 'Status', 'Type'],
    rows: [
{mand_modal_js}
    ]
  }},
  'total-pl': {{
    title: 'Parking Lots — {total_p} Total',
    tableLabel: 'Distribution across workstreams',
    headers: ['Module', 'Total PLs', 'Closed', 'Pending Client', 'Pending SAP'],
    rows: [
{total_pl_js}
    ]
  }},
  'pl-closed': {{
    title: 'Parking Lots Closed — {pl_closed} ({pl_clos_pct}%)',
    tableLabel: 'Closed parking lots by module',
    headers: ['Module', 'Closed', '% of Module', 'Total PLs'],
    rows: [
{pl_closed_js}
    ]
  }},
  'pl-pending': {{
    title: 'Parking Lots Pending — {pl_pending} Open',
    tableLabel: 'Items awaiting response — escalation required',
    headers: ['Module', 'Pending Client', 'Pending SAP', 'Total', 'Risk'],
    rows: [
{pl_pending_js}
    ]
  }},
  'interfaces': {{
    title: 'Interfaces &amp; Enhancements — {iface_d + enh_d} Items',
    tableLabel: '{iface_d} Interface-type + {enh_d} Enhancement-type deltas',
    headers: ['Delta ID', 'Module', 'Description', 'Type', 'Priority'],
    rows: [
{iface_modal_js}
    ]
  }},
  'ws-pp': {{
    title: 'PP — Production Planning · Parking Lots',
    tableLabel: 'All {len(pp_rows)} Parking Lots for this workstream',
    headers: ['PL ID', 'Module', 'Workstream', 'Description', 'Priority', 'Type', 'Status'],
    rows: [
{js_pl_rows(pp_rows)}
    ]
  }},
  'ws-ba': {{
    title: 'B&amp;A — Budget &amp; Analytics · Parking Lots',
    tableLabel: 'All {len(ba_rows)} Parking Lots for this workstream',
    headers: ['PL ID', 'Module', 'Workstream', 'Description', 'Priority', 'Type', 'Status'],
    rows: [
{js_pl_rows(ba_rows)}
    ]
  }},
  'ws-mmim': {{
    title: 'MM-IM — Inventory Mgmt · Parking Lots',
    tableLabel: 'All {len(mmim_rows)} Parking Lots for this workstream',
    headers: ['PL ID', 'Module', 'Workstream', 'Description', 'Priority', 'Type', 'Status'],
    rows: [
{js_pl_rows(mmim_rows)}
    ]
  }},
  'ws-co': {{
    title: 'CO — Controlling · Parking Lots',
    tableLabel: 'All {len(co_rows)} Parking Lots for this workstream',
    headers: ['PL ID', 'Module', 'Workstream', 'Description', 'Priority', 'Type', 'Status'],
    rows: [
{js_pl_rows(co_rows)}
    ]
  }},
  'ws-mmpur': {{
    title: 'MM-PUR — Purchasing · Parking Lots',
    tableLabel: 'All {len(mmpur_rows)} Parking Lots for this workstream',
    headers: ['PL ID', 'Module', 'Workstream', 'Description', 'Priority', 'Type', 'Status'],
    rows: [
{js_pl_rows(mmpur_rows)}
    ]
  }},
  'ws-sd': {{
    title: 'SD — Sales &amp; Distribution · Parking Lots',
    tableLabel: 'All {len(sd_rows)} Parking Lots for this workstream',
    headers: ['PL ID', 'Module', 'Workstream', 'Description', 'Priority', 'Type', 'Status'],
    rows: [
{js_pl_rows(sd_rows)}
    ]
  }},
  'ws-matcluster': {{
    title: 'Material Cluster — Situación Actual por País',
    tableLabel: 'Estado de implementación del template global de Material Cluster',
    headers: ['País / Región', 'Situación Actual', 'Template Global', 'Solución Usiminas', 'Next Step'],
    rows: [
{js_mc_rows()}
    ]
  }},

  // ── bar-drill: Deltas by Status ───────────────────────────────────────────
  'bar-status-Creado': {{
    title: 'Open Deltas — Creado ({open_d} items)',
    tableLabel: 'All deltas with status Creado (awaiting resolution)',
    headers: ['Delta ID', 'Module', 'Description', 'Priority', 'Status', 'Type'],
    rows: [
{status_rows_js('Creado')}
    ]
  }},
  'bar-status-Cerrado': {{
    title: 'Closed Deltas — Cerrado ({closed_d} items)',
    tableLabel: 'All deltas with status Cerrado (resolved)',
    headers: ['Delta ID', 'Module', 'Description', 'Priority', 'Status', 'Type'],
    rows: [
{status_rows_js('Cerrado')}
    ]
  }},
  'bar-status-Aprobado': {{
    title: 'Approved Deltas — Aprobado ({aprov_d} items)',
    tableLabel: 'All deltas with status Aprobado',
    headers: ['Delta ID', 'Module', 'Description', 'Priority', 'Status', 'Type'],
    rows: [
{status_rows_js('Aprobado')}
    ]
  }},
  'bar-status-Cancelado': {{
    title: 'Cancelled Deltas — Cancelado ({cancel_d} items)',
    tableLabel: 'All deltas with status Cancelado',
    headers: ['Delta ID', 'Module', 'Description', 'Priority', 'Status', 'Type'],
    rows: [
{cancelado_rows_js()}
    ]
  }},
  'bar-status-Rechazado': {{
    title: 'Rejected Deltas — Rechazado ({reject_d} items)',
    tableLabel: 'All deltas with status Rechazado',
    headers: ['Delta ID', 'Module', 'Description', 'Priority', 'Status', 'Type'],
    rows: [
{status_rows_js('Rechazado')}
    ]
  }},

  // ── bar-drill: Deltas by Type ─────────────────────────────────────────────
  'bar-type-enh': {{
    title: 'Enhancement Deltas ({enh_d} items)',
    tableLabel: 'All Enhancement-type deltas',
    headers: ['Delta ID', 'Module', 'Description', 'Status', 'Type'],
    rows: [
{type_rows_js(lambda r: 'nhace' in str(r[12] or ''))}
    ]
  }},
  'bar-type-iface': {{
    title: 'Interface Deltas ({iface_d} items)',
    tableLabel: 'All Interface-type deltas',
    headers: ['Delta ID', 'Module', 'Description', 'Status', 'Type'],
    rows: [
{type_rows_js(lambda r: 'Interfaz' in str(r[12] or ''))}
    ]
  }},
  'bar-type-newcfg': {{
    title: 'New Configuration Deltas ({newcfg_d} items)',
    tableLabel: 'All Nueva Configuración-type deltas',
    headers: ['Delta ID', 'Module', 'Description', 'Status', 'Type'],
    rows: [
{type_rows_js(lambda r: 'Nueva' in str(r[12] or ''))}
    ]
  }},
  'bar-type-rep': {{
    title: 'Report Deltas ({rep_d} items)',
    tableLabel: 'All Reporte-type deltas',
    headers: ['Delta ID', 'Module', 'Description', 'Status', 'Type'],
    rows: [
{type_rows_js(lambda r: 'Reporte' in str(r[12] or ''))}
    ]
  }},
  'bar-type-mig': {{
    title: 'Migration / Other Deltas ({mig_d} items)',
    tableLabel: 'All Conversion/Migración-type deltas',
    headers: ['Delta ID', 'Module', 'Description', 'Status', 'Type'],
    rows: [
{type_rows_js(lambda r: 'Migra' in str(r[12] or '') or 'Conver' in str(r[12] or ''))}
    ]
  }},

  // ── bar-drill: Deltas by Module (dynamic, one entry per module) ────────────
  {', '.join(f"'bar-mod-{s(m).replace(' ','_').replace('&amp;','AND').replace('&','AND')}': {{ title: '{s(m)} Module — {c} Deltas', tableLabel: 'All {c} deltas for module {s(m)}', headers: ['Delta ID','Module','Description','Priority','Status','Type'], rows: [{chr(10)}{mod_rows_js(m)}{chr(10)}    ] }}" for m,c in sorted(mod_d.items(), key=lambda x:-x[1]))},

  // ── bar-drill: Parking Lots by Status ─────────────────────────────────────
  {', '.join(f"'bar-plstat-{s(st).replace(' ','_').replace('-','_')}': {{ title: 'Parking Lots — {s(st)} ({c} items)', tableLabel: 'All PL items with status {s(st)}', headers: ['PL ID','Module','Workstream','Description','Priority','Type','Status'], rows: [{chr(10)}{pl_stat_rows_js('__cerrado__' if 'Cerrado' in st else st)}{chr(10)}    ] }}" for st,c in sorted(stat_p.items(), key=lambda x:-x[1]))},

  // ── Open Priority table: per-delta detail cards ───────────────────────────
  {', '.join(f"'delta-{s(r[0])}': {delta_detail_card_js(str(r[0] or '').strip())}" for r in open_p12 if r[0])},

  // ── Pending PL table: per-module pending detail cards ─────────────────────
  {_pending_mod_cards_js}
}};

const WS_CARDS = new Set(['ws-pp','ws-ba','ws-mmim','ws-co','ws-mmpur','ws-sd']);
const STATUS_COL = 6;

function barDrillStatusCol(cardId) {{
  // returns the column index that holds Status for filter-pills, or -1 if no pills
  if (WS_CARDS.has(cardId)) return 6;              // PL rows: col 6 = Status
  if (cardId.startsWith('bar-mod-')) return 4;      // delta rows: col 4 = Status
  if (cardId.startsWith('bar-type-')) return 3;     // delta rows: col 3 = Status
  if (cardId.startsWith('bar-plstat-')) return 6;   // PL rows: col 6 = Status
  if (cardId.startsWith('pl-pending-mod-')) return 9; // pending PL: col 9 = Status
  return -1;                                        // delta detail / no filter
}}
let currentCardId = null;
let currentRows = [];
let activeFilter = 'all';

function openModal(cardId) {{
  const data = CARD_DATA[cardId];
  if (!data) return;
  currentCardId = cardId;
  currentRows = data.rows;
  activeFilter = 'all';
  // wider modal for detail views and pending-PL tables
  const isDetail = cardId.startsWith('delta-') || cardId.startsWith('pl-pending-mod-');
  document.getElementById('modal-box').style.maxWidth = isDetail ? '960px' : '780px';
  document.getElementById('modal-title').textContent = cardId.startsWith('delta-')
    ? data.title
    : data.title + ' (' + data.rows.length + ' items)';
  document.getElementById('modal-table-label').textContent = data.tableLabel;
  const thead = document.getElementById('modal-thead');
  thead.innerHTML = '<tr>' + data.headers.map(h => '<th>' + h + '</th>').join('') + '</tr>';
  const filterBar = document.getElementById('modal-filter-bar');
  const stCol = barDrillStatusCol(cardId);
  if (stCol >= 0 && data.rows.length > 0) {{
    const statuses = [...new Set(data.rows.map(r => r[stCol]).filter(Boolean))].sort();
    const counts = {{}};
    statuses.forEach(s => {{ counts[s] = data.rows.filter(r => r[stCol] === s).length; }});
    const pillsEl = document.getElementById('modal-filter-pills');
    pillsEl.innerHTML =
      "<button class=\\"filter-pill active\\" data-status=\\"all\\" onclick=\\"setFilter('all')\\">All <span class=\\"pill-count\\">" + data.rows.length + "</span></button>" +
      statuses.map(s =>
        "<button class=\\"filter-pill\\" data-status=\\"" + s.replace(/"/g,'&quot;') + "\\" onclick=\\"setFilter(this.dataset.status)\\">" + s + " <span class=\\"pill-count\\">" + counts[s] + "</span></button>"
      ).join('');
    filterBar.style.display = 'flex';
  }} else {{
    filterBar.style.display = 'none';
  }}
  renderFilteredTable();
  const saved = localStorage.getItem('note_' + cardId) || '';
  const ta = document.getElementById('modal-notes');
  ta.value = saved;
  updateCounter(saved.length);
  document.getElementById('save-flash').classList.remove('show');
  document.getElementById('modal-backdrop').classList.add('open');
  document.body.style.overflow = 'hidden';
}}

function renderFilteredTable() {{
  const tbody = document.getElementById('modal-tbody');
  const countEl = document.getElementById('modal-count');
  const stCol = barDrillStatusCol(currentCardId);

  // Delta detail: key-value card — render with label column styled
  if (currentCardId && currentCardId.startsWith('delta-')) {{
    tbody.innerHTML = currentRows.map(row =>
      '<tr><td style="font-weight:700;width:160px;color:var(--sap-grey);white-space:nowrap">' + row[0] +
      '</td><td style="white-space:pre-wrap;word-break:break-word">' + row[1] + '</td></tr>'
    ).join('');
    if (countEl) countEl.textContent = '';
    return;
  }}

  if (activeFilter === 'all') {{
    if (stCol >= 0) {{
      const groups = {{}};
      currentRows.forEach(r => {{
        const st = r[stCol] || '—';
        if (!groups[st]) groups[st] = [];
        groups[st].push(r);
      }});
      const sorted = Object.keys(groups).sort();
      let html = '';
      sorted.forEach(st => {{
        const c = statusColor(st);
        const colspan = currentRows[0] ? currentRows[0].length : 7;
        html += '<tr><td colspan="' + colspan + '" style="background:' + c.bg + ';color:' + c.fg + ';font-weight:700;font-size:11px;text-transform:uppercase;letter-spacing:0.5px;padding:6px 14px;">' + st + ' <span style="font-weight:400;opacity:0.8">(' + groups[st].length + ')</span></td></tr>';
        groups[st].forEach(row => {{ html += '<tr>' + row.map(cell => '<td>' + cell + '</td>').join('') + '</tr>'; }});
      }});
      tbody.innerHTML = html;
    }} else {{
      tbody.innerHTML = currentRows.map(row => '<tr>' + row.map(cell => '<td>' + cell + '</td>').join('') + '</tr>').join('');
    }}
    if (countEl) countEl.textContent = currentRows.length + ' items';
  }} else {{
    const filtered = stCol >= 0 ? currentRows.filter(r => r[stCol] === activeFilter) : currentRows;
    tbody.innerHTML = filtered.map(row => '<tr>' + row.map(cell => '<td>' + cell + '</td>').join('') + '</tr>').join('');
    if (countEl) countEl.textContent = filtered.length + ' of ' + currentRows.length + ' items';
  }}
}}

function statusColor(st) {{
  if (st.startsWith('Cerrado')) return {{bg:'#E6F4EE',fg:'#107E3E'}};
  if (st === 'Pendiente Cliente') return {{bg:'#FCEAEA',fg:'#BB0000'}};
  if (st === 'Pendiente SAP') return {{bg:'#FEF3E2',fg:'#E76500'}};
  if (st === 'Aprobado') return {{bg:'#E8F3FF',fg:'#0070F2'}};
  return {{bg:'#F5F6F7',fg:'#6A6D70'}};
}}

function setFilter(status) {{
  activeFilter = status;
  document.querySelectorAll('#modal-filter-pills .filter-pill').forEach(btn => {{
    btn.classList.toggle('active', btn.dataset.status === status);
  }});
  renderFilteredTable();
}}

function closeModal() {{
  document.getElementById('modal-backdrop').classList.remove('open');
  document.body.style.overflow = '';
  currentCardId = null;
}}

function handleBackdropClick(e) {{
  if (e.target === document.getElementById('modal-backdrop')) closeModal();
}}

function updateCounter(len) {{
  const remaining = 500 - len;
  const el = document.getElementById('char-counter');
  el.textContent = remaining + ' character' + (remaining === 1 ? '' : 's') + ' remaining';
  el.classList.toggle('warn', remaining < 50);
}}

document.getElementById('modal-notes').addEventListener('input', function() {{
  updateCounter(this.value.length);
}});

function saveNote() {{
  if (!currentCardId) return;
  localStorage.setItem('note_' + currentCardId, document.getElementById('modal-notes').value);
  const flash = document.getElementById('save-flash');
  flash.classList.add('show');
  setTimeout(() => flash.classList.remove('show'), 2000);
}}

document.addEventListener('keydown', function(e) {{
  if (e.key === 'Escape') closeModal();
}});
</script>
</body>
</html>'''

with open(OUT, 'w', encoding='utf-8') as f:
    f.write(html)

print(f'✓ index.html rebuilt — {today}')
print(f'  Deltas: {total_d} total | {open_d} open | {closed_d} closed | {mand_d} mandatory')
print(f'  PLs:    {total_p} total | {pl_closed} closed ({pl_clos_pct}%) | {pl_pending} pending')
