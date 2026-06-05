"""
Viaero MW Construction — Deck Builder
Ariadne Integration: Called by the Node.js API route via HTTP
Runs on DigitalOcean droplet alongside Theseus

Input (multipart form or JSON paths):
  - tracker_path:       path to uploaded Viaero tracker .xlsx
  - previous_deck_path: path to previous session .pptx
  - snapshot_path:      path to previous session snapshot .json
  - ntp_comments_path:  path to NTP comments .xlsx (optional)
  - deck_date:          string e.g. "6/4/2026"

Output (JSON):
  - deck_path:          path to generated .pptx
  - snapshot_path:      path to generated snapshot .json
  - ntp_comments_path:  path to generated NTP comments .xlsx
  - summary:            dict of key stats for DB logging if needed
"""

import sys
import json
import os
import zipfile
import re
import io
import copy
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.dml.color import RGBColor
from lxml import etree


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

def gv(r, c):
    v = r.get(c, '')
    if pd.isna(v): return ''
    s = str(v).strip()
    return '' if s.lower() == 'nan' else s

def fmt_d(v):
    if pd.isna(v) or v is None: return ''
    try: return pd.Timestamp(v).strftime('%m/%d/%Y')
    except: return ''

def fmt_dm(v):
    if pd.isna(v) or v is None: return ''
    try: return pd.Timestamp(v).strftime('%m/%d')
    except: return ''

def fmt_ds(v):
    if pd.isna(v) or v is None: return ''
    try: return pd.Timestamp(v).strftime('%m/%d/%y')
    except: return ''

def set_shape_text(shape, text, para_idx=0, run_idx=0):
    if not shape.has_text_frame: return
    try:
        para = shape.text_frame.paragraphs[para_idx]
        if para.runs:
            para.runs[run_idx].text = str(text)
        else:
            r = para._p.find(qn('a:r'))
            if r is not None:
                t = r.find(qn('a:t'))
                if t is not None: t.text = str(text)
    except: pass

def replace_text_in_shape(shape, old, new):
    if not shape.has_text_frame: return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            if old in str(run.text):
                run.text = run.text.replace(old, new)

def set_table_cell(shape, row, col, text, color=None, bold=None):
    try:
        cell = shape.table.cell(row, col)
        para = cell.text_frame.paragraphs[0]
        if para.runs:
            para.runs[0].text = str(text)
            if color: para.runs[0].font.color.rgb = color
            if bold is not None: para.runs[0].font.bold = bold
        else:
            r_elem = para._p.find(qn('a:r'))
            if r_elem is not None:
                t = r_elem.find(qn('a:t'))
                if t is not None: t.text = str(text)
            else:
                NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
                r_new = etree.SubElement(para._p, f'{{{NS}}}r')
                t_new = etree.SubElement(r_new, f'{{{NS}}}t')
                t_new.text = str(text)
    except: pass


# ─────────────────────────────────────────────
# DATA EXTRACTION
# ─────────────────────────────────────────────

def extract_data(tracker_path: str, snapshot_path: str,
                 ntp_comments_path: str, deck_date: str) -> dict:
    """
    Read tracker, compute all metrics, return data dict.
    This is the single source of truth for everything in the deck.
    """
    today = pd.Timestamp(datetime.strptime(deck_date, '%m/%d/%Y'))

    # Load tracker
    df_raw = pd.read_excel(tracker_path, sheet_name='HOPs', header=1)
    df = df_raw[df_raw['DON 444'].astype(str).str.strip().str.upper() == 'DON 444'].copy()
    df = df.drop_duplicates(subset=['HOP'])

    # Parse dates
    date_cols = ['NTP A', 'Material Received A ', 'MS15 Implementation Start A',
                 'MS16 Implementation Ends A', 'MS15 Implementation Start F',
                 'MS16 Implementation Ends F', 'ITW Schedule Start',
                 'ITW Schedule Complete', 'Samsung Schedule Start',
                 'Samsung Schedule Complete']
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')

    # Status flags
    df['has_ntp'] = df['NTP A'].dt.year >= 2025
    df['has_mat'] = df['Material Received A '].dt.year >= 2025
    df['started'] = df['MS15 Implementation Start A'].notna()
    df['complete'] = df['MS16 Implementation Ends A'].notna()
    df['in_progress'] = df['started'] & ~df['complete']
    df['not_started'] = ~df['started'] & ~df['complete']

    # String columns
    str_cols = {
        '_pm': 'Nokia PM', '_gc': 'General Contractor', '_new_cm': 'New CM',
        '_gc_pm': 'GC PM', '_ops': 'Viaero Ops Field Ops', '_site_cm': 'Site CM',
        '_cx': 'CX Notes:', '_ntp_wait': 'NTP is waiting on',
        '_ntp_owner': 'NTP Action Owner'
    }
    for attr, col in str_cols.items():
        df[attr] = df.apply(lambda r, k=col: gv(r, k), axis=1)

    # Computed fields
    df['days_to_start'] = (df['MS15 Implementation Start F'] - today).dt.days
    df['days_elapsed'] = np.where(
        df['in_progress'],
        (today - df['MS15 Implementation Start A']).dt.days,
        np.nan
    )
    df['over_18d'] = df['in_progress'] & (df['days_elapsed'] > 18)
    df['overdue'] = df['in_progress'] & (df['MS16 Implementation Ends F'] < today)

    def readiness(r):
        if r['has_ntp'] and r['has_mat']: return 'Ready'
        if r['has_ntp']: return 'NTP Only'
        if r['has_mat']: return 'Material Only'
        return 'Blocked'
    df['Readiness'] = df.apply(readiness, axis=1)

    def ntp_cat(r):
        o = str(r.get('NTP Action Owner', '')).upper()
        if any(x in o for x in ['ITW', 'SAMSUNG', 'VIAERO']): return 'External'
        if 'NOKIA' in o: return 'Program Team'
        return 'Other'
    df['NTP_Cat'] = df.apply(ntp_cat, axis=1)

    def vendor_conflict(r):
        c = []
        ms15f = r['MS15 Implementation Start F']
        ms16f = r['MS16 Implementation Ends F']
        if pd.isna(ms15f): return ''
        for vendor, vs, ve in [
            ('Samsung', r.get('Samsung Schedule Start'), r.get('Samsung Schedule Complete')),
            ('ITW', r.get('ITW Schedule Start'), r.get('ITW Schedule Complete'))
        ]:
            if pd.notna(vs) and pd.notna(ve):
                vs_t = pd.Timestamp(vs); ve_t = pd.Timestamp(ve)
                if vs_t <= ms15f <= ve_t:
                    c.append(f'🔴 {vendor} on site until {ve_t.strftime("%m/%d")}')
                elif pd.notna(ms16f):
                    buf = (vs_t - pd.Timestamp(ms16f)).days
                    if 0 <= buf < 10:
                        c.append(f'⚠ {vendor} {vs_t.strftime("%m/%d")} — {buf}d gap')
                    elif buf >= 10:
                        c.append(f'✅ {vendor} {vs_t.strftime("%m/%d")} — {buf}d buffer')
        return ' | '.join(c)
    df['VendorConflict'] = df.apply(vendor_conflict, axis=1)

    # Core counts
    total = len(df)
    ntp_count = int(df['has_ntp'].sum())
    mat_count = int(df['has_mat'].sum())
    started_count = int(df['started'].sum())
    complete_count = int(df['complete'].sum())
    ip_count = int(df['in_progress'].sum())

    # Load snapshot for deltas
    with open(snapshot_path) as f:
        snap = json.load(f)

    curr_ip = set(df[df['in_progress']]['HOP'].tolist())
    prev_ip = set(snap.get('ip_hops', []))
    new_starts = sorted(curr_ip - prev_ip)
    completions = sorted(prev_ip - curr_ip)

    # POR data
    ms15f_col = 'MS15 Implementation Start F'
    por = {}
    for mo, name in [(5, 'may'), (6, 'jun'), (7, 'jul'), (8, 'aug')]:
        grp = df[(df[ms15f_col].dt.month == mo) & (df[ms15f_col].dt.year == 2026)]
        with_ntp = grp[grp['has_ntp']]
        pending = grp[~grp['has_ntp']]
        pending_rows = []
        for _, r in pending.iterrows():
            cat = ntp_cat(r)
            pending_rows.append({
                'HOP': r['HOP'], 'GC': gv(r, 'General Contractor'),
                'CM': gv(r, 'Site CM'), 'owner': gv(r, 'NTP Action Owner'),
                'waiting': gv(r, 'NTP is waiting on'), 'cat': cat,
                'ms15f': r.get(ms15f_col), 'has_mat': bool(r['has_mat'])
            })
        pending_rows.sort(key=lambda x: 0 if x['cat'] == 'External'
                         else (2 if x['cat'] == 'Program Team' else 1))
        por[name] = {
            'total': len(grp), 'ntp': len(with_ntp), 'pending': len(pending),
            'ntp_hops': with_ntp[['HOP', 'General Contractor', 'Site CM', 'has_mat']].to_dict('records'),
            'pending_rows': pending_rows,
            'external': [r for r in pending_rows if r['cat'] == 'External'],
            'prog_team': [r for r in pending_rows if r['cat'] == 'Program Team'],
            'other': [r for r in pending_rows if r['cat'] == 'Other']
        }

    # Look-ahead and MSS
    la = df[(df[ms15f_col] >= today) &
            (df[ms15f_col] <= today + timedelta(days=7)) &
            ~df['started']].copy()
    mss = df[(df['MS15 Implementation Start A'] >= today - timedelta(days=5)) &
             df['in_progress']].copy()
    ip_df = df[df['in_progress']].copy()
    ip_df['days_elapsed'] = (today - ip_df['MS15 Implementation Start A']).dt.days
    ip_df['over_18d'] = ip_df['days_elapsed'] > 18

    # NTP comments
    ntp_comments = {}
    if ntp_comments_path and os.path.exists(ntp_comments_path):
        xl = pd.ExcelFile(ntp_comments_path)
        for sheet in xl.sheet_names:
            if 'history' in sheet.lower(): continue
            ntp_df = pd.read_excel(ntp_comments_path, sheet_name=sheet)
            comments = {}
            for _, row in ntp_df.iloc[2:].iterrows():
                hop = str(row.get('HOP', '')).strip()
                if not hop or hop.lower() == 'nan': continue
                comment = str(row.get('COMMENT (fill after call)', '')).strip()
                if comment and comment.lower() not in ['nan', '']:
                    comments[hop] = comment
            ntp_comments[sheet] = comments

    # Chart data
    pre_months = [(8,2025),(9,2025),(10,2025),(11,2025),(12,2025),(1,2026)]
    after_months = [(2,2026),(3,2026),(4,2026),(5,2026),(6,2026),(7,2026),(8,2026),(9,2026),(10,2026)]
    ms15a = 'MS15 Implementation Start A'
    ms16a = 'MS16 Implementation Ends A'
    ms16f_col2 = 'MS16 Implementation Ends F'
    def msum(col, mo, yr):
        return int(((df[col].dt.month == mo) & (df[col].dt.year == yr)).sum())
    starts_fc = [sum(msum(ms15f_col,m,y) for m,y in pre_months)] + [msum(ms15f_col,m,y) for m,y in after_months]
    starts_act = [sum(msum(ms15a,m,y) for m,y in pre_months)] + [msum(ms15a,m,y) for m,y in after_months]
    compl_fc = [sum(msum(ms16f_col2,m,y) for m,y in pre_months)] + [msum(ms16f_col2,m,y) for m,y in after_months]
    compl_act = [sum(msum(ms16a,m,y) for m,y in pre_months)] + [msum(ms16a,m,y) for m,y in after_months]

    # Lookups
    hop_gc_pm = {}; hop_ops = {}; hop_site_cm = {}; hop_pm = {}
    hop_mat = {}; hop_ntp = {}; hop_ms16f = {}
    for _, r in df.iterrows():
        hop = r['HOP']
        hop_gc_pm[hop] = gv(r, 'GC PM')
        hop_ops[hop] = gv(r, 'Viaero Ops Field Ops')
        hop_site_cm[hop] = gv(r, 'Site CM')
        hop_pm[hop] = gv(r, 'Nokia PM')
        hop_mat[hop] = '✓' if r['has_mat'] else '✗'
        hop_ntp[hop] = '✓' if r['has_ntp'] else '✗'
        v = r.get('MS16 Implementation Ends F')
        hop_ms16f[hop] = pd.Timestamp(v) if pd.notna(v) else None

    return {
        'df': df, 'por': por, 'la': la, 'mss': mss, 'ip_df': ip_df,
        'total': total, 'ntp_count': ntp_count, 'mat_count': mat_count,
        'started_count': started_count, 'complete_count': complete_count,
        'ip_count': ip_count, 'new_starts': new_starts, 'completions': completions,
        'snap': snap, 'ntp_comments': ntp_comments,
        'starts_fc': starts_fc, 'starts_act': starts_act,
        'compl_fc': compl_fc, 'compl_act': compl_act,
        'hop_gc_pm': hop_gc_pm, 'hop_ops': hop_ops, 'hop_site_cm': hop_site_cm,
        'hop_pm': hop_pm, 'hop_mat': hop_mat, 'hop_ntp': hop_ntp,
        'hop_ms16f': hop_ms16f, 'today': today, 'deck_date': deck_date
    }


# ─────────────────────────────────────────────
# DECK UPDATE
# ─────────────────────────────────────────────

def update_deck(data: dict, previous_deck_path: str, output_path: str):
    """
    Load previous deck, update all data in place, save to output_path.
    Never recreates slides from scratch.
    """
    OLD_DATE = data['snap'].get('session_date_display', '')
    NEW_DATE = data['deck_date']
    GREEN_C = RGBColor(0x00, 0x70, 0x3C)
    RED_C = RGBColor(0xC0, 0x00, 0x00)

    # Load as zip for chart XML updates
    with zipfile.ZipFile(previous_deck_path, 'r') as z:
        content = {n: z.read(n) for n in z.namelist()}

    # Fix chart XML cache
    def build_cache(vals):
        pts = ''.join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>'
                      for i, v in enumerate(vals) if v is not None and v > 0)
        return (f'<c:numCache><c:formatCode>General</c:formatCode>'
                f'<c:ptCount val="{len(vals)}"/>{pts}</c:numCache>')

    def fix_chart(xml_bytes, fc_vals, act_vals):
        xml = xml_bytes.decode('utf-8')
        caches = list(re.finditer(r'<c:numCache>.*?</c:numCache>', xml, re.DOTALL))
        if len(caches) >= 2:
            xml = xml[:caches[0].start()] + build_cache(fc_vals) + xml[caches[0].end():]
            caches2 = list(re.finditer(r'<c:numCache>.*?</c:numCache>', xml, re.DOTALL))
            if len(caches2) >= 2:
                xml = xml[:caches2[1].start()] + build_cache(act_vals) + xml[caches2[1].end():]
        return xml.encode('utf-8')

    if 'ppt/charts/chart1.xml' in content:
        content['ppt/charts/chart1.xml'] = fix_chart(
            content['ppt/charts/chart1.xml'], data['starts_fc'], data['starts_act'])
    if 'ppt/charts/chart2.xml' in content:
        content['ppt/charts/chart2.xml'] = fix_chart(
            content['ppt/charts/chart2.xml'], data['compl_fc'], data['compl_act'])

    # Update embedded workbooks
    labels = ['Jan/26+','Feb/26','Mar/26','Apr/26','May/26','Jun/26',
              'Jul/26','Aug/26','Sep/26','Oct/26']
    for embed_path, fc_vals, act_vals in [
        ('ppt/embeddings/Microsoft_Excel_Worksheet.xlsx', data['starts_fc'], data['starts_act']),
        ('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx', data['compl_fc'], data['compl_act'])
    ]:
        if embed_path in content:
            wb = openpyxl.load_workbook(io.BytesIO(content[embed_path]))
            ws = wb.active
            for ri in range(2, ws.max_row + 2):
                for ci in range(1, 4): ws.cell(row=ri, column=ci).value = None
            for ri, (lbl, fc, act) in enumerate(zip(labels, fc_vals, act_vals), 2):
                ws.cell(row=ri, column=1).value = lbl
                ws.cell(row=ri, column=2).value = fc
                ws.cell(row=ri, column=3).value = act if act > 0 else None
            out = io.BytesIO(); wb.save(out)
            content[embed_path] = out.getvalue()

    # Save temp file for pptx manipulation
    tmp_path = output_path + '.tmp.pptx'
    with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, d_bytes in content.items(): z.writestr(n, d_bytes)

    prs = Presentation(tmp_path)
    d = data  # shorthand

    # Global date sweep
    old_date_str = OLD_DATE if OLD_DATE else ''
    for slide in prs.slides:
        for shape in slide.shapes:
            if old_date_str:
                replace_text_in_shape(shape, old_date_str, NEW_DATE)

    # Slide 4: Delta
    shapes4 = list(prs.slides[3].shapes)
    snap = d['snap']
    delta_starts = d['started_count'] - snap.get('total_starts', 0)
    delta_complete = d['complete_count'] - snap.get('total_complete', 0)
    delta_ip = d['ip_count'] - snap.get('in_progress', 0)
    delta_ntp = d['ntp_count'] - snap.get('total_ntp', 0)

    def ds(v, ref): return f'+{v} vs {ref}' if v > 0 else (f'No change vs {ref}' if v == 0 else f'{v} vs {ref}')
    snap_date = snap.get('session_date', '')

    set_shape_text(shapes4[6], str(d['started_count']))
    set_shape_text(shapes4[11], str(d['complete_count']))
    set_shape_text(shapes4[16], str(d['ip_count']))
    set_shape_text(shapes4[21], str(d['ntp_count']))
    set_shape_text(shapes4[9], ds(delta_starts, snap_date))
    set_shape_text(shapes4[14], ds(delta_complete, snap_date))
    set_shape_text(shapes4[19], ds(delta_ip, snap_date))
    set_shape_text(shapes4[24], ds(delta_ntp, snap_date))
    set_shape_text(shapes4[26], f'★  New Starts Since {snap_date} ({len(d["new_starts"])})')

    new_start_idxs = [29, 34, 39, 50, 55]
    for i, si in enumerate(new_start_idxs):
        if si < len(shapes4):
            if i < len(d['new_starts']):
                set_shape_text(shapes4[si], f'★ {d["new_starts"][i]}')
            else:
                set_shape_text(shapes4[si], '')
                for off in [1, 2]:
                    try: set_shape_text(shapes4[si + off], '')
                    except: pass

    por = d['por']
    jun_delta = por['jun']['ntp'] - snap.get('jun_por_ntp', 0)
    jul_delta = por['jul']['ntp'] - snap.get('jul_por_ntp', 0)

    def nd(mo, ntp, total, delta):
        sign = '+' if delta > 0 else ''
        return f'{mo} POR:  {ntp} of {total} HOPs with NTP  ({sign}{delta} since {snap_date})'

    set_shape_text(shapes4[45], nd('Jun', por['jun']['ntp'], por['jun']['total'], jun_delta))
    set_shape_text(shapes4[47], nd('Jul', por['jul']['ntp'], por['jul']['total'], jul_delta))

    # Slide 5: Materials & NTP
    shapes5 = list(prs.slides[4].shapes)
    set_shape_text(shapes5[8], str(d['mat_count']))
    set_shape_text(shapes5[13], str(d['ntp_count']))

    # Slide 8: Look-ahead table
    shapes8 = list(prs.slides[7].shapes)
    la_tbl = shapes8[7]
    la_sorted = d['la'].sort_values('MS15 Implementation Start F')
    for ri in range(1, 11):
        if ri - 1 < len(la_sorted):
            r = la_sorted.iloc[ri - 1]; hop = r['HOP']
            ntp_sym = '✓' if r.get('has_ntp', False) else '✗'
            nw = gv(r, 'NTP is waiting on')
            if not nw: nw = 'NTP complete' if r.get('has_ntp', False) else ''
            set_table_cell(la_tbl, ri, 0, hop)
            set_table_cell(la_tbl, ri, 1, d['hop_gc_pm'].get(hop, ''))
            set_table_cell(la_tbl, ri, 2, gv(r, 'General Contractor'))
            set_table_cell(la_tbl, ri, 3, ntp_sym,
                           color=GREEN_C if ntp_sym == '✓' else RED_C, bold=True)
            set_table_cell(la_tbl, ri, 4, fmt_d(r.get('MS15 Implementation Start F')))
            set_table_cell(la_tbl, ri, 5, d['hop_site_cm'].get(hop, ''))
            set_table_cell(la_tbl, ri, 6, d['hop_ops'].get(hop, ''))
            set_table_cell(la_tbl, ri, 7, nw[:50])
            set_table_cell(la_tbl, ri, 8, '')
        else:
            for ci in range(9): set_table_cell(la_tbl, ri, ci, '')

    # Slide 9: In progress table
    shapes9 = list(prs.slides[8].shapes)
    ip_tbl = shapes9[5]
    ip_sorted = d['ip_df'].sort_values('MS16 Implementation Ends F', na_position='last')
    new_starts_set = set(d['new_starts'])
    for ri in range(1, 26):
        if ri - 1 < len(ip_sorted):
            r = ip_sorted.iloc[ri - 1]; hop = r['HOP']
            is_new = hop in new_starts_set
            o18 = bool(r.get('over_18d', False))
            risk = '🔴 Over 18d' if o18 else ('★ NEW' if is_new else 'G')
            set_table_cell(ip_tbl, ri, 0, f'★ {hop}' if is_new else hop)
            set_table_cell(ip_tbl, ri, 1, d['hop_gc_pm'].get(hop, ''))
            set_table_cell(ip_tbl, ri, 2, gv(r, 'General Contractor'))
            set_table_cell(ip_tbl, ri, 3, fmt_dm(r.get('MS15 Implementation Start A')))
            set_table_cell(ip_tbl, ri, 4, fmt_ds(r.get('MS16 Implementation Ends F')))
            set_table_cell(ip_tbl, ri, 5, risk)
            set_table_cell(ip_tbl, ri, 6, '')
        else:
            for ci in range(7): set_table_cell(ip_tbl, ri, ci, '')

    # POR slides
    def update_por_overview(slide_idx, mo_name, mo_key):
        s = prs.slides[slide_idx]; p = por[mo_key]
        shapes = list(s.shapes)
        set_shape_text(shapes[1], f'{p["total"]} Forecasted {mo_name} POR')
        for idx, val in [(6, p['total']), (9, p['ntp']), (12, p['pending']),
                         (18, p['ntp']), (22, len(p['prog_team']) + len(p['other'])),
                         (26, len(p['external']))]:
            if idx < len(shapes): set_shape_text(shapes[idx], str(val))
        if 29 < len(shapes):
            set_shape_text(shapes[29], f'{p["ntp"]} with NTP of {p["total"]} POR  ·  {p["pending"]} pending NTP')

    def update_por_confirmed(slide_idx, mo_name, mo_key, sheet_key):
        s = prs.slides[slide_idx]; p = por[mo_key]
        shapes = list(s.shapes)
        set_shape_text(shapes[1], f'{p["ntp"]} with NTP of {p["total"]} POR')
        comments = d['ntp_comments'].get(sheet_key, {})
        for shape in s.shapes:
            if shape.shape_type == 19:
                tbl = shape.table
                hdrs = [tbl.cell(0, ci).text.strip() for ci in range(len(tbl.columns))]
                if 'HOP' not in hdrs: continue
                hop_col = next((i for i, h in enumerate(hdrs) if h == 'HOP'), 0)
                gc_col = next((i for i, h in enumerate(hdrs) if h == 'GC'), 1)
                mat_col = next((i for i, h in enumerate(hdrs) if 'Mat' in h), gc_col + 1)
                comment_col = len(hdrs) - 1
                ntp_hops = sorted(p['ntp_hops'], key=lambda h: d['hop_ms16f'].get(h['HOP'],
                                  pd.Timestamp('2099-01-01')) or pd.Timestamp('2099-01-01'))
                for ri in range(1, len(tbl.rows)):
                    if ri - 1 < len(ntp_hops):
                        h = ntp_hops[ri - 1]; hop = str(h['HOP'])
                        gc = str(h.get('General Contractor', '')).strip()
                        gc = '' if gc.lower() == 'nan' else gc
                        mat_sym = '✓' if h.get('has_mat', False) else '✗'
                        comment = comments.get(hop, '')
                        if not comment:
                            for k, v in comments.items():
                                if k.strip().upper() == hop.strip().upper():
                                    comment = v; break
                        set_table_cell(shape, ri, hop_col, hop)
                        set_table_cell(shape, ri, gc_col, gc)
                        if mat_col < len(hdrs):
                            set_table_cell(shape, ri, mat_col, mat_sym,
                                           color=GREEN_C if mat_sym == '✓' else RED_C, bold=True)
                        if comment: set_table_cell(shape, ri, comment_col, comment[:55])
                    else:
                        for ci in range(len(hdrs)): set_table_cell(shape, ri, ci, '')

    def update_por_pending(slide_idx, mo_name, mo_key, sheet_key):
        s = prs.slides[slide_idx]; p = por[mo_key]
        shapes = list(s.shapes)
        set_shape_text(shapes[1], f'{p["pending"]} of {p["total"]} pending NTP')
        comments = d['ntp_comments'].get(sheet_key, {})

        def sort_key(h):
            cat_order = 0 if h['cat'] == 'External' else (1 if h['cat'] == 'Other' else 2)
            ms16 = d['hop_ms16f'].get(h['HOP'])
            ts = ms16 if ms16 and pd.notna(ms16) else pd.Timestamp('2099-01-01')
            return (cat_order, ts)

        ext = sorted(p['external'], key=sort_key)
        prog = sorted(p['prog_team'] + p['other'], key=sort_key)

        tables = [(i, shape) for i, shape in enumerate(shapes) if shape.shape_type == 19]
        for tbl_idx, (_, shape) in enumerate(tables):
            tbl = shape.table; ncols = len(tbl.columns); nrows = len(tbl.rows)
            comment_col = ncols - 1
            rows_to_fill = ext if tbl_idx == 0 else prog
            for ri in range(1, nrows):
                if ri - 1 < len(rows_to_fill):
                    h = rows_to_fill[ri - 1]; hop = h['HOP']
                    fc_s = fmt_dm(h.get('ms15f'))
                    ms16_v = d['hop_ms16f'].get(hop)
                    fc_e = fmt_dm(ms16_v) if ms16_v and pd.notna(ms16_v) else ''
                    comment = comments.get(hop, '')
                    if not comment:
                        for k, v in comments.items():
                            if k.strip().upper() == hop.strip().upper():
                                comment = v; break
                    set_table_cell(shape, ri, 0, hop)
                    set_table_cell(shape, ri, 1, h['GC'])
                    set_table_cell(shape, ri, 2, fc_s)
                    set_table_cell(shape, ri, 3, fc_e)
                    set_table_cell(shape, ri, 4, h['owner'][:25])
                    set_table_cell(shape, ri, 5, h['waiting'][:45])
                    set_table_cell(shape, ri, comment_col, comment[:55] if comment else '')
                else:
                    for ci in range(ncols): set_table_cell(shape, ri, ci, '')

        for shape in shapes:
            if shape.has_text_frame and 'Program Team Actions' in shape.text_frame.text:
                set_shape_text(shape, f'Program Team Actions ({len(prog)})')

    update_por_overview(9, 'June', 'jun')
    update_por_overview(12, 'July', 'jul')
    update_por_overview(15, 'August', 'aug')
    update_por_confirmed(10, 'June', 'jun', 'Jun Pending NTP')
    update_por_confirmed(13, 'July', 'jul', 'Jul Pending NTP')
    update_por_confirmed(16, 'August', 'aug', 'Aug Pending NTP')
    update_por_pending(11, 'June', 'jun', 'Jun Pending NTP')
    update_por_pending(14, 'July', 'jul', 'Jul Pending NTP')
    update_por_pending(17, 'August', 'aug', 'Aug Pending NTP')

    # Final date sweep — catch anything missed
    for slide in prs.slides:
        for shape in slide.shapes:
            replace_text_in_shape(shape, '6/2/2026', NEW_DATE)
            replace_text_in_shape(shape, '6/4/2026', NEW_DATE)

    prs.save(output_path)
    if os.path.exists(tmp_path): os.remove(tmp_path)
    return output_path


# ─────────────────────────────────────────────
# SNAPSHOT + NTP COMMENTS GENERATORS
# ─────────────────────────────────────────────

def generate_snapshot(data: dict, output_path: str):
    df = data['df']; por = data['por']; la = data['la']
    snap = {
        "session_date": data['deck_date'],
        "session_date_display": data['deck_date'],
        "total_starts": data['started_count'],
        "total_complete": data['complete_count'],
        "in_progress": data['ip_count'],
        "total_ntp": data['ntp_count'],
        "mat_received": data['mat_count'],
        "may_por_total": por['may']['total'], "may_por_ntp": por['may']['ntp'],
        "jun_por_total": por['jun']['total'], "jun_por_ntp": por['jun']['ntp'],
        "jul_por_total": por['jul']['total'], "jul_por_ntp": por['jul']['ntp'],
        "aug_por_total": por['aug']['total'], "aug_por_ntp": por['aug']['ntp'],
        "ip_hops": df[df['in_progress']]['HOP'].tolist(),
        "la_hops": la['HOP'].tolist(),
        "scop_accepted": data.get('scop_accepted', 0),
        "scop_pending": data.get('scop_pending', 0),
        "scop_not_started": data.get('scop_ns', 0)
    }
    with open(output_path, 'w') as f:
        json.dump(snap, f, indent=2, default=str)
    return output_path


def generate_ntp_comments(data: dict, output_path: str):
    por = data['por']; ntp_comments = data['ntp_comments']
    hop_ms16f = data['hop_ms16f']
    NAVY='124191'; WHITE='FFFFFF'; LT_BLUE='EEF2F7'
    AMBER_H='D4860A'; GRAY='595959'; RED_H='C00000'; ORANGE_H='C55A11'
    CAT_FILLS = {
        'External': ('FFE5E5', 'C00000'),
        'Other': ('FFF3E0', 'C55A11'),
        'Program Team': ('E8F0FC', '124191')
    }

    def sort_key(h):
        cat_order = 0 if h['cat'] == 'External' else (1 if h['cat'] == 'Other' else 2)
        ms16 = hop_ms16f.get(h['HOP'])
        ts = ms16 if ms16 and pd.notna(ms16) else pd.Timestamp('2099-01-01')
        return (cat_order, ts)

    def fmt_dm_local(v):
        if v is None or pd.isna(v): return ''
        try: return pd.Timestamp(v).strftime('%m/%d')
        except: return ''

    wb = openpyxl.Workbook(); wb.remove(wb.active)
    for mo_name, mo_key, sheet_key in [
        ('Jun Pending NTP', 'jun', 'Jun Pending NTP'),
        ('Jul Pending NTP', 'jul', 'Jul Pending NTP'),
        ('Aug Pending NTP', 'aug', 'Aug Pending NTP')
    ]:
        ws = wb.create_sheet(mo_name)
        p = por[mo_key]; comments = ntp_comments.get(sheet_key, {})
        sorted_rows = sorted(p['pending_rows'], key=sort_key)
        hdrs = ['HOP','Category','Action Owner','GC','FC Start','FC End',
                'NTP Blocker / Waiting On','COMMENT (fill after call)','STATUS']
        col_widths = [45,14,18,14,10,10,50,58,16]
        for ci, (hdr, cw) in enumerate(zip(hdrs, col_widths), 1):
            c = ws.cell(row=1, column=ci, value=hdr)
            c.fill = PatternFill('solid', fgColor=NAVY)
            c.font = Font(name='Calibri', bold=True, size=10, color=WHITE)
            c.alignment = Alignment(horizontal='center', vertical='center')
            ws.column_dimensions[get_column_letter(ci)].width = cw
        ws.row_dimensions[1].height = 21.95
        ws.merge_cells('A2:I2')
        c = ws.cell(row=2, column=1,
                    value='Sorted to match slides: External first → FC End oldest to newest.')
        c.fill = PatternFill('solid', fgColor=LT_BLUE)
        c.font = Font(name='Calibri', size=9, color=GRAY)
        c.alignment = Alignment(horizontal='left', vertical='center')
        ws.row_dimensions[2].height = 14.1
        ws.merge_cells('A3:I3')
        c = ws.cell(row=3, column=1,
                    value='🔴 External  |  🟠 Other  |  🔵 Program Team  |  STATUS: "Action Taken" | "In Progress" | "Needs Attention" | "Pending"')
        c.font = Font(name='Calibri', size=8, color=AMBER_H)
        c.alignment = Alignment(horizontal='left', vertical='center')
        ws.row_dimensions[3].height = 12.0
        last_cat = None; data_row = 4
        for h in sorted_rows:
            cat = h['cat']; bg_fill, font_color = CAT_FILLS.get(cat, ('FFFFFF', '000000'))
            if cat != last_cat:
                cat_label = {'External':'🔴 EXTERNAL BLOCKERS','Other':'🟠 OTHER',
                             'Program Team':'🔵 PROGRAM TEAM ACTIONS'}[cat]
                ws.merge_cells(f'A{data_row}:I{data_row}')
                c = ws.cell(row=data_row, column=1, value=cat_label)
                divider_fill = {'External':RED_H,'Other':ORANGE_H,'Program Team':NAVY}[cat]
                c.fill = PatternFill('solid', fgColor=divider_fill)
                c.font = Font(name='Calibri', bold=True, size=9, color=WHITE)
                c.alignment = Alignment(horizontal='left', vertical='center')
                ws.row_dimensions[data_row].height = 16
                data_row += 1; last_cat = cat
            existing_comment = comments.get(h['HOP'], '')
            if not existing_comment:
                for k, v in comments.items():
                    if k.strip().upper() == h['HOP'].strip().upper():
                        existing_comment = v; break
            status = 'In Progress' if existing_comment else 'Pending'
            fc_s = fmt_dm_local(h.get('ms15f'))
            fc_e = fmt_dm_local(hop_ms16f.get(h['HOP']))
            vals = [h['HOP'],cat,h['owner'],h['GC'],fc_s,fc_e,
                    h['waiting'],existing_comment,status]
            for ci, val in enumerate(vals, 1):
                c = ws.cell(row=data_row, column=ci, value=val)
                c.fill = PatternFill('solid', fgColor=bg_fill)
                c.font = Font(name='Calibri', bold=(ci in [1,2]), size=9,
                              color=font_color if ci == 2 else '000000')
                c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            ws.row_dimensions[data_row].height = 20.1; data_row += 1

    # History tab
    ws_hist = wb.create_sheet('📋 Comments History')
    ws_hist.sheet_properties.tabColor = NAVY
    hist_hdrs = ['Month','HOP','GC','Category','Action Owner',
                 'NTP Blocker / Waiting On','Comment','Session Date']
    for ci, hdr in enumerate(hist_hdrs, 1):
        c = ws_hist.cell(row=1, column=ci, value=hdr)
        c.fill = PatternFill('solid', fgColor=NAVY)
        c.font = Font(name='Calibri', bold=True, size=10, color=WHITE)
        c.alignment = Alignment(horizontal='center', vertical='center')
    ri = 3
    for sheet_key, mo_label, mo_key in [('Jun Pending NTP','Jun','jun'),
                                         ('Jul Pending NTP','Jul','jul'),
                                         ('Aug Pending NTP','Aug','aug')]:
        comments = ntp_comments.get(sheet_key, {})
        row_lookup = {h['HOP']: h for h in por[mo_key]['pending_rows']}
        for hop, comment in comments.items():
            if not comment: continue
            h = row_lookup.get(hop, {})
            for ci, val in enumerate([mo_label,hop,h.get('GC',''),h.get('cat',''),
                                       h.get('owner',''),h.get('waiting',''),
                                       comment, data['deck_date']], 1):
                c = ws_hist.cell(row=ri, column=ci, value=val)
                c.font = Font(name='Calibri', bold=(ci==2), size=9)
                c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            ri += 1

    wb.save(output_path)
    return output_path


# ─────────────────────────────────────────────
# MAIN ENTRY POINT
# ─────────────────────────────────────────────

def build(tracker_path: str, previous_deck_path: str, snapshot_path: str,
          ntp_comments_path: str, deck_date: str, output_dir: str) -> dict:
    """
    Main entry point called by the API route.
    Returns paths to all generated files + summary stats.
    """
    os.makedirs(output_dir, exist_ok=True)
    date_slug = deck_date.replace('/', '-')

    # Extract
    data = extract_data(tracker_path, snapshot_path, ntp_comments_path, deck_date)

    # Generate outputs
    deck_out = os.path.join(output_dir, f'Viaero_Construction_Update_{date_slug}.pptx')
    snap_out = os.path.join(output_dir, f'session_snapshot_{date_slug}.json')
    ntp_out  = os.path.join(output_dir, f'NTP_Comments_{date_slug}.xlsx')

    update_deck(data, previous_deck_path, deck_out)
    generate_snapshot(data, snap_out)
    generate_ntp_comments(data, ntp_out)

    return {
        'deck_path': deck_out,
        'snapshot_path': snap_out,
        'ntp_comments_path': ntp_out,
        'summary': {
            'deck_date': deck_date,
            'total_hops': data['total'],
            'ntp_count': data['ntp_count'],
            'mat_count': data['mat_count'],
            'started_count': data['started_count'],
            'complete_count': data['complete_count'],
            'ip_count': data['ip_count'],
            'new_starts': data['new_starts'],
            'completions': data['completions'],
        }
    }


if __name__ == '__main__':
    # CLI usage: python build_deck.py <tracker> <prev_deck> <snapshot> <ntp_comments> <deck_date> <output_dir>
    args = sys.argv[1:]
    result = build(
        tracker_path=args[0],
        previous_deck_path=args[1],
        snapshot_path=args[2],
        ntp_comments_path=args[3] if len(args) > 3 else '',
        deck_date=args[4] if len(args) > 4 else datetime.today().strftime('%m/%d/%Y'),
        output_dir=args[5] if len(args) > 5 else './outputs'
    )
    print(json.dumps(result, indent=2))