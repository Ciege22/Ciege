"""
GR / Invoicing Tracker — data join logic (Part 6 of the GR Tracker build).

Joins the SPO Report (Supabase report_snapshots id='spo') against the latest
tracker snapshot (Supabase tracker_snapshot) to determine GR release status
per HOP/SOG line. Mirrors app/lib/grTracker.ts — keep both in sync.
"""

import json
from datetime import datetime

import pandas as pd

# SPO report rows are stored as arrays with no header row (see app/reports/page.tsx).
# These are the fixed column positions established by that page's SPO_COL_IDX mapping.
SPO_COL = {
    'customer_site_id': 7,
    'name': 8,
    'sog_name': 33,
    'spo_number': 40,
    'spo_vendor': 43,
    'spo_value': 47,
    'gr_date': 50,
}

# Same GC <> SPO-vendor alias table as app/lib/gcConfig.ts — keep in sync.
GC_CONFIG = [
    {'gc': 'Mastec',  'spo_match': ['mastec network solutions llc']},
    {'gc': 'MZI',     'spo_match': ['mzi group inc']},
    {'gc': 'NV Tel',  'spo_match': ['nv tel inc']},
    {'gc': 'Tech CX', 'spo_match': ['tech construction inc']},
    {'gc': 'Vikor',   'spo_match': ['sioux falls tower specialists inc.', 'sioux falls tower specialists inc']},
    {'gc': 'TCE',     'spo_match': ['tower communications experts. llc', 'tower communications experts llc']},
    {'gc': 'InSite',  'spo_match': ['insite telecom llc']},
]


def _row_get(row, idx):
    try:
        return row[idx]
    except (IndexError, TypeError):
        return None


def parse_date_any(val):
    """Parse Excel serials / ISO strings / MM-DD-YYYY strings into a naive datetime."""
    if val is None or val == '':
        return None
    if isinstance(val, datetime):
        return val.replace(tzinfo=None) if val.tzinfo else val
    if isinstance(val, (int, float)):
        if 40000 < val < 60000:
            try:
                return (pd.Timestamp('1899-12-30') + pd.Timedelta(days=val)).to_pydatetime()
            except Exception:
                return None
        return None
    s = str(val).strip()
    if not s or s.lower() in ('null', 'undefined', 'nan', 'nat'):
        return None
    try:
        ts = pd.Timestamp(s)
        if ts.tzinfo is not None:
            ts = ts.tz_convert(None)
        return ts.to_pydatetime()
    except Exception:
        return None


def normalize_tracker_hop(hop: str) -> str:
    """Tracker uses '<>' as the HOP separator, SPO report uses 'to'."""
    return hop.replace(' <> ', ' to ').replace('<>', 'to')


def match_key(s: str) -> str:
    return (s or '').strip().lower()


def gc_for_vendor(vendor) -> str:
    v = match_key(str(vendor or ''))
    for cfg in GC_CONFIG:
        if v in cfg['spo_match']:
            return cfg['gc']
    return str(vendor or '').strip()


def classify_sog(sog_name_raw: str):
    """SOG Name tier bucketing. 'init 20%' MUST be checked before plain '20%'."""
    s = (sog_name_raw or '').strip().lower()
    if 'init 20%' in s:
        return {'tier': 'init20', 'trigger': 'MS15A', 'cj_actionable': True,
                'tier_label': 'Init 20% — MS15A (Site Started)', 'is_decom_scop': False}
    if '60%' in s:
        return {'tier': '60', 'trigger': 'MS16A', 'cj_actionable': True,
                'tier_label': '60% — MS16A (CX Complete)', 'is_decom_scop': False}
    if '70%' in s:
        return {'tier': '70', 'trigger': 'MS16A', 'cj_actionable': True,
                'tier_label': '70% — MS16A (CX Complete, old split)', 'is_decom_scop': False}
    if '20%' in s and 'init' not in s:
        return {'tier': '20', 'trigger': None, 'cj_actionable': False,
                'tier_label': '20% — Decom/SCOP', 'is_decom_scop': True}
    if '30%' in s:
        return {'tier': '30', 'trigger': None, 'cj_actionable': False,
                'tier_label': '30% — Decom/SCOP (old split)', 'is_decom_scop': True}
    return {'tier': None, 'trigger': None, 'cj_actionable': False,
            'tier_label': sog_name_raw or 'Unknown', 'is_decom_scop': False}


def build_tracker_date_map(tracker_rows):
    """rows -> { normalized_hop: {ms15a, ms16a, nokia_pm} }. Mirrors buildTrackerDateMap in TS."""
    header_row = -1
    for i in range(min(10, len(tracker_rows))):
        if any(str(c).strip() == 'HOP' for c in (tracker_rows[i] or [])):
            header_row = i
            break
    date_map = {}
    if header_row == -1:
        return date_map

    headers = [str(h).strip() if h is not None else '' for h in tracker_rows[header_row]]

    def col(name):
        try:
            return headers.index(name)
        except ValueError:
            return -1

    hop_col = col('HOP')
    don444_col = col('DON 444')
    ms15a_col = col('MS15 Implementation Start A')
    ms16a_col = col('MS16 Implementation Ends A')
    nokia_pm_col = col('Nokia PM')

    for i in range(header_row + 1, len(tracker_rows)):
        row = tracker_rows[i] or []
        don = str(_row_get(row, don444_col) or '').strip().upper()
        if don != 'DON 444':
            continue
        hop = str(_row_get(row, hop_col) or '').strip()
        if not hop or hop == 'undefined':
            continue
        # Tracker HOP names use '<>' as a separator; SPO report Name uses 'to'.
        # Normalize here so the key matches the SPO row's raw Name column.
        key = match_key(normalize_tracker_hop(hop))
        if key in date_map:
            continue
        date_map[key] = {
            'ms15a': parse_date_any(_row_get(row, ms15a_col)),
            'ms16a': parse_date_any(_row_get(row, ms16a_col)),
            'nokia_pm': str(_row_get(row, nokia_pm_col) or '').strip(),
        }
    return date_map


def _trigger_met(trigger, dates):
    if not trigger or not dates:
        return False
    if trigger == 'MS15A':
        return dates.get('ms15a') is not None
    if trigger == 'MS16A':
        return dates.get('ms16a') is not None
    return False


def classify_row_type(sog_name_raw, spo_number_raw):
    """Foundational row filter — must run before any other row processing.

    Base PO = has SPO Number AND has SOG Name
    CR      = has SPO Number AND SOG Name is blank — displayed as-is, no "CR" label in UI
    Error   = has SOG Name but NO SPO Number — excluded entirely, never returned
    """
    has_spo = bool((spo_number_raw or '').strip())
    has_sog = bool((sog_name_raw or '').strip())
    if has_spo and has_sog:
        return 'base'
    if has_spo and not has_sog:
        return 'cr'
    return None


def _compute_status(gr_date, is_trigger_met):
    if gr_date:
        return 'GR Done'
    if is_trigger_met:
        return 'Ready to Release'
    return 'Awaiting Trigger'


def build_gr_rows(spo_rows, tracker_date_map):
    out = []
    for row in spo_rows:
        name_raw = str(_row_get(row, SPO_COL['name']) or '').strip()
        if not name_raw:
            continue

        sog_name_raw = str(_row_get(row, SPO_COL['sog_name']) or '').strip()
        spo_number_raw = str(_row_get(row, SPO_COL['spo_number']) or '').strip()
        row_type = classify_row_type(sog_name_raw, spo_number_raw)
        if not row_type:
            continue  # Error row (SOG, no SPO) or empty row — excluded entirely

        # CR rows have no SOG tier to classify, so they carry no trigger gate —
        # treated as always eligible for release until GR'd (business decision).
        if row_type == 'cr':
            sog = {'tier': None, 'trigger': None, 'cj_actionable': True,
                   'tier_label': '', 'is_decom_scop': False}
        else:
            sog = classify_sog(sog_name_raw)

        vendor = _row_get(row, SPO_COL['spo_vendor'])
        gc = gc_for_vendor(vendor)

        # SPO report Name column already uses 'to' — the tracker-side map was
        # normalized when it was built, so we match against the raw SPO name here.
        key = match_key(name_raw)
        dates = tracker_date_map.get(key)
        is_trigger_met = True if row_type == 'cr' else _trigger_met(sog['trigger'], dates)

        gr_date = parse_date_any(_row_get(row, SPO_COL['gr_date']))
        status = _compute_status(gr_date, is_trigger_met)

        if sog['trigger'] == 'MS15A':
            trigger_date = dates.get('ms15a') if dates else None
        elif sog['trigger'] == 'MS16A':
            trigger_date = dates.get('ms16a') if dates else None
        else:
            trigger_date = None

        raw_val = _row_get(row, SPO_COL['spo_value'])
        try:
            spo_value = float(raw_val) if raw_val not in (None, '') else 0.0
        except (TypeError, ValueError):
            try:
                spo_value = float(str(raw_val).replace('$', '').replace(',', ''))
            except (TypeError, ValueError):
                spo_value = 0.0

        out.append({
            'hop': key,
            'hop_display': name_raw,
            'path_id': str(_row_get(row, SPO_COL['customer_site_id']) or '').strip(),
            'gc': gc,
            'nokia_pm': (dates or {}).get('nokia_pm', ''),
            'sog_name': sog_name_raw,
            'spo_number': spo_number_raw,
            'spo_value': spo_value,
            'row_type': row_type,
            'trigger_date': trigger_date.strftime('%Y-%m-%d') if trigger_date else None,
            'trigger_met': is_trigger_met,
            'gr_date': gr_date.strftime('%Y-%m-%d') if gr_date else None,
            'status': status,
            'tier_label': sog['tier_label'],
            'cj_actionable': sog['cj_actionable'],
            'is_decom_scop': sog['is_decom_scop'],
        })
    return out


def compute_gr_breakdown(rows):
    base = [r for r in rows if r['row_type'] == 'base']
    cr = [r for r in rows if r['row_type'] == 'cr']
    return {
        'total_base_pos': len(base),
        'total_crs': len(cr),
        'base_po_value_grd': sum(r['spo_value'] for r in base if r['gr_date']),
        'base_po_value_pending': sum(r['spo_value'] for r in base if not r['gr_date'] and r['trigger_met']),
        'cr_value_grd': sum(r['spo_value'] for r in cr if r['gr_date']),
        'cr_value_pending': sum(r['spo_value'] for r in cr if not r['gr_date'] and r['trigger_met']),
    }


def group_gr_rows(rows, gc_filter=None):
    if gc_filter:
        rows = [r for r in rows if r['gc'] == gc_filter]
    ready = [r for r in rows if r['cj_actionable'] and r['status'] == 'Ready to Release']
    awaiting = [r for r in rows if r['cj_actionable'] and r['status'] == 'Awaiting Trigger']
    done = [r for r in rows if r['status'] == 'GR Done']
    awareness = [r for r in rows if r['is_decom_scop']]
    return {
        'ready': ready, 'awaiting': awaiting, 'done': done, 'awareness': awareness,
        'breakdown': compute_gr_breakdown(rows),
    }


def load_gr_data(supabase_client, gc_filter=None):
    """Loads SPO report + latest tracker snapshot from Supabase and returns grouped GR rows."""
    spo_result = supabase_client.table('report_snapshots').select('data').eq('id', 'spo').single().execute()
    spo_rows = json.loads(spo_result.data['data']) if spo_result.data and spo_result.data.get('data') else []

    tracker_result = supabase_client.table('tracker_snapshot').select('data').order('uploaded_at', desc=True).limit(1).single().execute()
    tracker_rows = json.loads(tracker_result.data['data']) if tracker_result.data and tracker_result.data.get('data') else []

    tracker_date_map = build_tracker_date_map(tracker_rows)
    gr_rows = build_gr_rows(spo_rows, tracker_date_map)
    return group_gr_rows(gr_rows, gc_filter=gc_filter)
