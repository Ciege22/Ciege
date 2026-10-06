"""
Deck Builder V2 — isolated rebuild of build_deck.py's /build pipeline.

Does NOT import anything that mutates build_deck.py's behavior — only reuses
its pure, already-proven helpers (gv, set_shape_text, set_table_cell,
expand_table_rows, replace_text_in_shape) and its extract_data() function,
which already computes every metric V2 needs (Cx Start/Complete Forecast +
Actual series, new_starts/completions, mss/look-ahead tables, hop lookups)
with zero Plan-related hardcoding — Plan is only ever touched in
update_deck()'s _BOSS_STARTS/_BOSS_COMPLETE block, which V2 never calls.

Behavioral differences from the original /build, all explicitly requested:
  1. No more Materials and NTP Status / On Air Acceptance Review / Final SCOP
     Close-Out / Crew Pipeline / HOP Duration Summary / Cycle Times slides —
     V2's template (deck_builder_v2_template.pptx, built once from CJ's own
     trimmed working deck) simply never has them, so update_deck_v2 never
     looks for them.
  2. POR months are a rolling window computed from deck_date (current month +
     the next few), skipping any month whose total is 0 — see
     compute_por_window(). Every included month gets an accurate +/- delta
     against the last build (not hardcoded to Jun/Jul).
  3. The Cx Start / Cx Complete charts' Plan series is read verbatim out of
     the previous deck's existing chart XML and written back unchanged —
     never recomputed from a hardcoded table. Forecast/Actual are still
     freshly computed every build from extract_data().
  4. The "What Changed" delta compares against a snapshot this module saves
     to Supabase (pm_updates_cache, id='deck-v2-last-snapshot') after every
     successful build, instead of an unrelated tracker-upload fallback.
  5. A new "New Starts & Completions Since Last Meeting" slide lists the
     actual HOP names (extract_data already computes these as
     new_starts/completions; the original builder discards the names and
     only shows a count).
  6. The Decom Dashboard's 3 slides and the SCOP Status Deck's 3 slides
     (already generated client-side in the browser — see app/lib/decomDeck.ts
     and app/lib/scopDeck.ts) are appended into the deck automatically via
     copy_slides_into(), instead of CJ pasting them in by hand.
"""

import os
import posixpath
import re
import copy
import json
import zipfile
import calendar
from datetime import datetime, timedelta

import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from lxml import etree

from build_deck import (
    extract_data, gv, fmt_d, fmt_dm, fmt_ds,
    set_shape_text, set_table_cell, expand_table_rows, replace_text_in_shape,
)

P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart'

GREEN_C = RGBColor(0x00, 0x70, 0x3C)
RED_C = RGBColor(0xC0, 0x00, 0x00)

EXT_CONTENT_TYPES = {
    'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg', 'gif': 'image/gif',
    'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
}


# ─────────────────────────────────────────────
# 1. SUPABASE SNAPSHOT (delta baseline) — replaces the tracker-upload fallback
# ─────────────────────────────────────────────

SNAPSHOT_PREFIX = 'deck-v2-snapshot-'


def call_tuesday(deck_date_str: str) -> datetime:
    """The Tuesday call this deck is for — the Tuesday nearest the picked date."""
    d = datetime.strptime(deck_date_str, '%m/%d/%Y')
    delta = (1 - d.weekday()) % 7
    if delta > 3:
        delta -= 7
    return d + timedelta(days=delta)


def load_prior_snapshot(supabase_client, call_dt: datetime) -> dict:
    """The saved numbers from the most recent Tuesday call before this one."""
    if not supabase_client:
        return {}
    try:
        res = supabase_client.table('pm_updates_cache').select('id,updates').like('id', f'{SNAPSHOT_PREFIX}%').execute()
        best_key, best = None, {}
        for row in res.data or []:
            day = row['id'][len(SNAPSHOT_PREFIX):]
            try:
                dt = datetime.strptime(day, '%Y-%m-%d')
            except ValueError:
                continue
            if dt < call_dt and (best_key is None or dt > best_key):
                best_key, best = dt, json.loads(row['updates'])
        return best
    except Exception as e:
        print(f'[v2-snapshot] load failed: {e}', flush=True)
    return {}


def save_snapshot(supabase_client, call_dt: datetime, snapshot: dict):
    if not supabase_client:
        return
    try:
        supabase_client.table('pm_updates_cache').upsert({
            'id': SNAPSHOT_PREFIX + call_dt.strftime('%Y-%m-%d'),
            'updates': json.dumps(snapshot, default=str),
            'updated_at': datetime.utcnow().isoformat(),
        }).execute()
    except Exception as e:
        print(f'[v2-snapshot] save failed: {e}', flush=True)


# ─────────────────────────────────────────────
# 2. ROLLING POR WINDOW — replaces the hardcoded may..dec / year==2026 loop
# ─────────────────────────────────────────────

MONTH_ABBR = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
MONTH_FULL = ['', 'January', 'February', 'March', 'April', 'May', 'June', 'July',
              'August', 'September', 'October', 'November', 'December']


def compute_por_window(df, deck_date: pd.Timestamp, window_months: int = 6, max_active: int = 3):
    """Starting at deck_date's month, walks forward up to window_months calendar
    months, keeps only the months with a non-zero POR total, and returns at
    most max_active of them in order (current month first). This replaces the
    old hardcoded may..dec/2026-only loop — also fixes the latent bug where
    POR silently went empty for any month outside that hardcoded year.
    """
    ms15f_col = 'MS15 Implementation Start F'
    por = {}
    order = []  # list of (mo_key, mo_name, sheet_key)
    y, m = deck_date.year, deck_date.month
    for _ in range(window_months):
        grp = df[(df[ms15f_col].dt.month == m) & (df[ms15f_col].dt.year == y)]
        total = len(grp)
        if total > 0:
            with_ntp = grp[grp['has_ntp']]
            pending = grp[~grp['has_ntp']]
            pending_rows = []
            for _, r in pending.iterrows():
                cat = r.get('NTP_Cat', 'Other')
                pending_rows.append({
                    'HOP': r['HOP'], 'General Contractor': gv(r, 'General Contractor'),
                    'New CM': gv(r, 'New CM'), '_ntp_owner': gv(r, '_ntp_owner'),
                    '_ntp_wait': gv(r, '_ntp_wait'), '_cx': gv(r, '_cx'), 'cat': cat,
                    'MS15 Implementation Start F': r.get(ms15f_col), 'has_mat': bool(r['has_mat']),
                    '_path_id': str(r.get('_path_id', '')).strip(), '_pm': gv(r, '_pm'),
                })
            pending_rows.sort(key=lambda x: 0 if x['cat'] == 'External'
                               else (2 if x['cat'] == 'Program Team' else 1))
            mo_key = f'{y}-{m:02d}'
            mo_name = MONTH_FULL[m]
            sheet_key = f'{MONTH_ABBR[m]} Pending NTP'
            por[mo_key] = {
                'total': total, 'ntp': len(with_ntp), 'pending': len(pending),
                'ntp_hops': with_ntp[[c for c in ['HOP', 'General Contractor', 'New CM', 'has_mat',
                                                   '_ntp_wait', '_cx', '_pm', '_ntp_owner',
                                                   'MS15 Implementation Start F', 'MS16 Implementation Ends F']
                                      if c in with_ntp.columns]].to_dict('records'),
                'pending_rows': pending_rows,
                'external': [r for r in pending_rows if r['cat'] == 'External'],
                'prog_team': [r for r in pending_rows if r['cat'] == 'Program Team'],
                'other': [r for r in pending_rows if r['cat'] == 'Other'],
            }
            order.append((mo_key, mo_name, sheet_key))
            if len(order) >= max_active:
                break
        m += 1
        if m > 12:
            m, y = 1, y + 1
    return por, order


# ─────────────────────────────────────────────
# 3. CHART PLAN PASS-THROUGH — read the existing Plan series instead of a
#    hardcoded table, so CJ's manual Plan edits survive every rebuild.
# ─────────────────────────────────────────────

def read_plan_series(chart_xml_bytes: bytes) -> dict:
    """Returns {category_label: plan_value} for whatever is currently in the
    chart's Plan series (column B per fix_chart_v2's convention below).
    Missing <c:pt> entries mean 0 (numCache only stores points > 0)."""
    try:
        root = etree.fromstring(chart_xml_bytes)
    except Exception:
        return {}
    out = {}
    for ser in root.findall(f'.//{{{C_NS}}}ser'):
        tx_vs = [e.text for e in ser.findall(f'.//{{{C_NS}}}tx//{{{C_NS}}}v') if e.text]
        is_plan = 'Plan' in tx_vs
        val_el = ser.find(f'{{{C_NS}}}val')
        cat_el = ser.find(f'{{{C_NS}}}cat')
        if val_el is None or cat_el is None:
            continue
        ref = val_el.find(f'{{{C_NS}}}numRef')
        col = None
        if ref is not None:
            f_el = ref.find(f'{{{C_NS}}}f')
            if f_el is not None and f_el.text:
                m = re.search(r'\$([A-Z])\$', f_el.text)
                col = m.group(1) if m else None
        if not (col == 'B' or is_plan):
            continue
        cat_cache = cat_el.find(f'.//{{{C_NS}}}strCache')
        labels = {}
        if cat_cache is not None:
            for pt in cat_cache.findall(f'{{{C_NS}}}pt'):
                idx = int(pt.get('idx'))
                v = pt.find(f'{{{C_NS}}}v')
                labels[idx] = v.text if v is not None else ''
        num_cache = val_el.find(f'.//{{{C_NS}}}numCache')
        count_el = num_cache.find(f'{{{C_NS}}}ptCount') if num_cache is not None else None
        n = int(count_el.get('val')) if count_el is not None else max(labels.keys(), default=-1) + 1
        values = [0] * n
        if num_cache is not None:
            for pt in num_cache.findall(f'{{{C_NS}}}pt'):
                idx = int(pt.get('idx'))
                v = pt.find(f'{{{C_NS}}}v')
                if idx < n and v is not None and v.text:
                    values[idx] = float(v.text)
        for i in range(n):
            label = labels.get(i)
            if label is not None:
                out[label] = values[i]
        break  # only one Plan series per chart
    return out


# ─────────────────────────────────────────────
# 4. CROSS-DECK SLIDE COPY — appends slides (with their charts/images/embeds)
#    from a donor .pptx (built client-side by decomDeck.ts / scopDeck.ts) into
#    the target zip `content` dict. Mirrors the zip/XML-surgery technique
#    update_deck() already uses for slide deletion, just the insert direction.
#    Donor slides are re-pointed at the TARGET's own (single) slide layout —
#    exactly what PowerPoint itself does when CJ pastes a slide in by hand
#    today (confirmed: every manually-pasted slide in his working deck already
#    sits on the target's one layout, not the donor's).
# ─────────────────────────────────────────────

def _ordered_slide_paths(content: dict):
    prs_root = etree.fromstring(content['ppt/presentation.xml'])
    rels_root = etree.fromstring(content['ppt/_rels/presentation.xml.rels'])
    rid_to_target = {r.get('Id'): r.get('Target') for r in rels_root.findall(f'{{{PKG_NS}}}Relationship')}
    sld_id_lst = prs_root.find(f'.//{{{P_NS}}}sldIdLst')
    paths = []
    for sld_id in (sld_id_lst.findall(f'{{{P_NS}}}sldId') if sld_id_lst is not None else []):
        rid = sld_id.get(f'{{{R_NS}}}id')
        target = rid_to_target.get(rid, '')
        path = target if target.startswith('ppt/') else 'ppt/' + target
        paths.append(path)
    return paths


def _drop_dangling_refs(xml: str, valid_rids: set) -> str:
    """Removes any graphic frame whose r:id points at a relationship that doesn't exist.
    A dangling id makes PowerPoint repair the file."""
    root = etree.fromstring(xml.encode('utf-8'))
    P_NS_ = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    for frame in list(root.iter(f'{{{P_NS_}}}graphicFrame')):
        refs = [v for el in frame.iter() for k, v in el.attrib.items() if k == f'{{{R_NS}}}id' or k == f'{{{R_NS}}}embed']
        if any(r not in valid_rids for r in refs):
            frame.getparent().remove(frame)
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True).decode('utf-8')


def _resolve_part(base_dir: str, target: str):
    """Resolves a relationship target (relative or absolute) to a package part name."""
    if target.startswith('/'):
        return target.lstrip('/')
    return posixpath.normpath(posixpath.join(base_dir, target))


def _rels_path_for(part_path: str) -> str:
    d, f = part_path.rsplit('/', 1)
    return f'{d}/_rels/{f}.rels'


def _ensure_content_type_default(content: dict, ext: str):
    if ext not in EXT_CONTENT_TYPES:
        return
    ct_root = etree.fromstring(content['[Content_Types].xml'])
    for d in ct_root.findall(f'{{{CT_NS}}}Default'):
        if d.get('Extension', '').lower() == ext.lower():
            return
    d = etree.SubElement(ct_root, f'{{{CT_NS}}}Default')
    d.set('Extension', ext)
    d.set('ContentType', EXT_CONTENT_TYPES[ext])
    content['[Content_Types].xml'] = etree.tostring(ct_root, xml_declaration=True, encoding='UTF-8', standalone=True)


def _add_content_type_override(content: dict, part_path: str, content_type: str):
    ct_root = etree.fromstring(content['[Content_Types].xml'])
    o = etree.SubElement(ct_root, f'{{{CT_NS}}}Override')
    o.set('PartName', '/' + part_path)
    o.set('ContentType', content_type)
    content['[Content_Types].xml'] = etree.tostring(ct_root, xml_declaration=True, encoding='UTF-8', standalone=True)


_CHART_CT = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
_SLIDE_CT = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
_REL_SLIDE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
_REL_LAYOUT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
_REL_CHART = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
_REL_PACKAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package'


def copy_slides_into(content: dict, donor_bytes: bytes, donor_slide_positions, insert_after_path: str):
    """Mutates `content` in place: appends the donor's slides (by 0-based
    position in the donor's own slide order) right after `insert_after_path`
    (an existing target slide's zip path), preserving their charts/tables.
    Returns the list of new target slide zip paths, in the order inserted.
    """
    with zipfile.ZipFile(__import__('io').BytesIO(donor_bytes)) as z:
        donor = {n: z.read(n) for n in z.namelist()}

    def _slide_cx(pkg):
        root = etree.fromstring(pkg['ppt/presentation.xml'])
        sz = root.find(f'{{{P_NS}}}sldSz')
        return int(sz.get('cx')) if sz is not None else None
    donor_cx = _slide_cx(donor)
    target_cx = _slide_cx(content)
    scale_f = (target_cx / donor_cx) if (donor_cx and target_cx) else 1.0

    donor_slide_paths = _ordered_slide_paths(donor)

    # Figure out the next free numeric suffix for every part type we might add,
    # so new parts never collide with the target's existing ones.
    def next_index(prefix_dir, prefix_name):
        nums = [0]
        for k in content:
            m = re.match(rf'{re.escape(prefix_dir)}/{re.escape(prefix_name)}(\d+)\.', k)
            if m:
                nums.append(int(m.group(1)))
        return max(nums) + 1

    target_prs_root = etree.fromstring(content['ppt/presentation.xml'])
    target_rels_root = etree.fromstring(content['ppt/_rels/presentation.xml.rels'])

    new_slide_paths = []

    for pos in donor_slide_positions:
        donor_slide_path = donor_slide_paths[pos]
        donor_slide_xml = donor[donor_slide_path]
        donor_rels_path = _rels_path_for(donor_slide_path)
        donor_rels = donor.get(donor_rels_path, b'')

        slide_xml_str = scale_slide_xml(donor_slide_xml.decode('utf-8'), scale_f)
        rid_remap = {}

        if donor_rels:
            d_rels_root = etree.fromstring(donor_rels)
            for rel in d_rels_root.findall(f'{{{PKG_NS}}}Relationship'):
                rtype, rid, target = rel.get('Type', ''), rel.get('Id'), rel.get('Target', '')
                if rtype.endswith('/slideLayout'):
                    continue  # re-pointed at the target's own layout below
                donor_part = _resolve_part('ppt/slides', target)
                if donor_part is None or donor_part not in donor:
                    continue
                ext = donor_part.rsplit('.', 1)[-1]
                if rtype.endswith('/chart'):
                    n = next_index('ppt/charts', 'chartV2_')
                    new_part = f'ppt/charts/chartV2_{n}.xml'
                    _add_content_type_override(content, new_part, _CHART_CT)
                    chart_xml = donor[donor_part]
                    chart_rels_path = _rels_path_for(donor_part)
                    chart_rels = donor.get(chart_rels_path, b'')
                    chart_rid_remap = {}
                    if chart_rels:
                        c_rels_root = etree.fromstring(chart_rels)
                        for crel in c_rels_root.findall(f'{{{PKG_NS}}}Relationship'):
                            ctarget = crel.get('Target', '')
                            cdonor_part = _resolve_part('ppt/charts', ctarget)
                            if cdonor_part is None or cdonor_part not in donor:
                                continue
                            cext = cdonor_part.rsplit('.', 1)[-1]
                            en = next_index('ppt/embeddings', 'embedV2_')
                            new_embed = f'ppt/embeddings/embedV2_{en}.{cext}'
                            _ensure_content_type_default(content, cext)
                            content[new_embed] = donor[cdonor_part]
                            chart_rid_remap[crel.get('Id')] = ('../' + new_embed.replace('ppt/', ''), crel.get('Type'))
                        new_chart_rels_root = etree.Element(c_rels_root.tag, nsmap=c_rels_root.nsmap)
                        for crel in c_rels_root.findall(f'{{{PKG_NS}}}Relationship'):
                            old_id = crel.get('Id')
                            if old_id in chart_rid_remap:
                                new_target, rtype2 = chart_rid_remap[old_id]
                                new_rel = etree.SubElement(new_chart_rels_root, f'{{{PKG_NS}}}Relationship')
                                new_rel.set('Id', old_id)
                                new_rel.set('Type', rtype2)
                                new_rel.set('Target', new_target)
                        content[_rels_path_for(new_part)] = etree.tostring(
                            new_chart_rels_root, xml_declaration=True, encoding='UTF-8', standalone=True)
                    content[new_part] = chart_xml
                    rid_remap[rid] = (f'../charts/chartV2_{n}.xml', rtype)
                elif ext in EXT_CONTENT_TYPES:
                    n = next_index('ppt/media', 'imageV2_')
                    new_part = f'ppt/media/imageV2_{n}.{ext}'
                    _ensure_content_type_default(content, ext)
                    content[new_part] = donor[donor_part]
                    rid_remap[rid] = (f'../media/imageV2_{n}.{ext}', rtype)

        # Write the new slide's own rels: remapped parts + the target's layout.
        new_rels_root = etree.Element(f'{{{PKG_NS}}}Relationships')
        for old_id, (new_target, rtype) in rid_remap.items():
            rel = etree.SubElement(new_rels_root, f'{{{PKG_NS}}}Relationship')
            rel.set('Id', old_id)
            rel.set('Type', rtype)
            rel.set('Target', new_target)
        # Point the new slide at the target's own layout — reuse an existing
        # target slide's layout relationship id/target verbatim, same as
        # PowerPoint does when a slide is pasted in by hand.
        LAYOUT_RID_IN_NEW_SLIDE = 'rIdLayoutV2'
        layout_rel = etree.SubElement(new_rels_root, f'{{{PKG_NS}}}Relationship')
        layout_rel.set('Id', LAYOUT_RID_IN_NEW_SLIDE)
        layout_rel.set('Type', _REL_LAYOUT)
        existing_slide_paths = _ordered_slide_paths(content)
        sample_slide = existing_slide_paths[0]
        sample_rels = etree.fromstring(content[_rels_path_for(sample_slide)])
        sample_layout_target = None
        for r in sample_rels.findall(f'{{{PKG_NS}}}Relationship'):
            if r.get('Type', '').endswith('/slideLayout'):
                sample_layout_target = r.get('Target')
                break
        layout_rel.set('Target', sample_layout_target or '../slideLayouts/slideLayout1.xml')

        n_slide = next_index('ppt/slides', 'slideV2_')
        new_slide_path = f'ppt/slides/slideV2_{n_slide}.xml'

        # Rewrite rIds in the slide XML: old rId -> LAYOUT_RID_IN_NEW_SLIDE for the
        # layout ref (if present), and old chart/image rIds -> their new ids
        # (ids are kept the same, only Targets changed, so no text rewrite needed
        # for those — only the layout relationship, which pptxgenjs always emits
        # as the slide's one non-chart/image relationship, typically "rId1").
        if donor_rels:
            d_rels_root = etree.fromstring(donor_rels)
            for rel in d_rels_root.findall(f'{{{PKG_NS}}}Relationship'):
                if rel.get('Type', '').endswith('/slideLayout'):
                    old_layout_rid = rel.get('Id')
                    slide_xml_str = slide_xml_str.replace(f'r:id="{old_layout_rid}"', f'r:id="{LAYOUT_RID_IN_NEW_SLIDE}"')

        slide_xml_str = _drop_dangling_refs(slide_xml_str, set(rid_remap) | {LAYOUT_RID_IN_NEW_SLIDE})
        content[new_slide_path] = slide_xml_str.encode('utf-8')
        content[_rels_path_for(new_slide_path)] = etree.tostring(
            new_rels_root, xml_declaration=True, encoding='UTF-8', standalone=True)
        _add_content_type_override(content, new_slide_path, _SLIDE_CT)

        # Insert into presentation.xml's sldIdLst + presentation.xml.rels right
        # after `insert_after_path`.
        max_rid_num = max(
            [int(m.group(1)) for rid in [r.get('Id') for r in target_rels_root.findall(f'{{{PKG_NS}}}Relationship')]
             if rid and (m := re.match(r'rId(\d+)', rid))],
            default=0)
        new_rid = f'rId{max_rid_num + 1}'
        rel_el = etree.SubElement(target_rels_root, f'{{{PKG_NS}}}Relationship')
        rel_el.set('Id', new_rid)
        rel_el.set('Type', _REL_SLIDE)
        rel_el.set('Target', new_slide_path.replace('ppt/', ''))
        content['ppt/_rels/presentation.xml.rels'] = etree.tostring(
            target_rels_root, xml_declaration=True, encoding='UTF-8', standalone=True)

        sld_id_lst = target_prs_root.find(f'.//{{{P_NS}}}sldIdLst')
        max_sld_id = max([int(s.get('id')) for s in sld_id_lst.findall(f'{{{P_NS}}}sldId')], default=255)
        new_sld_id = etree.SubElement(sld_id_lst, f'{{{P_NS}}}sldId')
        new_sld_id.set('id', str(max_sld_id + 1))
        new_sld_id.set(f'{{{R_NS}}}id', new_rid)

        # Move it into position: right after insert_after_path's <p:sldId>.
        anchor_rid = None
        for r in target_rels_root.findall(f'{{{PKG_NS}}}Relationship'):
            if r.get('Target') == insert_after_path.replace('ppt/', ''):
                anchor_rid = r.get('Id')
                break
        if anchor_rid:
            sld_ids = list(sld_id_lst)
            anchor_idx = next((i for i, s in enumerate(sld_ids) if s.get(f'{{{R_NS}}}id') == anchor_rid), len(sld_ids) - 2)
            sld_id_lst.remove(new_sld_id)
            sld_id_lst.insert(anchor_idx + 1, new_sld_id)
            insert_after_path = new_slide_path  # next donor slide goes after this one

        content['ppt/presentation.xml'] = etree.tostring(
            target_prs_root, xml_declaration=True, encoding='UTF-8', standalone=True)

        new_slide_paths.append(new_slide_path)

    return new_slide_paths


# ─────────────────────────────────────────────
# 5. HEADER-DRIVEN TABLE HELPERS
#
# CJ's actual working template has already diverged from the original
# builder's hardcoded shape positions (e.g. his MSS Readiness slide now has
# one table instead of two, and both the Look-ahead and In Progress tables
# have had their PM/GC columns removed). Resolving columns by header text —
# the same technique update_por_confirmed() already uses in build_deck.py —
# makes V2 resilient to that kind of template edit instead of assuming a
# fixed layout that's already proven to drift.
# ─────────────────────────────────────────────

def find_table(slide, required_header_keywords):
    """Finds the first table on `slide` whose header row contains every
    keyword in required_header_keywords (case-insensitive substring match).
    Returns (shape, header_texts) or (None, None)."""
    for sh in slide.shapes:
        if not sh.has_table:
            continue
        hdrs = [sh.table.cell(0, c).text.strip() for c in range(len(sh.table.columns))]
        if all(any(kw.lower() in h.lower() for h in hdrs) for kw in required_header_keywords):
            return sh, hdrs
    return None, None


def col_idx(hdrs, *keywords, exact=False):
    """First column index whose header matches any of `keywords`. Tries an
    exact case-insensitive match first, then substring, so e.g. 'NTP' doesn't
    accidentally match 'NTP Waiting On' when an exact 'NTP' column exists."""
    for i, h in enumerate(hdrs):
        if any(h.strip().lower() == kw.lower() for kw in keywords):
            return i
    if exact:
        return None
    for i, h in enumerate(hdrs):
        if any(kw.lower() in h.lower() for kw in keywords):
            return i
    return None


def clear_extra_rows(shape, hdrs, from_row):
    for ri in range(from_row, len(shape.table.rows)):
        for ci in range(len(hdrs)):
            set_table_cell(shape, ri, ci, '')


MONTHS_RE = r'(January|February|March|April|May|June|July|August|September|October|November|December)'
TABLE_BOTTOM_EMU = 4850000  # footer bar starts at 4885403


def relabel_months(slide, mo_name):
    """Rewrites any 'Month POR' / 'Month Plan of Record' label on the slide to mo_name."""
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        for para in sh.text_frame.paragraphs:
            for run in para.runs:
                run.text = re.sub(MONTHS_RE + r'(?=\s+(?:POR|Plan of Record))', mo_name, run.text)


def set_para_text(para, text):
    runs = para.runs
    if runs:
        runs[0].text = text
        for r in runs[1:]:
            r.text = ''
    else:
        para.text = text


def _tighten_margins(tr):
    for tc in tr.findall(qn('a:tc')):
        tcPr = tc.find(qn('a:tcPr'))
        if tcPr is not None:
            tcPr.set('marT', '22860')
            tcPr.set('marB', '22860')


def fit_table(shape, bottom_emu=TABLE_BOTTOM_EMU, min_row_emu=160000):
    """Shrinks row heights so a table with extra rows stops above the footer."""
    if shape.top + shape.height <= bottom_emu:
        return
    trs = shape.table._tbl.findall(qn('a:tr'))
    if len(trs) < 2:
        return
    for tr in trs:
        _tighten_margins(tr)
    avail = bottom_emu - shape.top
    hdr_h = int(trs[0].get('h', min_row_emu))
    per = max(min_row_emu, (avail - hdr_h) // (len(trs) - 1))
    for tr in trs[1:]:
        tr.set('h', str(per))
    shape.height = hdr_h + per * (len(trs) - 1)


def stack_fit(tables, bottom_emu=TABLE_BOTTOM_EMU, gap_emu=120000, min_row_emu=160000):
    """Two stacked tables on one slide: shares the vertical space between them."""
    tables = sorted(tables, key=lambda t: t.top)
    total_rows = sum(len(t.table.rows) for t in tables)
    avail = bottom_emu - tables[0].top - gap_emu * (len(tables) - 1)
    per = max(min_row_emu, avail // total_rows)
    y = tables[0].top
    for t in tables:
        trs = t.table._tbl.findall(qn('a:tr'))
        for tr in trs:
            _tighten_margins(tr)
            tr.set('h', str(per))
        t.top = y
        t.height = per * len(trs)
        y += t.height + gap_emu


def fit_deck_tables(prs, bottom_emu=TABLE_BOTTOM_EMU):
    for sl in prs.slides:
        tbls = [sh for sh in sl.shapes if sh.has_table]
        if not tbls:
            continue
        if len(tbls) == 2 and max(t.top + t.height for t in tbls) > bottom_emu:
            stack_fit(tbls, bottom_emu)
        else:
            for t in tbls:
                fit_table(t, bottom_emu)


def scale_slide_xml(xml: str, f: float) -> str:
    """Scales a donor slide's geometry and font sizes by f (donor width -> target width)."""
    if abs(f - 1.0) < 1e-6:
        return xml
    def scale_tag(m, attrs):
        tag = m.group(0)
        for a in attrs:
            tag = re.sub(rf'\b{a}="(-?\d+)"', lambda mm: f'{a}="{int(round(int(mm.group(1)) * f))}"', tag)
        return tag
    xml = re.sub(r'<a:(?:off|chOff)\b[^>]*>', lambda m: scale_tag(m, ['x', 'y']), xml)
    xml = re.sub(r'<a:(?:ext|chExt)\b[^>]*>', lambda m: scale_tag(m, ['cx', 'cy']), xml)
    xml = re.sub(r'<a:gridCol\b[^>]*>', lambda m: scale_tag(m, ['w']), xml)
    xml = re.sub(r'<a:tr\b[^>]*>', lambda m: scale_tag(m, ['h']), xml)
    xml = re.sub(r'<a:(?:rPr|defRPr|endParaRPr)\b[^>]*>', lambda m: scale_tag(m, ['sz']), xml)
    return xml


def delete_slide_content(content: dict, slide_path: str):
    P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
    prs_root = etree.fromstring(content['ppt/presentation.xml'])
    rels_root = etree.fromstring(content['ppt/_rels/presentation.xml.rels'])
    rid = None
    for r in rels_root.findall(f'{{{PKG_NS}}}Relationship'):
        t = r.get('Target', '')
        full = t if t.startswith('ppt/') else 'ppt/' + t
        if full == slide_path:
            rid = r.get('Id')
            rels_root.remove(r)
            break
    if rid is None:
        return
    sld_id_lst = prs_root.find(f'.//{{{P_NS}}}sldIdLst')
    for sid in list(sld_id_lst.findall(f'{{{P_NS}}}sldId')):
        if sid.get(f'{{{R_NS}}}id') == rid:
            sld_id_lst.remove(sid)
            break
    content['ppt/presentation.xml'] = etree.tostring(prs_root, xml_declaration=True, encoding='UTF-8', standalone=True)
    content['ppt/_rels/presentation.xml.rels'] = etree.tostring(rels_root, xml_declaration=True, encoding='UTF-8', standalone=True)
    content.pop(slide_path, None)
    content.pop(_rels_path_for(slide_path), None)
    ct_root = etree.fromstring(content['[Content_Types].xml'])
    for ov in list(ct_root.findall(f'{{{CT_NS}}}Override')):
        if ov.get('PartName') == '/' + slide_path:
            ct_root.remove(ov)
            break
    content['[Content_Types].xml'] = etree.tostring(ct_root, xml_declaration=True, encoding='UTF-8', standalone=True)


# ─────────────────────────────────────────────
# 6. MAIN BUILD
# ─────────────────────────────────────────────

def update_deck_v2(data: dict, previous_deck_path: str, output_path: str,
                    last_snapshot: dict, supabase_client=None,
                    decom_pptx_bytes: bytes = None, scop_pptx_bytes: bytes = None,
                    call_dt: datetime = None):
    d = data
    NEW_DATE = d['deck_date']

    with zipfile.ZipFile(previous_deck_path, 'r') as z:
        content = {n: z.read(n) for n in z.namelist()}

    # The date to sweep away is whatever date is ACTUALLY written in the
    # previous deck — not the snapshot's session_date. Those two can get out
    # of sync (e.g. the very first V2 build has no snapshot yet, so nothing
    # gets swept that build, leaving the template's original date baked in;
    # trusting the snapshot on the next build would then search for a date
    # string that was never actually written). Scanning the deck itself for
    # its most common exact "M/D/YYYY"-shaped text is the ground truth.
    _date_counts = {}
    for _sp, _sb in content.items():
        if _sp.startswith('ppt/slides/slide') and _sp.endswith('.xml'):
            for _m in re.finditer(r'<a:t>(\d{1,2}/\d{1,2}/\d{4})</a:t>', _sb.decode('utf-8', errors='ignore')):
                _date_counts[_m.group(1)] = _date_counts.get(_m.group(1), 0) + 1
    OLD_DATE = max(_date_counts, key=_date_counts.get) if _date_counts else last_snapshot.get('session_date', '')

    # ── Cx Start / Cx Complete charts: Plan pass-through + fresh Forecast/Actual ──
    def _parse_rels(rels_bytes):
        try:
            root = etree.fromstring(rels_bytes)
            return [(r.get('Type', ''), r.get('Id', ''), r.get('Target', '')) for r in root]
        except Exception:
            return []

    def _rid_map(rels_bytes):
        return {rid: tgt for _, rid, tgt in _parse_rels(rels_bytes)}

    def _slide_chart(slide_path):
        slide_name = slide_path.split('/')[-1]
        for _type, _rid, _tgt in _parse_rels(content.get(f'ppt/slides/_rels/{slide_name}.rels', b'')):
            if '/chart' in _tgt or 'chart' in _type.lower():
                return f'ppt/charts/{_tgt.split("/")[-1]}'
        return None

    def _chart_embed(chart_path):
        chart_name = chart_path.split('/')[-1]
        for _type, _rid, _tgt in _parse_rels(content.get(f'ppt/charts/_rels/{chart_name}.rels', b'')):
            if '/embeddings/' in _tgt:
                return f'ppt/embeddings/{_tgt.split("/")[-1]}'
        return None

    ordered_slides = _ordered_slide_paths(content)
    starts_chart_path = complete_chart_path = None
    for sp in ordered_slides:
        sb = content.get(sp, b'')
        cp = _slide_chart(sp)
        if cp is None:
            continue
        if starts_chart_path is None and b'Cx Starts' in sb:
            starts_chart_path = cp
        if complete_chart_path is None and (b'Construction Complete' in sb or b'Cx Complete' in sb):
            complete_chart_path = cp
    if starts_chart_path is None or complete_chart_path is None:
        chart_list = list(dict.fromkeys(c for p in ordered_slides for c in [_slide_chart(p)] if c))
        if len(chart_list) >= 2:
            starts_chart_path = starts_chart_path or chart_list[0]
            complete_chart_path = complete_chart_path or chart_list[1]
    starts_chart_path = starts_chart_path or 'ppt/charts/chart1.xml'
    complete_chart_path = complete_chart_path or 'ppt/charts/chart2.xml'
    starts_embed_path = _chart_embed(starts_chart_path) or 'ppt/embeddings/Microsoft_Excel_Worksheet.xlsx'
    complete_embed_path = _chart_embed(complete_chart_path) or 'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx'

    def build_plan_vals(chart_path, month_labels):
        existing = read_plan_series(content.get(chart_path, b''))
        return [existing.get(lbl, 0) for lbl in month_labels]

    starts_plan = build_plan_vals(starts_chart_path, d['starts_labels'])
    complete_plan = build_plan_vals(complete_chart_path, d['complete_labels'])

    def fix_chart_v2(xml_bytes, fc_vals, act_vals, plan_vals, month_labels):
        xml = xml_bytes.decode('utf-8')
        n = len(month_labels)

        def build_str_cache(labels):
            pts = ''.join(f'<c:pt idx="{i}"><c:v>{l}</c:v></c:pt>' for i, l in enumerate(labels))
            return f'<c:strCache><c:ptCount val="{len(labels)}"/>{pts}</c:strCache>'

        str_caches = list(re.finditer(r'<c:strCache>.*?</c:strCache>', xml, re.DOTALL))
        for sc in reversed(str_caches):
            if len(re.findall(r'<c:pt ', sc.group())) > 1:
                xml = xml[:sc.start()] + build_str_cache(month_labels) + xml[sc.end():]
        xml = re.sub(r'(\$[A-Z]\$2:\$[A-Z]\$)\d+', lambda m: m.group(1) + str(n + 1), xml)

        try:
            lroot = etree.fromstring(xml.encode('utf-8'))

            def set_numdata(container, vals, tag):
                old = container.find(f'{{{C_NS}}}{tag}')
                if old is not None:
                    container.remove(old)
                el = etree.SubElement(container, f'{{{C_NS}}}{tag}')
                fce = etree.SubElement(el, f'{{{C_NS}}}formatCode'); fce.text = 'General'
                pce = etree.SubElement(el, f'{{{C_NS}}}ptCount'); pce.set('val', str(n))
                for i, v in enumerate(vals):
                    if v is not None and v > 0:
                        pt = etree.SubElement(el, f'{{{C_NS}}}pt'); pt.set('idx', str(i))
                        ve = etree.SubElement(pt, f'{{{C_NS}}}v'); ve.text = str(v)

            unresolved = []
            for ser in lroot.findall(f'.//{{{C_NS}}}ser'):
                tx_vs = [e.text for e in ser.findall(f'.//{{{C_NS}}}tx//{{{C_NS}}}v') if e.text]
                is_plan = 'Plan' in tx_vs
                val_el = ser.find(f'{{{C_NS}}}val')
                if val_el is None:
                    continue
                ref = val_el.find(f'{{{C_NS}}}numRef')
                lit = val_el.find(f'{{{C_NS}}}numLit')
                col = None
                if ref is not None:
                    fe = ref.find(f'{{{C_NS}}}f')
                    if fe is not None and fe.text:
                        cm = re.search(r'\$([A-Z])\$', fe.text)
                        col = cm.group(1) if cm else None
                if col == 'B' or is_plan:
                    nv = plan_vals
                elif col == 'C':
                    nv = fc_vals
                elif col == 'D':
                    nv = act_vals
                else:
                    if not is_plan:
                        idx_el = ser.find(f'{{{C_NS}}}idx')
                        oi = int(idx_el.get('val', '99')) if idx_el is not None else 99
                        unresolved.append((oi, ser, ref, lit))
                    continue
                con = ref if ref is not None else lit
                if con is not None:
                    set_numdata(con, nv, 'numCache' if ref is not None else 'numLit')
                if not is_plan:
                    dlbls_el = ser.find(f'{{{C_NS}}}dLbls')
                    if dlbls_el is not None:
                        for dlbl in list(dlbls_el.findall(f'{{{C_NS}}}dLbl')):
                            dlbls_el.remove(dlbl)

            unresolved.sort(key=lambda x: x[0])
            for fi, (oi, ser, ref, lit) in enumerate(unresolved[:2]):
                fv = fc_vals if fi == 0 else act_vals
                con = ref if ref is not None else lit
                if con is not None:
                    set_numdata(con, fv, 'numCache' if ref is not None else 'numLit')

            decl = xml[:xml.index('?>') + 2] if xml.startswith('<?xml') else ''
            xml = etree.tostring(lroot, encoding='unicode')
            if decl and not xml.startswith('<?xml'):
                xml = decl + xml
        except Exception as e:
            print(f'[v2-chart] numCache update failed: {e}', flush=True)

        xml = re.sub(r'<c:externalData\b[^>]*/>', '', xml)
        xml = re.sub(r'<c:externalData\b[^>]*>.*?</c:externalData>', '', xml, flags=re.DOTALL)
        return xml.encode('utf-8')

    if starts_chart_path in content:
        content[starts_chart_path] = fix_chart_v2(
            content[starts_chart_path], d['cx_starts_fc'], d['cx_starts_act'], starts_plan, d['starts_labels'])
    if complete_chart_path in content:
        content[complete_chart_path] = fix_chart_v2(
            content[complete_chart_path], d['cx_complete_fc'], d['cx_complete_act'], complete_plan, d['complete_labels'])

    import io as _io
    import openpyxl as _openpyxl
    for ep, fc_v, act_v, plan_v, mo_lbl in [
        (starts_embed_path, d['cx_starts_fc'], d['cx_starts_act'], starts_plan, d['starts_labels']),
        (complete_embed_path, d['cx_complete_fc'], d['cx_complete_act'], complete_plan, d['complete_labels']),
    ]:
        if ep in content:
            wb = _openpyxl.load_workbook(_io.BytesIO(content[ep]))
            ws = wb.active
            for ri in range(2, ws.max_row + 2):
                for ci in range(1, 5):
                    ws.cell(row=ri, column=ci).value = None
            ws.cell(row=1, column=4).value = 'Plan'
            for ri, (lbl, fc, act, pv) in enumerate(zip(mo_lbl, fc_v, act_v, plan_v), 2):
                ws.cell(row=ri, column=1).value = lbl
                ws.cell(row=ri, column=2).value = pv if pv > 0 else None
                ws.cell(row=ri, column=3).value = fc if fc > 0 else None
                ws.cell(row=ri, column=4).value = act if act > 0 else None
            out = _io.BytesIO(); wb.save(out)
            content[ep] = out.getvalue()

    # ── POR rolling window ──
    por, por_order = compute_por_window(d['df'], pd.Timestamp(datetime.strptime(NEW_DATE, '%m/%d/%Y')))

    # Save to a tmp file for python-pptx-level edits.
    tmp_path = output_path + '.tmp1.pptx'
    with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, b in content.items():
            z.writestr(n, b)

    prs = Presentation(tmp_path)

    # Global date sweep: swaps the old build date inside text, and also any
    # date-stamp box that holds only a date (stale stamps from older builds).
    if OLD_DATE:
        for slide in prs.slides:
            for shape in slide.shapes:
                replace_text_in_shape(shape, OLD_DATE, NEW_DATE)
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame and re.fullmatch(r'\d{1,2}/\d{1,2}/\d{4}', shape.text_frame.text.strip() or 'x'):
                set_shape_text(shape, NEW_DATE)

    def slide_by_title(fragment):
        for s in prs.slides:
            for sh in s.shapes:
                if sh.has_text_frame and fragment.lower() in sh.text_frame.text.lower():
                    return s
        return None

    # ── Delta slide ("What Changed Dashboard") ──
    prev_session = last_snapshot.get('session_date', '')
    delta_slide = slide_by_title('What Changed')
    if delta_slide is not None:
        shapes4 = list(delta_slide.shapes)
        delta_starts = d['started_count'] - last_snapshot.get('total_starts', 0)
        delta_complete = d['complete_count'] - last_snapshot.get('total_complete', 0)
        delta_ip = d['ip_count'] - last_snapshot.get('in_progress', 0)
        delta_ntp = d['ntp_count'] - last_snapshot.get('total_ntp', 0)
        ref = prev_session or 'first build'

        def ds(v, ref):
            return f'+{v} vs {ref}' if v > 0 else (f'No change vs {ref}' if v == 0 else f'{v} vs {ref}')

        for idx, val in [(6, d['started_count']), (11, d['complete_count']), (16, d['ip_count']), (21, d['ntp_count'])]:
            if idx < len(shapes4):
                set_shape_text(shapes4[idx], str(val))
        for idx, val in [(9, delta_starts), (14, delta_complete), (19, delta_ip), (24, delta_ntp)]:
            if idx < len(shapes4):
                set_shape_text(shapes4[idx], ds(val, ref))

        for sh in delta_slide.shapes:
            if not sh.has_text_frame:
                continue
            t = sh.text_frame.text
            if 'Compared to' in t:
                set_shape_text(sh, f'Compared to {prev_session or "first build"}  ·  {NEW_DATE}')
            elif 'Construction Completed' in t:
                total = d['total']
                done = d['complete_count']
                pct = round(100 * done / total) if total else 0
                new_t = re.sub(r'\d+ of \d+ HOPs \(\d+%\)', f'{done} of {total} HOPs ({pct}%)', t)
                new_t = re.sub(r'– \d+ HOPs', f'– {total - done} HOPs', new_t)
                set_shape_text(sh, new_t)

        # "NTP Start Changes by Month" lines — one per active POR month, in order.
        month_lines = []
        for shp in delta_slide.shapes:
            if not shp.has_text_frame:
                continue
            for para in shp.text_frame.paragraphs:
                if re.search(r'POR:\s+\d+ of \d+', para.text) or re.search(r'POR:\s+\d+ of \d+ Hops', para.text):
                    month_lines.append(para)
        for k, para in enumerate(month_lines):
            if k < len(por_order):
                mo_key, mo_name, _ = por_order[k]
                pp = por[mo_key]
                set_para_text(para, f'{mo_name} POR:  {pp["ntp"]} of {pp["total"]} HOPs with NTP')
            else:
                set_para_text(para, '')

    # ── POR slides: fill whatever trio-groups exist in the template, in
    #    rolling-window order; blank any trailing group beyond the active
    #    month count (handles the project having fewer months left than the
    #    template has slide-groups for). ──
    por_title_slides = [s for s in prs.slides if any(
        sh.has_text_frame and re.search(r'Forecasted\s+\S+\s+POR', sh.text_frame.text) for sh in s.shapes)]

    def update_por_overview(slide, mo_name, mo_key):
        p = por[mo_key]
        shapes = list(slide.shapes)
        for sh in shapes:
            if sh.has_text_frame and re.search(r'\d+\s+Forecasted\s+\S+\s+POR', sh.text_frame.text):
                set_shape_text(sh, f'{p["total"]} Forecasted {mo_name} POR')
                break
        for sh in shapes:
            if sh.has_text_frame and re.search(r'\S+\s+Plan of Record', sh.text_frame.text):
                set_shape_text(sh, f'{mo_name} Plan of Record')
                break
        texts = [sh for sh in shapes if sh.has_text_frame]
        # Numeric KPI cards: match by the static label text directly beneath
        # each number (label shapes don't change build to build).
        label_to_val = {
            'POR HOPs': p['total'], 'With NTP': p['ntp'], 'Pending NTP': p['pending'],
            'Green - High Confidence': p['ntp'], 'Yellow - Medium Risk': len(p['prog_team']) + len(p['other']),
        }
        for sh in texts:
            t = sh.text_frame.text.strip()
            for label, val in label_to_val.items():
                if t == label:
                    # the number lives in the text shape immediately preceding this one in z-order
                    idx = shapes.index(sh)
                    if idx > 0 and shapes[idx - 1].has_text_frame:
                        set_shape_text(shapes[idx - 1], str(val))
        for sh in texts:
            if re.search(r'\d+\s+with\s+NTP\s+of\s+\d+\s+POR', sh.text_frame.text, re.I):
                set_shape_text(sh, f'{p["ntp"]} with NTP of {p["total"]} POR  ·  {p["pending"]} pending NTP')

    def update_por_confirmed(slide, mo_name, mo_key, sheet_key):
        p = por[mo_key]
        for sh in slide.shapes:
            if sh.has_text_frame and re.search(r'\d+\s+with\s+NTP\s+of\s+\d+\s+POR', sh.text_frame.text, re.I):
                set_shape_text(sh, f'{p["ntp"]} with NTP of {p["total"]} POR')
        comments = d['ntp_comments'].get(sheet_key, {})
        tbl_shape, hdrs = find_table(slide, ['HOP'])
        if tbl_shape is None:
            return
        hop_col = col_idx(hdrs, 'HOP')
        pm_col = col_idx(hdrs, 'PM')
        gc_col = col_idx(hdrs, 'GC')
        mat_col = col_idx(hdrs, 'Mat')
        ntp_col = col_idx(hdrs, 'NTP', exact=True)
        start_col = col_idx(hdrs, 'Start')
        end_col = col_idx(hdrs, 'End')
        owner_col = col_idx(hdrs, 'Owner', 'Action')
        comment_col = col_idx(hdrs, 'Note', 'Status', 'Comment', 'Wait')
        ntp_hops = sorted(p['ntp_hops'], key=lambda h: (
            pd.Timestamp(h['MS15 Implementation Start F'])
            if h.get('MS15 Implementation Start F') and pd.notna(h.get('MS15 Implementation Start F'))
            else pd.Timestamp('2099-01-01')))
        expand_table_rows(tbl_shape, len(ntp_hops))
        for ri in range(1, len(tbl_shape.table.rows)):
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
                ntp_wait = str(h.get('_ntp_wait', '')).strip()
                pm_val = str(h.get('_pm', '')).strip(); pm_val = '' if pm_val.lower() == 'nan' else pm_val
                owner_val = str(h.get('_ntp_owner', '')).strip(); owner_val = '' if owner_val.lower() == 'nan' else owner_val
                ms15f_val = h.get('MS15 Implementation Start F')
                ms16f_val = h.get('MS16 Implementation Ends F')
                if hop_col is not None: set_table_cell(tbl_shape, ri, hop_col, hop)
                if pm_col is not None: set_table_cell(tbl_shape, ri, pm_col, pm_val)
                if gc_col is not None: set_table_cell(tbl_shape, ri, gc_col, gc)
                if mat_col is not None:
                    set_table_cell(tbl_shape, ri, mat_col, mat_sym, color=GREEN_C if mat_sym == '✓' else RED_C, bold=True)
                if ntp_col is not None:
                    set_table_cell(tbl_shape, ri, ntp_col, '✓', color=GREEN_C, bold=True)
                if start_col is not None:
                    set_table_cell(tbl_shape, ri, start_col, fmt_dm(ms15f_val) if pd.notna(ms15f_val) else '')
                if end_col is not None:
                    set_table_cell(tbl_shape, ri, end_col, fmt_dm(ms16f_val) if pd.notna(ms16f_val) else '')
                if owner_col is not None: set_table_cell(tbl_shape, ri, owner_col, owner_val)
                if comment_col is not None:
                    cell_note = comment or ntp_wait
                    set_table_cell(tbl_shape, ri, comment_col, cell_note[:55] if cell_note else '')
            else:
                clear_extra_rows(tbl_shape, hdrs, ri)

    def update_por_pending(slide, mo_name, mo_key, sheet_key):
        p = por[mo_key]
        subtitle = next((sh for sh in slide.shapes if sh.has_text_frame and
                          re.search(r'\d+\s+of\s+\d+\s+pending', sh.text_frame.text, re.I)), None)
        if subtitle is not None:
            set_shape_text(subtitle, f'{p["pending"]} of {p["total"]} pending NTP')
        for sh in slide.shapes:
            if sh.has_text_frame and re.search(r'External Blockers', sh.text_frame.text):
                set_shape_text(sh, f'External Blockers ({len(p["external"])})  —  ITW · Samsung · Viaero')
            if sh.has_text_frame and re.search(r'Program Team Actions', sh.text_frame.text):
                set_shape_text(sh, f'Program Team Actions ({len(p["prog_team"]) + len(p["other"])})')
        tbl_shapes = [sh for sh in slide.shapes if sh.has_table]
        for tbl_shape in tbl_shapes:
            hdrs = [tbl_shape.table.cell(0, c).text.strip() for c in range(len(tbl_shape.table.columns))]
            if 'HOP' not in hdrs:
                continue
            hop_col = col_idx(hdrs, 'HOP')
            gc_col = col_idx(hdrs, 'GC')
            owner_col = col_idx(hdrs, 'Owner', 'Action')
            wait_col = col_idx(hdrs, 'Wait', 'Note')
            rows = p['external'] if any(sh.has_text_frame and 'External' in sh.text_frame.text for sh in slide.shapes) else p['prog_team'] + p['other']
            # Heuristic: first table = external, second = program team/other —
            # matches the slide's left/right layout convention.
            rows = p['external'] if tbl_shapes.index(tbl_shape) == 0 else (p['prog_team'] + p['other'])
            expand_table_rows(tbl_shape, len(rows))
            for ri in range(1, len(tbl_shape.table.rows)):
                if ri - 1 < len(rows):
                    r = rows[ri - 1]
                    if hop_col is not None: set_table_cell(tbl_shape, ri, hop_col, r['HOP'])
                    if gc_col is not None: set_table_cell(tbl_shape, ri, gc_col, r.get('General Contractor', ''))
                    if owner_col is not None: set_table_cell(tbl_shape, ri, owner_col, r.get('_ntp_owner', ''))
                    if wait_col is not None: set_table_cell(tbl_shape, ri, wait_col, (r.get('_ntp_wait', '') or '')[:60])
                else:
                    clear_extra_rows(tbl_shape, hdrs, ri)

    # Group POR-related slides into trios by their position (overview, then
    # the next slide with an NTP-count-of-POR subtitle but no "pending"
    # wording = confirmed list, then the "pending" one).
    por_overview_slides = [s for s in prs.slides if any(
        sh.has_text_frame and re.search(r'\d+\s+Forecasted\s+\S+\s+POR', sh.text_frame.text) for sh in s.shapes)]
    all_slides = list(prs.slides)
    blank_trio_bases = []
    for i, ov_slide in enumerate(por_overview_slides):
        base_idx = all_slides.index(ov_slide)
        confirmed_slide = all_slides[base_idx + 1] if base_idx + 1 < len(all_slides) else None
        pending_slide = all_slides[base_idx + 2] if base_idx + 2 < len(all_slides) else None
        if i < len(por_order):
            mo_key, mo_name, sheet_key = por_order[i]
            for sl in [ov_slide, confirmed_slide, pending_slide]:
                if sl is not None:
                    relabel_months(sl, mo_name)
            update_por_overview(ov_slide, mo_name, mo_key)
            if confirmed_slide is not None:
                update_por_confirmed(confirmed_slide, mo_name, mo_key, sheet_key)
            if pending_slide is not None:
                update_por_pending(pending_slide, mo_name, mo_key, sheet_key)
        else:
            # No active month for this template group (project nearly done) —
            # the whole 3-slide group is removed after the save below.
            blank_trio_bases.append(base_idx)

    if por_order and por_overview_slides:
        relabel_months(slide_by_title('Agenda') or prs.slides[1], por_order[0][1])

    # ── MSS Readiness / Look-ahead / In Progress — header-driven, resilient
    #    to CJ's already-simplified column sets on these tables. ──
    mss_slide = slide_by_title('MSS Readiness')
    if mss_slide is not None:
        mss_sorted = d['mss'].sort_values('MS15 Implementation Start A', ascending=True)
        mss_count = len(mss_sorted)
        for sh in mss_slide.shapes:
            if sh.has_text_frame and re.search(r'started \(last 7 days\)', sh.text_frame.text):
                set_shape_text(sh, re.sub(r'^\d+', str(mss_count), sh.text_frame.text))
            if sh.has_text_frame and re.search(r'Sites Started', sh.text_frame.text):
                set_shape_text(sh, f'Sites Started — MSS Ready or Ready Soon ({mss_count})')
        tbl_shape, hdrs = find_table(mss_slide, ['HOP'])
        if tbl_shape is not None:
            hop_col = col_idx(hdrs, 'HOP')
            gc_col = col_idx(hdrs, 'GC', 'Contractor')
            ops_col = col_idx(hdrs, 'Ops')
            status_col = col_idx(hdrs, 'Status', 'Readiness')
            start_col = col_idx(hdrs, 'Cx Start', 'Start')
            notes_col = col_idx(hdrs, 'Notes', 'Status', 'Comments')
            expand_table_rows(tbl_shape, mss_count)
            for ri in range(1, len(tbl_shape.table.rows)):
                if ri - 1 < mss_count:
                    r = mss_sorted.iloc[ri - 1]; hop = r['HOP']
                    if hop_col is not None: set_table_cell(tbl_shape, ri, hop_col, hop)
                    if gc_col is not None: set_table_cell(tbl_shape, ri, gc_col, gv(r, 'General Contractor'))
                    if ops_col is not None: set_table_cell(tbl_shape, ri, ops_col, d['hop_ops'].get(hop, ''))
                    if status_col is not None: set_table_cell(tbl_shape, ri, status_col, r.get('Readiness', ''))
                    if start_col is not None: set_table_cell(tbl_shape, ri, start_col, fmt_d(r.get('MS15 Implementation Start A')))
                else:
                    clear_extra_rows(tbl_shape, hdrs, ri)

    la_slide = slide_by_title('Construction Start Look ahead')
    if la_slide is not None:
        la_sorted = d['la'].sort_values('MS15 Implementation Start F')
        la_count = len(la_sorted)
        for sh in la_slide.shapes:
            if sh.has_text_frame and re.search(r'\d+.*start', sh.text_frame.text, re.I):
                set_shape_text(sh, re.sub(r'\d+', str(la_count), sh.text_frame.text, count=1))
                break
        tbl_shape, hdrs = find_table(la_slide, ['HOP'])
        if tbl_shape is not None:
            hop_col = col_idx(hdrs, 'HOP')
            gc_col = col_idx(hdrs, 'GC', 'Contractor')
            ntp_col = col_idx(hdrs, 'NTP', exact=True)
            start_col = col_idx(hdrs, 'Fc Start', 'Start')
            cm_col = col_idx(hdrs, 'Site CM', 'CM')
            ops_col = col_idx(hdrs, 'Ops')
            wait_col = col_idx(hdrs, 'Waiting')
            comment_col = col_idx(hdrs, 'Comment')
            expand_table_rows(tbl_shape, la_count)
            for ri in range(1, len(tbl_shape.table.rows)):
                if ri - 1 < la_count:
                    r = la_sorted.iloc[ri - 1]; hop = r['HOP']
                    ntp_sym = '✓' if r.get('has_ntp', False) else '✗'
                    if hop_col is not None: set_table_cell(tbl_shape, ri, hop_col, hop)
                    if gc_col is not None: set_table_cell(tbl_shape, ri, gc_col, gv(r, 'General Contractor'))
                    if ntp_col is not None:
                        set_table_cell(tbl_shape, ri, ntp_col, ntp_sym, color=GREEN_C if ntp_sym == '✓' else RED_C, bold=True)
                    if start_col is not None: set_table_cell(tbl_shape, ri, start_col, fmt_d(r.get('MS15 Implementation Start F')))
                    if cm_col is not None: set_table_cell(tbl_shape, ri, cm_col, d['hop_site_cm'].get(hop, ''))
                    if ops_col is not None: set_table_cell(tbl_shape, ri, ops_col, d['hop_ops'].get(hop, ''))
                    if wait_col is not None: set_table_cell(tbl_shape, ri, wait_col, gv(r, '_ntp_wait')[:200])
                    if comment_col is not None: set_table_cell(tbl_shape, ri, comment_col, gv(r, '_cx')[:200])
                else:
                    clear_extra_rows(tbl_shape, hdrs, ri)

    ip_slide = slide_by_title('Hops in Progress')
    if ip_slide is not None:
        ip_sorted = d['ip_df'].sort_values('MS16 Implementation Ends F', na_position='last')
        for sh in ip_slide.shapes:
            if sh.has_text_frame and re.search(r'HOPs started', sh.text_frame.text):
                set_shape_text(sh, f'{d["ip_count"]} HOPs started · Green=On Track · Yellow=At Risk · Red=Escalation · ⚑=Needs Attention')
        tbl_shape, hdrs = find_table(ip_slide, ['HOP'])
        if tbl_shape is not None:
            hop_col = col_idx(hdrs, 'HOP')
            start_col = col_idx(hdrs, 'Cx Start', 'Start')
            end_col = col_idx(hdrs, 'Fc End', 'End')
            risk_col = col_idx(hdrs, 'Risk')
            notes_col = col_idx(hdrs, 'Comments', 'Live Status', 'Notes')
            expand_table_rows(tbl_shape, len(ip_sorted))
            new_starts_set = set(d['new_starts'])
            for ri in range(1, len(tbl_shape.table.rows)):
                if ri - 1 < len(ip_sorted):
                    r = ip_sorted.iloc[ri - 1]; hop = r['HOP']
                    is_new = hop in new_starts_set
                    o18 = bool(r.get('over_18d', False))
                    risk = '🔴 Over 18d' if o18 else ('★ NEW' if is_new else 'G')
                    if hop_col is not None: set_table_cell(tbl_shape, ri, hop_col, f'★ {hop}' if is_new else hop)
                    if start_col is not None: set_table_cell(tbl_shape, ri, start_col, fmt_dm(r.get('MS15 Implementation Start A')))
                    if end_col is not None: set_table_cell(tbl_shape, ri, end_col, fmt_ds(r.get('MS16 Implementation Ends F')))
                    if risk_col is not None: set_table_cell(tbl_shape, ri, risk_col, risk)
                    if notes_col is not None: set_table_cell(tbl_shape, ri, notes_col, str(r.get('_cx', ''))[:120])
                else:
                    clear_extra_rows(tbl_shape, hdrs, ri)

    # ── New Starts & Completions Since Last Meeting ──
    ns_slide = slide_by_title('New Starts & Completions')
    ns_slide_idx = next((i for i, sl in enumerate(prs.slides) if ns_slide is not None and sl.slide_id == ns_slide.slide_id), None)
    if ns_slide is not None:
        tbl_shapes = [sh for sh in ns_slide.shapes if sh.has_table]
        new_starts_rows = [{'HOP': h, 'GC': d['hop_gc_pm'].get(h, ''), 'CM': d['hop_site_cm'].get(h, '')}
                            for h in d['new_starts']]
        completions_rows = [{'HOP': h, 'GC': d['hop_gc_pm'].get(h, ''), 'CM': d['hop_site_cm'].get(h, '')}
                             for h in d['completions']]
        for i, tbl_shape in enumerate(sorted(tbl_shapes, key=lambda s: s.left)):
            rows = new_starts_rows if i == 0 else completions_rows
            expand_table_rows(tbl_shape, max(len(rows), 1))
            for ri in range(1, len(tbl_shape.table.rows)):
                if ri - 1 < len(rows):
                    r = rows[ri - 1]
                    set_table_cell(tbl_shape, ri, 0, r['HOP'])
                    set_table_cell(tbl_shape, ri, 1, r['GC'])
                    set_table_cell(tbl_shape, ri, 2, r['CM'])
                else:
                    clear_extra_rows(tbl_shape, ['HOP', 'GC', 'CM'], ri)
        subtitle = next((sh for sh in ns_slide.shapes if sh.has_text_frame and 'HOPs that started' in sh.text_frame.text), None)
        if subtitle is not None:
            set_shape_text(subtitle, f'{len(d["new_starts"])} new starts, {len(d["completions"])} completions since the last build')

    prs.save(output_path + '.tmp2.pptx')
    os.remove(tmp_path)

    # ── Append the Decom Dashboard + SCOP Status Deck slides ──
    with zipfile.ZipFile(output_path + '.tmp2.pptx', 'r') as z:
        content2 = {n: z.read(n) for n in z.namelist()}
    os.remove(output_path + '.tmp2.pptx')
    cur_paths = _ordered_slide_paths(content2)
    remove_idx = set()
    for b in blank_trio_bases:
        remove_idx.update({b, b + 1, b + 2})
    if ns_slide_idx is not None:
        remove_idx.add(ns_slide_idx)
    for k in sorted(remove_idx):
        if k < len(cur_paths):
            delete_slide_content(content2, cur_paths[k])
        # paths shift after deletion; nothing else depends on them below

    # Reopen once to map title -> slide index by position.
    tmp3 = output_path + '.tmp3.pptx'
    with zipfile.ZipFile(tmp3, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, b in content2.items():
            z.writestr(n, b)
    prs3 = Presentation(tmp3)
    titles = []
    for s in prs3.slides:
        t = ''
        for sh in s.shapes:
            if sh.has_text_frame and sh.text_frame.text.strip():
                t = sh.text_frame.text.strip()
                break
        titles.append(t)
    paths2 = _ordered_slide_paths(content2)
    os.remove(tmp3)

    # A previous build's output already carries these merged slides — drop any
    # copies so the fresh ones below are the only set in the deck.
    MERGED_TITLES = {
        'Decom Status — Program Overview', 'Overall SCOP Status — Program Summary',
        'Top Aging Sites — Drop Off Pending', 'GC Decom Accountability — Status by Contractor',
        'Pathwave SCOP — Contractor Action Items', 'QuickBase Status',
    }
    stale = [i for i, t in enumerate(titles) if t.split('\n')[0].strip() in MERGED_TITLES]
    for i in sorted(stale, reverse=True):
        delete_slide_content(content2, paths2[i])
    if stale:
        titles = [t for i, t in enumerate(titles) if i not in set(stale)]
        paths2 = _ordered_slide_paths(content2)

    action_items_idx = next((i for i, t in enumerate(titles) if 'Action Items Log' in t), None)

    def _title_for(fragment, exact=False):
        """Re-derives a slide's current zip path by title, fresh — needed
        after any insertion shifts every later slide's position, so a stale
        pre-insertion index can't be reused (that was the original bug: Thank
        You's index from before group 1 was inserted no longer pointed at
        Thank You once 2 slides had been spliced in ahead of it)."""
        tmp = output_path + '.tmp4.pptx'
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z:
            for n, b in content2.items():
                z.writestr(n, b)
        p = Presentation(tmp)
        cur_paths = _ordered_slide_paths(content2)
        result = None
        for i, s in enumerate(p.slides):
            t = ''
            for sh in s.shapes:
                if sh.has_text_frame and sh.text_frame.text.strip():
                    t = sh.text_frame.text.strip()
                    break
            if (t.strip() == fragment) if exact else (fragment in t):
                result = cur_paths[i]
                break
        os.remove(tmp)
        return result

    if decom_pptx_bytes and scop_pptx_bytes and action_items_idx is not None:
        # Group 1: Decom slide 0 (Program Overview) + SCOP slide 2 (Overall
        # Summary) — spliced together right before Action Items Log.
        insert_after = paths2[action_items_idx - 1]
        new_paths = copy_slides_into(content2, decom_pptx_bytes, [0], insert_after)
        insert_after = new_paths[-1]
        copy_slides_into(content2, scop_pptx_bytes, [2], insert_after)

        # Group 2: Decom slides 1,2 (Top Aging, GC Accountability) + SCOP
        # slides 1,0 (Pathwave SCOP, QuickBase Status) — appended after
        # Thank You, matching CJ's confirmed working order. Thank You's path
        # is re-derived fresh since group 1's insertion shifted it.
        insert_after = _title_for('Thank You', exact=True)
        if insert_after:
            new_paths = copy_slides_into(content2, decom_pptx_bytes, [1, 2], insert_after)
            insert_after = new_paths[-1]
            copy_slides_into(content2, scop_pptx_bytes, [1, 0], insert_after)

    final_tmp = output_path + '.final.pptx'
    with zipfile.ZipFile(final_tmp, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, b in content2.items():
            z.writestr(n, b)
    final_prs = Presentation(final_tmp)
    fit_deck_tables(final_prs)
    final_prs.save(output_path)
    os.remove(final_tmp)

    # Persist this build's headline numbers as the baseline for next time.
    snapshot = {
        'session_date': NEW_DATE,
        'total_starts': d['started_count'], 'total_complete': d['complete_count'],
        'in_progress': d['ip_count'], 'total_ntp': d['ntp_count'],
        'ip_hops': d['ip_df']['HOP'].tolist(),
    }
    for mo_key, mo_name, sheet_key in por_order:
        snapshot[f'por_{mo_key}_ntp'] = por[mo_key]['ntp']
    save_snapshot(supabase_client, call_dt or datetime.strptime(NEW_DATE, '%m/%d/%Y'), snapshot)

    return output_path


def build(tracker_path: str = '', previous_deck_path: str = '', deck_date: str = '',
          output_dir: str = '', tracker_rows=None, supabase_client=None,
          decom_pptx_bytes: bytes = None, scop_pptx_bytes: bytes = None) -> dict:
    os.makedirs(output_dir, exist_ok=True)
    call_dt = call_tuesday(deck_date)
    deck_date = call_dt.strftime('%m/%d/%Y')
    date_slug = deck_date.replace('/', '-')

    # The prior Tuesday's saved numbers, fed to extract_data via a throwaway
    # JSON file matching its snapshot_path format exactly, so new_starts and
    # completions compare call to call.
    last_snapshot = load_prior_snapshot(supabase_client, call_dt)
    snapshot_json_path = os.path.join(output_dir, '_last_snapshot.json')
    with open(snapshot_json_path, 'w') as f:
        json.dump(last_snapshot, f)

    data = extract_data(tracker_path, snapshot_json_path, '', deck_date, tracker_rows=tracker_rows, prev_snapshot_data=None)
    os.remove(snapshot_json_path)

    deck_out = os.path.join(output_dir, f'Viaero_Construction_Update_V2_{date_slug}.pptx')
    update_deck_v2(data, previous_deck_path, deck_out, last_snapshot, supabase_client, decom_pptx_bytes, scop_pptx_bytes, call_dt)

    return {
        'deck_path': deck_out,
        'summary': {
            'deck_date': deck_date,
            'total_hops': data['total'], 'ntp_count': data['ntp_count'],
            'started_count': data['started_count'], 'complete_count': data['complete_count'],
            'ip_count': data['ip_count'], 'new_starts': len(data['new_starts']),
            'completions': len(data['completions']),
        }
    }
