import io
import os
import sys
import json
import base64
import shutil
import zipfile
import tempfile
import traceback
from datetime import datetime, timedelta

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
from supabase import create_client

# Ensure backend package path is importable when running this file directly
sys.path.append(os.path.dirname(__file__))
import build_deck
import build_deck_v2
import ntp_comments_v2

app = Flask(__name__)
CORS(app)
# Trackers run well past 10MB for this program; cap well above that so a
# legitimately large upload never gets cut off, while still bounding worst-case
# memory per request.
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024

SUPABASE_URL = os.environ.get('SUPABASE_URL', '')
SUPABASE_KEY = os.environ.get('SUPABASE_KEY', '')
supabase_client = create_client(SUPABASE_URL, SUPABASE_KEY) if SUPABASE_URL and SUPABASE_KEY else None


def _save_uploaded_file(uploaded, dest_dir, field_name):
	if not uploaded:
		return ''
	filename = secure_filename(uploaded.filename) or field_name
	out_path = os.path.join(dest_dir, filename)
	uploaded.save(out_path)
	return out_path


@app.route('/build', methods=['POST'])
def build_endpoint():
	tmpdir = tempfile.mkdtemp(prefix='ciege_build_')
	try:
		# Read uploaded files
		tracker = request.files.get('tracker')
		previous_deck = request.files.get('previous_deck')
		snapshot = request.files.get('snapshot')
		ntp_comments = request.files.get('ntp_comments')
		deck_date = request.form.get('deck_date') or request.args.get('deck_date')

		if not previous_deck or not deck_date:
			return jsonify({'error': 'Missing required fields: previous_deck, deck_date'}), 400
		if not tracker:
			# Stage 6: no more Supabase tracker_snapshot fallback — the
			# tracker must come from the request every time, same as
			# /build_v2. `snapshot` (the previous session's numbers, for
			# the "what changed" comparison) stays optional: without one,
			# build_deck.build() just skips that comparison, same as a
			# first-ever build always has.
			return jsonify({'error': 'Missing tracker: upload a .xlsx file with this request.'}), 400

		tracker_path = _save_uploaded_file(tracker, tmpdir, 'tracker.xlsx')
		previous_deck_path = _save_uploaded_file(previous_deck, tmpdir, 'previous_deck.pptx')
		snapshot_path = _save_uploaded_file(snapshot, tmpdir, 'snapshot.json') if snapshot else ''
		ntp_comments_path = _save_uploaded_file(ntp_comments, tmpdir, 'ntp_comments.xlsx') if ntp_comments else ''

		# Normalize deck_date to expected format if possible (MM/DD/YYYY)
		try:
			parsed = datetime.strptime(deck_date, '%m/%d/%Y')
			deck_date_str = parsed.strftime('%m/%d/%Y')
		except Exception:
			deck_date_str = deck_date

		# Call build_deck.build
		out = build_deck.build(
			tracker_path=tracker_path,
			previous_deck_path=previous_deck_path,
			snapshot_path=snapshot_path,
			ntp_comments_path=ntp_comments_path,
			deck_date=deck_date_str,
			output_dir=tmpdir,
			tracker_rows=None,
			prev_snapshot_data=None,
		)

		# Create ZIP with outputs
		zip_buffer = io.BytesIO()
		with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as z:
			for key in ('deck_path', 'snapshot_path', 'ntp_comments_path', 'debug_slides_path'):
				path = out.get(key)
				if path and os.path.exists(path):
					z.write(path, arcname=os.path.basename(path))

		zip_buffer.seek(0)
		filename = f'ciege_outputs_{deck_date_str.replace("/","-")}.zip'
		resp = send_file(
			zip_buffer,
			mimetype='application/zip',
			as_attachment=True,
			download_name=filename,
		)
		if out.get('ntp_emails'):
			encoded = base64.b64encode(json.dumps(out['ntp_emails']).encode()).decode()
			resp.headers['X-Ntp-Emails'] = encoded
			resp.headers['Access-Control-Expose-Headers'] = 'X-Ntp-Emails'
		return resp

	except Exception as e:
		tb = traceback.format_exc()
		return jsonify({'error': str(e), 'traceback': tb}), 500

	finally:
		try:
			shutil.rmtree(tmpdir)
		except Exception:
			pass


@app.route('/build_v2', methods=['POST'])
def build_v2_endpoint():
	"""Deck Builder V2 — fully separate from /build above (which it never
	calls into, and never modifies any state /build also reads). See
	build_deck_v2.py's module docstring for the full list of behavioral
	differences this exists to deliver."""
	tmpdir = tempfile.mkdtemp(prefix='ciege_build_v2_')
	try:
		tracker = request.files.get('tracker')
		previous_deck = request.files.get('previous_deck')
		decom_pptx = request.files.get('decom_pptx')
		scop_pptx = request.files.get('scop_pptx')
		deck_date = request.form.get('deck_date') or request.args.get('deck_date')

		if not previous_deck or not deck_date:
			return jsonify({'error': 'Missing required fields: previous_deck, deck_date'}), 400
		if not decom_pptx or not scop_pptx:
			return jsonify({'error': 'Missing required fields: decom_pptx, scop_pptx (generated client-side before upload)'}), 400
		if not tracker:
			# Stage 6: no more Supabase tracker_snapshot fallback.
			return jsonify({'error': 'Missing tracker: upload a .xlsx file with this request.'}), 400

		tracker_path = _save_uploaded_file(tracker, tmpdir, 'tracker.xlsx')
		previous_deck_path = _save_uploaded_file(previous_deck, tmpdir, 'previous_deck.pptx')
		decom_pptx_bytes = decom_pptx.read()
		scop_pptx_bytes = scop_pptx.read()

		try:
			parsed = datetime.strptime(deck_date, '%m/%d/%Y')
			deck_date_str = parsed.strftime('%m/%d/%Y')
		except Exception:
			deck_date_str = deck_date

		out = build_deck_v2.build(
			tracker_path=tracker_path,
			previous_deck_path=previous_deck_path,
			deck_date=deck_date_str,
			output_dir=tmpdir,
			tracker_rows=None,
			supabase_client=supabase_client,
			decom_pptx_bytes=decom_pptx_bytes,
			scop_pptx_bytes=scop_pptx_bytes,
		)

		# Read into memory before the `finally` block's tmpdir cleanup runs —
		# send_file() given a path can still be mid-stream when that cleanup
		# fires, since Flask sends the body after this function returns (the
		# existing /build route sidesteps this the same way, via an in-memory
		# zip buffer instead of a path).
		with open(out['deck_path'], 'rb') as f:
			deck_bytes = f.read()
		buffer = io.BytesIO(deck_bytes)
		return send_file(
			buffer,
			mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
			as_attachment=True,
			download_name=os.path.basename(out['deck_path']),
		)

	except Exception as e:
		tb = traceback.format_exc()
		return jsonify({'error': str(e), 'traceback': tb}), 500

	finally:
		try:
			shutil.rmtree(tmpdir)
		except Exception:
			pass


@app.route('/ntp_emails', methods=['POST'])
def ntp_emails_endpoint():
	tmpdir = tempfile.mkdtemp(prefix='ciege_ntp_emails_')
	try:
		ntp_file = request.files.get('ntp_comments')
		if not ntp_file:
			return jsonify({'error': 'No NTP Comments file uploaded'}), 400
		ntp_path = _save_uploaded_file(ntp_file, tmpdir, 'ntp_comments.xlsx')
		emails = build_deck.generate_ntp_emails_from_file(ntp_path)
		return jsonify(emails)
	except Exception as e:
		tb = traceback.format_exc()
		return jsonify({'error': str(e), 'traceback': tb}), 500
	finally:
		try:
			shutil.rmtree(tmpdir)
		except Exception:
			pass


# ─────────────────────────────────────────────
# NTP comments (Deck Builder V2) — the NTP tab reads these, edits save here,
# and the Excel export is generated on demand for the customer.
# ─────────────────────────────────────────────

def _ntp_months(deck_date_str: str, tracker_path: str):
	"""Pending-NTP HOPs for every month in the window around the call, with the
	saved comment and status merged in. Same POR logic as the V2 slides."""
	import pandas as pd
	if not supabase_client:
		raise RuntimeError('Supabase is not configured')
	call_dt = build_deck_v2.call_tuesday(deck_date_str)
	data = build_deck.extract_data(tracker_path, '', '', call_dt.strftime('%m/%d/%Y'), tracker_rows=None, prev_snapshot_data=None)
	# A fixed 90-day lookback silently dropped any pending-NTP HOP whose
	# forecast start is older than that — confirmed on a real HOP forecast in
	# June while the call was in October (~115 days back), which had no month
	# tab to land in at all and vanished from the tab entirely, not just
	# undercounted. Start from whichever is earlier: 90 days back, or the
	# oldest pending HOP's own forecast month, so nothing overdue falls out.
	window_start = pd.Timestamp(call_dt - timedelta(days=90))
	ms15f_col = 'MS15 Implementation Start F'
	pending_dates = data['df'].loc[~data['df']['has_ntp'], ms15f_col].dropna()
	if len(pending_dates) > 0:
		earliest_pending = pd.Timestamp(pending_dates.min()).normalize().replace(day=1)
		if earliest_pending < window_start:
			window_start = earliest_pending
	months_span = (call_dt.year - window_start.year) * 12 + (call_dt.month - window_start.month) + 6
	por, order = build_deck_v2.compute_por_window(data['df'], window_start, window_months=max(12, months_span), max_active=24)
	store = ntp_comments_v2.load_store(supabase_client)
	months = []
	for mo_key, mo_name, sheet in order:
		rows = ntp_comments_v2.month_rows(por[mo_key]['pending_rows'], store.get(sheet, {}))
		months.append({'sheet': sheet, 'label': mo_name, 'rows': rows})
	return call_dt, months


@app.route('/ntp_comments', methods=['POST'])
def ntp_comments_get():
	tmpdir = tempfile.mkdtemp(prefix='ciege_ntp_comments_')
	try:
		tracker = request.files.get('tracker')
		if not tracker:
			return jsonify({'error': 'Missing tracker: upload a .xlsx file with this request.'}), 400
		tracker_path = _save_uploaded_file(tracker, tmpdir, 'tracker.xlsx')
		deck_date = request.form.get('deck_date') or request.args.get('deck_date') or datetime.now().strftime('%m/%d/%Y')
		call_dt, months = _ntp_months(deck_date, tracker_path)
		return jsonify({
			'deck_date': call_dt.strftime('%m/%d/%Y'),
			'statuses': ntp_comments_v2.STATUS_OPTIONS,
			'months': months,
		})
	except Exception as e:
		traceback.print_exc()
		return jsonify({'error': str(e)}), 500
	finally:
		try:
			shutil.rmtree(tmpdir)
		except Exception:
			pass


@app.route('/ntp_comments/save', methods=['POST'])
def ntp_comments_save():
	"""Saves one or more HOP edits for one month sheet. Reports success only
	after the saved row is read back and matches."""
	try:
		body = request.get_json(force=True) or {}
		sheet = body.get('sheet')
		entries = body.get('entries') or []
		if not sheet or not entries:
			return jsonify({'error': 'sheet and entries are required'}), 400
		for e in entries:
			if not e.get('hop'):
				return jsonify({'error': 'every entry needs a hop'}), 400
			if 'status' in e and e['status'] not in ntp_comments_v2.STATUS_OPTIONS:
				return jsonify({'error': 'unknown status'}), 400
		return jsonify(ntp_comments_v2.apply_batch(supabase_client, sheet, entries))
	except Exception as e:
		traceback.print_exc()
		return jsonify({'error': str(e)}), 500


@app.route('/ntp_comments/export', methods=['POST'])
def ntp_comments_export():
	tmpdir = tempfile.mkdtemp(prefix='ciege_ntp_export_')
	try:
		tracker = request.files.get('tracker')
		if not tracker:
			return jsonify({'error': 'Missing tracker: upload a .xlsx file with this request.'}), 400
		tracker_path = _save_uploaded_file(tracker, tmpdir, 'tracker.xlsx')
		deck_date = request.form.get('deck_date') or request.args.get('deck_date') or datetime.now().strftime('%m/%d/%Y')
		call_dt, months = _ntp_months(deck_date, tracker_path)
		xlsx = ntp_comments_v2.export_workbook([(m['sheet'], m['rows']) for m in months])
		return send_file(io.BytesIO(xlsx), mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
		                 as_attachment=True, download_name=f'NTP_Pending_{call_dt.strftime("%m-%d-%Y")}.xlsx')
	except Exception as e:
		traceback.print_exc()
		return jsonify({'error': str(e)}), 500
	finally:
		try:
			shutil.rmtree(tmpdir)
		except Exception:
			pass


if __name__ == '__main__':
	port = int(os.environ.get('PORT', 8000))
	app.run(host='0.0.0.0', port=port, debug=False)
