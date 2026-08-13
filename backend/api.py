import io
import os
import sys
import json
import base64
import shutil
import zipfile
import tempfile
import traceback
from datetime import datetime

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
from supabase import create_client

# Ensure backend package path is importable when running this file directly
sys.path.append(os.path.dirname(__file__))
import build_deck
import gr_tracker

app = Flask(__name__)
CORS(app)

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

		tracker_path = _save_uploaded_file(tracker, tmpdir, 'tracker.xlsx') if tracker else ''
		previous_deck_path = _save_uploaded_file(previous_deck, tmpdir, 'previous_deck.pptx')
		snapshot_path = _save_uploaded_file(snapshot, tmpdir, 'snapshot.json') if snapshot else ''
		ntp_comments_path = _save_uploaded_file(ntp_comments, tmpdir, 'ntp_comments.xlsx') if ntp_comments else ''

		# Fall back to Supabase if tracker or snapshot not uploaded
		tracker_rows = None
		prev_snapshot_data = None
		if (not tracker or not snapshot) and supabase_client:
			try:
				if not tracker:
					latest = supabase_client.table('tracker_snapshot').select('*').order('uploaded_at', desc=True).limit(1).single().execute()
					if latest.data:
						tracker_rows = json.loads(latest.data['data'])
				if not snapshot:
					previous = supabase_client.table('tracker_snapshot').select('*').order('uploaded_at', desc=True).range(1, 1).execute()
					if previous.data:
						prev_snapshot_data = {
							'rows': json.loads(previous.data[0]['data']),
							'uploaded_at': previous.data[0]['uploaded_at']
						}
			except Exception as e:
				print(f'Supabase fetch error: {e}')

		if not tracker_path and tracker_rows is None:
			return jsonify({'error': 'Missing tracker: upload a .xlsx file or ensure Supabase has a snapshot'}), 400

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
			tracker_rows=tracker_rows,
			prev_snapshot_data=prev_snapshot_data,
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


@app.route('/ntp_emails', methods=['POST'])
def ntp_emails_endpoint():
	try:
		ntp_file = request.files.get('ntp_comments')
		if not ntp_file:
			return jsonify({'error': 'No NTP Comments file uploaded'}), 400
		with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
			ntp_file.save(tmp.name)
			ntp_path = tmp.name
		emails = build_deck.generate_ntp_emails_from_file(ntp_path)
		os.unlink(ntp_path)
		return jsonify(emails)
	except Exception as e:
		tb = traceback.format_exc()
		return jsonify({'error': str(e), 'traceback': tb}), 500


@app.route('/gr_data', methods=['POST'])
def gr_data_endpoint():
	try:
		if not supabase_client:
			return jsonify({'error': 'Supabase not configured'}), 500
		body = request.get_json(silent=True) or {}
		gc_filter = body.get('gc') or None
		groups = gr_tracker.load_gr_data(supabase_client, gc_filter=gc_filter)
		return jsonify(groups)
	except Exception as e:
		tb = traceback.format_exc()
		return jsonify({'error': str(e), 'traceback': tb}), 500


if __name__ == '__main__':
	port = int(os.environ.get('PORT', 8000))
	app.run(host='0.0.0.0', port=port, debug=False)
