import io
import os
import sys
import json
import shutil
import zipfile
import tempfile
import traceback
from datetime import datetime

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename

# Ensure backend package path is importable when running this file directly
sys.path.append(os.path.dirname(__file__))
import build_deck

app = Flask(__name__)
CORS(app)


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
		# Support both file upload and JSON data from Supabase
		tracker_json = request.form.get('tracker_json')
		prev_snapshot_json = request.form.get('prev_snapshot_json')

		tracker_rows = json.loads(tracker_json) if tracker_json else None
		prev_snapshot_data = json.loads(prev_snapshot_json) if prev_snapshot_json else None

		previous_deck = request.files.get('previous_deck')
		ntp_comments = request.files.get('ntp_comments')
		deck_date = request.form.get('deck_date') or request.args.get('deck_date')

		# Fall back to file uploads for tracker/snapshot if JSON not provided
		tracker = request.files.get('tracker') if tracker_rows is None else None
		snapshot = request.files.get('snapshot') if prev_snapshot_data is None else None

		if not previous_deck or not deck_date:
			return jsonify({'error': 'Missing required fields: previous_deck, deck_date'}), 400
		if tracker_rows is None and tracker is None:
			return jsonify({'error': 'Missing tracker: provide tracker_json or tracker file'}), 400

		tracker_path = _save_uploaded_file(tracker, tmpdir, 'tracker.xlsx') if tracker else ''
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
			tracker_rows=tracker_rows,
			prev_snapshot_data=prev_snapshot_data,
		)

		# Create ZIP with outputs
		zip_buffer = io.BytesIO()
		with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as z:
			for key in ('deck_path', 'snapshot_path', 'ntp_comments_path'):
				path = out.get(key)
				if path and os.path.exists(path):
					z.write(path, arcname=os.path.basename(path))

		zip_buffer.seek(0)
		filename = f'ciege_outputs_{deck_date_str.replace("/","-")}.zip'
		return send_file(
			zip_buffer,
			mimetype='application/zip',
			as_attachment=True,
			download_name=filename,
		)

	except Exception as e:
		tb = traceback.format_exc()
		return jsonify({'error': str(e), 'traceback': tb}), 500

	finally:
		try:
			shutil.rmtree(tmpdir)
		except Exception:
			pass


if __name__ == '__main__':
	port = int(os.environ.get('PORT', 8000))
	app.run(host='0.0.0.0', port=port, debug=False)
