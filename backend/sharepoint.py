"""
SharePoint reads for the Railway backend — Stage 1 of the manual-upload-to-
SharePoint migration (see app/lib/sharepoint.ts for the frontend's half).

This is a SEPARATE auth path from the frontend's: the frontend uses MSAL
delegated login (a signed-in Nokia user's own browser session), but two
backend endpoints (/gr_data and the NTP comments flow) have no per-request
file upload and need to fetch Nokia source data on their own, server-side,
with nobody's browser in the loop. For that, this module uses the Azure AD
app-only client-credentials flow (AZURE_TENANT_ID/AZURE_CLIENT_ID/
AZURE_CLIENT_SECRET on Railway) — a token tied to the app registration
itself, not any one person's account.

Not wired into any existing endpoint yet. This stage only proves the token
exchange and file read work; /gr_data and /ntp_comments* keep reading from
Supabase until a later stage.

SITES/FILES below must stay in sync with app/lib/sharepoint.ts by hand —
there's no way to share one source of truth across the Python/TypeScript
boundary here.
"""

import os
import time
from urllib.parse import quote

import requests

GRAPH_BASE = 'https://graph.microsoft.com/v1.0'

AZURE_TENANT_ID = os.environ.get('AZURE_TENANT_ID', '')
AZURE_CLIENT_ID = os.environ.get('AZURE_CLIENT_ID', '')
AZURE_CLIENT_SECRET = os.environ.get('AZURE_CLIENT_SECRET', '')

SITES = {
	'viaeroWireless': 'nokianam.sharepoint.com/sites/viaero-wireless2',
	'namds': 'nokia.sharepoint.com/sites/NAMDS',
}

FILES = {
	'tracker': {
		'site': SITES['viaeroWireless'],
		'path': '/Shared Documents/Team PM/CJ Ojeda/Viaero - Construction Schedule Tracker - WIP 2.0.xlsx',
	},
	'spo': {
		'site': SITES['namds'],
		'path': '/Shared Documents/NDPd (SiteForge)/US Majors/US Majors NDPd Reports/Viaero NDPd Reports/Current/NDPd_NAM_003_Viaero_SPO_Report.xlsx',
	},
	'decom': {
		'site': SITES['viaeroWireless'],
		'path': '/Shared Documents/Team Decom/Decom Tracker Updates/Nokia Decom Status_Report_LIVE - NEW.xlsx',
	},
	'cr': {
		'site': SITES['viaeroWireless'],
		'path': '/Shared Documents/Team CQT/Viaero - CR Claims Tracker.xlsx',
	},
	'scop': {
		'site': SITES['viaeroWireless'],
		'path': '/Shared Documents/Team Close Out/Viaero COP Tracker_LIVE.xlsx',
	},
}


class SharePointConfigError(Exception):
	"""AZURE_* env vars missing — distinct from an auth/network failure so
	the caller can give CJ a config-specific error instead of a confusing
	401."""
	pass


class SharePointNotFoundError(Exception):
	def __init__(self, file_key, file_path):
		self.file_key = file_key
		self.file_path = file_path
		super().__init__(f'Could not find "{file_key}" on SharePoint at: {file_path}')


class SharePointRateLimitError(Exception):
	def __init__(self, retry_after_seconds):
		self.retry_after_seconds = retry_after_seconds
		super().__init__(f'SharePoint/Graph API rate limited — retry after {retry_after_seconds}s.')


# Cached in module-level state (not per-request) — a client-credentials
# token is good for about an hour and this process serves many requests,
# so fetching a fresh one every call would be both slow and wasteful.
_token_cache = {'access_token': None, 'expires_at': 0}


def _get_graph_token() -> str:
	if not (AZURE_TENANT_ID and AZURE_CLIENT_ID and AZURE_CLIENT_SECRET):
		raise SharePointConfigError(
			'AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET not configured on the backend.'
		)

	# 60s safety margin so a token doesn't expire mid-request.
	if _token_cache['access_token'] and time.time() < _token_cache['expires_at'] - 60:
		return _token_cache['access_token']

	resp = requests.post(
		f'https://login.microsoftonline.com/{AZURE_TENANT_ID}/oauth2/v2.0/token',
		data={
			'grant_type': 'client_credentials',
			'client_id': AZURE_CLIENT_ID,
			'client_secret': AZURE_CLIENT_SECRET,
			'scope': 'https://graph.microsoft.com/.default',
		},
		timeout=30,
	)
	if not resp.ok:
		raise Exception(f'Graph token request failed ({resp.status_code}): {resp.text[:500]}')

	body = resp.json()
	_token_cache['access_token'] = body['access_token']
	_token_cache['expires_at'] = time.time() + body.get('expires_in', 3600)
	return _token_cache['access_token']


def _site_segment(site: str) -> str:
	# SITES values are "{hostname}/{server-relative-path}" — Graph's
	# path-based site addressing wants "{hostname}:{server-relative-path}"
	# as one segment.
	hostname, _, site_path = site.partition('/')
	return f'{hostname}:/{site_path}'


def read_sharepoint_file(site_url: str, file_path: str) -> bytes:
	"""Reads a file from SharePoint via Graph using the app's own
	client-credentials token (not any one user's login)."""
	token = _get_graph_token()
	# quote(..., safe='/') — keeps path separators intact while
	# percent-encoding spaces/parens/etc. these real file paths are full of.
	url = f'{GRAPH_BASE}/sites/{_site_segment(site_url)}:/drive/root:{quote(file_path, safe="/")}:/content'
	resp = requests.get(url, headers={'Authorization': f'Bearer {token}'}, timeout=60)

	if resp.status_code == 404:
		raise SharePointNotFoundError(file_path, file_path)
	if resp.status_code == 429:
		retry_after = int(resp.headers.get('Retry-After', 30))
		raise SharePointRateLimitError(retry_after)
	if not resp.ok:
		raise Exception(f'SharePoint read failed ({resp.status_code}): {resp.text[:500]}')

	return resp.content


def read_named_file(key: str) -> bytes:
	"""Convenience wrapper over read_sharepoint_file for one of the named
	FILES entries."""
	entry = FILES.get(key)
	if not entry:
		raise ValueError(f'Unknown SharePoint file key: {key}')
	try:
		return read_sharepoint_file(entry['site'], entry['path'])
	except SharePointNotFoundError:
		raise SharePointNotFoundError(key, entry['path'])
