import os
import json
import uuid
import base64
import requests
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from docx import Document
from docx.shared import RGBColor, Pt
from flask import Flask, request, render_template_string, redirect, session, flash

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "advocare_secure_session_key")
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
OUTPUT_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Client_Rosters')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

BASE_HTML = """
<!DOCTYPE html>
<html>
<head>
    <title>Advocare SA Roster Mailer</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background-color: #f4f4f9; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        .navbar { background-color: #800080; }
        .navbar-brand, .nav-link { color: white !important; font-weight: bold; }
        .card-header { background-color: #800080; color: white; font-weight: bold; }
        .btn-primary { background-color: #800080; border: none; }
        .btn-primary:hover { background-color: #5c005c; }
        .btn-gold { background-color: #C5A059; color: white; font-weight: bold; }
        .btn-gold:hover { background-color: #a38242; color: white; }
    </style>
</head>
<body>
    <nav class="navbar navbar-expand-lg mb-4">
        <div class="container">
            <a class="navbar-brand" href="/">Advocare SA Web Mailer</a>
            <div class="d-flex">
                <a class="nav-link me-3" href="/">Home</a>
                <a class="nav-link me-3" href="/emails">Manage Emails</a>
                <a class="nav-link" href="/logout">Logout</a>
            </div>
        </div>
    </nav>
    <div class="container pb-5">
        {% with messages = get_flashed_messages(with_categories=true) %}
          {% if messages %}
            {% for category, message in messages %}
              <div class="alert alert-{{ 'success' if category == 'success' else 'danger' }}">{{ message }}</div>
            {% endfor %}
          {% endif %}
        {% endwith %}
        {% block content %}{% endblock %}
    </div>
</body>
</html>
"""

LOGIN_HTML = BASE_HTML.replace('{% block content %}{% endblock %}', """
<div class="row justify-content-center">
    <div class="col-md-5">
        <div class="card shadow-sm mt-5">
            <div class="card-header text-center">Admin Login Required</div>
            <div class="card-body p-4">
                <form method="POST" action="/login">
                    <div class="mb-3">
                        <label class="form-label">System Password</label>
                        <input type="password" name="password" class="form-control" required>
                    </div>
                    <button type="submit" class="btn btn-primary w-100">Secure Login</button>
                </form>
            </div>
        </div>
    </div>
</div>
""").replace('<nav class="navbar navbar-expand-lg mb-4">', '<nav class="navbar navbar-expand-lg mb-4" style="display:none;">')

INDEX_HTML = BASE_HTML.replace('{% block content %}{% endblock %}', """
<div class="row">
    <!-- Primary Sync Option -->
    <div class="col-md-6 mb-4">
        <div class="card shadow-sm h-100">
            <div class="card-header" style="background-color: #C5A059;">1. Auto-Sync from ShiftCare</div>
            <div class="card-body p-4">
                <h5 class="card-title mb-3 text-center">Pull Latest Published Shifts</h5>
                <p class="text-muted mb-4 text-center">Select the exact date range to fetch from the ShiftCare system.</p>
                <form method="POST" action="/sync">
                    <div class="row g-2 mb-3">
                        <div class="col-6">
                            <label class="form-label small fw-bold">Start Date</label>
                            <input type="date" name="sync_start" id="sync_start" class="form-control" required>
                        </div>
                        <div class="col-6">
                            <label class="form-label small fw-bold">End Date</label>
                            <input type="date" name="sync_end" id="sync_end" class="form-control" required>
                        </div>
                    </div>
                    <button type="submit" class="btn btn-gold btn-lg w-100" id="syncBtn" onclick="if(this.form.checkValidity()){ this.innerHTML='<span class=\\'spinner-border spinner-border-sm\\'></span> Syncing...'; this.classList.add('disabled'); }">Sync Now</button>
                </form>
            </div>
        </div>
    </div>

    <!-- Secondary Backup Option -->
    <div class="col-md-6 mb-4">
        <div class="card shadow-sm h-100">
            <div class="card-header">2. Manual CSV Backup</div>
            <div class="card-body p-4 text-center">
                <h5 class="card-title mb-3">Upload Export Files</h5>
                <p class="text-muted mb-4">Use this as a fallback. You can select multiple CSV files at once to combine months.</p>
                <form method="POST" action="/upload" enctype="multipart/form-data">
                    <div class="mb-3">
                        <input class="form-control" type="file" name="schedule_files" accept=".csv, text/csv, application/vnd.ms-excel, text/comma-separated-values" multiple required>
                    </div>
                    <button type="submit" class="btn btn-outline-primary w-100">Process CSV</button>
                </form>
            </div>
        </div>
    </div>
</div>
<script>
    document.addEventListener('DOMContentLoaded', function() {
        let today = new Date();
        let startInput = document.getElementById('sync_start');
        let endInput = document.getElementById('sync_end');
        if(startInput && endInput) {
            // Format YYYY-MM-DD for the HTML date inputs
            startInput.value = today.toISOString().split('T')[0];
            
            // Calculate the last day of the current month
            let endOfMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0);
            
            // Handle timezone offset to ensure the local date is picked correctly
            let endOfMonthLocal = new Date(endOfMonth.getTime() - (endOfMonth.getTimezoneOffset() * 60000));
            endInput.value = endOfMonthLocal.toISOString().split('T')[0];
        }
    });
</script>
""")

EMAILS_HTML = BASE_HTML.replace('{% block content %}{% endblock %}', """
<div class="card shadow-sm">
    <div class="card-header">Client Email Directory</div>
    <div class="card-body">
        <form method="POST" action="/emails" class="mb-4 bg-light p-3 border rounded" id="emailForm">
            <h5>Add or Update Client Email</h5>
            <div class="row g-2 align-items-center">
                <div class="col-md-5">
                    <input type="text" name="client_name" id="client_name_input" class="form-control" placeholder="Exact Client Name" required>
                </div>
                <div class="col-md-5">
                    <input type="text" name="client_emails" id="client_emails_input" class="form-control" placeholder="Emails (comma separated)" required>
                </div>
                <div class="col-md-2">
                    <button type="submit" class="btn btn-primary w-100" id="submitBtn">Save</button>
                </div>
            </div>
            <div class="row mt-2" id="cancelEditRow" style="display: none;">
                <div class="col-12 text-end">
                    <button type="button" class="btn btn-sm btn-secondary" onclick="cancelEdit()">Cancel Edit</button>
                </div>
            </div>
        </form>
        <hr>
        <div class="table-responsive">
            <table class="table table-striped align-middle">
                <thead><tr><th>Client Name</th><th>Configured Emails</th><th>Action</th></tr></thead>
                <tbody>
                    {% for name, emails in directory.items() %}
                    <tr>
                        <td>{{ name }}</td>
                        <td>{{ emails | join(', ') }}</td>
                        <td>
                            <button type="button" class="btn btn-sm btn-info text-white me-2" onclick="editClient('{{ name }}', '{{ emails | join(', ') }}')">Edit</button>
                            <form method="POST" action="/delete_email" style="display:inline;">
                                <input type="hidden" name="client_name" value="{{ name }}">
                                <button type="submit" class="btn btn-sm btn-danger">Remove</button>
                            </form>
                        </td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>
</div>
<script>
    function editClient(name, emails) {
        document.getElementById('client_name_input').value = name;
        document.getElementById('client_emails_input').value = emails;
        document.getElementById('submitBtn').textContent = 'Update';
        document.getElementById('cancelEditRow').style.display = 'block';
        document.getElementById('client_name_input').focus();
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    function cancelEdit() {
        document.getElementById('client_name_input').value = '';
        document.getElementById('client_emails_input').value = '';
        document.getElementById('submitBtn').textContent = 'Save';
        document.getElementById('cancelEditRow').style.display = 'none';
    }
</script>
""")

SELECT_HTML = BASE_HTML.replace('{% block content %}{% endblock %}', """
<div class="card shadow-sm">
    <div class="card-header">Select Clients to Process</div>
    <div class="card-body">
        <form method="POST" action="/process">
            <input type="hidden" name="filename" value="{{ filename }}">
            
            <div class="d-flex justify-content-between mb-3 align-items-center">
                <div>
                    <button type="button" class="btn btn-sm btn-outline-secondary" onclick="toggleAll(true)">Select All</button>
                    <button type="button" class="btn btn-sm btn-outline-secondary" onclick="toggleAll(false)">Deselect All</button>
                </div>
                <div class="form-check form-switch fs-5">
                    <input class="form-check-input" type="checkbox" name="is_amended" id="amendedCheck" value="yes">
                    <label class="form-check-label text-purple fw-bold" for="amendedCheck" style="color: #800080;">This is an Amended Roster</label>
                </div>
            </div>
            
            <div class="border rounded p-3 mb-4" style="max-height: 400px; overflow-y: auto; background-color: white;">
                {% for client in clients %}
                <div class="form-check mb-2">
                    <input class="form-check-input client-checkbox" type="checkbox" name="selected_clients" value="{{ client.name }}" id="chk_{{ loop.index }}" {% if not client.has_email %}disabled{% endif %}>
                    <label class="form-check-label" for="chk_{{ loop.index }}" style="{% if not client.has_email %}color: red;{% endif %}">
                        <strong>{{ client.name }}</strong> - {{ client.shifts }} shifts 
                        <span class="text-muted">{% if client.has_email %}[{{ client.emails }}]{% else %}[NO EMAIL CONFIGURED]{% endif %}</span>
                    </label>
                </div>
                {% endfor %}
            </div>
            <button type="submit" class="btn btn-gold btn-lg w-100" id="submitBtn">Generate & Send Selected</button>
        </form>
    </div>
</div>
<script>
    function toggleAll(state) {
        document.querySelectorAll('.client-checkbox:not([disabled])').forEach(cb => cb.checked = state);
    }
    document.querySelector('form').addEventListener('submit', function() {
        document.getElementById('submitBtn').innerHTML = '<span class="spinner-border spinner-border-sm" role="status" aria-hidden="true"></span> Processing & Sending... Please wait.';
        document.getElementById('submitBtn').classList.add('disabled');
    });
</script>
""")

RESULT_HTML = BASE_HTML.replace('{% block content %}{% endblock %}', """
<div class="card shadow-sm">
    <div class="card-header">Batch Complete</div>
    <div class="card-body">
        <h4 class="text-success mb-3">{{ success_count }} Emails Successfully Dispatched</h4>
        <div class="bg-dark text-success p-3 rounded" style="font-family: monospace; max-height: 400px; overflow-y: auto;">
            {% for log in logs %}
                <div>{{ log }}</div>
            {% endfor %}
        </div>
        <a href="/" class="btn btn-primary mt-4">Start New Batch</a>
    </div>
</div>
""")

def get_google_sheet():
    creds_json = os.environ.get('GOOGLE_CREDENTIALS')
    if not creds_json:
        raise Exception("GOOGLE_CREDENTIALS environment variable is missing.")
    creds_dict = json.loads(creds_json)
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(credentials)
    return client.open("Advocare Client Emails").sheet1

def get_emails():
    try:
        sheet = get_google_sheet()
        records = sheet.get_all_records()
        directory = {}
        for row in records:
            name = str(row.get('Client Name', '')).strip()
            if name:
                directory[name] = [e.strip() for e in str(row.get('Emails', '')).split(',') if e.strip()]
        return directory
    except Exception as e:
        print(f"Google Sheets Error: {e}")
        return {}

def save_emails(directory):
    try:
        sheet = get_google_sheet()
        sheet.clear()
        rows = [['Client Name', 'Emails']]
        for name, emails in sorted(directory.items()):
            rows.append([name, ", ".join(emails)])
        sheet.update('A1', rows)
    except Exception as e:
        print(f"Google Sheets Error: {e}")

def send_brevo_email(to_emails, subject, html_content):
    api_key = os.environ.get('BREVO_API_KEY')
    if not api_key:
        raise Exception("BREVO_API_KEY missing.")
        
    url = "https://api.brevo.com/v3/smtp/email"
    
    payload = {
        "sender": {"name": "Advocare SA", "email": "rostering@advocaresa.com.au"},
        "to": [{"email": e} for e in to_emails],
        "subject": subject,
        "htmlContent": html_content
    }
        
    headers = {
        "accept": "application/json",
        "api-key": api_key,
        "content-type": "application/json"
    }
    response = requests.post(url, json=payload, headers=headers)
    if response.status_code not in [201, 202]:
        raise Exception(f"Brevo API Error: {response.text}")

def check_auth():
    return session.get('logged_in') == True

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        admin_password = os.environ.get('ADMIN_PASSWORD', 'Admin123')
        if request.form.get('password') == admin_password:
            session['logged_in'] = True
            return redirect('/')
        else:
            flash("Incorrect Password", "danger")
    return render_template_string(LOGIN_HTML)

@app.route('/logout')
def logout():
    session.clear()
    return redirect('/login')

@app.route('/')
def index():
    if not check_auth(): return redirect('/login')
    return render_template_string(INDEX_HTML)

@app.route('/emails', methods=['GET', 'POST'])
def manage_emails():
    if not check_auth(): return redirect('/login')
    directory = get_emails()
    
    if request.method == 'POST':
        name = request.form.get('client_name').strip()
        emails_raw = request.form.get('client_emails').split(',')
        emails = [e.strip() for e in emails_raw if e.strip()]
        if name and emails:
            directory[name] = emails
            save_emails(directory)
            flash(f"Emails saved for {name} to Google Sheets", "success")
        return redirect('/emails')
        
    return render_template_string(EMAILS_HTML, directory=directory)

@app.route('/delete_email', methods=['POST'])
def delete_email():
    if not check_auth(): return redirect('/login')
    name = request.form.get('client_name')
    directory = get_emails()
    if name in directory:
        del directory[name]
        save_emails(directory)
        flash(f"Removed {name} from directory", "success")
    return redirect('/emails')

@app.route('/sync', methods=['POST'])
def sync():
    if not check_auth(): return redirect('/login')
    
    account_id = os.environ.get('SHIFTCARE_ACCOUNT_ID')
    api_key = os.environ.get('SHIFTCARE_API_KEY')
    if not api_key or not account_id:
        flash("ShiftCare credentials missing. Add SHIFTCARE_ACCOUNT_ID and SHIFTCARE_API_KEY to Render.", "danger")
        return redirect('/')
        
    sync_start = request.form.get('sync_start')
    sync_end = request.form.get('sync_end')
    
    if not sync_start or not sync_end:
        flash("Please provide both start and end dates.", "danger")
        return redirect('/')
        
    try:
        tz = 'Australia/Adelaide'
        # Parse start and end dates from the HTML form to valid Adelaide bounds
        start_dt = pd.to_datetime(sync_start).tz_localize(tz)
        end_dt = pd.to_datetime(sync_end).tz_localize(tz) + pd.Timedelta(days=1, seconds=-1)
        
        start_iso = start_dt.isoformat()
        end_iso = end_dt.isoformat()
        
        auth = (account_id, api_key)
        headers = {"Accept": "application/json"}
        
        print("\n--- SHIFTCARE SYNC INITIATED ---")
        print(f"Requesting shifts from {start_iso} to {end_iso}")
        
        params = {
            "from_date": start_iso,
            "to_date": end_iso,
            "include_clients": "true",
            "include_staff": "true",
            "include_metadata": "true",
            "per_page": 20,
            "page": 1
        }
        
        all_shifts = []
        base_url = "https://api.shiftcare.com/api/v3/shifts"
        
        while True:
            response = requests.get(base_url, auth=auth, headers=headers, params=params)
            if response.status_code != 200:
                print(f"API HTTP Error: {response.status_code} - {response.text}")
                flash(f"ShiftCare API Error: {response.text}", "danger")
                return redirect('/')
            
            data = response.json()
            
            if isinstance(data, dict):
                all_shifts.extend(data.get('shifts', []))
            else:
                all_shifts.extend(data)
            
            current_page = int(response.headers.get('Current-Page', 1))
            total_pages = int(response.headers.get('Total-Pages', 1))
            
            if current_page >= total_pages:
                break
            params['page'] = current_page + 1
        
        print(f"Total raw shifts fetched from API: {len(all_shifts)}")
        
        mapped_data = []
        for shift in all_shifts:
            if not shift.get('published', True):
                continue
                
            staff_name = "Unassigned"
            if shift.get('staff') and len(shift['staff']) > 0:
                s = shift['staff'][0]
                staff_name = s.get('name') or s.get('full_name') or s.get('staff_name')
                if not staff_name and s.get('first_name'):
                    staff_name = f"{s.get('first_name', '')} {s.get('family_name', s.get('last_name', ''))}".strip()
                if not staff_name:
                    staff_name = "Unassigned"
                
            # ROBUST CLIENT NAME EXTRACTION (DROPPING NDIS TAGS)
            for c in shift.get('clients', []):
                fname = str(c.get('first_name', '')).strip()
                lname_raw = str(c.get('family_name', c.get('last_name', ''))).strip()
                
                lname = ""
                if lname_raw:
                    first_word = lname_raw.split()[0]
                    if 'NDIS' not in first_word.upper():
                        lname = first_word

                if fname or lname:
                    client_name = f"{fname} {lname}".strip()
                else:
                    client_name = str(c.get('name') or c.get('full_name') or 'Unknown').split(' NDIS')[0].strip()
                    
                mapped_data.append({
                    'Name': client_name,
                    'Start Date Time': shift.get('start_at') or shift.get('start_time'),
                    'End Date Time': shift.get('end_at') or shift.get('end_time'),
                    'Staff': staff_name,
                    'Shift Status': 'published'
                })
        
        print(f"Total shifts after mapping clients and skipping drafts: {len(mapped_data)}")
            
        if not mapped_data:
            flash(f"Sync successful, but no published shifts were found between {sync_start} and {sync_end}.", "warning")
            return redirect('/')
            
        df = pd.DataFrame(mapped_data)
        
        filename = f"schedule_api_{uuid.uuid4().hex}.csv"
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        df.to_csv(filepath, index=False)
        
        df['Start Date Time'] = pd.to_datetime(df['Start Date Time'], errors='coerce', utc=True).dt.tz_convert('Australia/Adelaide')
        df['End Date Time'] = pd.to_datetime(df['End Date Time'], errors='coerce', utc=True).dt.tz_convert('Australia/Adelaide')
        df = df.dropna(subset=['Start Date Time', 'End Date Time'])
        
        # Enforce exact bounds requested by user
        df = df[(df['Start Date Time'] >= start_dt) & (df['Start Date Time'] <= end_dt)]
        
        print(f"Total shifts surviving the exact-date filter: {len(df)}")
        
        clients_data = []
        directory = get_emails()
        
        for name, data in df.groupby('Name'):
            if pd.notna(name):
                has_email = name in directory
                emails_str = ", ".join(directory.get(name, []))
                clients_data.append({
                    'name': name,
                    'shifts': len(data),
                    'has_email': has_email,
                    'emails': emails_str
                })
        
        clients_data = sorted(clients_data, key=lambda x: x['name'])
        print("--- SYNC COMPLETE ---\n")
        return render_template_string(SELECT_HTML, clients=clients_data, filename=filename)
        
    except Exception as e:
        print(f"CRITICAL SYNC ERROR: {str(e)}")
        flash(f"An error occurred during sync: {str(e)}", "danger")
        return redirect('/')

@app.route('/upload', methods=['POST'])
def upload():
    if not check_auth(): return redirect('/login')
    
    files = request.files.getlist('schedule_files')
    if not files or files[0].filename == '':
        flash("No file selected", "danger")
        return redirect('/')
        
    dfs = []
    for file in files:
        if file.filename:
            temp_filename = f"temp_{uuid.uuid4().hex}.csv"
            temp_filepath = os.path.join(UPLOAD_FOLDER, temp_filename)
            file.save(temp_filepath)
            try:
                dfs.append(pd.read_csv(temp_filepath))
            except Exception as e:
                print(f"Error reading {file.filename}: {e}")
            os.remove(temp_filepath)
            
    if not dfs:
        flash("No valid CSV files could be read.", "danger")
        return redirect('/')
        
    try:
        df = pd.concat(dfs, ignore_index=True)
        
        filename = f"schedule_{uuid.uuid4().hex}.csv"
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        df.to_csv(filepath, index=False)
        
        df = df[df['Shift Status'].str.lower() != 'cancelled']
        df = df[~df['Staff'].astype(str).str.contains("CANCELLED", case=False, na=False)]
        
        df['Start Date Time'] = pd.to_datetime(df['Start Date Time'], errors='coerce', utc=True).dt.tz_convert('Australia/Adelaide')
        df['End Date Time'] = pd.to_datetime(df['End Date Time'], errors='coerce', utc=True).dt.tz_convert('Australia/Adelaide')
        df = df.dropna(subset=['Start Date Time', 'End Date Time'])
        
        today_start = pd.Timestamp.now('Australia/Adelaide').normalize()
        df = df[df['Start Date Time'] >= today_start]
        
        # Clean Name column in CSV exports to match the sync logic
        df['Name'] = df['Name'].astype(str).apply(lambda x: x.split(' NDIS')[0].strip())
        
        clients_data = []
        directory = get_emails()
        
        for name, data in df.groupby('Name'):
            if pd.notna(name):
                has_email = name in directory
                emails_str = ", ".join(directory.get(name, []))
                clients_data.append({
                    'name': name,
                    'shifts': len(data),
                    'has_email': has_email,
                    'emails': emails_str
                })
        
        clients_data = sorted(clients_data, key=lambda x: x['name'])
        return render_template_string(SELECT_HTML, clients=clients_data, filename=filename)
        
    except Exception as e:
        flash(f"Error processing CSV files: {str(e)}", "danger")
        return redirect('/')

@app.route('/process', methods=['POST'])
def process():
    if not check_auth(): return redirect('/login')
    
    filename = request.form.get('filename')
    selected = request.form.getlist('selected_clients')
    is_amended = request.form.get('is_amended') == 'yes'
    
    if not selected:
        flash("No clients selected.", "danger")
        return redirect('/')
        
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    df = pd.read_csv(filepath)
    df = df[df['Shift Status'].str.lower() != 'cancelled']
    df = df[~df['Staff'].astype(str).str.contains("CANCELLED", case=False, na=False)]
    
    df['Start Date Time'] = pd.to_datetime(df['Start Date Time'], errors='coerce', utc=True).dt.tz_convert('Australia/Adelaide')
    df['End Date Time'] = pd.to_datetime(df['End Date Time'], errors='coerce', utc=True).dt.tz_convert('Australia/Adelaide')
    df = df.dropna(subset=['Start Date Time', 'End Date Time'])
    
    today_start = pd.Timestamp.now('Australia/Adelaide').normalize()
    df = df[df['Start Date Time'] >= today_start]
    
    # Ensure CSV names are clean during final generation too
    df['Name'] = df['Name'].astype(str).apply(lambda x: x.split(' NDIS')[0].strip())
    
    directory = get_emails()
    logs = []
    success_count = 0
    
    amended_word = "amended " if is_amended else ""
    subject_prefix = "Amended - " if is_amended else ""
    file_prefix = "Amended_" if is_amended else ""

    for name in selected:
        try:
            client_data = df[df['Name'] == name].sort_values(by='Start Date Time')
            if client_data.empty: continue
            
            unique_months = client_data['Start Date Time'].dt.strftime('%B').unique().tolist()
            if len(unique_months) > 1:
                if len(unique_months) == 2:
                    month_name = f"{unique_months[0]} and {unique_months[1]}"
                else:
                    month_name = f"{unique_months[0]} to {unique_months[-1]}"
                date_phrase = f"for {month_name}"
                doc_month_label = f"{unique_months[0]}-{unique_months[-1]}"
            else:
                month_name = unique_months[0]
                date_phrase = f"until the end of {month_name}"
                doc_month_label = month_name
            
            doc = Document()
            style = doc.styles['Normal']
            font = style.font
            font.name = 'Calibri'
            font.size = Pt(11) 
            
            doc.add_paragraph(f"Dear {name},") 
            doc.add_paragraph(f"Please find below your {amended_word}roster of support {date_phrase} for approval:\n") 
            
            table = doc.add_table(rows=1, cols=4)
            table.style = 'Normal Table'
            hdr_cells = table.rows[0].cells
            for i, header_text in enumerate(['Date', 'Support Team Member', 'Start', 'End']):
                run = hdr_cells[i].paragraphs[0].add_run(header_text)
                run.bold = True
                run.font.color.rgb = RGBColor(128, 0, 128)
            
            current_week = None
            html_rows = ""
            
            for index, row in client_data.iterrows():
                shift_week = row['Start Date Time'].isocalendar()[1]
                if current_week is not None and shift_week != current_week:
                    table.add_row() 
                    html_rows += "<tr><td colspan='4' style='padding: 6px; border: none;'></td></tr>\n"
                current_week = shift_week
                
                date_str = row['Start Date Time'].strftime('%a, %d %b %Y')
                staff_str = str(row['Staff'])
                start_str = row['Start Date Time'].strftime('%I:%M %p').lstrip('0')
                end_str = row['End Date Time'].strftime('%I:%M %p').lstrip('0')

                row_cells = table.add_row().cells
                row_cells[0].text = date_str
                row_cells[1].text = staff_str
                row_cells[2].text = start_str
                row_cells[3].text = end_str
                
                html_rows += f"""
                <tr>
                    <td style="font-family: 'Calibri', sans-serif; font-size: 11pt; padding: 6px 8px; border: 1px solid #cccccc; text-align: left;">{date_str}</td>
                    <td style="font-family: 'Calibri', sans-serif; font-size: 11pt; padding: 6px 8px; border: 1px solid #cccccc; text-align: left;">{staff_str}</td>
                    <td style="font-family: 'Calibri', sans-serif; font-size: 11pt; padding: 6px 8px; border: 1px solid #cccccc; text-align: left;">{start_str}</td>
                    <td style="font-family: 'Calibri', sans-serif; font-size: 11pt; padding: 6px 8px; border: 1px solid #cccccc; text-align: left;">{end_str}</td>
                </tr>"""
            
            doc.add_paragraph("\nAll look OK?") 
            doc.add_paragraph("If you have any concerns or queries, please don't hesitate to contact us  🙂") 
            
            safe_name = str(name).replace("/", "_").replace("\\", "_")
            doc_filename_short = f"{file_prefix}{safe_name}_{doc_month_label}_Roster.docx"
            doc_filename = os.path.join(OUTPUT_FOLDER, doc_filename_short)
            doc.save(doc_filename)
            
            target_emails = directory.get(name)
            if not target_emails:
                logs.append(f"SKIPPED {name}: No email configured.")
                continue
            
            subject = f"{subject_prefix}Advocare SA - Roster of Support ({month_name})"
            
            html_body = f"""
            <div style="font-family: 'Calibri', sans-serif; font-size: 11pt; color: #000000; line-height: 1.2;">
            <p style="margin-top: 0; margin-bottom: 10px;">Dear {name},</p>
            <p style="margin-top: 0; margin-bottom: 10px;">Please find <b>below</b> your {amended_word}roster of support {date_phrase} for approval:</p>
            <br>
            <table style="border-collapse: collapse; width: 100%; max-width: 650px; margin-top: 5px; margin-bottom: 15px; border: 1px solid #cccccc;">
                <tr>
                    <th style="font-family: 'Calibri', sans-serif; font-size: 11pt; padding: 6px 8px; border: 1px solid #cccccc; background-color: #f9f9f9; text-align: left; color: #800080;">Date</th>
                    <th style="font-family: 'Calibri', sans-serif; font-size: 11pt; padding: 6px 8px; border: 1px solid #cccccc; background-color: #f9f9f9; text-align: left; color: #800080;">Support Team Member</th>
                    <th style="font-family: 'Calibri', sans-serif; font-size: 11pt; padding: 6px 8px; border: 1px solid #cccccc; background-color: #f9f9f9; text-align: left; color: #800080;">Start</th>
                    <th style="font-family: 'Calibri', sans-serif; font-size: 11pt; padding: 6px 8px; border: 1px solid #cccccc; background-color: #f9f9f9; text-align: left; color: #800080;">End</th>
                </tr>
                {html_rows}
            </table>
            <br>
            <p style="margin-top: 0; margin-bottom: 10px;">All look OK?<br>If you have any concerns or queries, please don't hesitate to contact us 🙂</p>
            <br>
            <p style="margin: 0;">Kind regards,</p>
            <br>
            <p style="margin: 0;">Admin Team 😊</p>
            <br>
            <p style="margin: 0; color: #7030A0; font-weight: bold; font-size: 12pt;">Advocare SA</p>
            <p style="margin: 0; font-style: italic;">Compassionate community care. Your Way.</p>
            <p style="margin: 0;">Email: <a href="mailto:rostering@advocaresa.com.au" style="color: #0563C1;">rostering@advocaresa.com.au</a></p>
            <p style="margin: 0;">Phone: 0421 813 560</p>
            <p style="margin: 0;">Post: 286 Flinders St, Adelaide SA 5000</p>
            </div>
            """
            
            send_brevo_email(
                to_emails=target_emails,
                subject=subject,
                html_content=html_body
            )
            
            logs.append(f"SUCCESS: Sent roster to {name} ({', '.join(target_emails)})")
            success_count += 1
            
        except Exception as e:
            logs.append(f"ERROR processing {name}: {str(e)}")
    
    try:
        admin_email = os.environ.get('ADMIN_EMAIL', 'rostering@advocaresa.com.au')
        confirm_subject = f"Roster Run Complete - {success_count} Sent"
        confirm_body = f"<p>Roster Dispatch Summary:</p><pre>" + "\n".join(logs) + "</pre>"
        send_brevo_email([admin_email], confirm_subject, confirm_body)
    except Exception as e:
        print(f"Failed to send admin confirmation: {e}")

    return render_template_string(RESULT_HTML, success_count=success_count, logs=logs)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
