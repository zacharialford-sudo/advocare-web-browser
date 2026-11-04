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
<div class="card shadow-sm mb-4">
    <div class="card-header">Upload Schedule (CSV)</div>
    <div class="card-body">
        <form method="POST" action="/upload" enctype="multipart/form-data">
            <div class="mb-3">
                <input class="form-control" type="file" name="schedule_file" accept=".csv" required>
            </div>
            <button type="submit" class="btn btn-primary">Process Schedule</button>
        </form>
    </div>
</div>
""")

EMAILS_HTML = BASE_HTML.replace('{% block content %}{% endblock %}', """
<div class="card shadow-sm">
    <div class="card-header">Client Email Directory</div>
    <div class="card-body">
        <form method="POST" action="/emails" class="mb-4 bg-light p-3 border rounded">
            <h5>Add or Update Client Email</h5>
            <div class="row g-2 align-items-center">
                <div class="col-md-5">
                    <input type="text" name="client_name" class="form-control" placeholder="Exact Client Name" required>
                </div>
                <div class="col-md-5">
                    <input type="text" name="client_emails" class="form-control" placeholder="Emails (comma separated)" required>
                </div>
                <div class="col-md-2">
                    <button type="submit" class="btn btn-primary w-100">Save</button>
                </div>
            </div>
        </form>
        <hr>
        <div class="table-responsive">
            <table class="table table-striped">
                <thead><tr><th>Client Name</th><th>Configured Emails</th><th>Action</th></tr></thead>
                <tbody>
                    {% for name, emails in directory.items() %}
                    <tr>
                        <td>{{ name }}</td>
                        <td>{{ emails | join(', ') }}</td>
                        <td>
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

def send_brevo_email(to_emails, subject, html_content, attachment_path=None, filename=None):
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
    
    if attachment_path and filename:
        with open(attachment_path, "rb") as f:
            encoded_file = base64.b64encode(f.read()).decode('utf-8')
        payload["attachment"] = [{"name": filename, "content": encoded_file}]
        
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

@app.route('/upload', methods=['POST'])
def upload():
    if not check_auth(): return redirect('/login')
    file = request.files.get('schedule_file')
    if not file or file.filename == '':
        flash("No file selected", "danger")
        return redirect('/')
        
    filename = f"schedule_{uuid.uuid4().hex}.csv"
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)
    
    try:
        df = pd.read_csv(filepath)
        df = df[df['Shift Status'].str.lower() != 'cancelled']
        df = df[~df['Staff'].astype(str).str.contains("CANCELLED", case=False, na=False)]
        df = df.dropna(subset=['Start Date Time', 'End Date Time'])
        
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
        flash(f"Error processing CSV: {str(e)}", "danger")
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
            
            month_name = client_data.iloc[0]['Start Date Time'].strftime('%B')
            
            doc = Document()
            style = doc.styles['Normal']
            font = style.font
            font.name = 'Calibri'
            font.size = Pt(11) 
            
            doc.add_paragraph(f"Dear {name},") 
            doc.add_paragraph(f"Please find below your {amended_word}roster of support until the end of {month_name} for approval:\n") 
            
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
            doc_filename_short = f"{file_prefix}{safe_name}_{month_name}_Roster.docx"
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
            <p style="margin-top: 0; margin-bottom: 10px;">Please find <b>below</b> your {amended_word}roster of support until the end of {month_name} for approval:</p>
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
