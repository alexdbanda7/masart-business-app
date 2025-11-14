from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from jinja2 import Template
from docx import Document
import os
from dotenv import load_dotenv
from datetime import datetime
import requests

app = Flask(__name__)
app.secret_key = os.environ.get('e28dae81fd1d2c85b09bd3af3386c871', 'dev_secret_key')

# ============================
# RESEND EMAIL CONFIGURATION
# ============================
load_dotenv()  # loads variables from .env
RESEND_API_KEY = os.environ.get("re_ariaGQ4K_AkyHhdGwajBAnuxahtM5BcEj")
EMAIL_SENDER = os.environ.get("masartngs@gmail.com")  # Your verified Resend sender email
RECEIVER_EMAIL = os.environ.get("masartngs@gmail.com")  # Where submissions should be sent


# ============================
# ROUTES
# ============================

@app.route('/')
def welcome():
    return render_template('welcome.html')

@app.route('/services/business')
def services_business():
    return render_template('choose_business_type.html')

@app.route('/services/graphic')
def services_graphic():
    return render_template('graphic_design_form.html')

@app.route('/services/other')
def services_other():
    return render_template('general_request_form.html')

@app.route('/other_services/general_request')
def other_services_general_request():
    return render_template('general_request.html')


# Ensure storage folder exists
if not os.path.exists("generated_docs"):
    os.makedirs("generated_docs")


@app.route('/form/<service_type>')
def show_form(service_type):
    service_type = service_type.lower()
    match service_type:
        case 'business_plan':
            return render_template('business_plan_form.html')
        case 'business_profile':
            return render_template('business_profile_form.html')
        case 'graphic_design':
            return render_template('graphic_design_form.html')
        case 'general_request':
            return render_template('general_request_form.html')
        case _:
            flash("Invalid service selected.")
            return redirect(url_for('welcome'))


@app.route('/submit/<service_type>', methods=['POST'])
def submit_form(service_type):
    service_type = service_type.lower()

    # Common form fields
    data = {
        'client_name': request.form.get('client_name'),
        'phone_number': request.form.get('phone_number'),
        'email': request.form.get('email'),
        'submission_date': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }

    # Service-specific fields
    template_file = None

    if service_type == 'business_plan':
        template_file = 'business_plan_template.txt'
        keys = [
            'business_name', 'owner_name', 'mission', 'vision',
            'products', 'target_market', 'competitors',
            'marketing_strategy', 'revenue', 'expenses',
            'funding', 'conclusion'
        ]
        for k in keys:
            data[k] = request.form.get(k)

    elif service_type == 'business_profile':
        template_file = 'business_profile_template.txt'
        keys = [
            'business_name', 'business_type', 'established_year',
            'location', 'services_offered', 'achievements',
            'staff_count', 'contact_info', 'additional_notes'
        ]
        for k in keys:
            data[k] = request.form.get(k)

    elif service_type == 'graphic_design':
        template_file = 'graphic_design_template.txt'
        keys = [
            'project_name', 'design_type', 'details',
            'deadline', 'budget', 'additional_notes'
        ]
        for k in keys:
            data[k] = request.form.get(k)

    elif service_type == 'general_request':
        template_file = 'general_request_template.txt'
        keys = ['request_type', 'details', 'delivery']
        for k in keys:
            data[k] = request.form.get(k)

    else:
        flash("Invalid service submission.")
        return redirect(url_for('welcome'))

    # Generate Document
    file_path = generate_doc(data, template_file, service_type)

    # Send Email via Resend
    try:
        send_email_via_resend(data, file_path, service_type)
        flash("Form submitted successfully! Your document has been emailed.")
    except Exception as e:
        print(f"❌ Email sending failed: {e}")
        flash("Form submitted, but email sending failed. We will contact you soon.")

    return render_template('success.html', file_name=os.path.basename(file_path))


@app.route('/download/<file_name>')
def download_file(file_name):
    path = os.path.join('generated_docs', file_name)
    return send_file(path, as_attachment=True)


# ============================
# DOCUMENT GENERATION
# ============================

def generate_doc(data, template_file, service_type):
    """Generate a .docx file from a text template."""
    template_path = os.path.join('templates', template_file)
    with open(template_path, 'r', encoding='utf-8') as f:
        template_text = f.read()

    template = Template(template_text)
    rendered = template.render(data)

    doc = Document()
    for line in rendered.split('\n'):
        doc.add_paragraph(line)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = data.get('client_name', 'client').replace(' ', '_')

    file_name = f"{service_type}_{safe_name}_{timestamp}.docx"
    file_path = os.path.join('generated_docs', file_name)

    doc.save(file_path)
    return file_path


# ============================
# EMAIL (RESEND API)
# ============================

def send_email_via_resend(data, file_path, service_type):
    if not RESEND_API_KEY:
        raise Exception("RESEND_API_KEY is not set in environment variables.")
    if not EMAIL_SENDER:
        raise Exception("EMAIL_SENDER is not set.")
    if not RECEIVER_EMAIL:
        raise Exception("RECEIVER_EMAIL is not set.")

    url = "https://api.resend.com/emails"

    with open(file_path, "rb") as f:
        file_data = f.read()

    filename = os.path.basename(file_path)

    payload = {
        "from": EMAIL_SENDER,
        "to": [RECEIVER_EMAIL],
        "subject": f"New {service_type.replace('_', ' ').title()} Submission",
        "text": f"""
New submission received.

Client: {data.get('client_name')}
Phone: {data.get('phone_number')}
Email: {data.get('email')}
Submitted At: {data.get('submission_date')}

Document attached.
        """,
        "attachments": [
            {
                "filename": filename,
                "content": file_data.decode("latin1"),
                "disposition": "attachment"
            }
        ]
    }

    headers = {
        "Authorization": f"Bearer {RESEND_API_KEY}",
        "Content-Type": "application/json"
    }

    response = requests.post(url, json=payload, headers=headers)

    if response.status_code >= 400:
        raise Exception(f"Resend API error: {response.text}")

    print("✅ Email successfully sent via Resend.")


# ============================
# RUN APP
# ============================

if __name__ == '__main__':
    app.run(debug=True)
