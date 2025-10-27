from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from jinja2 import Template
from docx import Document
import os
from datetime import datetime
import smtplib
from email.message import EmailMessage

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'dev_secret_key')  # Needed for flashing messages

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


# Ensure folder exists
if not os.path.exists("generated_docs"):
    os.makedirs("generated_docs")

# =======================
# EMAIL CONFIGURATION
# =======================
EMAIL_ADDRESS = os.environ.get("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD")
RECEIVER_EMAIL = os.environ.get("RECEIVER_EMAIL", EMAIL_ADDRESS)
# =======================


@app.route('/form/<service_type>')
def show_form(service_type):
    service_type = service_type.lower()
    if service_type == 'business_plan':
        return render_template('business_plan_form.html')
    elif service_type == 'business_profile':
        return render_template('business_profile_form.html')
    elif service_type == 'graphic_design':
        return render_template('graphic_design_form.html')
    else:
        flash("Invalid service selected.")
        return redirect(url_for('welcome'))


@app.route('/submit/<service_type>', methods=['POST'])
def submit_form(service_type):
    service_type = service_type.lower()

    # Common fields
    data = {
        'client_name': request.form.get('client_name'),
        'phone_number': request.form.get('phone_number'),
        'email': request.form.get('email'),
        'submission_date': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }

    if service_type == 'business_plan':
        data.update({
            'business_name': request.form.get('business_name'),
            'owner_name': request.form.get('owner_name'),
            'mission': request.form.get('mission'),
            'vision': request.form.get('vision'),
            'products': request.form.get('products'),
            'target_market': request.form.get('target_market'),
            'competitors': request.form.get('competitors'),
            'marketing_strategy': request.form.get('marketing_strategy'),
            'revenue': request.form.get('revenue'),
            'expenses': request.form.get('expenses'),
            'funding': request.form.get('funding'),
            'conclusion': request.form.get('conclusion'),
        })
        template_file = 'business_plan_template.txt'

    elif service_type == 'business_profile':
        data.update({
            'business_name': request.form.get('business_name'),
            'business_type': request.form.get('business_type'),
            'established_year': request.form.get('established_year'),
            'location': request.form.get('location'),
            'services_offered': request.form.get('services_offered'),
            'achievements': request.form.get('achievements'),
            'staff_count': request.form.get('staff_count'),
            'contact_info': request.form.get('contact_info'),
            'additional_notes': request.form.get('additional_notes'),
        })
        template_file = 'business_profile_template.txt'

    elif service_type == 'graphic_design':
        data.update({
            'project_name': request.form.get('project_name'),
            'design_type': request.form.get('design_type'),
            'details': request.form.get('details'),
            'deadline': request.form.get('deadline'),
            'budget': request.form.get('budget'),
            'additional_notes': request.form.get('additional_notes'),
        })
        template_file = 'graphic_design_template.txt'

    elif service_type == 'general_request':
        data.update({
            'request_type': request.form.get('request_type'),
            'details': request.form.get('details'),
            'delivery': request.form.get('delivery'),
        })
        template_file = 'general_request_template.txt'

    else:
        flash("Invalid service submission.")
        return redirect(url_for('welcome'))

    # Generate document
    file_path = generate_doc(data, template_file, service_type)

    # Send email with attachment, handle errors gracefully
    try:
        send_email_with_attachment(data, file_path, service_type)
        flash("Form submitted successfully! Check your email for confirmation.")
    except Exception as e:
        print(f"❌ Email sending failed: {e}")
        flash("Form submitted but failed to send email. We will contact you soon.")

    return render_template('success.html', file_name=os.path.basename(file_path))


@app.route('/download/<file_name>')
def download_file(file_name):
    path = os.path.join('generated_docs', file_name)
    return send_file(path, as_attachment=True)


def generate_doc(data, template_file, service_type):
    with open(os.path.join('templates', template_file), 'r', encoding='utf-8') as f:
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


def send_email_with_attachment(data, file_path, service_type):
    if not EMAIL_ADDRESS or not EMAIL_PASSWORD:
        raise Exception("Email credentials are not set in environment variables.")

    msg = EmailMessage()
    msg['Subject'] = f"New {service_type.replace('_', ' ').title()} Submission from {data.get('client_name')}"
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = RECEIVER_EMAIL

    # Set Reply-To so you can reply directly to the client
    if data.get('email'):
        msg['Reply-To'] = data.get('email')

    # Email body
    body = f"""
New {service_type.replace('_', ' ').title()} submission received.

Client Name: {data.get('client_name')}
Phone Number: {data.get('phone_number')}
Email: {data.get('email')}
Submitted At: {data.get('submission_date')}

Please find the attached document for more details.
    """
    msg.set_content(body)

    # Attach generated document
    with open(file_path, 'rb') as f:
        file_data = f.read()
    filename = os.path.basename(file_path)
    msg.add_attachment(
        file_data,
        maintype='application',
        subtype='vnd.openxmlformats-officedocument.wordprocessingml.document',
        filename=filename
    )

    # Send email
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(msg)
        print(f"✅ Email sent to {RECEIVER_EMAIL}")


if __name__ == '__main__':
    app.run(debug=True)
