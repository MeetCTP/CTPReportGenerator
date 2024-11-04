from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for
from sqlalchemy import create_engine, text
import msal
import pandas as pd
from flask_talisman import Talisman
from generators.appointment_match_agora import generate_appointment_agora_report
from generators.appointment_match_insight import generate_appointment_insight_report
from generators.active_contacts_report import generate_active_contacts_report
from generators.late_conversions_report import generate_late_conversions_report
from generators.no_show_late_cancel_report import generate_no_show_late_cancel_report
from generators.provider_sessions_report import generate_provider_sessions_report
from generators.provider_connections_report import generate_provider_connections_report
from generators.forty_eight_conversions import generate_unconverted_time_report
from generators.forty_eight_conversions import reminder_email
from generators.client_cancellation_report import generate_client_cancel_report
from generators.util_tracker import generate_util_tracker
from generators.certification_expiration import generate_cert_exp_report
from flask_cors import CORS
from datetime import datetime
from io import BytesIO
from logging.handlers import RotatingFileHandler
import datetime as dt
import json
import urllib.parse
import logging
import traceback
import tempfile
import os

csp = {
    'default-src': "'self'",
    'style-src': [
        "'self'", 
        "'unsafe-inline'",
    ]
}

app = Flask(__name__)
CORS(app)
#Talisman(app, content_security_policy=csp, force_https=True)

# Make the WSGI interface available at the top level so wfastcgi can get it.
wsgi_app = app.wsgi_app

connection_string = f"mssql+pymssql://MeetCTP\Administrator:$Unlock01@CTP-DB/CTPHOME"
engine = create_engine(connection_string)

#People
lisa = f"Lisa.A.Kowalski"
dan = f"Dan.A.Costello"
admin = f"Administrator"
fabian = f"Fabian.Legarreta"
josh = f"Joshua.Bliven"
aaron = f"Aaron.A.Robertson"
linda = f"Linda.Brooks"
eileen = f"Eileen.H.Council"
cari = f"Cari.Tomczyk"
amy = f"Amy.P.Ronen"
stacey = f"Stacey.A.Nardo"
christi = f"Christina.K.Sampson"
greg = f"Gregory.T.Hughes"
jesse = f"Jesse.Petrecz"
olivia = f"Olivia.a.DiPasquale"
megan = f"Megan.Leighton"

#Groups
admin_group = [lisa, admin, cari]
recruiting_group = [amy, stacey]
clinical_group = [megan, jesse]
accounting_group = [eileen, greg, cari]
student_services_group = [eileen, christi, olivia]
human_resources_group = [aaron, linda]
it_group = [fabian, dan]
testing_group = [josh, fabian]
site_mod_group = [josh, fabian, lisa, admin, eileen, aaron, amy]

def handle_submit_form_data(table, data):
    if table == 'News_Posts':
        query = text("""
        INSERT INTO dbo.News_Posts (Title, Body, CreatedBy, RowModifiedAt) 
        VALUES (:Title, :Body, :CreatedBy, :RowModifiedAt)
        """)
    elif table == 'Notifications':
        query = text("""
        INSERT INTO dbo.Notifications (EventDate, Body, CreatedBy, RowModifiedAt) 
        VALUES (:EventDate, :Body, :CreatedBy, :RowModifiedAt)
        """)
    elif table == 'WeeklyQA':
        query = text("""
        INSERT INTO dbo.WeeklyQA (Body, CreatedBy, RowModifiedAt) 
        VALUES (:Body, :CreatedBy, :RowModifiedAt)
        """)

    try:
        with engine.connect() as connection:
            connection.execute(query, data)
            connection.commit()
        print(f"Query executed successfully: {query}, with data: {data}")
    except Exception as e:
        return print(e)
    
def handle_delete_homepage_item(table, id):
    try:
        # Ensure the id is an integer
        id = int(id)
    except ValueError:
        print(f"Invalid id value: {id}")
        return

    if table == 'News_Posts':
        query = text("""
        DELETE FROM dbo.News_Posts WHERE NewsId = :id
        """)
    elif table == 'Notifications':
        query = text("""
        DELETE FROM dbo.Notifications WHERE NotifId = :id
        """)
    elif table == 'WeeklyQA':
        query = text("""
        DELETE FROM dbo.WeeklyQA WHERE QAId = :id;
        DELETE FROM dbo.WeeklyQAResponses WHERE QuestionId = :id;
        """)

    try:
        with engine.connect() as connection:
            connection.execute(query, {'id': id})
            connection.commit()
        print(f"Query executed successfully: {query}, with data: {id}")
    except Exception as e:
        return print(e)

# Define a function to fetch data from the database
def fetch_data():
    news_query = text("SELECT NewsId, Title, Body FROM dbo.News_Posts")
    notifications_query = text("SELECT NotifId, EventDate, Body FROM dbo.Notifications")
    weekly_qa_query = text("SELECT QAId, Body FROM dbo.WeeklyQA")
    responses_query = text("""
        SELECT QuestionId, ResponseBody, CreatedBy, CreatedAt
        FROM dbo.QAResponseView
    """)

    with engine.connect() as connection:
        news_articles = connection.execute(news_query).fetchall()
        notifications = connection.execute(notifications_query).fetchall()
        weekly_qas = connection.execute(weekly_qa_query).fetchall()
        responses = connection.execute(responses_query).fetchall()

    notifications.reverse()
    weekly_qas.reverse()
    news_articles.reverse()
    #datetime.strftime(notifications.EventDate, '%m/%d/%Y')

    # Create a dictionary to map QAId to their responses
    qa_dict = {qa.QAId: {'Id': qa.QAId, 'Body': qa.Body, 'responses': []} for qa in weekly_qas}
    
    for response in responses:
        if response.QuestionId in qa_dict:
            qa_dict[response.QuestionId]['responses'].append({
                'ResponseBody': response.ResponseBody,
                'CreatedBy': response.CreatedBy,
                'CreatedAt': datetime.strftime(response.CreatedAt, '%m/%d/%Y %I:%M%p')
            })

    return news_articles, notifications, list(qa_dict.values())

@app.route('/test-user')
def test_user():
    username = request.environ.get('REMOTE_USER')
    username = str(username).split('\\')[-1]
    if username in admin_group:
        return f"Hello, {username}"
    else:
        return "Hello, anonymous user!"

@app.route('/')
def home():
    username = request.environ.get('REMOTE_USER')
    username = str(username).split('\\')[-1]
    news_articles, notifications, weekly_qas = fetch_data()
    is_mod = username in site_mod_group
    return render_template('home.html', 
                           news_articles=news_articles, 
                           notifications=notifications, 
                           weekly_qas=weekly_qas,
                           is_mod=is_mod)

@app.route('/submit-response', methods=['POST'])
def submit_response():
    data = request.get_json()
    question_id = data['questionId']
    response_body = data['responseBody']
    username = request.environ.get('REMOTE_USER')
    username = str(username).split('\\')[-1]

    response_data = {
        'QuestionId': question_id,
        'ResponseBody': response_body,
        'CreatedBy': username,
        'CreatedAt': datetime.now()
    }

    query = text("""
        INSERT INTO dbo.WeeklyQAResponses (QuestionId, ResponseBody, CreatedBy, CreatedAt)
        VALUES (:QuestionId, :ResponseBody, :CreatedBy, :CreatedAt)
    """)

    try:
        with engine.connect() as connection:
            connection.execute(query, response_data)
            connection.commit()
        return jsonify({'success': True, 'createdBy': username}), 200
    except Exception as e:
        print(f"Error submitting response: {e}")
        return jsonify({'success': False, 'message': str(e)}), 500
    
@app.route('/delete-home-item', methods=['POST'])
def delete_item():
    data = request.get_json()
    item_id = int(data['id'])
    table_name = data['table']
    
    try:
        handle_delete_homepage_item(table_name, item_id)
        return jsonify({"status": "success"}), 200
    except Exception as e:
        print(f"Error: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/set')
def site_mod_page():
    username = request.environ.get('REMOTE_USER')
    username = str(username).split('\\')[-1]
    if username in site_mod_group:
        return render_template('site_mod.html')
    else:
        return redirect(url_for('access_denied'))

@app.route('/set/submit-form-data', methods=['POST'])
def submit_form_data():
    username = request.environ.get('REMOTE_USER')
    username = str(username).split('\\')[-1]
    form_data = request.form.to_dict()
    form_type = form_data.pop('form_type', None)
    
    form_data.update({
        'CreatedBy': username,
        'RowModifiedAt': datetime.now()
    })

    if 'filename' in request.files:
        file = request.files['filename']
        file_data = file.read()
        form_data['Attachment'] = file_data
    else:
        form_data['Attachment'] = None

    if form_type == 'news':
        handle_submit_form_data('News_Posts', form_data)
    elif form_type == 'notification':
        handle_submit_form_data('Notifications', form_data)
    elif form_type == 'qa':
        handle_submit_form_data('WeeklyQA', form_data)

    return jsonify({'message': 'Form submitted successfully'}), 200

@app.route('/report-generator')
def reports():
    """Renders the reports home page template to the /reports url."""
    username = request.environ.get('REMOTE_USER')
    username = str(username).split('\\')[-1]
    if username in admin_group:
        return render_template('all-prod-reports.html')
    elif username in clinical_group:
        return render_template('clinical-reports.html')
    elif username in student_services_group:
        return render_template('student-services-reports.html')
    elif username in accounting_group:
        return render_template('accounting-reports.html')
    elif username in human_resources_group:
        return render_template('human-resources-reports.html')
    elif username in testing_group:
        return render_template('all-reports.html')
    elif username in it_group:
        return render_template('all-reports.html')
    else:
        return redirect(url_for('access_denied'))
    
@app.route('/hipaa-training')
def hipaa_stuff():
    return render_template('hipaa-stuff.html')

@app.route('/calendar')
def calendar():
    return render_template('holiday-calendar.html')
    
@app.route('/access-denied')
def access_denied():
    return render_template('access-denied.html')

@app.route('/report-generator/agora-match')
def agora_match_report():
    return render_template('agora-report.html')

@app.route('/report-generator/insight-match')
def insight_match_report():
    return render_template('insight-report.html')

@app.route('/report-generator/active-contacts')
def active_contacts_report():
    return render_template('active-contacts.html')

@app.route('/report-generator/late-conversions')
def late_conversions_report():
    return render_template('late-conversions.html')

@app.route('/report-generator/no-show-late-cancel')
def no_show_late_cancel_report():
    return render_template('no-show-late-cancel.html')

@app.route('/report-generator/provider-sessions')
def provider_sessions():
    return render_template('provider-sessions.html')

@app.route('/report-generator/provider-connections')
def provider_connections():
    return render_template('provider-connections.html')

company_role_options = ["CompanyRole: Employee", "CompanyRole: Contractor"]

@app.route('/report-generator/forty-eight-hour-warning')
def forty_eight_warning():
    return render_template('forty-eight.html', company_roles=company_role_options)

@app.route('/report-generator/client-cancels')
def client_cancellations():
    return render_template('client-cancel.html')

@app.route('/report-generator/clinical-util-tracker')
def clinical_util():
    return render_template('clinical-util.html')

@app.route('/report-generator/certification-expiration')
def certification_expiration():
    return render_template('certification-expiration.html')

@app.route('/report-generator/agora-match/generate-report', methods=['POST'])
def generate_report():
    """Generates the Agora Match report."""
    if request.headers['Content-Type'] == 'application/json':
        # Get start and end date values from the request body
        data = request.get_json()
        start_date = data['start_date']
        end_date = data['end_date']

        try:
            # Call your Python function to generate the report
            excel_file = generate_appointment_agora_report(start_date, end_date)

            # Return the Excel file as a download to the browser
            return send_file(
                excel_file,
                as_attachment=True,
                download_name=f"Agora_Appointment_Match_Report_{start_date}-{end_date}.xlsx"
            )
        except Exception as e:
            print('Exception occurred: ', e)
            return jsonify({'error': str(e)}), 500
    else:
        return jsonify({'error': 'Unsupported Media Type'}), 415
    
@app.route('/report-generator/insight-match/generate-report', methods=['POST'])
def handle_generate_insight_report():
    """Generates the Insight Match report."""
    if request.headers['Content-Type'] == 'application/json':
        # Get start and end date values from the request body
        data = request.get_json()
        start_date = data['start_date']
        end_date = data['end_date']

        try:
            # Call your Python function to generate the report
            excel_file = generate_appointment_insight_report(start_date, end_date)

            # Return the Excel file as a download to the browser
            return send_file(
                excel_file,
                as_attachment=True,
                download_name=f"Insight_Appointment_Match_Report_{start_date}-{end_date}.xlsx"
            )
        except Exception as e:
            print('Exception occurred: ', e)
            return jsonify({'error': str(e)}), 500
    else:
        return jsonify({'error': 'Unsupported Media Type'}), 415
    
@app.route('/report-generator/active-contacts/generate-report', methods=['POST'])
def handle_generate_active_contacts_report():
    """Generates the Active Contacts report."""
    if request.headers['Content-Type'] == 'application/json':
        data = request.get_json()
        status = data['status']
        pg_type = data['pg_type']
        service_types = data['service_type']

        try:
            # Call your Python function to generate the report
            excel_file = generate_active_contacts_report(status, pg_type, service_types)

            # Return the Excel file as a download to the browser
            return send_file(
                excel_file,
                as_attachment=True,
                download_name=f"Active_Contacts_Report_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
            )
        except Exception as e:
            print('Exception occurred: ', e)
            return jsonify({'error': str(e)}), 500
    else:
        return jsonify({'error': 'Unsupported Media Type'}), 415
    
@app.route('/report-generator/late-conversions/generate-report', methods=['POST'])
def handle_generate_late_conversions_report():
    """Generates the Late Conversions report."""
    if request.headers['Content-Type'] == 'application/json':
        data = request.get_json()
        app_start = datetime.strptime(data['app_start'], '%Y-%m-%d')
        app_end = datetime.strptime(data['app_end'], '%Y-%m-%d')
        converted_after = datetime.strptime(data['converted_after'], '%Y-%m-%d')

        try:
            # Call your Python function to generate the report
            excel_file = generate_late_conversions_report(app_start, app_end, converted_after)

            # Return the Excel file as a download to the browser
            return send_file(
                excel_file,
                as_attachment=True,
                download_name=f"Late_Conversions_Report_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
            )
        except Exception as e:
            print('Exception occurred: ', e)
            return jsonify({'error': str(e)}), 500
    else:
        return jsonify({'error': 'Unsupported Media Type'}), 415
    
@app.route('/report-generator/no-show-late-cancel/generate-report', methods=['POST'])
def handle_generate_no_show_late_cancel_report():
    """Generates the No Show/Late Cancel report."""
    if request.headers['Content-Type'] == 'application/json':
        data = request.get_json()
        app_start = datetime.strptime(data['app_start'], '%Y-%m-%d')
        app_end = datetime.strptime(data['app_end'], '%Y-%m-%d')
        provider = data['provider']
        client = data['client']
        school = data['school']

        try:
            # Call your Python function to generate the report
            excel_file = generate_no_show_late_cancel_report(app_start, app_end, provider, client, school)

            # Return the Excel file as a download to the browser
            return send_file(
                excel_file,
                as_attachment=True,
                download_name=f"No_Show_Late_Cancel_Report_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
            )
        except Exception as e:
            print('Exception occurred: ', e)
            return jsonify({'error': str(e)}), 500
    else:
        return jsonify({'error': 'Unsupported Media Type'}), 415
    
@app.route('/report-generator/provider-sessions/generate-report', methods=['POST'])
def handle_generate_provider_sessions_report():
    """Generates the Provider Sessions report."""
    if request.headers['Content-Type'] == 'application/json':
        data = request.get_json()
        range_start = datetime.strptime(data['range_start'], '%Y-%m-%d')
        range_end = datetime.strptime(data['range_end'], '%Y-%m-%d')
        supervisors = data['supervisors']
        status = data['status']

        try:
            # Call your Python function to generate the report
            excel_file = generate_provider_sessions_report(range_start, range_end, supervisors, status)

            # Return the Excel file as a download to the browser
            return send_file(
                excel_file,
                as_attachment=True,
                download_name=f"Provider_Sessions_Report_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
            )
        except Exception as e:
            print('Exception occurred: ', e)
            return jsonify({'error': str(e)}), 500
    else:
        return jsonify({'error': 'Unsupported Media Type'}), 415
    
@app.route('/report-generator/provider-connections/generate-report', methods=['POST'])
def handle_generate_provider_connections_report():
    if request.headers['Content-Type'] == 'application/json':
        try:
            excel_file = generate_provider_connections_report()
            return send_file(
                excel_file,
                as_attachment=True,
                download_name=f"Provider_Sessions_Report_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
            )
        except Exception as e:
            print('Exception occurred: ', e)
            return jsonify({'error': str(e)}), 500
    else:
        return jsonify({'error': 'Unsupported Media Type'}), 415
    
@app.route('/report-generator/forty-eight-hour-warning/generate-report', methods=['POST'])
def handle_generate_forty_eight_hour_report():
    if request.headers['Content-Type'] == 'application/json':
        data = request.json
        selected_roles = data.get('company_roles', [])
        start_date = data.get('start_date')
        end_date = data.get('end_date')
        
        try:
            _, _, _, excel_file = generate_unconverted_time_report(selected_roles, start_date, end_date)
            return send_file(
                excel_file,
                as_attachment=True,
                download_name=f"48_Hour_Late_Conversions_Report_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
            )
        except Exception as e:
            print('Exception occurred: ', e)
            return jsonify({'error': str(e)}), 500
    else:
        return jsonify({'error': 'Unsupported Media Type'}), 415

cached_warning_list = []
cached_non_payment_list = []

@app.route('/report-generator/forty-eight-hour-warning/get-mailing-list', methods=['POST'])
def get_mailing_list():
    global cached_warning_list, cached_non_payment_list

    if request.headers['Content-Type'] == 'application/json':
        data = request.json
        selected_roles = data.get('company_roles', [])
        start_date = data.get('start_date')
        end_date = data.get('end_date')
        
        try:
            mailing_list, warning_list, non_payment_list, _ = generate_unconverted_time_report(selected_roles, start_date, end_date)
            cached_warning_list = warning_list
            cached_non_payment_list = non_payment_list
            return jsonify({
                'mailing_list': mailing_list,
                'warning_list': warning_list,
                'non_payment_list': non_payment_list
            })
        except Exception as e:
            print('Exception occurred: ', e)
            return jsonify({'error': str(e)}), 500
    else:
        return jsonify({'error': 'Unsupported Media Type'}), 415

@app.route('/report-generator/forty-eight-hour-warning/send-emails', methods=['POST'])
def send_emails():
    global cached_warning_list, cached_non_payment_list
    if request.headers['Content-Type'] == 'application/json':
        data = request.json
        selected_providers = data.get('selectedProviders', {})

        try:
            reminder_email(selected_providers, cached_warning_list, cached_non_payment_list)
            return jsonify({'message': 'Emails sent successfully!'}), 200
        except Exception as e:
            print('Exception occurred: ', e)
            return jsonify({'error': str(e)}), 500
    else:
        return jsonify({'error': 'Unsupported Media Type'}), 415
    
@app.route('/report-generator/client-cancels/generate-report', methods=['POST'])
def handle_generate_client_cancel_report():
    if request.headers['Content-Type'] == 'application/json':
        data = request.get_json()
        range_start = data.get('range_start')
        provider = data.get('provider')
        client = data.get('client')
        cancel_reasons = data.get('cancel_reasons', [])

        if not cancel_reasons:
            return jsonify({"error": "At least one cancellation reason must be selected"}), 400
        
        try:
            report_file = generate_client_cancel_report(provider, client, cancel_reasons, range_start)
            return send_file(
                report_file,
                as_attachment=True,
                download_name=f"Client-Cancellation-Report.xlsx"
            )
        except Exception as e:
            return jsonify({"error": str(e)}), 500
    else:
        return jsonify({'error': 'Unsupported Media Type'}), 415

@app.route('/report-generator/clinical-util-tracker/generate-report', methods=['POST'])
def handle_generate_clinical_util_report():
    if request.headers['Content-Type'] == 'application/json':
        data = request.get_json()
        start_date = data.get('start_date')
        end_date = data.get('end_date')
        provider = data.get('provider')

        if not provider:
            return jsonify({"error": "Provider must be specified."})

        try:
            report_file = generate_util_tracker(start_date, end_date, provider)
            return send_file(
                report_file,
                as_attachment=True,
                download_name=f"Clinical_Util_Tracker_'{provider}'_'{start_date}'-'{end_date}'.xlsx"
            )
        except Exception as e:
            return jsonify({"error": str(e)}), 500
    else:
        return jsonify({'error': 'Unsupported Media Type'}), 415

@app.route('/report-generator/certification-expiration/generate-report', methods=["POST"])
def handle_generate_cert_exp_report():
    if request.headers['Content-Type'] == 'application/json':
        data = request.get_json()
        timeframe = data.get('timeframe')

        try:
            report_file = generate_cert_exp_report(timeframe)
            return send_file(
                report_file,
                as_attachment=True,
                download_name=f"Certification_Expiration.xlsx"
            )
        except Exception as e:
            return jsonify({'error': str(e)}), 500
    else:
        return jsonify({'error': 'Unsupported Media Type'}), 415