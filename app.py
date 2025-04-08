from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for
from sqlalchemy import create_engine, text
import msal
import pandas as pd
import openpyxl
from flask_talisman import Talisman
from generators.appointment_match_agora import generate_appointment_agora_report
from generators.appointment_match_insight import generate_appointment_insight_report
from generators.active_contacts_report import generate_active_contacts_report
from generators.late_conversions_report import generate_late_conversions_report
from generators.no_show_late_cancel_report import generate_no_show_late_cancel_report
from generators.no_show_late_cancel_report import generate_no_show_late_cancel_report_single
from generators.provider_sessions_report import generate_provider_sessions_report
from generators.provider_connections_report import generate_provider_connections_report
from generators.forty_eight_conversions import generate_unconverted_time_report
from generators.forty_eight_conversions import reminder_email
from generators.client_cancellation_report import generate_client_cancel_report
from generators.util_tracker import generate_util_tracker
from generators.util_tracker import calculate_cancellation_percentage
from generators.certification_expiration import generate_cert_exp_report
from generators.pad_indirect_time_report import generate_pad_indirect
from generators.monthly_active_users_report import generate_monthly_active_users
from generators.appt_overlap_pandas import generate_appt_overlap_report
from generators.original_agora_report import generate_original_agora_report
from generators.original_insight_report import generate_original_insight_report
from generators.monthly_at_report import get_all_at_tables
from generators.valid_emails import generate_valid_email_report
from generators.no_contact_list import get_no_contact_list
from generators.code_look_up import code_search
from flask_cors import CORS
from datetime import datetime
from io import BytesIO
from logging.handlers import RotatingFileHandler
import datetime as dt
import base64
import pymssql
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
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024
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
christi = f"Christina.K.Sampson"
greg = f"Gregory.T.Hughes"
jesse = f"Jesse.Petrecz"
olivia = f"Olivia.a.DiPasquale"
megan = f"Megan.Leighton"
deborah = f"Deborah.Debrule"
kim = f"Kimberly.D.Trate"

#Groups
admin_group = [lisa, admin, dan]
recruiting_group = [amy, kim]
clinical_group = []
accounting_group = [eileen, greg, deborah]
student_services_group = [eileen, christi, olivia]
human_resources_group = [aaron, linda]
testing_group = [josh, fabian, megan, cari]
site_mod_group = [josh, fabian, lisa, admin, eileen, aaron, amy]

def handle_submit_form_data(table, data):
    if table == 'News_Posts':
        query = text("""
            INSERT INTO dbo.News_Posts (Title, Body, Attachment, CreatedBy, RowModifiedAt) 
            VALUES (:Title, :Body, :Attachment, :CreatedBy, :RowModifiedAt)
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
        DELETE FROM dbo.WeeklyQAResponses WHERE QuestionId = :id;
        DELETE FROM dbo.WeeklyQA WHERE QAId = :id;
        """)
    elif table == 'WeeklyQAResponses':
        query = text("""
        DELETE FROM dbo.WeeklyQAResponses WHERE Id = :id;
        """)

    try:
        with engine.connect() as connection:
            connection.execute(query, {'id': id})
            connection.commit()
        print(f"Query executed successfully: {query}, with data: {id}")
    except Exception as e:
        return print(e)

def get_image_as_base64(image_data):
    return base64.b64encode(image_data).decode('utf-8') if image_data else None

def fetch_data():
    news_query = text("SELECT NewsId, Title, Body, Attachment, RowModifiedAt FROM dbo.News_Posts ORDER BY RowModifiedAt DESC")
    notifications_query = text("SELECT NotifId, EventDate, Body FROM dbo.Notifications ORDER BY RowModifiedAt DESC")
    weekly_qa_query = text("SELECT QAId, Body FROM dbo.WeeklyQA ORDER BY RowModifiedAt DESC")
    responses_query = text("""
        SELECT ResponseId, QuestionId, ResponseBody, CreatedBy, CreatedAt
        FROM dbo.QAResponseView
    """)

    with engine.connect() as connection:
        news_articles = connection.execute(news_query).fetchall()
        notifications = connection.execute(notifications_query).fetchall()
        weekly_qas = connection.execute(weekly_qa_query).fetchall()
        responses = connection.execute(responses_query).fetchall()

    #datetime.strftime(notifications.EventDate, '%m/%d/%Y')

    news_articles = [
        {
            'NewsId': article.NewsId,
            'Title': article.Title,
            'Body': article.Body,
            'ImageBase64': get_image_as_base64(article.Attachment),
            'RowModifiedAt': article.RowModifiedAt.strftime('%m/%d/%Y %I:%M%p')
        }
        for article in news_articles
    ]

    notifications = [
        {
            'NotifId': notif.NotifId,
            'EventDate': notif.EventDate.strftime('%m/%d/%Y'),
            'Body': notif.Body
        }
        for notif in notifications
    ]

    qa_dict = {qa.QAId: {'Id': qa.QAId, 'Body': qa.Body, 'responses': []} for qa in weekly_qas}
    
    for response in responses:
        if response.QuestionId in qa_dict:
            qa_dict[response.QuestionId]['responses'].append({
                'ResponseId': response.ResponseId,
                'ResponseBody': response.ResponseBody,
                'CreatedBy': response.CreatedBy,
                'CreatedAt': datetime.strftime(response.CreatedAt, '%m/%d/%Y %I:%M%p')
            })

    return news_articles, notifications, list(qa_dict.values())

def search_table(sheet, query):
        results = []

        for row in sheet.iter_rows(min_row=2, values_only=True):
            code_name, desc, role, session_type, service_type, interaction, *details = row
            
            if (code_name and query in str(code_name).lower()) or (desc and query in str(desc).lower()):
                results.append(row)
        
        return results

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

@app.route('/submit-search', methods=['POST'])
def search_service_codes():
    try:
        data = request.get_json()
        look_up = data.get('query', '').strip().lower()

        path = os.path.join('static', 'servicecodetable.xlsx')
        if not os.path.exists(path):
            return jsonify({'error': 'File not found.'}), 500

        workbook = openpyxl.load_workbook(path)
        sheet = workbook.active

        search_results = search_table(sheet, look_up)

        if search_results:
            return jsonify({'results': search_results})
        else:
            return jsonify({'message': 'No matches found.', 'results': []}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

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
    try:
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
            form_data['Attachment'] = file.read()
        else:
            form_data['Attachment'] = None

        if form_type == 'news':
            handle_submit_form_data('News_Posts', form_data)
        elif form_type == 'notification':
            handle_submit_form_data('Notifications', form_data)
        elif form_type == 'qa':
            handle_submit_form_data('WeeklyQA', form_data)

        return jsonify({'message': 'Form submitted successfully'}), 200

    except Exception as e:
        print(f"Error processing form submission: {e}")
        return jsonify({'error': 'Internal Server Error', 'details': str(e)}), 500

@app.route('/report-generator')
def reports():
    """Renders the reports home page template to the /reports url."""
    username = request.environ.get('REMOTE_USER')
    username = str(username).split('\\')[-1]
    if username in admin_group:
        return render_template('all-prod-reports.html')
    elif username in clinical_group:
        return render_template('all-prod-reports.html')
    elif username in student_services_group:
        return render_template('all-prod-reports.html')
    elif username in accounting_group:
        return render_template('all-prod-reports.html')
    elif username in human_resources_group:
        return render_template('all-prod-reports.html')
    elif username in recruiting_group:
        return render_template('all-prod-reports.html')
    elif username in testing_group:
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

@app.route('/report-generator/pad-indirect')
def pad_indirect():
    return render_template('pad-indirect.html')

@app.route('/report-generator/monthly-active')
def monthly_active():
    return render_template('monthly-active.html')

@app.route('/report-generator/school-matching')
def school_match():
    return render_template('school-match.html')

@app.route('/report-generator/appt-overlap')
def appt_overlap():
    return render_template('appt-overlap.html')

@app.route('/report-generator/school-matching/generate-report', methods=['POST'])
def handle_generate_school_matching_report():
    start_date = request.form.get('start_date')
    end_date = request.form.get('end_date')
    school_choice = request.form.get('school')
    pg_type = request.form.get('pg_type')
    excel_file = request.files.get('excel_file')

    # Validate inputs
    if not start_date or not end_date:
        return jsonify({'error': 'Start and end dates are required.'}), 400
    if not school_choice:
        return jsonify({'error': 'School choice is required.'}), 400

    try:
        if school_choice == 'Agora':
            report_file = generate_appointment_agora_report(start_date, end_date, excel_file, pg_type)
        elif school_choice == 'Insight':
            report_file = generate_appointment_insight_report(start_date, end_date, excel_file, pg_type)
        else:
            return jsonify({'error': 'Invalid school choice.'}), 400

        return send_file(
            report_file,
            as_attachment=True,
            download_name=f"{school_choice}_Appointment_Match_Report_{start_date}-{end_date}.xlsx"
        )
    except Exception as e:
        print('Exception occurred:', e)
        return jsonify({'error': str(e)}), 500

@app.route('/report-generator/agora-match/generate-report', methods=['POST'])
def generate_report():
    if request.headers['Content-Type'] == 'application/json':
        data = request.get_json()
        start_date = data['start_date']
        end_date = data['end_date']

        try:
            excel_file = generate_original_agora_report(start_date, end_date)

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
        data = request.get_json()
        start_date = data['start_date']
        end_date = data['end_date']

        try:
            excel_file = generate_original_insight_report(start_date, end_date)

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
            if data['single'] != 1:
                excel_file = generate_no_show_late_cancel_report(app_start, app_end, provider, client, school)
            else:
                excel_file = generate_no_show_late_cancel_report_single(app_start, app_end, provider, client, school)
            
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
        start_date = data.get('start_date')
        end_date = data.get('end_date')
        provider = data.get('provider')
        client = data.get('client')
        cancel_reasons = data.get('cancel_reasons', [])

        if not cancel_reasons:
            return jsonify({"error": "At least one cancellation reason must be selected"}), 400
        
        try:
            report_file = generate_client_cancel_report(provider, client, cancel_reasons, start_date, end_date)
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
        company_role = data.get('company_role')

        file_name = f"Clinical_Util_Tracker_'{company_role}'_'{start_date}'-'{end_date}'.xlsx"

        try:
            report_file = generate_util_tracker(start_date, end_date, company_role)
            return send_file(
                report_file,
                as_attachment=True,
                download_name=file_name
            )
        except Exception as e:
            return jsonify({"error": str(e)}), 500
    else:
        return jsonify({'error': 'Unsupported Media Type'}), 415

@app.route('/report-generator/certification-expiration/generate-report', methods=["POST"])
def handle_generate_cert_exp_report():
    if request.headers['Content-Type'] == 'application/json':
        data = request.get_json()
        status = data.get('status')
        timeframe = data.get('timeframe')
        provider = data.get('provider')

        try:
            report_file = generate_cert_exp_report(status, timeframe, provider)
            return send_file(
                report_file,
                as_attachment=True,
                download_name=f"Certification_Expiration.xlsx"
            )
        except Exception as e:
            return jsonify({'error': str(e)}), 500
    else:
        return jsonify({'error': 'Unsupported Media Type'}), 415

@app.route('/report-generator/pad-indirect/generate-report', methods=["POST"])
def handle_generate_pad_indirect_report():
    if request.headers['Content-Type'] == 'application/json':
        data = request.get_json()
        start_date = data.get('start_date')
        end_date = data.get('end_date')
        
        try:
            report_file = generate_pad_indirect(start_date, end_date)
            return send_file(
                report_file,
                as_attachment=True,
                download_name=f"PAD_Indirect_Time_Report_'{start_date}'-'{end_date}'.xlsx"
            )
        except Exception as e:
            return jsonify({'error': str(e)}), 500
    else:
        return jsonify({'error': 'Unsupported Media Type'}), 415

@app.route('/report-generator/monthly-active/generate-report', methods=["POST"])
def handle_generate_monthly_active_users_report():
    if request.headers['Content-Type'] == 'application/json':
        data = request.get_json()
        start_date = data.get('start_date')
        end_date = data.get('end_date')
        
        try:
            report_file = generate_monthly_active_users(start_date, end_date)
            return send_file(
                report_file,
                as_attachment=True,
                download_name=f"Monthly_Active_Users_Report_'{start_date}'-'{end_date}'.xlsx"
            )
        except Exception as e:
            return jsonify({'error': str(e)}), 500
    else:
        return jsonify({'error': 'Unsupported Media Type'}), 415

@app.route('/airtable-test')
def render_airtable_page():
    return render_template('airtable-test.html')

@app.route('/report-generator/valid-emails')
def valid_emails_page():
    return render_template('valid-emails.html')

@app.route('/airtable-test/generate-report', methods=["POST"])
def airtable_test_page():
    if request.headers['Content-Type'] == 'application/json':
        data = request.get_json()
        start_date = data['start_date']
        end_date = data['end_date']

        try:
            report_file = get_all_at_tables(start_date, end_date)
            return send_file(
                report_file,
                as_attachment=True,
                download_name=f"All_Airtables.xlsx"
            )
        except Exception as e:
                return jsonify({'error': str(e)}), 500
    else:
        return jsonify({'error': 'Unsupported Media Type'}), 415
    
@app.route('/report-generator/valid-emails/generate-report', methods=["POST"])
def handle_generate_valid_email_report():
    try:
        report_file = get_no_contact_list()
        return send_file(
            report_file,
            as_attachment=True,
            download_name=f"Valid_Email_Addresses.xlsx"
        )
    except Exception as e:
            return jsonify({'error': str(e)}), 500

@app.route('/report-generator/appt-overlap/generate-report', methods=['POST'])
def handle_generate_appt_overlap_report():
    if request.headers['Content-Type'] == 'application/json':
        data = request.get_json()
        start_date = data.get('start_date')
        end_date = data.get('end_date')
        provider = data.get('provider')
        client = data.get('client')
        
        try:
            report_file = generate_appt_overlap_report(start_date, end_date, provider, client)
            return send_file(
                report_file,
                as_attachment=True,
                download_name=f"Overlapping_Appointment_Report_'{start_date}'-'{end_date}'.xlsx"
            )
        except Exception as e:
            return jsonify({'error': str(e)}), 500
    else:
        return jsonify({'error': 'Unsupported Media Type'}), 415