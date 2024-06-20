from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for
from sqlalchemy import create_engine, text
import pandas as pd
from generators.appointment_match_agora import generate_appointment_agora_report
from generators.appointment_match_insight import generate_appointment_insight_report
from generators.active_contacts_report import generate_active_contacts_report
from generators.late_conversions_report import generate_late_conversions_report
from generators.no_show_late_cancel_report import generate_no_show_late_cancel_report
from generators.provider_sessions_report import generate_provider_sessions_report
from generators.provider_connections_report import generate_provider_connections_report
from generators.forty_eight_conversions import generate_unconverted_time_report
from generators.util_tracker import generate_util_tracker
from flask_cors import CORS
from datetime import datetime
import urllib.parse
import logging
import traceback
import tempfile
import os

app = Flask(__name__)
CORS(app)

# Make the WSGI interface available at the top level so wfastcgi can get it.
wsgi_app = app.wsgi_app

connection_string = f"mssql+pymssql://MeetCTP\Joshua.Bliven:$Unlock03@CTP-DB/CTPHOME"
engine = create_engine(connection_string)

#People
lisa = f"Lisa.Kowalski"
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
christie = f"Christina.K.Sampson"
greg = f"Gregory.T.Hughes"
bree = f"Brianna.T.Peterson"
jesse = f"Jesse.Petracz"
saqoya = f"Saqoya.S.Weldon"
capriest = f"Capriest.T.Parker"
olivia = f"Olivia.a.DiPasquale"

#Groups
admin_group = [josh, fabian, lisa, admin]
recruiting_group = [amy, stacey]
clinical_group = [saqoya, jesse]
accounting_group = [eileen, greg, cari]
student_services_group = [eileen, cari, christie, bree, capriest, olivia]
human_resources_group = [aaron, linda]
it_group = [josh, fabian, dan]
testing_group = [josh, fabian]
site_mod_group = [josh, fabian, lisa, admin, aaron, eileen]

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

    # Execute the query with parameters
    try:
        with engine.connect() as connection:
            connection.execute(query, data)
        print(f"Query executed successfully: {query}, with data: {data}")
    except Exception as e:
        print(f"Error executing query: {e}")

# Define a function to fetch data from the database
def fetch_data():
    news_query = text("SELECT Title, Body FROM dbo.News_Posts")
    notifications_query = text("SELECT EventDate, Body FROM dbo.Notifications")
    weekly_qa_query = text("SELECT Body FROM dbo.WeeklyQA")

    with engine.connect() as connection:
        news_articles = connection.execute(news_query).fetchall()
        notifications = connection.execute(notifications_query).fetchall()
        weekly_qas = connection.execute(weekly_qa_query).fetchall()

    return news_articles, notifications, weekly_qas

@app.route('/test-user')
def test_user():
    username = request.environ.get('REMOTE_USER')
    username = str(username).split('\\')[-1]
    if username in admin_group:
        return f"Hello, {username}"
    else:
        return "Hello, anonymous user!"

def set_site_data():
    username = request.environ.get('REMOTE_USER')
    username = str(username).split('\\')[-1]
    if username in site_mod_group:
        if request.method == 'POST':
            try:
                form_type = request.form.get('form_type')
                if form_type == 'news':
                    title = request.form['Title']
                    body = request.form['Body']
                    attachment_file = request.files['filename'] if 'filename' in request.files else None
                    attachment = attachment_file.read() if attachment_file else None
                    data = {
                        'Title': title,
                        'Body': body,
                        'Attachment': attachment,
                        'CreatedBy': username,
                        'RowModifiedAt': datetime.now()
                    }
                    handle_submit_form_data('News_Posts', data)
                
                elif form_type == 'notification':
                    event_date = request.form['Date']
                    body = request.form['Body']
                    data = {
                        'EventDate': event_date,
                        'Body': body,
                        'CreatedBy': username,
                        'RowModifiedAt': datetime.now()
                    }
                    handle_submit_form_data('Notifications', data)
                
                elif form_type == 'qa':
                    body = request.form['Body']
                    data = {
                        'Body': body,
                        'CreatedBy': username,
                        'RowModifiedAt': datetime.now()
                    }
                    handle_submit_form_data('WeeklyQA', data)

                return redirect(url_for('home'))
            except Exception as e:
                print(f"Error: {e}")
                return f"An internal error occurred: {e}", 500
        return render_template('site_mod.html')
    else:
        return redirect(url_for('access_denied'))

@app.route('/')
def home():
    news_articles, notifications, weekly_qas = fetch_data()
    return render_template('home.html', 
                           news_articles=news_articles, 
                           notifications=notifications, 
                           weekly_qas=weekly_qas)

@app.route('/report-generator')
def reports():
    """Renders the reports home page template to the /reports url."""
    username = request.environ.get('REMOTE_USER')
    username = str(username).split('\\')[-1]
    if username in admin_group:
        return render_template('all-reports.html')
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
            _, excel_file = generate_unconverted_time_report(selected_roles, start_date, end_date)
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
    
@app.route('/report-generator/forty-eight-hour-warning/get-mailing-list', methods=['POST'])
def get_mailing_list():
    if request.headers['Content-Type'] == 'application/json':
        data = request.json
        selected_roles = data.get('company_roles', [])
        start_date = data.get('start_date')
        end_date = data.get('end_date')
        
        try:
            mailing_list, _ = generate_unconverted_time_report(selected_roles, start_date, end_date)
            return jsonify(mailing_list)
        except Exception as e:
            print('Exception occurred: ', e)
            return jsonify({'error': str(e)}), 500
    else:
        return jsonify({'error': 'Unsupported Media Type'}), 415