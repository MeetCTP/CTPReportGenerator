from flask import Flask, render_template, request, jsonify, send_file
from generators.appointment_match_agora import generate_appointment_agora_report
from flask_cors import CORS
import logging
import os

app = Flask(__name__)

CORS(app)

logging.basicConfig(filename='app.log', level=logging.DEBUG)

# Make the WSGI interface available at the top level so wfastcgi can get it.
wsgi_app = app.wsgi_app


@app.route('/')
def home():
    """Renders a sample page."""
    return render_template('home.html')

@app.route('/report-generator')
def reports():
    """Renders the reports home page template to the /reports url."""
    return render_template('reports.html')

@app.route('/report-generator/agora-match')
def agora_match_report():
    return render_template('agora-report.html')

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
                download_name=f"Agora_Appointment_Match_Report_{start_date}-{end_date}_Success.xlsx"
            )
        except Exception as e:
            print('Exception occurred: ', e)
            return jsonify({'error': str(e)}), 500
    else:
        return jsonify({'error': 'Unsupported Media Type'}), 415
