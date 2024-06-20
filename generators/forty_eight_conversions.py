import pandas as pd
from sqlalchemy import create_engine
import pymssql
import openpyxl
from datetime import datetime, timedelta
from collections import defaultdict
import sys
import os
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def generate_unconverted_time_report(company_roles, start_date_str, end_date_str):
    try:
        user_name = os.getlogin()
        documents_path = f"C:/Users/{user_name}/Documents/"
        connection_string = f"mssql+pymssql://MeetCTP\Joshua.Bliven:$Unlock03@CTP-DB/CRDB"
        engine = create_engine(connection_string)
        today = datetime.now()

        query = f"""
            SELECT
                Provider,
                Client,
                ProviderEmail,
                ServiceDate,
                AppStart,
                AppEnd,
                Status
            FROM Sessions_New
            WHERE CONVERT(DATE, ServiceDate, 101) BETWEEN '{start_date_str}' AND '{end_date_str}';
        """
        data = pd.read_sql_query(query, engine)

        data['AppStart'] = pd.to_datetime(data['AppStart'], format='%m/%d/%Y %I:%M%p')
        data['AppEnd'] = pd.to_datetime(data['AppEnd'], format='%m/%d/%Y %I:%M%p')
        
        provider_query = f"""
            SELECT
                ProviderFullName,
                ProviderId,
                ProviderEmployeeType
            FROM Provider_Raw_New;
        """
        provider_data = pd.read_sql_query(provider_query, engine)
        
        company_role_query = f"""
            SELECT name, contactId
            FROM CECL_Join
            WHERE name IN ({', '.join([f"'{s}'" for s in company_roles])});
        """
        company_role_data = pd.read_sql_query(company_role_query, engine)

        late_conversion_query = f"""
            SELECT
                ConvertedDT,
                AppStart,
                AppEnd,
                ProviderName,
                ClientName
            FROM Conversions_New
            WHERE CONVERT(DATE, AppStart, 101) BETWEEN '{start_date_str}' AND '{end_date_str}';
        """
        late_conversion_data = pd.read_sql_query(late_conversion_query, engine)
        
        data = pd.merge(data, provider_data, left_on='Provider', right_on='ProviderFullName', how='inner')
        data = pd.merge(data, company_role_data, left_on='ProviderId', right_on='contactId', how='inner')
        data = pd.merge(data, late_conversion_data, left_on=['Provider', 'Client', 'AppStart', 'AppEnd'], right_on=['ProviderName', 'ClientName', 'AppStart', 'AppEnd'], how='left')
        data = data.drop(columns=['contactId', 'ProviderFullName', 'ProviderId'])

        data['ServiceDate'] = pd.to_datetime(data['ServiceDate'])
        data['ConvertedDT'] = pd.to_datetime(data['ConvertedDT'])
        data['IsLate'] = (data['ConvertedDT'] - data['ServiceDate']).dt.total_seconds() > 48 * 3600

        data = data[(data['Status'] == 'Un-Converted') | (data['IsLate'])]
        
        contractor_data = data[data['name'] == 'CompanyRole: Contractor']
        non_employee_contractor_data = contractor_data[contractor_data['ProviderEmployeeType'] == 'nonEmployee']
        
        mailing_list = defaultdict(lambda: {'name': '', 'email': '', 'appointments': set()})

        for _, row in non_employee_contractor_data.iterrows():
            # Extract provider, client, email, and service date from the row
            provider = row['Provider']
            email = row['ProviderEmail']
            service_date = pd.to_datetime(row['ServiceDate']).strftime('%m-%d-%Y')
        
            mailing_list[email]['name'] = provider
            mailing_list[email]['email'] = email
            mailing_list[email]['appointments'].add(service_date)
        
        for provider_info in mailing_list.values():
            if provider_info['email']:
                provider_info['appointments'] = ', '.join(provider_info['appointments'])

        output_file = io.BytesIO()
        data.to_excel(output_file, index=False)
        output_file.seek(0)
        
        return mailing_list, output_file
    
    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

    finally:
        engine.dispose()
    

    #Email function

def reminder_email(selected_providers):
    smtp_server = 'smtp.office365.com'
    port = 587
    sender_email = 'cari.tomczyk@meetctp.com'
    password = 'Jlennon7**'      #Get password from Cari
    
    server = smtplib.SMTP(smtp_server, port)
    server.starttls()
    server.login(sender_email, password)
    
    for email, obj in selected_providers.items():
        # Extract details from the object
        email_address = email
        name = obj['name']
        appointments = ', '.join(obj['appointments'])
    
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = email_address
        msg['Subject'] = 'Reminder: Unconverted Appointments'

        body = f"Hi {name},\n\nCan you please go into CR ASAP and convert your appointments on the following dates - they are all UNconverted and need to be converted as they are late.\n\nPlease follow the 48 hour rule moving forward or they will be flagged by billing and will result in pay delays.\n\n{appointments}\n\nBest regards,\nService Coordinator"
        msg.attach(MIMEText(body, 'plain'))
    
        server.send_message(msg)
    
    server.quit()