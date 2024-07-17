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
        connection_string = f"mssql+pymssql://MeetCTP\Joshua.Bliven:$Unlock03@CTP-DB/CRDB2"
        engine = create_engine(connection_string)
        today = datetime.now()

        query = f"""
            SELECT
                *
            FROM FortyEightHourReportView
            WHERE (CONVERT(DATE, ServiceDate, 101) BETWEEN '{start_date_str}' AND '{end_date_str}') AND (CompanyRole IN ({', '.join([f"'{s}'" for s in company_roles])}));
        """
        data = pd.read_sql_query(query, engine)

        data['AppStart'] = pd.to_datetime(data['AppStart'], format='%m/%d/%Y %I:%M%p')
        data['AppEnd'] = pd.to_datetime(data['AppEnd'], format='%m/%d/%Y %I:%M%p')

        data['ServiceDate'] = pd.to_datetime(data['ServiceDate'])
        data['ConvertedDT'] = pd.to_datetime(data['ConvertedDT'])
        data['IsLate'] = (data['ConvertedDT'] - data['ServiceDate']).dt.total_seconds() > 48 * 3600

        data = data[(data['Status'] == 'Un-Converted') | (data['IsLate'])]
        
        contractor_data = data[data['CompanyRole'] == 'CompanyRole: Contractor']

        warning_list = pd.read_sql_query("SELECT * FROM WarnedProviders", engine)
        non_payment_list = pd.read_sql_query("SELECT * FROM ProviderNonPayment", engine)
        
        warning_set = set((row['ProviderEmail'], row['AppStart']) for _, row in warning_list.iterrows())
        non_payment_set = set((row['ProviderEmail'], row['AppStart']) for _, row in non_payment_list.iterrows())

        new_warning_list = []
        new_non_payment_list = []

        for _, row in data.iterrows():
            provider_email = row['ProviderEmail']
            app_start = row['AppStart']
            provider_name = row['Provider']
            client = row['Client']
            service_date = row['ServiceDate']
            app_end = row['AppEnd']
            status = row['Status']
            company_role = row['CompanyRole']
            converted_dt = row['ConvertedDT']

            if (provider_email, app_start) not in warning_set:
                new_warning_list.append((provider_name, provider_email, client, service_date, app_start, app_end, status, company_role, converted_dt))
                warning_set.add((provider_email, app_start))
            elif (provider_email, app_start) not in non_payment_set:
                new_non_payment_list.append((provider_name, provider_email, client, service_date, app_start, app_end, status, company_role, converted_dt))
                non_payment_set.add((provider_email, app_start))

        # Insert new warnings and non-payments into the database
        if new_warning_list:
            warning_df = pd.DataFrame(new_warning_list, columns=['Provider', 'Client', 'ProviderEmail', 'ServiceDate', 'AppStart', 'AppEnd', 'Status', 'CompanyRole', 'ConvertedDT'])
            warning_df.to_sql('WarnedProviders', engine, if_exists='append', index=False)
        
        if new_non_payment_list:
            non_payment_df = pd.DataFrame(new_non_payment_list, columns=['Provider', 'Client', 'ProviderEmail', 'ServiceDate', 'AppStart', 'AppEnd', 'Status', 'CompanyRole', 'ConvertedDT'])
            non_payment_df.to_sql('ProviderNonPayment', engine, if_exists='append', index=False)

        # Prepare mailing list
        contractor_data = data[data['CompanyRole'] == 'CompanyRole: Contractor']
        
        mailing_list = defaultdict(lambda: {'name': '', 'email': '', 'appointments': set()})

        for _, row in contractor_data.iterrows():
            provider = row['Provider']
            email = row['ProviderEmail']
            service_date = row['ServiceDate'].strftime('%m-%d-%Y')
        
            mailing_list[email]['name'] = provider
            mailing_list[email]['email'] = email
            mailing_list[email]['appointments'].add(service_date)
        
        for provider_info in mailing_list.values():
            if provider_info['email']:
                provider_info['appointments'] = ', '.join(provider_info['appointments'])

        # Prepare the excel file
        output_file = io.BytesIO()
        data.to_excel(output_file, index=False)
        output_file.seek(0)
        
        return mailing_list, warning_list.to_dict(orient='records'), non_payment_list.to_dict(orient='records'), output_file
    
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
    password = 'password'
    
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



    """
    First warning email:
        {name},

        The following appointments are past the 48 hour deadline to convert: {appointments}
        Please cancel or convert these appointments immediately. This is the only outreach you will receive.
        Next time this occurs, you will not be paid for appointments converted beyond the 48 hour deadline.
        Thank you.
        -THIS IS AN AUTOMATED EMAIL, DO NOT REPLY-

    Second nonpayment email:
        {name},

        The following appointments are past the 48 hour deadline to convert: {appointments}
        Since this is your second warning, you will not receive payment for these appointments,
        however, you are still expected to convert or cancel these appointments as outlined in your contract.
        To avoid further non-payment, please convert sessions within 48 hours of the appointment time.
        Thank you.
        -THIS IS AN AUTOMATED EMAIL, DO NOT REPLY-
    """