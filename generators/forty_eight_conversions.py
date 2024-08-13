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
        connection_string = f"mssql+pymssql://MeetCTP\Administrator:$Unlock01@CTP-DB/CRDB2"
        engine = create_engine(connection_string)
        today = datetime.now()

        query = f"""
            SELECT DISTINCT
                *
            FROM FortyEightHourReportView
            WHERE (CONVERT(DATE, ServiceDate, 101) BETWEEN '{start_date_str}' AND '{end_date_str}') AND (CompanyRole IN ({', '.join([f"'{s}'" for s in company_roles])}));
        """
        data = pd.read_sql_query(query, engine)

        # Convert relevant columns to datetime
        data['ServiceDate'] = pd.to_datetime(data['ServiceDate'])
        data['ConvertedDT'] = pd.to_datetime(data['ConvertedDT'])
        data['IsLate'] = (data['ConvertedDT'] - data['ServiceDate']).dt.days > 2

        # Filter data
        data = data[(data['Status'] == 'Un-Converted') | (data['IsLate'])]

        # Fetch existing data from WarnedProviders and ProviderNonPayment tables
        warning_list = pd.read_sql_query("SELECT * FROM WarnedProviders", engine)
        non_payment_list = pd.read_sql_query("SELECT * FROM ProviderNonPayment", engine)

        # Check if WarnedProviders is empty
        is_first_run = warning_list.empty

        # Relevant columns for WarnedProviders and ProviderNonPayment
        relevant_columns = ['Provider', 'Client', 'ProviderEmail', 'ServiceDate', 'AppStart', 'AppEnd', 'Status', 'CompanyRole', 'ConvertedDT']

        if is_first_run:
            data[relevant_columns].to_sql('WarnedProviders', engine, if_exists='append', index=False)
        else:
            # Load existing data from WarnedProviders and ProviderNonPayment tables
            existing_warned_df = pd.read_sql('SELECT Provider, AppStart FROM WarnedProviders', engine)
            existing_non_payment_df = pd.read_sql('SELECT Provider, AppStart FROM ProviderNonPayment', engine)

            # Create sets for efficient lookup
            warned_providers_set = set(zip(existing_warned_df['Provider'], existing_warned_df['AppStart']))
            non_payment_providers_set = set(zip(existing_non_payment_df['Provider'], existing_non_payment_df['AppStart']))

            # Function to check if a row is in a set
            def is_in_set(row, check_set):
                return (row['Provider'], row['AppStart']) in check_set

            # Lists to hold new rows to be inserted
            new_warning_list = []
            new_non_payment_list = []

            for _, row in data.iterrows():
                provider = row['Provider']
                app_start = row['AppStart']
                
                if is_in_set(row, warned_providers_set):
                    # Skip rows already in WarnedProviders
                    continue
                elif is_in_set(row, non_payment_providers_set):
                    # Skip rows already in ProviderNonPayment
                    continue
                elif provider in existing_warned_df['Provider'].values:
                    # Add to non-payment list if provider already warned (but row not in any list yet)
                    new_non_payment_list.append(row)
                    non_payment_providers_set.add((provider, app_start))
                else:
                    # Add to warned list if provider is new
                    new_warning_list.append(row)
                    warned_providers_set.add((provider, app_start))

            # Convert lists to DataFrames and add to the database
            if new_warning_list:
                warning_df = pd.DataFrame(new_warning_list, columns=data.columns)
                warning_df[relevant_columns].to_sql('WarnedProviders', engine, if_exists='append', index=False)

            if new_non_payment_list:
                non_payment_df = pd.DataFrame(new_non_payment_list, columns=data.columns)
                non_payment_df[relevant_columns].to_sql('ProviderNonPayment', engine, if_exists='append', index=False)

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

        data.drop(columns='IsLate', inplace=True)

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