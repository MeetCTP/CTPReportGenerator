import pandas as pd
from sqlalchemy import create_engine
import requests
import json
import msal
import pymssql
import openpyxl
from datetime import datetime, timedelta
from collections import defaultdict
import sys
import os
import io
from flask_mail import Message
from flask import url_for

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
            WHERE (CONVERT(DATE, ServiceDate, 101) BETWEEN '{start_date_str}' AND '{end_date_str}')
        """
        if company_roles:
            query += f""" AND (CompanyRole IN ({', '.join([f"'{s}'" for s in company_roles])}))"""
        data = pd.read_sql_query(query, engine)

        data['ServiceDate'] = pd.to_datetime(data['ServiceDate'], errors='coerce')
        data['ConvertedDT'] = pd.to_datetime(data['ConvertedDT'], errors='coerce')
        data['IsLate'] = data.apply(
            lambda row: True if pd.isna(row['ConvertedDT']) else (row['ConvertedDT'] - row['ServiceDate']).days > 2,
            axis=1
        )

        data = data[(data['Status'] == 'Un-Converted') | (data['IsLate'])]

        warning_list = pd.read_sql_query("SELECT * FROM WarnedProviders", engine)
        non_payment_list = pd.read_sql_query("SELECT * FROM ProviderNonPayment", engine)

        is_first_run = warning_list.empty

        relevant_columns = ['Provider', 'Client', 'ProviderEmail', 'ServiceDate', 'AppStart', 'AppEnd', 'Status', 'CompanyRole', 'ConvertedDT']

        if is_first_run:
            data[relevant_columns].to_sql('WarnedProviders', engine, if_exists='append', index=False)
        else:
            existing_warned_df = pd.read_sql('SELECT Provider, AppStart FROM WarnedProviders', engine)
            existing_non_payment_df = pd.read_sql('SELECT Provider, AppStart FROM ProviderNonPayment', engine)

            warned_providers_set = set(zip(existing_warned_df['Provider'], existing_warned_df['AppStart']))
            non_payment_providers_set = set(zip(existing_non_payment_df['Provider'], existing_non_payment_df['AppStart']))

            def is_in_set(row, check_set):
                return (row['Provider'], row['AppStart']) in check_set

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
            service_date = row['ServiceDate'].strftime('%m-%d-%Y') if pd.notna(row['ServiceDate']) else 'N/A'
            
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
        #email_test()
        
        return mailing_list, warning_list.astype(str).to_dict(orient='records'), non_payment_list.astype(str).to_dict(orient='records'), output_file
    
    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

    finally:
        engine.dispose()
    

#Email function

def reminder_email(selected_providers, warning_list, non_payment_list):
    # OAuth2 setup
    client_id = '4900abdd-dfe5-4297-8a0a-fd9ac2cee73a'
    tenant_id = '91d63f97-21f4-43a6-945e-afd92bb2695f'
    client_secret = os.getenv('AZURE_CLIENT_SECRET')

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    scope = ["https://outlook.office365.com/.default"]

    oauth_app = msal.ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret
    )

    def get_oauth2_token():
        result = oauth_app.acquire_token_for_client(scopes=scope)
        if "access_token" in result:
            return result['access_token']
        else:
            raise Exception("Could not obtain access token")

    def send_mail_with_graph(subject, recipient, body):
        late_conversions = 'lateconversions@meetctp.com'
        me = 'joshua.bliven'
        access_token = get_oauth2_token()

        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }

        email_data = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "Text",
                    "content": body
                },
                "from": {
                    "emailAddress": {
                        "address": late_conversions
                    }
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": recipient
                        }
                    }
                ]
            },
            "saveToSentItems": "true"
        }

        user_endpoint = f'https://graph.microsoft.com/v1.0/users/{me}/sendMail'
        response = requests.post(
            user_endpoint,
            headers=headers,
            data=json.dumps(email_data)
        )
        if response.status_code != 202:
            raise Exception(f"Error sending email: {response.status_code} - {response.text}")
    
    for provider in selected_providers:
        email_address = provider['Email']
        name = provider['Name']
        subject = 'Unconverted Appointments'

        appointments = []
        if any(p['Provider'] == name for p in non_payment_list):
            appointments = [p['AppStart'] for p in non_payment_list if p['Provider'] == name]
        elif any(p['Provider'] == name for p in warning_list):
            appointments = [p['AppStart'] for p in warning_list if p['Provider'] == name]

        provider['Appointments'] = appointments
    
        if any(p['Provider'] == name for p in non_payment_list):
            # Send non-payment email
            body = f"""
                -TEST EMAIL-
                -DO NOT REPLY-
                {name}
                Appointments: {appointments}
            """
            send_mail_with_graph(subject, email_address, body)
        elif any(p['Provider'] == name for p in warning_list):
            # Send warned email
            body = f"""
                -TEST EMAIL-
                -DO NOT REPLY-
                {name}
                Appointments: {appointments}
            """
            send_mail_with_graph(subject, email_address, body)
        else:
            # Skip providers not in either list
            continue



    """
    First warning email:
        Hello {name},

            The following appointments are past the 48 hour deadline to convert: {appointments}
            Please cancel or convert these appointments immediately. This is the only outreach 
            you will receive.
            Next time this occurs, you will not be paid for appointments converted beyond the 48 
            hour deadline.
            Thank you.
            -THIS IS AN AUTOMATED EMAIL, DO NOT REPLY-

    Second nonpayment email:
        {name},

            The following appointments are past the 48 hour deadline to convert: {appointments}
            Since this is your second warning, you will not receive payment for these 
            appointments, however, you are still expected to convert or cancel these appointments 
            as outlined in your contract.
            To avoid further non-payment, please convert sessions within 48 hours of the 
            appointment time.
            Thank you.
            -THIS IS AN AUTOMATED EMAIL, DO NOT REPLY-
    """