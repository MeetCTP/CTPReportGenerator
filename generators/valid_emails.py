import pandas as pd
from email_validator import validate_email, EmailNotValidError
from pandas import ExcelWriter
from dotenv import load_dotenv
import io
import re
import os
from pyairtable import Api

def generate_valid_email_report(selected_tables):
    try:
        load_dotenv()

        api_key = os.getenv("AT_API_KEY")

        api = Api(api_key)

        counselors_social_table = api.table('applyILT6MqcpyHWU', 'tblcISPJ1KskmFJ3V')
        bcba_lbs_table = api.table('app9O5xkhfInyGoip', 'tbl0YfBacdKvvNqpq')
        wilson_table = api.table('appACRGeTxgqokzXT', 'tblfk2P4ZiZsy2ZAV')
        speech_table = api.table('appFwul5GLBW3XXkA', 'tblkZq8PRZCGPykBi')
        sped_table = api.table('appGj6OWRMqrdcydL', 'tblDWEdcCkYnXGNcb')
        paras_table = api.table('app5obuWU6q9BKfiL', 'tblWylyMP4shyRhpM')
        mobile_table = api.table('app27nPo3s0RmlPyW', 'tbl42Un3FVBkJGXpe')
        archived_para_21_22 = api.table('appJMe2I9C9NMSu9d', 'tblwpjA57QX8h8xj6')
        archived_para_19_21 = api.table('appkZep4g2h0AGfR9', 'tblwDYMALGzp1Gfbl')
        archived_para_22 = api.table('appCsoodShQ4P4JrV', 'tbltCys3NfScMbLyW')
        not_to_use = api.table('appGL58BLgeQts6DX', 'tblxVfcrGegYqz8KY')

        tables = [
            (counselors_social_table, "Counselors and Social Workers"),
            (bcba_lbs_table, "BCBA and LBS"),
            (wilson_table, "Wilson Reading Instructors"),
            (speech_table, "Speech Therapists"),
            (sped_table, "SPED Teachers and Tutors"),
            (paras_table, "Paraprofessional"),
            (mobile_table, "Mobile Therapist"),
            (archived_para_21_22, "Archived Para Apps 2021-2022"),
            (archived_para_19_21, "Archived Para Apps 2019-2021"),
            (archived_para_22, "Archived Para Apps 08.15.2022"),
            (not_to_use, "Simple Tracker (Not to use)")
        ]
        
        valid_emails = []  # List to hold valid emails
        invalid_emails = []  # List to hold invalid emails

        output_file = io.BytesIO()

        # Create a new ExcelWriter object to write data to the output_file
        with ExcelWriter(output_file, engine='openpyxl') as writer:
            for tbl, sheet_name in tables:
                if sheet_name not in selected_tables:
                    continue
                # Get all records from Airtable
                records = tbl.all()

                if not records:
                    continue

                # Convert records to a DataFrame
                try:
                    data = [record['fields'] for record in records]
                except KeyError:
                    continue
                df = pd.DataFrame(data)
                if df.empty:
                    continue

                if 'Re-engagement?' not in df.columns:
                    continue

                df['Re-engagement?'] = df['Re-engagement?'].astype(str)
                df = df[df['Re-engagement?'].str.lower() == 'yes']
                
                # Check if there is an "Email Address" column
                if 'Email Address' in df.columns:
                    # Create a new DataFrame with 'Email Address' and corresponding table name
                    email_df = df[['Email Address']].copy()
                    email_df['Table'] = sheet_name
                    
                    email_df = email_df[email_df['Email Address'].notna()]  # Remove rows with NaN
                    email_df = email_df[email_df['Email Address'] != '']  # Remove empty strings
                    email_df['Email Address'] = email_df['Email Address'].astype(str)

                    # Apply email validation to create the 'IsValidEmail' column
                    email_df['IsValidEmail'] = email_df['Email Address'].apply(lambda email: email_validation(email))

                    # Separate valid and invalid emails
                    valid_df = email_df[email_df['IsValidEmail'] == True].drop(columns=['IsValidEmail'])
                    invalid_df = email_df[email_df['IsValidEmail'] == False].drop(columns=['IsValidEmail'])

                    # Append valid and invalid emails to their respective lists
                    valid_emails.append(valid_df)
                    invalid_emails.append(invalid_df)

            # Concatenate the valid and invalid email dataframes
            if valid_emails:
                valid_df = pd.concat(valid_emails, ignore_index=True)
                valid_df = valid_df.sort_values(by='Email Address', ascending=True)
                valid_df.to_excel(writer, sheet_name="Valid Emails", index=False)

            if invalid_emails:
                invalid_df = pd.concat(invalid_emails, ignore_index=True)
                invalid_df.to_excel(writer, sheet_name="Invalid Emails", index=False)

        # Ensure the file pointer is at the beginning of the file before returning
        output_file.seek(0)

        # Return the output file as an attachment
        return output_file
    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e
    
def validate_and_filter_emails(email):
    # Lowercase the email to avoid domain casing issues
    email = email.lower()

    # Remove any trailing or leading whitespace
    email = email.strip()

    # Remove any spaces in the email
    email = email.replace(" ", "")

    # Attempt to validate the email
    try:
        # Validate the email (this does not fix errors, only validates)
        valid = validate_email(email, check_deliverability=True)
        return valid.email, True
    except EmailNotValidError as e:
        print(f"Invalid email address: {email}, Error: {e}")
        return email, False
    
def email_validation(email):
    email = email.strip()
    email_regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    if email is None:
        return False
    return bool(re.match(email_regex, email))