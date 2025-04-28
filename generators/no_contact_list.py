import pandas as pd
from sqlalchemy import create_engine
from datetime import datetime
from pyairtable import Api
import numpy as np
import requests
import time
import io
import os
import re

API_URL = "https://api.airtable.com/v0/app4EYPWzbGdr6Lxz/tblpdl8tBk8jHOElr"
API_KEY = "patpaS7kXYs546WpG.0c6e11f5836a4c6610260c377c861980a3d0373e0796246ef26a7a59b95c02fa"
HEADERS = {
    "Authorization": f"Bearer {API_KEY}",
    "Content-Type": "application/json"
}

def get_inactive_employee_list():
    try:
        connection_string = f"mssql+pymssql://MeetCTP\Administrator:$Unlock01@CTP-DB:1433/CRDB2"
        engine = create_engine(connection_string)

        query = f"""
            SELECT
                *
            FROM NoContactView
            WHERE Status = 'Inactive' AND Type = 'employee'
        """
        data = pd.read_sql_query(query, engine)

        data.rename(columns={'Email': 'Email Address'}, inplace=True)

        data.drop_duplicates(inplace=True)

        return data

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

def get_no_contact_list():
    try:

        api = Api(API_KEY)

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

        no_contact_list = pd.DataFrame()

        for table, sheet_name in tables:
            # Get all records from Airtable
            records = table.all()

            # Convert records to a DataFrame
            data = [record['fields'] for record in records]
            df = pd.DataFrame(data)

            df['Status'] = df['Status'].astype(str)
            #df = df[~df['Status'].str.lower().isin(['no hire', 'no contact', 'ncns- no hire'])]
            # Alternate way:
            df = df[(df['Status'].str.lower() == 'no contact') | (df['Status'].str.lower() == 'no hire') | (df['Status'].str.lower() == 'ncns- no hire')]

            no_contact_list = pd.concat([no_contact_list, df], ignore_index=True)
        
        return no_contact_list

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e
    
def merge_and_push_NC():
    try:
        # Step 1: Retrieve the two dataframes (using your existing function)
        at_table = get_no_contact_list()  # Assuming this returns a DataFrame
        inactive_employee_df = get_inactive_employee_list()  # Assuming this returns a DataFrame
        
        # Step 2: Merge the two dataframes on 'Email Address' (and possibly other columns)
        merged_df = pd.merge(at_table, inactive_employee_df, on='Email Address', how='outer', suffixes=('_AT', '_CR'))
        
        # Step 3: Remove duplicates based on 'Email Address' (or other relevant columns)
        merged_df = merged_df.drop_duplicates(subset=['Email Address'], keep='first')
        
        df_cleaned = merged_df.copy()
        print(df_cleaned.columns)
        
        def get_full_name(row):
            if pd.notnull(row.get('FirstName')) and pd.notnull(row.get('LastName')):
                return f"{row['FirstName']} {row['LastName']}"
            return row.get('Name')

        df_cleaned['Full Name'] = df_cleaned.apply(get_full_name, axis=1)
        
        def get_combined_position(row):
            pos = row.get('Position')
            job_title = row.get('JobTitle')

            # Convert list to string if necessary
            if isinstance(pos, list):
                pos = ', '.join(pos)

            if pd.notnull(pos) and str(pos).strip():
                return str(pos).strip()
            elif pd.notnull(job_title):
                return str(job_title).strip()
            else:
                return ''
            
        df_cleaned['CombinedPosition'] = df_cleaned.apply(get_combined_position, axis=1)
        df_cleaned['Notes_CR'] = df_cleaned['Notes_CR'].astype(str)

        # Build new DataFrame
        final_columns = {
            'Name': df_cleaned['Full Name'],
            'Position': df_cleaned['CombinedPosition'],
            'Status': df_cleaned['Status_AT'],
            'Interviewer': df_cleaned['Interviewer'],
            'Interview Notes': df_cleaned['Interview Notes'],
            'Notes': df_cleaned['Notes_AT'],
            'CR_Notes': df_cleaned['Notes_CR']
            
        }

        # Create the final DataFrame
        df_final = pd.DataFrame(final_columns)
        
        for col in df_final.columns:
            df_final[col] = df_final[col].apply(lambda x: str(x) if isinstance(x, list) else x)

        df_final.drop_duplicates(inplace=True)
        df_final['Name'] = df_final['Name'].astype(str)
        df_final['Name'] = df_final['Name'].apply(lambda x: x.strip().title() if isinstance(x, str) else x)
        df_final = df_final.sort_values(by='Name', ascending=True)

        df_final = df_final.replace([np.inf, -np.inf], np.nan)
        df_final = df_final.fillna("None")
        
        """output_file = io.BytesIO()
        df_final.to_excel(output_file, index=False)
        output_file.seek(0)
        return output_file"""
        
        # Step 4: Prepare the merged dataframe to be pushed to Airtable
        records_to_push = df_final.to_dict(orient='records')
        
        # Step 5: Push the records to Airtable (batch upload)
        batch_size = 10  # Airtable API limit: 10 records per batch
        for i in range(0, len(records_to_push), batch_size):
            batch = records_to_push[i:i+batch_size]
            payload = {
                "records": [{"fields": record} for record in batch]
            }
            response = requests.post(API_URL, json=payload, headers=HEADERS)
            
            # Check if the request was successful
            if response.status_code == 200:
                print(f"Batch {i//batch_size + 1} uploaded successfully.")
            else:
                print(f"Error uploading batch {i//batch_size + 1}: {response.text}")

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e


merge_and_push_NC()