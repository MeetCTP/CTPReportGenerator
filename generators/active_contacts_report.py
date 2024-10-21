import pandas as pd
from sqlalchemy import create_engine
from datetime import datetime
import pymssql
import io
import os
import re
import openpyxl

def generate_active_contacts_report(status, pg_type, service_types):
    try:
        user_name = os.getlogin()
        connection_string = f"mssql+pymssql://MeetCTP\Administrator:$Unlock01@CTP-DB/CRDB2"
        today_dt = datetime.now()
        today = datetime.strftime(today_dt, '%Y-%m-%d')
        engine = create_engine(connection_string)

        active_contacts_query = f"""
            SELECT
                *
            FROM ActiveContacts
        """
        conditions = []
        if status:
            conditions.append(f"""Status IN ({', '.join([f"'{s}'" for s in status])})""")
        if service_types:
            conditions.append(f"""ServiceType IN ({', '.join([f"'{st}'" for st in service_types])})""")
        if pg_type:
            conditions.append(f"""Type IN ({', '.join([f"'{t}'" for t in pg_type])})""")
        if conditions:
            active_contacts_query += " WHERE " + " AND ".join(conditions)

        report_data = pd.read_sql_query(active_contacts_query, engine)
        report_data.drop_duplicates(subset=['ContactId'], inplace=True)

        if 'Email' in report_data.columns:
            report_data['IsValidEmail'] = report_data['Email'].apply(lambda email: email_validation(email))
        
        report_data = report_data[report_data['IsValidEmail'] == True]
        report_data.drop(columns=['IsValidEmail'], inplace=True)

        output_file = io.BytesIO()
        report_data.to_excel(output_file, index=False)
        output_file.seek(0)
        return output_file

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

    finally:
        engine.dispose()

def email_validation(email):
    email_regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    if email is None:
        return False
    return bool(re.match(email_regex, email))