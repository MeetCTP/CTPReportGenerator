import pandas as pd
from sqlalchemy import create_engine
from datetime import datetime
import pymssql
import io
import os
import openpyxl

def generate_active_contacts_report(status, pg_type, service_types):
    try:
        user_name = os.getlogin()
        connection_string = f"mssql+pymssql://MeetCTP\Joshua.Bliven:$Unlock03@CTP-DB/CRDB"
        today_dt = datetime.now()
        today = datetime.strftime(today_dt, '%Y-%m-%d')
        engine = create_engine(connection_string)

        provider_query = f"""
            SELECT
                id,
                fullname,
                Status,
                email,
                ServiceType,
                gender,
                phonehome,
                phonecell,
                phonework,
                phoneworkext,
                phonefax,
                HomeAddress1,
                HomeAddress2,
                HomeCity,
                County,
                HomeState,
                HomeZip,
                Region
            FROM Contacts_New_Test
        """
        if status and service_types:
            provider_query += f""" WHERE Status IN ({', '.join([f"'{s}'" for s in status])}) AND ServiceType IN ({', '.join([f"'{st}'" for st in service_types])})"""
        if not service_types and status:
            provider_query += f""" WHERE Status IN ({', '.join([f"'{s}'" for s in status])})"""
        if not status and service_types:
            provider_query += f""" WHERE ServiceType IN ({', '.join([f"'{st}'" for st in service_types])})"""
        provider_data = pd.read_sql_query(provider_query, engine)
        
        provider_raw_query = f"""
            SELECT
                ProviderFullName,
                ProviderEmployeeType
            FROM Provider_Raw_New
        """
        if pg_type:
            provider_raw_query += f""" WHERE ProviderEmployeeType IN ({', '.join([f"'{t}'" for t in pg_type])})"""
        provider_raw_data = pd.read_sql_query(provider_raw_query, engine)
        
        report_data = pd.merge(provider_data, provider_raw_data, left_on='fullname', right_on='ProviderFullName', how='inner')
        report_data = report_data.drop(columns=['ProviderFullName'])

        output_file = io.BytesIO()
        report_data.to_excel(output_file, index=False)
        output_file.seek(0)  # Reset the file pointer to the beginning of the BytesIO object
        return output_file

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

    finally:
        engine.dispose()