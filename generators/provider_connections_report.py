import pandas as pd
from sqlalchemy import create_engine
from datetime import datetime
import pymssql
import io
import os
import openpyxl

def generate_provider_connections_report():
    try:
        os.getlogin()
        connection_string = f"mssql+pymssql://MeetCTP\Joshua.Bliven:$Unlock03@CTP-DB/CRDB"
        today_dt = datetime.now()
        today = datetime.strftime(today_dt, '%Y-%m-%d')
        engine = create_engine(connection_string)

        cecl_query = f"""
            SELECT
                name,
                contactId
            FROM CECL_Join
            WHERE name = 'CompanyRole: Employee' OR name = 'CompanyRole: Contractor';
        """
        cecl_data = pd.read_sql_query(cecl_query, engine)

        prov_sess_query = f"""
            SELECT DISTINCT
                Provider,
                ProviderId,
                Client,
                ClientId,
                School
            FROM ProviderSessions_New
            WHERE CONVERT(date, ServiceDate, 101) > '2024-01-01';
        """
        prov_sess_data = pd.read_sql_query(prov_sess_query, engine)

        contacts_query = f"""
            SELECT
                id,
                County
            FROM Contacts_New_Test;
        """
        contacts_data = pd.read_sql_query(contacts_query, engine)

        report_data = pd.merge(prov_sess_data, cecl_data, left_on="ProviderId", right_on="contactId", how="inner")
        report_data = pd.merge(report_data, contacts_data, left_on="contactId", right_on="id", how="inner")

        columns_to_drop = ['contactId', 'id']
        report_data.drop(columns=columns_to_drop, inplace=True)

        output_file = io.BytesIO()
        report_data.to_excel(output_file, index=False)
        output_file.seek(0)
        return output_file

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

    finally:
        engine.dispose()