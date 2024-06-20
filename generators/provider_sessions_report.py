import pandas as pd
from sqlalchemy import create_engine
from datetime import datetime
import pymssql
import openpyxl
import io
import os

def generate_provider_sessions_report(range_start, range_end, supervisor, status):
    try:
        user_name = os.getlogin()
        documents_path = f"C:/Users/{user_name}/Documents/"
        connection_string = f"mssql+pymssql://MeetCTP\Joshua.Bliven:$Unlock03@CTP-DB/CRDB"
        engine = create_engine(connection_string)
        range_start_101 = datetime.strftime(range_start, '%Y-%m-%d')
        range_end_101 = datetime.strftime(range_end, '%Y-%m-%d')
        
        provider_sessions_query = f"""
            SELECT
                Provider,
                ProviderEmail,
                Client,
                School,
                SessionName,
                BillingCode,
                BillingDesc,
                ServiceDate,
                AppStart,
                AppEnd,
                AppMinutes,
                AppHours,
                Mileage,
                Supervisor,
                ProviderSigned,
                Status,
                CancellationReason
            FROM ProviderSessions_New
            WHERE CONVERT(DATE, ServiceDate, 101) BETWEEN '{range_start_101}' AND '{range_end_101}'
        """
        if supervisor:
            provider_sessions_query += f""" AND Supervisor IN ({', '.join([f"'{su}'" for su in supervisor])})"""
        if status:
            provider_sessions_query += f""" AND Status IN ({', '.join([f"'{st}'" for st in status])})"""
        provider_sessions_data = pd.read_sql_query(provider_sessions_query, engine)
    
        output_file = io.BytesIO()
        provider_sessions_data.to_excel(output_file, index=False)
        output_file.seek(0)  # Reset the file pointer to the beginning of the BytesIO object
        return output_file

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

    finally:
        engine.dispose()