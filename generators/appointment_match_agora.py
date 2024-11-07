import pandas as pd
from sqlalchemy import create_engine
import pymssql
import openpyxl
from datetime import datetime
import os
import io

def generate_appointment_agora_report(range_start, range_end):
    try:
        user_name = os.getlogin()
        connection_string = f"mssql+pymssql://MeetCTP\Administrator:$Unlock01@CTP-DB/CRDB2"
        engine = create_engine(connection_string)
        range_start_dt = pd.to_datetime(range_start)
        range_end_dt = pd.to_datetime(range_end)
        range_start_101 = datetime.strptime(range_start, '%Y-%m-%d')
        range_end_101 = datetime.strptime(range_end, '%Y-%m-%d')

        appointment_match_query = f"""
            SELECT DISTINCT *
            FROM EasyTracSessionComparison
            WHERE CONVERT(DATE, Date, 101) BETWEEN '{range_start_101}' AND '{range_end_101}'
        """
        appointment_match_data = pd.read_sql_query(appointment_match_query, engine)

        appointment_match_data.drop_duplicates(inplace=True)
        appointment_match_data.drop('School', axis=1, inplace=True)
        
        output_file = io.BytesIO()
        appointment_match_data.to_excel(output_file, index=False)
        output_file.seek(0)
        return output_file

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

    finally:
        engine.dispose()