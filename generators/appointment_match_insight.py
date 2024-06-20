import pandas as pd
from sqlalchemy import create_engine
import pymssql
import openpyxl
from datetime import datetime
import io
import os

def generate_appointment_insight_report(range_start, range_end):
    try:
        user_name = os.getlogin()
        connection_string = f"mssql+pymssql://MeetCTP\Joshua.Bliven:$Unlock03@CTP-DB/CRDB"
        engine = create_engine(connection_string)
        range_start_dt = pd.to_datetime(range_start)
        range_end_dt = pd.to_datetime(range_end)
        range_start_101 = datetime.strptime(range_start, '%Y-%m-%d')
        range_end_101 = datetime.strptime(range_end, '%Y-%m-%d')
        
        appointment_match_query = f"""
            SELECT *
            FROM Appointment_Match_Insight
            WHERE CONVERT(DATE, Date, 101) BETWEEN '{range_start_101}' AND '{range_end_101}'
        """
        appointment_match_data = pd.read_sql_query(appointment_match_query, engine)
    

        output_file = io.BytesIO()
        appointment_match_data.to_excel(output_file, index=False)
        output_file.seek(0)  # Reset the file pointer to the beginning of the BytesIO object
        return output_file

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

    finally:
        engine.dispose()