import pandas as pd
from sqlalchemy import create_engine
from datetime import datetime
import pymssql
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os
import io

def generate_util_tracker(start_date, end_date, provider):
    try:
        user_name = os.getlogin()
        documents_path = f"C:/Users/{user_name}/Documents/"
        connection_string = f"mssql+pymssql://MeetCTP\Administrator:$Unlock01@CTP-DB/CRDB2"
        engine = create_engine(connection_string)
        
        query = f"""
            SELECT *
            FROM ClinicalUtilizationTracker
            WHERE (Provider = '{provider}') AND (CONVERT(DATE, AppStart, 101) BETWEEN '{start_date}' AND '{end_date}')
        """
        raw_data = pd.read_sql_query(query, engine)

        raw_data.drop_duplicates(inplace=True)

        output_file = io.BytesIO()
        raw_data.to_excel(output_file, index=False)
        output_file.seek(0)
        return output_file

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

    finally:
        engine.dispose()