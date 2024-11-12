import pandas as pd
import numpy as np
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
            WHERE (Provider = '{provider}') AND (CONVERT(DATE, AppStart, 101) BETWEEN '{start_date}' AND DATEADD(day, 1, '{end_date}'))
        """
        data = pd.read_sql_query(query, engine)

        indirect_categories = {
            'ProgressReports': ['Progress', 'Report'],
            'MakeUpTime': ['Make Up'],
            'Meeting/OtherIndirect': ['IEP', 'School Personnel']
        }

        general_indirect_kw = ['Indirect']

        def categorize_service(desc):
            for subcategory, keywords in indirect_categories.items():
                if any(keyword in desc for keyword in keywords):
                    return 'Indirect', subcategory

            if any(keyword in desc for keyword in general_indirect_kw):
                return 'Indirect', 'Indirect Time'

            return 'Direct', 'Direct Time'

        data[['Category', 'Subcategory']] = data['ServiceCodeDescription'].apply(
            lambda desc: pd.Series(categorize_service(desc))
        )

        data['CompletedPercentage'] = np.where(
            data['AuthHours'] == 0, np.nan, (data['EventHours'] / data['AuthHours']) * 100
        )

        #if data['Subcategory'] == 'MakeUpTime'

        data = data.sort_values(by='LastName', ascending=True)
        data.drop_duplicates(inplace=True)

        output_file = io.BytesIO()
        data.to_excel(output_file, index=False)
        output_file.seek(0)
        return output_file

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

    finally:
        engine.dispose()