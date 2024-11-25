import pandas as pd
import numpy as np
from sqlalchemy import create_engine
from datetime import datetime, timedelta
import pymssql
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os
import io

def calculate_indirect_time(data):
    indirect_allotment = {
        '60': '15',
        '45': '11',
        '30': '7',
        '15': '3'
    }
    
    filtered_data = data[data['Subcategory'].isin(['Direct Time', 'Indirect Time'])]

    grouped = filtered_data.groupby(['Provider', 'Client', 'Subcategory']).agg(
        TotalMinutes=('EventMinutes', 'sum')
    )

    idx = pd.MultiIndex.from_product(
        [
            grouped.index.get_level_values('Provider').unique(),
            grouped.index.get_level_values('Client').unique(),
            ['Direct Time', 'Indirect Time']
        ],
        names=['Provider', 'Client', 'Subcategory']
    )
    grouped = grouped.reindex(idx, fill_value=0).unstack('Subcategory', fill_value=0)

    grouped.columns = grouped.columns.droplevel(0)

    def calculate_allowed_indirect(direct_minutes):
        allowed_indirect = 0
        for direct_time, indirect_time in sorted(indirect_allotment.items(), key=lambda x: int(x[0]), reverse=True):
            direct_time = int(direct_time)
            indirect_time = int(indirect_time)
            if direct_minutes >= direct_time:
                multiplier = direct_minutes // direct_time
                allowed_indirect += multiplier * indirect_time
                direct_minutes %= direct_time
        return allowed_indirect

    grouped['AllowedIndirectMinutes'] = grouped['Direct Time'].apply(calculate_allowed_indirect)

    grouped['ExceedsAllowedIndirect'] = grouped['Indirect Time'] > grouped['AllowedIndirectMinutes']

    grouped.reset_index(inplace=True)

    final_data = pd.merge(data, grouped, on=['Provider', 'Client'], how='left')

    return final_data

def generate_pad_indirect(start_date, end_date):
    try:
        connection_string = f"mssql+pymssql://MeetCTP\Administrator:$Unlock01@CTP-DB/CRDB2"
        engine = create_engine(connection_string)
        
        query = f"""
            SELECT *
            FROM PADIndirectTimeView
            WHERE (CONVERT(DATE, ServiceDate, 101) BETWEEN '{start_date}' AND DATEADD(day, 1, '{end_date}'))
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
        
        data = calculate_indirect_time(data)
        
        data.drop_duplicates(inplace=True)
        data = data.sort_values(by=['Provider', 'Client', 'ServiceDate'], ascending=True)
        data['AppStart'] = data['AppStart'].dt.strftime('%m/%d/%Y %I:%M%p')
        data['AppEnd'] = data['AppEnd'].dt.strftime('%m/%d/%Y %I:%M%p')

        output_file = io.BytesIO()
        data.to_excel(output_file, index=False)
        output_file.seek(0)
        return output_file
        
    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

    finally:
        engine.dispose()