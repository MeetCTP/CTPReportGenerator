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

    grouped_minutes = filtered_data.groupby(['Provider', 'Client', 'Subcategory']).agg(
        TotalMinutes=('EventMinutes', 'sum'),
    ).reset_index()

    idx = pd.MultiIndex.from_product(
        [
            grouped_minutes['Provider'].unique(),
            grouped_minutes['Client'].unique(),
            ['Direct Time', 'Indirect Time']
        ],
        names=['Provider', 'Client', 'Subcategory']
    )

    grouped_minutes = grouped_minutes.set_index(['Provider', 'Client', 'Subcategory']).reindex(idx, fill_value=0).unstack('Subcategory', fill_value=0)

    grouped_minutes.columns = grouped_minutes.columns.droplevel(0)

    grouped_minutes = grouped_minutes.reset_index()

    service_code_indicators = filtered_data.groupby(['Provider', 'Client']).agg(
        HasCounselor=('ServiceCodeDescription', lambda x: any('Counselor' in str(desc) for desc in x))
    ).reset_index()

    grouped = grouped_minutes.merge(service_code_indicators, on=['Provider', 'Client'], how='left')

    def calculate_allowed_indirect(direct_minutes, service_code_descriptions):
        allowed_indirect = 0

        for index, direct_time in enumerate(service_code_descriptions):
            if "Counselor" in direct_time:
                allowed_indirect += 15
            else:
                for time, indirect_time in sorted(indirect_allotment.items(), key=lambda x: int(x[0]), reverse=True):
                    time = int(time)
                    indirect_time = int(indirect_time)
                    if direct_minutes >= time:
                        multiplier = direct_minutes // time
                        allowed_indirect += multiplier * indirect_time
                        direct_minutes %= time
        return allowed_indirect

    grouped['AllowedIndirectMinutes'] = grouped.apply(
        lambda row: calculate_allowed_indirect(
            row['Direct Time'], 
            data.loc[(data['Provider'] == row['Provider']) & (data['Client'] == row['Client']), 'ServiceCodeDescription']
        ), 
        axis=1
    )

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
        data.drop(columns='HasCounselor', inplace=True)
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