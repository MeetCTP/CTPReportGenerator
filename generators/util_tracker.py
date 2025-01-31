import pandas as pd
import numpy as np
from sqlalchemy import create_engine
from datetime import datetime
import pymssql
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from pandas import ExcelWriter
import os
import io

def calculate_completion_percentage(data):
    valid_data = data[data['SchedulingCancelledReason'].isna()]

    grouped = valid_data.groupby(['Client', 'ServiceCode', 'AuthType'])

    total_event_hours = grouped['EventHours'].transform('sum')

    data['CompletedPercentage'] = np.nan

    data.loc[valid_data.index, 'CompletedPercentage'] = np.where(
        valid_data['AuthHours'] == 0, 
        np.nan, 
        (total_event_hours / valid_data['AuthHours']) * 100
    )

    return data

def calculate_cancellation_percentage(data):
    grouped = data.groupby(['Client', 'ServiceCode', 'AuthType'])

    cancellation_stats = grouped.agg(
        TotalSessions=('SchedulingCancelledReason', 'size'),
        CancelledSessions=('SchedulingCancelledReason', lambda x: x.notnull().sum())
    ).reset_index()

    cancellation_stats['CancellationPercentage'] = (
        cancellation_stats['CancelledSessions'] / cancellation_stats['TotalSessions']
    ) * 100

    data = pd.merge(
        data,
        cancellation_stats[['Client', 'ServiceCode', 'AuthType', 'CancellationPercentage']],
        on=['Client', 'ServiceCode', 'AuthType'],
        how='left'
    )

    data['CancellationPercentage'] = data['CancellationPercentage'].fillna(0)

    return data

def generate_util_tracker(start_date, end_date, provider, client):
    try:
        user_name = os.getlogin()
        connection_string = f"mssql+pymssql://MeetCTP\Administrator:$Unlock01@CTP-DB/CRDB2"
        engine = create_engine(connection_string)
        
        query = f"""
            SELECT *
            FROM ClinicalUtilizationTracker
        """
        conditions = []
        if provider:
            conditions.append(f"""Provider = '{provider}'""")
        if client:
            conditions.append(f"""Client = '{client}'""")
        if conditions:
            query += " WHERE " + " AND ".join(conditions)

        query += f""" AND (CONVERT(DATE, AppStart, 101) BETWEEN '{start_date}' AND DATEADD(day, 1, '{end_date}'))"""

        data = pd.read_sql_query(query, engine)

        ins_query = f"""
            SELECT *
            FROM InsuranceClinicalUtil
        """
        if conditions:
            ins_query += " WHERE " + " AND ".join(conditions)

        ins_query += f""" AND (CONVERT(DATE, AppStart, 101) BETWEEN '{start_date}' AND DATEADD(day, 1, '{end_date}'))"""

        ins_data = pd.read_sql_query(ins_query, engine)

        indirect_categories = {
            'ProgressReports': ['Progress', 'Report'],
            'MakeUpTime': ['Make Up'],
            'Meeting/OtherIndirect': ['IEP', 'School Personnel']
        }

        general_indirect_kw = ['Indirect']

        def categorize_service(desc):
            if pd.isnull(desc) or desc == '':
                return 'Undefined', 'Undefined'

            for subcategory, keywords in indirect_categories.items():
                if any(keyword in desc for keyword in keywords):
                    return 'Indirect', subcategory

            if any(keyword in desc for keyword in general_indirect_kw):
                return 'Indirect', 'Indirect Time'

            return 'Direct', 'Direct Time'
        
        if not data.empty:
            data[['Category', 'Subcategory']] = data['ServiceCodeDescription'].apply(
                lambda desc: pd.Series(categorize_service(desc))
            )

            data['SchedulingCancelledReason'] = data['SchedulingCancelledReason'].replace('', np.nan)
            data = calculate_completion_percentage(data)
            data = calculate_cancellation_percentage(data)

        if not ins_data.empty:
            ins_data[['Category', 'Subcategory']] = ins_data['ServiceCodeDescription'].apply(
                lambda desc: pd.Series(categorize_service(desc))
            )
        
            ins_data['SchedulingCancelledReason'] = ins_data['SchedulingCancelledReason'].replace('', np.nan)
            ins_data = calculate_completion_percentage(ins_data)
            ins_data = calculate_cancellation_percentage(ins_data)

        #if data['Subcategory'] == 'MakeUpTime'

        data = data.sort_values(by='LastName', ascending=True)
        data.drop_duplicates(inplace=True)

        ins_data = ins_data.sort_values(by='LastName', ascending=True)
        ins_data.drop_duplicates(inplace=True)

        data = pd.concat([data, ins_data], ignore_index=True)
        data.drop_duplicates(inplace=True)

        output_file = io.BytesIO()
        
        with ExcelWriter(output_file, engine='openpyxl') as writer:
            data.to_excel(writer, sheet_name='Schools', index=False)
            #ins_data.to_excel(writer, sheet_name='Insurance', index=False)
        
        output_file.seek(0)
        return output_file

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

    finally:
        engine.dispose()