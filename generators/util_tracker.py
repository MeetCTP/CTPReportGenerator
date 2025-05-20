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

def generate_util_tracker(start_date, end_date, company_role):
    try:
        user_name = os.getlogin()
        db_user = os.getenv("DB_USER")
        db_pw = os.getenv("DB_PW")
        connection_string = f"mssql+pymssql://{db_user}:{db_pw}@CTP-DB/CRDB2"
        engine = create_engine(connection_string)

        employee_providers = [
            'Cathleen DiMaria',
            'Christine Veneziale',
            'Jacqui Maxwell',
            'Jessica Trudeau',
            'Kaitlin Konopka',
            'Kristie Girten',
            'Nicole Morrison', 
            'Roseanna Vellner',
            'Terri Ahern'
        ]
        
        query = f"""
            SELECT *
            FROM ClinicalUtilizationTracker
        """

        ins_query = f"""
            SELECT *
            FROM InsuranceClinicalUtil
        """

        if company_role == 'Employee':
            query += f"""WHERE Provider IN ({', '.join([f"'{p}'" for p in employee_providers])})"""
            ins_query += f"""WHERE Provider IN ({', '.join([f"'{p}'" for p in employee_providers])})"""
        else:
            query += f"""WHERE Provider = 'Nicole Morrison'"""
            ins_query += f"""WHERE Provider = 'Nicole Morrison'"""

        query += f""" AND (CONVERT(DATE, AppStart, 101) BETWEEN '{start_date}' AND DATEADD(day, 1, '{end_date}'))"""

        data = pd.read_sql_query(query, engine)

        ins_query += f""" AND (CONVERT(DATE, AppStart, 101) BETWEEN '{start_date}' AND DATEADD(day, 1, '{end_date}'))"""

        ins_data = pd.read_sql_query(ins_query, engine)

        indirect_categories = {
            'Progress Reports': ['Progress', 'Report'],
            'Meeting/OtherIndirect': ['IEP', 'School Personnel']
        }

        direct_categories = {
            'MakeUpTime': ['Make Up']
        }

        general_indirect_kw = ['Indirect']

        def categorize_service(row):
            if row['Status'] == 'Cancelled':
                return 'Cancelled', row['SchedulingCancelledReason'] if pd.notnull(row['SchedulingCancelledReason']) else 'No Reason Provided'

            if pd.isnull(row['ServiceCodeDescription']) or row['ServiceCodeDescription'] == '':
                return 'Undefined', 'Undefined'

            # Check Indirect categories
            for subcategory, keywords in indirect_categories.items():
                if any(keyword in row['ServiceCodeDescription'] for keyword in keywords):
                    row['AuthType'] = 'monthly'
                    return 'Indirect', subcategory

            # Check general indirect keywords
            if any(keyword in row['ServiceCodeDescription'] for keyword in general_indirect_kw):
                return 'Indirect', 'Indirect Time'

            # Check Direct categories (MakeUpTime now belongs here)
            for subcategory, keywords in direct_categories.items():
                if any(keyword in row['ServiceCodeDescription'] for keyword in keywords):
                    return 'Direct', subcategory

            return 'Direct', 'Direct Time'
        
        if not data.empty:
            data[['Category', 'Subcategory']] = data.apply(
                lambda row: pd.Series(categorize_service(row)), axis=1
            )

            data['SchedulingCancelledReason'] = data['SchedulingCancelledReason'].replace('', np.nan)
            data = calculate_completion_percentage(data)
            data = calculate_cancellation_percentage(data)

        if not ins_data.empty:
            ins_data[['Category', 'Subcategory']] = ins_data.apply(
                lambda row: pd.Series(categorize_service(row)), axis=1
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
        data.drop_duplicates(subset=["Client", "AuthType", "Provider", "AppStart", "AppEnd", "Status"], inplace=True)

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