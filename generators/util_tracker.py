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

def get_first_valid_index(group):
    valid = group['AuthHours'].fillna(0) > 0
    return group[valid].head(1).index if valid.any() else pd.Index([])
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
            query += f"""WHERE Provider IN ({', '.join([f"'{s}'" for s in employee_providers])})"""
            ins_query += f"""WHERE Provider IN ({', '.join([f"'{s}'" for s in employee_providers])})"""
        elif company_role == 'Contractor':
            query += f"""WHERE Provider NOT IN ({', '.join([f"'{s}'" for s in employee_providers])})"""
            ins_query += f"""WHERE Provider NOT IN ({', '.join([f"'{s}'" for s in employee_providers])})"""

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

            for subcategory, keywords in direct_categories.items():
                if any(keyword in row['ServiceCodeDescription'] for keyword in keywords):
                    return 'Direct', subcategory

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

        if not ins_data.empty:
            ins_data[['Category', 'Subcategory']] = ins_data.apply(
                lambda row: pd.Series(categorize_service(row)), axis=1
            )

            ins_data['SchedulingCancelledReason'] = ins_data['SchedulingCancelledReason'].replace('', np.nan)

        #if data['Subcategory'] == 'MakeUpTime'

        data = data.sort_values(by='LastName', ascending=True)
        data.drop_duplicates(inplace=True)

        ins_data = ins_data.sort_values(by='LastName', ascending=True)
        ins_data.drop_duplicates(inplace=True)

        data = pd.concat([data, ins_data], ignore_index=True)
        data.drop_duplicates(subset=["Client", "AuthType", "Provider", "AppStart", "AppEnd", "Status"], inplace=True)
        valid_indices = data.groupby(['Client', 'ServiceCode']).apply(get_first_valid_index).explode().dropna().astype(int)
        data.loc[~data.index.isin(valid_indices), 'AuthHours'] = 0
        data.loc[data['Subcategory'] == 'Indirect Time', 'AuthType'] = 'monthly'
        data['Provider'] = data['Provider'].astype(str).str.replace(' ', '_')
        data['School'] = data['School'].astype(str).str.replace(' ', '_')

        data = data[['Client',
                     'LastName',
                     'School',
                     'AuthType',
                     'AuthHours',
                     'ServiceCode',
                     'ServiceCodeDescription',
                     'Provider',
                     'AppStart',
                     'AppEnd',
                     'EventHours',
                     'Status',
                     'SchedulingCancelledReason',
                     'PayorName',
                     'PayorPlanName',
                     'Coordinator',
                     'Category',
                     'Subcategory']] 

        data['ProviderInitial'] = data['Provider'].str[0].str.upper()
        aj_data = data[data['ProviderInitial'].between('A', 'J')]
        kz_data = data[data['ProviderInitial'].between('K', 'Z')]

        aj_data = aj_data.drop(columns=['ProviderInitial'])
        kz_data = kz_data.drop(columns=['ProviderInitial'])
        
        output_file = io.BytesIO()
        
        with ExcelWriter(output_file, engine='openpyxl') as writer:
            aj_data.to_excel(writer, sheet_name='A-J', index=False)
            kz_data.to_excel(writer, sheet_name='K-Z', index=False)
        
        output_file.seek(0)
        return output_file

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

    finally:
        engine.dispose()