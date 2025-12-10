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

def generate_monthly_nums(start_date, end_date, company_role):
    try:
        user_name = os.getlogin()
        db_user = os.getenv("DB_USER")
        db_pw = os.getenv("DB_PW")
        connection_string = f"mssql+pymssql://{db_user}:{db_pw}@CTP-DB/CRDB2"
        engine = create_engine(connection_string)

        start_dt = pd.to_datetime(start_date)
        end_dt = pd.to_datetime(end_date)

        employee_providers = [
            'Cathleen DiMaria',
            'Christine Veneziale',
            'Jacqui Maxwell',
            'Jessica Trudeau',
            'Kaitlin Konopka',
            'Kait Konopka',
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

        query += f""" AND (CONVERT(DATE, AppStart, 101) BETWEEN '{start_date}' AND '{end_date}')"""

        data = pd.read_sql_query(query, engine)

        ins_query += f""" AND (CONVERT(DATE, AppStart, 101) BETWEEN '{start_date}' AND '{end_date}')"""

        ins_data = pd.read_sql_query(ins_query, engine)

        indirect_categories = {
            'Progress Reports': ['Progress', 'Report'],
            'Meeting/OtherIndirect': ['IEP', 'School Personnel']
        }

        direct_categories = {
            'MakeUpTime': ['Make Up', 'Makeup', 'Make up', 'Make']
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
                     'Subcategory',
                     'ChangeNote']] 
        
        data = data[data["ServiceCode"] != "GPAT"]

        # week numbers relative to the month
        unique_weeks = data.assign(Week=data['AppStart'].dt.isocalendar().week).Week.nunique()

        weeks_in_month = unique_weeks if unique_weeks > 0 else 4

        g = data.groupby("Provider")

        monthly_nums = pd.DataFrame({
            "Provider": g.size().index,

            # Total hours
            "TotalHours": g.apply(lambda x: x.loc[x["Status"] != "Cancelled", "EventHours"].sum()),

            # Average hours per week
            "AverageHoursPerWeek": g.apply(
                lambda x: x.loc[x["Status"] != "Cancelled", "EventHours"].sum() / weeks_in_month
            ),

            # Direct hours
            "CompletedDirect": g.apply(
                lambda x: x.loc[
                    (x["Category"] == "Direct") & (x["Status"] != "Cancelled"), 
                    "EventHours"
                ].sum()
            ),

            # Direct percentage
            "%ofDirect": g.apply(
                lambda x: (
                    (
                        x.loc[
                            (x["Category"] == "Direct") & (x["Status"] != "Cancelled"),
                            "EventHours"
                        ].sum()
                    )
                    /
                    (
                        x.loc[x["Category"] == "Direct", "AuthHours"].sum()
                        * weeks_in_month
                    )
                    * 100
                ) if x.loc[x["Category"] == "Direct", "AuthHours"].sum() > 0 else 0
            ),

            # Indirect hours
            "IndirectUsed": g.apply(
                lambda x: x.loc[
                    (x["Category"] == "Indirect") & (x["Status"] != "Cancelled"), 
                    "EventHours"
                ].sum()
            ),

            # Indirect percentage
            "%ofIndirect": g.apply(
                lambda x: (
                    # Branch 1: PA Distance uses WEEKLY indirect auth hours
                    (
                        x.loc[
                            (x["Category"] == "Indirect") & (x["Status"] != "Cancelled"),
                            "EventHours"
                        ].sum()
                        /
                        (
                            x.loc[x["Category"] == "Indirect", "AuthHours"].sum()
                            * weeks_in_month
                        )
                        * 100
                    )
                    if x["School"].iloc[0] == "School:_PA_Distance_Learning_Charter"
                    else
                    # Branch 2: Normal schools use MONTHLY indirect auth hours
                    (
                        x.loc[
                            (x["Category"] == "Indirect") & (x["Status"] != "Cancelled"),
                            "EventHours"
                        ].sum()
                        /
                        x.loc[x["Category"] == "Indirect", "AuthHours"].sum()
                        * 100
                    )
                )
                if x.loc[x["Category"] == "Indirect", "AuthHours"].sum() > 0
                else 0
            ),

            # Evaluations based on description keywords
            "Evals": g.apply(
                lambda x: x.loc[
                    (x["Status"] != "Cancelled")
                    & (x["ServiceCodeDescription"].str.contains(
                        "Assessment|Evaluation|Test|Functional|Mapp|Promoting", case=False, na=False
                    )),
                    "EventHours"
                ].sum()
            ),

            # Provider cancels
            "ProviderCancel": g.apply(
                lambda x: x.loc[x["SchedulingCancelledReason"] == "Provider Cancel", "EventHours"].sum()
            ),

            # MakeUpTime
            "MakeUp": g.apply(
                lambda x: x.loc[
                    (x["Subcategory"] == "MakeUpTime")
                    & (x["Status"] != "Cancelled"),
                    "EventHours"
                ].sum()
            ),
        })

        # ------------------------ WRITE BOTH SHEETS ------------------------
        output_file = io.BytesIO()
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            data.to_excel(writer, sheet_name="Raw Data", index=False)
            monthly_nums.to_excel(writer, sheet_name="Monthly Numbers", index=False)

        output_file.seek(0)
        return output_file

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

    finally:
        engine.dispose()