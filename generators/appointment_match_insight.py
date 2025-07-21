import pandas as pd
import numpy as np
from sqlalchemy import create_engine
from pandas import ExcelWriter
from fuzzywuzzy import fuzz
import re
import pymssql
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime, timedelta
from werkzeug.utils import secure_filename
import os
import io

def generate_appointment_insight_report(range_start, range_end, rsm_file, employment_type):
    try:
        user_name = os.getlogin()
        db_user = os.getenv("DB_USER")
        db_pw = os.getenv("DB_PW")
        connection_string = f"mssql+pymssql://{db_user}:{db_pw}@CTP-DB/CRDB2"
        engine = create_engine(connection_string)
        range_start_101 = datetime.strptime(range_start, '%Y-%m-%d')
        range_end_101 = datetime.strptime(range_end, '%Y-%m-%d')
        
        appointment_match_query = f"""
            SELECT DISTINCT *
            FROM InsightSessionComparison
            WHERE CONVERT(DATE, [Date of Service], 101) BETWEEN '{range_start_101}' AND '{range_end_101}'
        """
        appointment_match_data = pd.read_sql_query(appointment_match_query, engine)

        appointment_match_data.drop_duplicates(inplace=True)
        appointment_match_data.drop('School', axis=1, inplace=True)
        
        if employment_type:
            appointment_match_data = appointment_match_data[appointment_match_data['EmploymentType'].isin(employment_type)]

        if rsm_file:
            rsm_file.seek(0)
            rsm_data = pd.read_excel(rsm_file)
            rsm_data = rsm_data.sort_values(by=['Therapist', 'First Name'], ascending=True)
            
            appointment_match_data['Service Date'] = pd.to_datetime(appointment_match_data['Service Date']).dt.normalize().astype(object)
            appointment_match_data['Start Time'] = pd.to_datetime(appointment_match_data['Start Time']).dt.strftime('%H:%M:%S').astype('object')
            appointment_match_data['End Time'] = pd.to_datetime(appointment_match_data['End Time']).dt.strftime('%H:%M:%S').astype('object')
            
            rsm_data['Student Name'] = rsm_data['Student Name'].astype('object')
            rsm_data['Service Name'] = rsm_data['Service Name'].astype('object')
            rsm_data['Delivery Status'] = rsm_data['Delivery Status'].astype('object')
            rsm_data['Service Date'] = pd.to_datetime(rsm_data['Service Date']).dt.normalize().astype(object)
            rsm_data['ID Number'] = rsm_data['ID Student'].astype('object')
            rsm_data['Start Time'] = pd.to_datetime(rsm_data['Start Time']).dt.strftime('%H:%M:%S').astype('object')
            rsm_data['End Time'] = pd.to_datetime(rsm_data['End Time']).dt.strftime('%H:%M:%S').astype('object')

            rsm_data = pd.merge(rsm_data, appointment_match_data[["Therapist", "EmploymentType"]], 
                                on='Therapist', how='left')
            
            if employment_type:
                rsm_data = rsm_data[rsm_data['EmploymentType'] == employment_type]

            for df in [appointment_match_data, rsm_data]:
                for col in df.select_dtypes(include=['object']).columns:
                    df[col] = df[col].astype(str).str.strip()
                    df[col] = df[col].astype('object')
            
                for col in ['Provider', 'StudentFirstName', 'StudentLastName']:
                    if col in df.columns:
                        df[col] = df[col].astype(str).str.upper()
                        df[col] = df[col].str.replace(r'\s*(JR\.|SR\.|III|II|IV)\s*$', '', regex=True)
                        df[col] = df[col].astype('object')
            
            time_diffs, missing_from = find_time_discrepancies(appointment_match_data, rsm_data)

            appointment_match_data.drop_duplicates(inplace=True)
            rsm_data.drop_duplicates(inplace=True)
            
            output_file = io.BytesIO()
            with ExcelWriter(output_file, engine='openpyxl') as writer:
                appointment_match_data.to_excel(writer, sheet_name="CR Data", index=False)
                rsm_data.to_excel(writer, sheet_name="RSM Data", index=False)
                missing_from.to_excel(writer, sheet_name="Missing From RSM", index=False)
                time_diffs.to_excel(writer, sheet_name="Time Discrepancies", index=False)

            output_file.seek(0)
            return output_file
        else:
            appointment_match_data['Therapy Start Time'] = appointment_match_data['Therapy Start Time'].dt.strftime('%m/%d/%Y %I:%M%p')
            appointment_match_data['Therapy End Time'] = appointment_match_data['Therapy End Time'].dt.strftime('%m/%d/%Y %I:%M%p')
            
            output_file = io.BytesIO()
            appointment_match_data.to_excel(output_file, index=False)
            output_file.seek(0)
            return output_file

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

    finally:
        engine.dispose()

def find_time_discrepancies(cr_data, rsm_data):
    aligned_cr_data = cr_data[
        ['Therapist', 'Student ID', 'First Name', 'Last Name', 'Date of Service', 'Therapy Start Time', 'Therapy End Time', 'Status', 'CancellationReason']    
    ]
    aligned_rsm_data = rsm_data[
        ['Therapist', 'Student ID', 'First Name', 'Last Name', 'Date of Service', 'Therapy Start Time', 'Therapy End Time']        
    ]

    time_diffs = find_time_diffs(aligned_cr_data, aligned_rsm_data)
    missing_from  = find_missing_from(aligned_cr_data, aligned_rsm_data, time_diffs)

    return time_diffs, missing_from

def find_missing_from(aligned_match_data, aligned_et_data, time_diffs):
    merged_df = pd.merge(aligned_match_data, aligned_et_data, on=["Therapist", "First Name", "Last Name", "Student ID", "Date of Service", "Therapy End Time"], how="left", suffixes=("_CR", "_RSM"))
    
    missing_from_et_df = merged_df[merged_df["Therapy Start Time_RSM"].isna()]

    missing_from_et_df["DiscrepancyType"] = "Missing from RSM"
    
    missing_from_et_df['Date of Service'] = pd.to_datetime(missing_from_et_df['Date of Service']).dt.strftime('%m/%d/%Y').astype(object)
    
    discrepancy_df = merged_df = pd.merge(missing_from_et_df, time_diffs, on=["Therapist", "First Name", "Last Name", "Student ID", "Date of Service", "Therapy Start Time_CR"], how="left", indicator=True)
    
    discrepancy_df = discrepancy_df[merged_df['_merge'] == 'left_only']
    
    discrepancy_df.drop(columns=['_merge'], inplace=True)

    discrepancy_df = discrepancy_df[["Therapist", "First Name", "Last Name", 
                                         "Student ID", "Status_x", "Date of Service",
                                         "Therapy Start Time_CR", "Therapy End Time"]]
    
    discrepancy_df = discrepancy_df[discrepancy_df['Status_x'] != 'Un-Converted']
    
    discrepancy_df.drop_duplicates(inplace=True)
    
    discrepancy_df['Date of Service'] = pd.to_datetime(discrepancy_df['Date of Service']).dt.strftime('%m/%d/%Y').astype(object)
    
    discrepancy_df = discrepancy_df.sort_values(by=['Therapist', 'First Name', 'Date of Service'], ascending=True)
    
    return discrepancy_df

def find_time_diffs(aligned_cr_data, aligned_rsm_data):
    cr_copy = aligned_cr_data.copy()
    rsm_copy = aligned_rsm_data.copy()

    cr_copy["Therapy Start Time_Hour"] = pd.to_datetime(cr_copy["Therapy Start Time"]).dt.hour
    cr_copy["Therapy Start Time_Minute"] = pd.to_datetime(cr_copy["Therapy Start Time"]).dt.minute
    cr_copy["Therapy End Time_Hour"] = pd.to_datetime(cr_copy["Therapy End Time"]).dt.hour
    cr_copy["Therapy End Time_Minute"] = pd.to_datetime(cr_copy["Therapy End Time"]).dt.minute
    
    rsm_copy["Therapy Start Time_Hour"] = pd.to_datetime(rsm_copy["Therapy Start Time"]).dt.hour
    rsm_copy["Therapy Start Time_Minute"] = pd.to_datetime(rsm_copy["Therapy Start Time"]).dt.minute
    rsm_copy["Therapy End Time_Hour"] = pd.to_datetime(rsm_copy["Therapy End Time"]).dt.hour
    rsm_copy["Therapy End Time_Minute"] = pd.to_datetime(rsm_copy["Therapy End Time"]).dt.minute

    aligned_cr_data.drop_duplicates(inplace=True)
    aligned_rsm_data.drop_duplicates(inplace=True)

    merged = pd.merge(cr_copy, rsm_copy, on=['Therapist', 'Student ID', 'First Name', 'Last Name', 'Date of Service', "Therapy Start Time_Hour", "Therapy End Time_Hour"], how="outer", suffixes=("_CR", "_RSM"))

    minute_match_condition = (
        (merged["Therapy Start Time_Minute_CR"] == merged["Therapy Start Time_Minute_RSM"]) |
        (merged["Therapy End Time_Minute_CR"] == merged["Therapy End Time_Minute_RSM"])
    )

    merged = merged.dropna(subset=["Therapy Start Time_CR", "Therapy Start Time_RSM", "Therapy End Time_CR", "Therapy End Time_RSM"])
    
    merged["DiscrepancyType"] = None

    start_time_minute_diff = (merged["Therapy Start Time_Minute_CR"] != merged["Therapy Start Time_Minute_RSM"]) & minute_match_condition
    merged.loc[start_time_minute_diff, "DiscrepancyType"] = "Time(Start)"

    end_time_minute_diff = (merged["Therapy End Time_Minute_CR"] != merged["Therapy End Time_Minute_RSM"]) & minute_match_condition
    merged.loc[end_time_minute_diff, "DiscrepancyType"] = "Time(End)"

    merged["DiscrepancyType"] = merged["DiscrepancyType"].fillna("No Discrepancy")
    
    discrepancy_df = merged[["Therapist", "First Name", "Last Name", "Student ID", "Date of Service", 'Status', 'CancellationReason', 
                             "Therapy Start Time_CR", "Therapy Start Time_RSM", "Therapy End Time_CR", "Therapy End Time_RSM", "DiscrepancyType"]]

    time_diffs = discrepancy_df[discrepancy_df['DiscrepancyType'] != "No Discrepancy"]
    
    time_diffs['Date of Service'] = pd.to_datetime(time_diffs['Date of Service']).dt.strftime('%m/%d/%Y').astype(object)

    time_diffs = time_diffs.sort_values(by=['Therapist', 'First Name', 'Date of Service'], ascending=True)
    time_diffs.drop_duplicates(inplace=True)

    return time_diffs