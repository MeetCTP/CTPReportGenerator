import pandas as pd
import numpy as np
from sqlalchemy import create_engine
from pandas import ExcelWriter
import pymssql
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
from werkzeug.utils import secure_filename
import os
import io

def generate_appointment_insight_report(range_start, range_end, rsm_file, employment_type):
    try:
        user_name = os.getlogin()
        connection_string = f"mssql+pymssql://MeetCTP\Administrator:$Unlock01@CTP-DB/CRDB2"
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
            appointment_match_data = appointment_match_data[appointment_match_data['EmploymentType'] == employment_type]

        if rsm_file:
            rsm_file.seek(0)
            rsm_data = pd.read_excel(rsm_file)
            rsm_data = rsm_data.sort_values(by=['Therapist', 'First Name'], ascending=True)
            
            appointment_match_data['Date of Service'] = pd.to_datetime(appointment_match_data['Date of Service']).dt.normalize().astype(object)
            appointment_match_data['Therapy Start Time'] = pd.to_datetime(appointment_match_data['Therapy Start Time']).dt.strftime('%H:%M:%S').astype('object')
            appointment_match_data['Therapy End Time'] = pd.to_datetime(appointment_match_data['Therapy End Time']).dt.strftime('%H:%M:%S').astype('object')
            
            rsm_data['Date of Service'] = pd.to_datetime(rsm_data['Date of Service']).dt.normalize().astype(object)
            rsm_data['Student ID'] = rsm_data['Student ID'].astype('object')

            rsm_data = pd.merge(rsm_data, appointment_match_data[["Therapist", "EmploymentType"]], 
                                on='Therapist', how='left')
            
            if employment_type:
                rsm_data = rsm_data[rsm_data['EmploymentType'] == employment_type]
            
            time_diffs, missing_from = find_time_discrepancies(rsm_data, appointment_match_data)

            appointment_match_data.drop_duplicates(inplace=True)
            rsm_data.drop_duplicates(inplace=True)
            
            output_file = io.BytesIO()
            with ExcelWriter(output_file, engine='openpyxl') as writer:
                appointment_match_data.to_excel(writer, sheet_name="CR Data", index=False)
                rsm_data.to_excel(writer, sheet_name="RSM Data", index=False)
                aligned_cr_data.to_excel(writer, sheet_name="Aligned CR Data", index=False)
                aligned_rsm_data.to_excel(writer, sheet_name="Aligned RSM Data", index=False)
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
        ['Therapist', 'Student ID', 'First Name', 'Last Name', 'Date of Service', 'Therapy Start Time', 'Therapy End Time']    
    ]
    aligned_rsm_data = rsm_data[
        ['Therapist', 'Student ID', 'First Name', 'Last Name', 'Date of Service', 'Therapy Start Time', 'Therapy End Time']        
    ]

    for df in [aligned_cr_data, aligned_rsm_data]:
        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].astype(str).str.strip().str.lower()
            df[col] = df[col].astype('object')

    aligned_cr_data.drop_duplicates(inplace=True)
    aligned_rsm_data.drop_duplicates(inplace=True)

    combined_data = pd.concat([aligned_cr_data, aligned_rsm_data], keys=['CR', 'RSM'], names=['Source'])
    duplicates = combined_data[combined_data.duplicated(keep=False)]
    time_diffs = combined_data.drop_duplicates(keep=False).reset_index(drop=True)
    time_diffs['Date of Service'] = pd.to_datetime(time_diffs['Date of Service']).dt.strftime('%m/%d/%Y %I:%M%p').astype(object)

    time_diffs = time_diffs.sort_values(by=['Therapist', 'First Name', 'Date of Service'], ascending=True)

    return time_diffs, aligned_rsm_data, aligned_cr_data
    
    match_data_types = aligned_cr_data.dtypes.to_dict()
    print("Data types of aligned_match_data columns:")
    for column, dtype in match_data_types.items():
        print(f"{column}: {dtype}")
                
    rsm_data_types = aligned_rsm_data.dtypes.to_dict()
    print('Data types of rsm data:')
    for column, dtype in rsm_data_types.items():
        print(f"{column}: {dtype}")
        
    return aligned_rsm_data, aligned_cr_data