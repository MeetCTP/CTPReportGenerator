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

def generate_appointment_agora_report(range_start, range_end, et_file):
    try:
        user_name = os.getlogin()
        connection_string = f"mssql+pymssql://MeetCTP\Administrator:$Unlock01@CTP-DB/CRDB2"
        engine = create_engine(connection_string)
        range_start_101 = datetime.strptime(range_start, '%Y-%m-%d')
        range_end_101 = datetime.strptime(range_end, '%Y-%m-%d')

        appointment_match_query = f"""
            SELECT DISTINCT *
            FROM EasyTracSessionComparison
            WHERE CONVERT(DATE, ServiceDate, 101) BETWEEN '{range_start_101}' AND '{range_end_101}'
        """
        appointment_match_data = pd.read_sql_query(appointment_match_query, engine)

        appointment_match_data.drop_duplicates(inplace=True)
        appointment_match_data.drop('School', axis=1, inplace=True)
        
        if et_file:
            et_file.seek(0)
            et_data = pd.read_excel(et_file)
            et_data = et_data.sort_values(by=['Provider', 'StudentFirstName'], ascending=True)

            appointment_match_data['Provider'] = appointment_match_data['Provider'].astype('object')
            appointment_match_data['StudentFirstName'] = appointment_match_data['StudentFirstName'].astype('object')
            appointment_match_data['StudentLastName'] = appointment_match_data['StudentLastName'].astype('object')
            appointment_match_data['StudentCode'] = appointment_match_data['StudentCode'].astype('object')
            appointment_match_data['ServiceDate'] = pd.to_datetime(appointment_match_data['ServiceDate']).dt.normalize().astype(object)
            appointment_match_data['StartTime'] = appointment_match_data['StartTime'].dt.strftime('%H:%M:%S').astype('object')
            appointment_match_data['EndTime'] = appointment_match_data['EndTime'].dt.strftime('%H:%M:%S').astype('object')
            appointment_match_data['Mileage'] = appointment_match_data['Mileage'].astype('object')

            et_data['ServiceDate'] = pd.to_datetime(et_data['ServiceDate']).dt.normalize().astype(object)
            et_data['StudentCode'] = et_data['StudentCode'].astype('object')

            mile_diffs, et_virtual = find_mileage_discrepancies(et_data, appointment_match_data)
            time_diffs = find_time_discrepancies(et_data, appointment_match_data, et_virtual)
            high_miles, high_times = find_high_mileage(appointment_match_data, et_data)

            output_file = io.BytesIO()
            with ExcelWriter(output_file, engine='openpyxl') as writer:
                appointment_match_data.to_excel(writer, sheet_name="CR Data", index=False)
                et_data.to_excel(writer, sheet_name="ET Data", index=False)
                #cr_mileage.to_excel(writer, sheet_name="CR Mileage Entries", index=False)
                #et_mileage.to_excel(writer, sheet_name="ET Mileage Entries", index=False)
                #duplicates.to_excel(writer, sheet_name="Exact Matches", index=False)
                mile_diffs.to_excel(writer, sheet_name="Mileage Discrepancies", index=False)
                time_diffs.to_excel(writer, sheet_name="Time Discrepancies", index=False)

            output_file.seek(0)
            return output_file
        else:
            appointment_match_data['StartTime'] = appointment_match_data['StartTime'].dt.strftime('%m/%d/%Y %I:%M%p')
            appointment_match_data['EndTime'] = appointment_match_data['EndTime'].dt.strftime('%m/%d/%Y %I:%M%p')
            
            output_file = io.BytesIO()
            appointment_match_data.to_excel(output_file, index=False)
            output_file.seek(0)
            return output_file

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

    finally:
        engine.dispose()

def find_mileage_discrepancies(et_data, cr_data):
    et_mileage = et_data[et_data['Type'].str.contains('Mileage', case=False, na=False)]
    et_virtual = et_data[~et_data['Type'].str.contains('Mileage', case=False, na=False)]
    cr_mileage = cr_data[
        (cr_data['Mileage'] != 0) & (cr_data['Mileage'].notnull())
    ].copy()

    et_mileage = et_mileage[
        ['Provider', 'StudentFirstName', 'StudentLastName', 'StudentCode', 'Type', 'ServiceDate', 'StartTime', 'EndTime']
    ]

    cr_mileage = cr_mileage[
        ['Provider', 'StudentFirstName', 'StudentLastName', 'StudentCode', 'Mileage', 'ServiceDate', 'StartTime', 'EndTime']
    ]

    et_mileage = et_mileage.sort_values(by=['Provider', 'StudentFirstName', 'ServiceDate'], ascending=True)
    cr_mileage = cr_mileage.sort_values(by=['Provider', 'StudentFirstName', 'ServiceDate'], ascending=True)

    for frame in [et_mileage, cr_mileage]:
        for column in frame.select_dtypes(include=['object']).columns:
            frame[column] = frame[column].astype(str).str.strip().str.lower()
            frame[column] = frame[column].astype('object')

    merged_mileage = et_mileage.merge(
        cr_mileage,
        on=['Provider', 'StudentCode', 'ServiceDate'],
        how='outer',
        indicator=True
    )

    mile_diffs = merged_mileage[merged_mileage['_merge'] != 'both'].copy()
    mile_diffs.drop('_merge', axis=1, inplace=True)
    mile_diffs.reset_index(drop=True, inplace=True)
    mile_diffs.drop_duplicates(inplace=True)
    return mile_diffs, et_virtual

def find_time_discrepancies(et_data, appointment_match_data, et_virtual):
    aligned_match_data = appointment_match_data[
        ['Provider', 'StudentFirstName', 'StudentLastName', 'StudentCode', 'ServiceDate', 'StartTime', 'EndTime']
    ]
    aligned_et_data = et_virtual[
        ['Provider', 'StudentFirstName', 'StudentLastName', 'StudentCode', 'ServiceDate', 'StartTime', 'EndTime']
    ]
    
    for df in [aligned_match_data, aligned_et_data]:
        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].astype(str).str.strip().str.lower()
            df[col] = df[col].astype('object')

    aligned_match_data.drop_duplicates(inplace=True)
    aligned_et_data.drop_duplicates(inplace=True)

    combined_data = pd.concat([aligned_match_data, aligned_et_data], keys=['CR', 'ET'], names=['Source'])
    duplicates = combined_data[combined_data.duplicated(keep=False)]
    time_diffs = combined_data.drop_duplicates(keep=False).reset_index(drop=True)
    time_diffs['ServiceDate'] = pd.to_datetime(time_diffs['ServiceDate']).dt.strftime('%m/%d/%Y %I:%M%p').astype(object)

    time_diffs = time_diffs.sort_values(by=['Provider', 'StudentFirstName', 'ServiceDate'], ascending=True)

    return time_diffs

def find_high_mileage(cr_data, et_data):
    

    """
    CODE TO DISPLAY DATATYPES FOR TYPE MATCHING

    match_data_types = aligned_match_data.dtypes.to_dict()
    print("Data types of aligned_match_data columns:")
    for column, dtype in match_data_types.items():
        print(f"{column}: {dtype}")
    """