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

def generate_appointment_agora_report(range_start, range_end, et_file, employment_type):
    try:
        user_name = os.getlogin()
        db_user = os.getenv("DB_USER")
        db_pw = os.getenv("DB_PW")
        connection_string = f"mssql+pymssql://{db_user}:{db_pw}@CTP-DB/CRDB2"
        engine = create_engine(connection_string)
        range_start_101 = datetime.strptime(range_start, '%Y-%m-%d')
        range_end_101 = datetime.strptime(range_end, '%Y-%m-%d')

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

        appointment_match_query = f"""
            SELECT DISTINCT *
            FROM EasyTracSessionComparison
            WHERE CONVERT(DATE, ServiceDate, 101) BETWEEN '{range_start_101}' AND '{range_end_101}'
        """
        appointment_match_data = pd.read_sql_query(appointment_match_query, engine)

        appointment_match_data.drop_duplicates(inplace=True)
        appointment_match_data.drop('School', axis=1, inplace=True)
        
        if employment_type:
            appointment_match_data = appointment_match_data[appointment_match_data['EmploymentType'].isin(employment_type)]
        
        if et_file:
            et_file.seek(0)
            et_data = pd.read_excel(et_file)
            et_data = et_data.sort_values(by=['Provider', 'StudentFirstName'], ascending=True)
            et_data['Status'] = et_data['Type'].apply(lambda x: 'Cancelled' if 'Absent' in str(x) else 'Converted')

            appointment_match_data['Provider'] = appointment_match_data['Provider'].astype('object')
            appointment_match_data['StudentFirstName'] = appointment_match_data['StudentFirstName'].astype('object')
            appointment_match_data['StudentLastName'] = appointment_match_data['StudentLastName'].astype('object')
            appointment_match_data['StudentCode'] = appointment_match_data['StudentCode'].astype('object')
            appointment_match_data['ServiceDate'] = pd.to_datetime(appointment_match_data['ServiceDate']).dt.strftime('%m/%d/%Y').astype(object)
            appointment_match_data['StartTime'] = appointment_match_data['StartTime'].dt.strftime('%I:%M%p').astype('object')
            appointment_match_data['EndTime'] = appointment_match_data['EndTime'].dt.strftime('%I:%M%p').astype('object')
            appointment_match_data['Mileage'] = appointment_match_data['Mileage'].astype('object')

            et_data['ServiceDate'] = pd.to_datetime(et_data['ServiceDate']).dt.strftime('%m/%d/%Y').astype(object)
            et_data['StudentCode'] = et_data['StudentCode'].astype('object')
            et_data['StartTime'] = pd.to_datetime(et_data['StartTime'], format='%H:%M:%S').dt.strftime('%I:%M%p').astype('object')
            et_data['EndTime'] = pd.to_datetime(et_data['EndTime'], format='%H:%M:%S').dt.strftime('%I:%M%p').astype('object')
            et_data['DateTimeSigned'] = pd.to_datetime(et_data['DateTimeSigned'], format='%m/%d/%Y %I:%M:%S %p')
            
            et_data = pd.merge(et_data, appointment_match_data[['Provider', 'EmploymentType']], 
                        on='Provider', how='left')
            
            if employment_type:
                et_data = et_data[et_data['EmploymentType'].isin(employment_type)]
                
            for df in [appointment_match_data, et_data]:
                for col in df.select_dtypes(include=['object']).columns:
                    df[col] = df[col].astype(str).str.strip()
                    df[col] = df[col].astype('object')
            
                for col in ['Provider', 'StudentFirstName', 'StudentLastName']:
                    if col in df.columns:
                        df[col] = df[col].astype(str).str.upper()
                        df[col] = df[col].str.replace(r'\s*(JR\.|SR\.|III|II|IV)\s*$', '', regex=True)
                        df[col] = df[col].str.replace('-', ' ', regex=False)
                        df[col] = df[col].astype('object')
                        
            et_data = et_data.loc[et_data.groupby(['Provider', 'Student', 'ServiceDate', 'StartTime', 'EndTime'])['DateTimeSigned'].idxmax()]

            mile_diffs, et_virtual, cr_mileage, et_mileage = find_mileage_discrepancies(et_data, appointment_match_data)
            time_diffs, missing_from, status_diffs = find_time_discrepancies(et_data, appointment_match_data, et_virtual)
            high_miles, high_times = find_high_mileage(cr_mileage, et_mileage)
            
            appointment_match_data.drop_duplicates(inplace=True)
            et_data.drop_duplicates(inplace=True)

            output_file = io.BytesIO()
            with ExcelWriter(output_file, engine='openpyxl') as writer:
                appointment_match_data.to_excel(writer, sheet_name="CR Data", index=False)
                et_data.to_excel(writer, sheet_name="ET Data", index=False)
                missing_from.to_excel(writer, sheet_name="Missing From ET", index=False)
                time_diffs.to_excel(writer, sheet_name="Time Discrepancies", index=False)
                status_diffs.to_excel(writer, sheet_name="Status Discrepancies", index=False)
                mile_diffs.to_excel(writer, sheet_name="Mileage Discrepancies", index=False)
                high_miles.to_excel(writer, sheet_name="Mileage over 60", index=False)
                high_times.to_excel(writer, sheet_name="Over 60 minute Drive Time", index=False)
                #type_diffs.to_excel(writer, sheet_name='Type Discrepancies', index=False)
                #end_time_diffs.to_excel(writer, sheet_name='EndTimeDiffs', index=False)

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
    cr_data['Mileage'] = cr_data['Mileage'].astype('float')
    et_mileage = et_data[et_data['Type'].str.contains('Mileage', case=False, na=False)]
    et_virtual = et_data[~et_data['Type'].str.contains('Mileage', case=False, na=False)]
    cr_mileage = cr_data[
        (cr_data['Mileage'] != 0.0) & (cr_data['Mileage'].notnull())
    ].copy()

    et_mileage = et_mileage[
        ['Provider', 'StudentFirstName', 'StudentLastName', 'StudentCode', 'Minutes', 'ServiceDate', 'StartTime', 'EndTime']
    ]

    cr_mileage = cr_mileage[
        ['Provider', 'StudentFirstName', 'StudentLastName', 'StudentCode', 'Mileage', 'ServiceDate', 'Status', 'CancellationReason', 'StartTime', 'EndTime']
    ]

    et_mileage = et_mileage.sort_values(by=['Provider', 'StudentFirstName', 'ServiceDate'], ascending=True)
    cr_mileage = cr_mileage.sort_values(by=['Provider', 'StudentFirstName', 'ServiceDate'], ascending=True)

    merged_mileage = pd.merge(
        et_mileage, cr_mileage,
        on=['Provider', 'ServiceDate'],
        how='outer',
        indicator=True,
        suffixes=("_ET", "_CR")
    )

    mile_diffs = merged_mileage[merged_mileage['_merge'] == 'right_only'].copy()
    mile_diffs.drop('_merge', axis=1, inplace=True)
    mile_diffs.reset_index(drop=True, inplace=True)
    mile_diffs.drop_duplicates(inplace=True)
    return mile_diffs, et_virtual, cr_mileage, et_mileage

def find_time_discrepancies(et_data, appointment_match_data, et_virtual):
    aligned_match_data = appointment_match_data[
        ['Provider', 'StudentFirstName', 'StudentLastName', 'StudentCode', 'BillingCode', 'BillingDesc', 'ServiceDate', 'StartTime', 'EndTime', 'Status', 'CancellationReason']
    ]
    aligned_et_data = et_virtual[
        ['Provider', 'StudentFirstName', 'StudentLastName', 'StudentCode', 'ServiceDate', 'StartTime', 'EndTime', 'Status', 'Type']
    ]
    
    """for df in [aligned_match_data, aligned_et_data]:
        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].astype('object')
            
    for col in ['Provider', 'StudentFirstName', 'StudentLastName']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.upper()
            df[col] = df[col].str.replace(r'\s*(JR\.|SR\.|III|II|IV)\s*$', '', regex=True)
            df[col] = df[col].astype('object')"""

    aligned_match_data.drop_duplicates(inplace=True)
    aligned_et_data.drop_duplicates(inplace=True)
    
    #start_time_diffs = find_start_time_diffs(aligned_match_data, aligned_et_data)
    #end_time_diffs = find_end_time_diffs(aligned_match_data, aligned_et_data)
    #time_differences = pd.concat([start_time_diffs, end_time_diffs], keys=['CR', 'ET'], names=['Source'])

    time_differences = find_time_diffs(aligned_match_data, aligned_et_data)

    missing_from = find_missing_from(aligned_match_data, aligned_et_data, time_differences)

    status_diffs = find_status_diffs(aligned_match_data, aligned_et_data, missing_from)

    missing_from = merged_df = pd.merge(missing_from, status_diffs, on=["Provider", "StudentFirstName", "StudentLastName", "ServiceDate", "EndTime"], how="left", indicator=True, suffixes=("_CR", "_ET"))
    
    missing_from = missing_from[merged_df['_merge'] == 'left_only']

    missing_from.drop(columns=['_merge', 'Type', 'Status_ET', 'CancellationReason', 'Status_CR', 'StartTime', 'BillingDesc', 'StudentCode_ET', 'BillingCode'], inplace=True)

    #type_diffs = find_type_diffs(aligned_match_data, aligned_et_data)

    time_differences = time_differences.style.applymap(highlight_diff_type_cells, subset=['DiscrepancyType'])
    
    return time_differences, missing_from, status_diffs

def find_high_mileage(cr_data, et_data):
    cr_data = cr_data[
        (cr_data['Mileage'].astype(float) >= 60)
    ].copy()
    
    et_data = et_data[
        (et_data['Minutes'].astype(int) >= 60)    
    ].copy()
    
    return cr_data, et_data

def find_time_diffs(cr_copy, et_copy):
    cr_copy = cr_copy.copy()
    et_copy = et_copy.copy()

    start_merge = find_start_time_diffs(cr_copy, et_copy)
    end_merge = find_end_time_diffs(cr_copy, et_copy)

    cr_copy["StartTime_Hour"] = pd.to_datetime(cr_copy["StartTime"]).dt.hour
    cr_copy["StartTime_Minute"] = pd.to_datetime(cr_copy["StartTime"]).dt.minute
    cr_copy["EndTime_Hour"] = pd.to_datetime(cr_copy["EndTime"]).dt.hour
    cr_copy["EndTime_Minute"] = pd.to_datetime(cr_copy["EndTime"]).dt.minute
    
    et_copy["StartTime_Hour"] = pd.to_datetime(et_copy["StartTime"]).dt.hour
    et_copy["StartTime_Minute"] = pd.to_datetime(et_copy["StartTime"]).dt.minute
    et_copy["EndTime_Hour"] = pd.to_datetime(et_copy["EndTime"]).dt.hour
    et_copy["EndTime_Minute"] = pd.to_datetime(et_copy["EndTime"]).dt.minute

    merged = pd.merge(cr_copy, et_copy, on=["Provider", "StudentFirstName", "StudentLastName", "StudentCode", "Status", "ServiceDate", "StartTime_Hour", "EndTime_Hour"], how="outer", suffixes=("_CR", "_ET"))

    minute_match_condition = (
        (merged["StartTime_Minute_CR"] == merged["StartTime_Minute_ET"]) |
        (merged["EndTime_Minute_CR"] == merged["EndTime_Minute_ET"])
    )

    merged = merged.dropna(subset=["StartTime_CR", "StartTime_ET", "EndTime_CR", "EndTime_ET"])
    
    merged["DiscrepancyType"] = None

    start_time_minute_diff = (merged["StartTime_Minute_CR"] != merged["StartTime_Minute_ET"]) & minute_match_condition
    merged.loc[start_time_minute_diff, "DiscrepancyType"] = "Time(Start)"

    end_time_minute_diff = (merged["EndTime_Minute_CR"] != merged["EndTime_Minute_ET"]) & minute_match_condition
    merged.loc[end_time_minute_diff, "DiscrepancyType"] = "Time(End)"

    merged["DiscrepancyType"] = merged["DiscrepancyType"].fillna("No Discrepancy")
    
    discrepancy_df = merged[["Provider", "StudentFirstName", "StudentLastName", "StudentCode", "BillingCode", "ServiceDate", 'Status', 'CancellationReason', 
                         "StartTime_CR", "StartTime_ET", "EndTime_CR", "EndTime_ET", "DiscrepancyType"]]

    discrepancy_df = discrepancy_df[discrepancy_df['DiscrepancyType'] != "No Discrepancy"]
    
    start_merge.rename(columns={
        'StartTime_CR': 'StartTime_CR_start_merge', 
        'StartTime_ET': 'StartTime_ET_start_merge',
        'EndTime': 'EndTime_start_merge'
    }, inplace=True)

    end_merge.rename(columns={
        'StartTime': 'StartTime_end_merge',
        'EndTime_CR': 'EndTime_CR_end_merge',
        'EndTime_ET': 'EndTime_ET_end_merge'
    }, inplace=True)

    discrepancy_df = pd.merge(
        discrepancy_df, 
        start_merge, 
        on=["Provider", "StudentFirstName", "StudentLastName", "StudentCode", "BillingCode", "Status", "ServiceDate", "CancellationReason"],
        how='outer'
    )

    discrepancy_df = pd.merge(
        discrepancy_df, 
        end_merge, 
        on=["Provider", "StudentFirstName", "StudentLastName", "StudentCode", "BillingCode", "Status", "ServiceDate", "CancellationReason"],
        how='outer'
    )
    discrepancy_df.drop_duplicates(subset=['Provider', 'StudentFirstName', 'StudentCode', 'BillingCode', 'ServiceDate', 'Status', 'CancellationReason', 'DiscrepancyType'], inplace=True)

    discrepancy_df['StartTime_CR'] = discrepancy_df['StartTime_CR'].fillna(discrepancy_df['StartTime_CR_start_merge']).fillna(discrepancy_df['StartTime_end_merge'])
    discrepancy_df['StartTime_ET'] = discrepancy_df['StartTime_ET'].fillna(discrepancy_df['StartTime_ET_start_merge']).fillna(discrepancy_df['StartTime_end_merge'])
    discrepancy_df['EndTime_CR'] = discrepancy_df['EndTime_CR'].fillna(discrepancy_df['EndTime_CR_end_merge']).fillna(discrepancy_df['EndTime_start_merge'])
    discrepancy_df['EndTime_ET'] = discrepancy_df['EndTime_ET'].fillna(discrepancy_df['EndTime_ET_end_merge']).fillna(discrepancy_df['EndTime_start_merge'])
    discrepancy_df['DiscrepancyType'] = discrepancy_df['DiscrepancyType'].fillna(discrepancy_df['DiscrepancyType_x']).fillna(discrepancy_df['DiscrepancyType_y'])

    discrepancy_df.drop(columns=['DiscrepancyType_x', 'StartTime_CR_start_merge', 'StartTime_ET_start_merge', 'EndTime_start_merge', 'DiscrepancyType_y', 'StartTime_end_merge', 'EndTime_CR_end_merge', 'EndTime_ET_end_merge'], inplace=True) 

    overlapping_times = find_overlapping_appointments(cr_copy, et_copy)

    discrepancy_df = pd.merge(discrepancy_df, overlapping_times, on=['Provider', 'StudentFirstName', 'StudentLastName', 'ServiceDate', 'Status', 'StartTime_CR', 'EndTime_CR', 'StartTime_ET', 'EndTime_ET'], how='outer')

    discrepancy_df['DiscrepancyType'] = np.where(
        discrepancy_df['StartTime_CR'] == discrepancy_df['StartTime_ET'], "Time(End)", 
        np.where(
            discrepancy_df['EndTime_CR'] == discrepancy_df['EndTime_ET'], "Time(Start)", 
            "Overlapping"
        )
    )

    discrepancy_df['BillingCode'] = discrepancy_df['BillingCode_x'].fillna(discrepancy_df['BillingCode_y'])
    discrepancy_df['StudentCode'] = discrepancy_df['StudentCode'].fillna(discrepancy_df['StudentCode_CR']).fillna(discrepancy_df['StudentCode_ET'])
    discrepancy_df['CancellationReason'] = discrepancy_df['CancellationReason_x'].fillna(discrepancy_df['CancellationReason_y'])

    discrepancy_df.drop(columns=['BillingCode_x', 'BillingCode_y', 'CancellationReason_x', 'StudentCode_CR', 'StudentCode_ET', 'CancellationReason_y'], inplace=True)

    discrepancy_df = discrepancy_df[["Provider", "StudentFirstName", "StudentLastName", "StudentCode", "BillingCode", "ServiceDate", 'Status', 'CancellationReason', 
                     "StartTime_CR", "StartTime_ET", "EndTime_CR", "EndTime_ET", "DiscrepancyType"]]
    
    #discrepancy_df['ServiceDate'] = pd.to_datetime(discrepancy_df['ServiceDate']).dt.strftime('%m/%d/%Y').astype(object)
    discrepancy_df = discrepancy_df.sort_values(by=['Provider', 'StudentFirstName', 'ServiceDate'], ascending=True)
    discrepancy_df.drop_duplicates(inplace=True)
    
    return discrepancy_df

def find_start_time_diffs(cr_copy, et_copy):
    merged_on_end = pd.merge(
        cr_copy, et_copy,
        on=["Provider", "StudentFirstName", "StudentLastName", "StudentCode", "ServiceDate", "EndTime", "Status"],
        how="outer",
        suffixes=("_CR", "_ET")
    )

    # Initialize discrepancy column
    merged_on_end["DiscrepancyType"] = pd.NA

    # Condition for mismatched start times
    mask = (
        (merged_on_end["StartTime_CR"] != merged_on_end["StartTime_ET"])
        & merged_on_end["StartTime_CR"].notna()
        & merged_on_end["StartTime_ET"].notna()
    )
    merged_on_end.loc[mask, "DiscrepancyType"] = "Time(Start)"

    # Keep only relevant columns
    discrepancy_df = merged_on_end[
        ["Provider", "StudentFirstName", "StudentLastName", "StudentCode",
         "BillingCode", "ServiceDate", "Status", "CancellationReason",
         "StartTime_CR", "StartTime_ET", "EndTime", "DiscrepancyType"]
    ]

    # Keep only rows where discrepancy exists
    discrepancy_df = discrepancy_df[discrepancy_df["DiscrepancyType"].notna()].reset_index(drop=True)

    # Format date and sort
    discrepancy_df["ServiceDate"] = pd.to_datetime(discrepancy_df["ServiceDate"]).dt.strftime("%m/%d/%Y").astype(object)
    discrepancy_df = discrepancy_df.sort_values(by=["Provider", "StudentFirstName", "ServiceDate"], ascending=True)
    discrepancy_df.drop_duplicates(inplace=True)

    return discrepancy_df

def find_end_time_diffs(cr_copy, et_copy):
    merged_on_start = pd.merge(
        cr_copy, et_copy,
        on=["Provider", "StudentFirstName", "StudentLastName", "StudentCode", "ServiceDate", "StartTime", "Status"],
        how="outer",
        suffixes=("_CR", "_ET")
    )

    # Create the discrepancy flag
    merged_on_start["DiscrepancyType"] = pd.NA
    mask = (
        (merged_on_start["EndTime_CR"] != merged_on_start["EndTime_ET"])
        & merged_on_start["EndTime_CR"].notna()
        & merged_on_start["EndTime_ET"].notna()
    )
    merged_on_start.loc[mask, "DiscrepancyType"] = "Time(End)"

    # Keep only relevant columns
    discrepancy_df = merged_on_start[
        ["Provider", "StudentFirstName", "StudentLastName", "StudentCode",
         "BillingCode", "ServiceDate", "Status", "CancellationReason",
         "StartTime", "EndTime_CR", "EndTime_ET", "DiscrepancyType"]
    ]

    # Drop rows where DiscrepancyType is still NA
    discrepancy_df = discrepancy_df[discrepancy_df["DiscrepancyType"].notna()].reset_index(drop=True)

    # Format and sort
    discrepancy_df["ServiceDate"] = pd.to_datetime(discrepancy_df["ServiceDate"]).dt.strftime("%m/%d/%Y").astype(object)
    discrepancy_df = discrepancy_df.sort_values(by=["Provider", "StudentFirstName", "ServiceDate"], ascending=True)
    discrepancy_df.drop_duplicates(inplace=True)

    return discrepancy_df

def find_missing_from(aligned_match_data, aligned_et_data, time_diffs):
    merged_df = pd.merge(aligned_match_data, aligned_et_data, on=["Provider", "StudentFirstName", "StudentLastName", "Status", "StudentCode", "ServiceDate", "EndTime"], how="left", suffixes=("_CR", "_ET"))
    
    missing_from_et_df = merged_df[merged_df["StartTime_ET"].isna()]

    missing_from_et_df["DiscrepancyType"] = "Missing from ET"
    
    missing_from_et_df['ServiceDate'] = pd.to_datetime(missing_from_et_df['ServiceDate']).dt.strftime('%m/%d/%Y').astype(object)
    
    discrepancy_df = merged_df = pd.merge(missing_from_et_df, time_diffs, on=["Provider", "StudentFirstName", "StudentLastName", "StudentCode", "Status", "ServiceDate", "StartTime_CR"], how="left", indicator=True, suffixes=("_CR", "_ET"))
    
    discrepancy_df = discrepancy_df[merged_df['_merge'] == 'left_only']
    
    discrepancy_df.drop(columns=['_merge'], inplace=True)

    discrepancy_df = discrepancy_df[["Provider", "StudentFirstName", "StudentLastName", 
                                         "StudentCode", "BillingCode_CR", "Status", "ServiceDate",
                                         "StartTime_CR", "EndTime", "CancellationReason_CR"]]
    
    discrepancy_df = discrepancy_df[discrepancy_df['Status'] != 'Un-Converted']
    
    discrepancy_df.drop_duplicates(inplace=True)
    
    #discrepancy_df['ServiceDate'] = pd.to_datetime(discrepancy_df['ServiceDate']).dt.strftime('%m/%d/%Y').astype(object)
    
    discrepancy_df = discrepancy_df.sort_values(by=['Provider', 'StudentFirstName', 'ServiceDate'], ascending=True)

    #discrepancy_df = discrepancy_df[discrepancy_df['StudentFirstName'] != 'AGORA CLASSROOM']
    discrepancy_df = discrepancy_df[discrepancy_df['StudentFirstName'] != 'AGORA SEL GROUP']
    #discrepancy_df = discrepancy_df[discrepancy_df['StudentFirstName'] != 'AGORA SOCIAL SKILLS']
    
    return discrepancy_df

def find_overlapping_appointments(cr_copy, et_copy):
    merged = pd.merge(cr_copy, et_copy, on=["Provider", "StudentFirstName", "StudentLastName", "Status", "ServiceDate"], 
                      suffixes=('_CR', '_ET'), how='outer')
    
    merged["StartTime_CR"] = pd.to_datetime(merged["StartTime_CR"], format='%I:%M%p').dt.strftime('%H:%M')
    merged["EndTime_CR"] = pd.to_datetime(merged["EndTime_CR"], format='%I:%M%p').dt.strftime('%H:%M')
    merged["StartTime_ET"] = pd.to_datetime(merged["StartTime_ET"], format='%I:%M%p').dt.strftime('%H:%M')
    merged["EndTime_ET"] = pd.to_datetime(merged["EndTime_ET"], format='%I:%M%p').dt.strftime('%H:%M')
    
    overlap_condition = (
        (merged["StartTime_CR"] < merged["EndTime_ET"]) & (merged["StartTime_CR"] > merged["StartTime_ET"]) |
        (merged["EndTime_CR"] < merged["EndTime_ET"]) & (merged["EndTime_CR"] > merged["StartTime_ET"]) | 
        (merged["StartTime_ET"] < merged["EndTime_CR"]) & (merged["StartTime_ET"] > merged["StartTime_CR"]) |
        (merged["EndTime_ET"] < merged["EndTime_CR"]) & (merged["EndTime_ET"] > merged["StartTime_CR"])
    )

    overlap_df = merged[overlap_condition]
    
    overlap_df["StartTime_CR"] = pd.to_datetime(overlap_df["StartTime_CR"], format='%H:%M').dt.strftime('%I:%M%p').astype('object')
    overlap_df["EndTime_CR"] = pd.to_datetime(overlap_df["EndTime_CR"], format='%H:%M').dt.strftime('%I:%M%p').astype('object')
    overlap_df["StartTime_ET"] = pd.to_datetime(overlap_df["StartTime_ET"], format='%H:%M').dt.strftime('%I:%M%p').astype('object')
    overlap_df["EndTime_ET"] = pd.to_datetime(overlap_df["EndTime_ET"], format='%H:%M').dt.strftime('%I:%M%p').astype('object')
    
    return overlap_df

def find_status_diffs(cr_copy, et_copy, missing_from):
    #missing_from['StartTime'] = missing_from['StartTime_CR']
    merged = pd.merge(cr_copy, et_copy, on=['Provider', 'StudentFirstName', 'StudentLastName', 'StudentCode', 'ServiceDate', 'StartTime', 'EndTime'], how='inner', suffixes=('_CR', '_ET'))
    status_diffs = merged[merged['Status_CR'] != merged['Status_ET']]
    return status_diffs

def highlight_diff_type_cells(val):
    if val == 'Time(Start)':
        return 'background-color: #99ff99'
    elif val == 'Time(End)':
        return 'background-color: #9999ff'
    elif val == 'Overlapping':
        return 'background-color: #ff9999'
    else:
        return ''

category_keywords = {
    'Direct': ['face to face', 'in-person', 'scheduled', 'one-on-one'],
    'Direct: Make-Up Session': ['make-up session', 'make up session', 'rescheduled session', 'make-up'],
    'Direct: Virtual': ['virtual', 'online session', 'remote session', 'telehealth'],
    'Direct: Virtual Make-Up Session': ['virtual make up session', 'virtual make up'],
    'Indirect': ['indirect', 'review', 'progress report', 'school', 'personnel', 'IEP', 'Meeting']
}

def find_type_diffs(cr_copy, et_copy):
    cr_converted = cr_copy[cr_copy['Status'] == 'Converted']
    et_converted = et_copy[et_copy['Status'] == 'Converted']

    type_diffs = pd.merge(cr_converted, et_converted, on=['Provider', 'StudentFirstName', 'StudentLastName', 'ServiceDate', 'StartTime', 'EndTime'], how="inner", suffixes=('_CR', '_ET'))

    type_values = ['Direct', 'Direct: Make-Up Session', 'Direct: Virtual', 'Direct: Virtual Make-Up Session', 'Indirect']

    type_diffs['MatchedType_CR'] = type_diffs['BillingDesc'].apply(find_matching_type, args=(category_keywords, type_values))

    discrepancies = type_diffs[type_diffs['Type'] != type_diffs['MatchedType_CR']]

    return discrepancies

def find_matching_type(billing_desc, type_keywords, type_values):
        billing_desc = str(billing_desc).lower()  # Normalize case for comparison
        
        # First check if the billing description matches any keyword categories
        for appointment_type, keywords in type_keywords.items():
            if any(re.search(r'\b' + re.escape(keyword) + r'\b', billing_desc) for keyword in keywords):
                return appointment_type
            
        best_match_type = None
        highest_score = 0
        for type_value in type_values:
            score = fuzz.ratio(billing_desc, type_value.lower())
            if score > highest_score:
                highest_score = score
                best_match_type = type_value
        
        return best_match_type if highest_score > 50 else None  # Adjust threshold as needed

"""
CODE TO DISPLAY DATATYPES FOR TYPE MATCHING FOR DEBUGGING

match_data_types = aligned_match_data.dtypes.to_dict()
print("Data types of aligned_match_data columns:")
for column, dtype in match_data_types.items():
    print(f"{column}: {dtype}")
"""