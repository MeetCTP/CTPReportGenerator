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

def generate_appointment_agora_report(range_start, range_end, et_file, employment_type):
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
        
        if employment_type:
            appointment_match_data = appointment_match_data[appointment_match_data['EmploymentType'] == employment_type]
        
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
            
            et_data = pd.merge(et_data, appointment_match_data[['Provider', 'EmploymentType']], 
                        on='Provider', how='left')
            
            if employment_type:
                et_data = et_data[et_data['EmploymentType'] == employment_type]

            mile_diffs, et_virtual, cr_mileage, et_mileage = find_mileage_discrepancies(et_data, appointment_match_data)
            time_diffs, missing_from = find_time_discrepancies(et_data, appointment_match_data, et_virtual)
            high_miles, high_times = find_high_mileage(cr_mileage, et_mileage)
            
            appointment_match_data.drop_duplicates(inplace=True)
            et_data.drop_duplicates(inplace=True)

            output_file = io.BytesIO()
            with ExcelWriter(output_file, engine='openpyxl') as writer:
                appointment_match_data.to_excel(writer, sheet_name="CR Data", index=False)
                et_data.to_excel(writer, sheet_name="ET Data", index=False)
                #cr_mileage.to_excel(writer, sheet_name="CR Mileage Entries", index=False)
                #et_mileage.to_excel(writer, sheet_name="ET Mileage Entries", index=False)
                #duplicates.to_excel(writer, sheet_name="Exact Matches", index=False)
                mile_diffs.to_excel(writer, sheet_name="Mileage Discrepancies", index=False)
                high_miles.to_excel(writer, sheet_name="Mileage over 60", index=False)
                high_times.to_excel(writer, sheet_name="Over 60 minute Drive Time", index=False)
                time_diffs.to_excel(writer, sheet_name="Time Discrepancies", index=False)
                missing_from.to_excel(writer, sheet_name="Missing From", index=False)

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
        ['Provider', 'StudentFirstName', 'StudentLastName', 'StudentCode', 'Minutes', 'ServiceDate', 'StartTime', 'EndTime']
    ]

    cr_mileage = cr_mileage[
        ['Provider', 'StudentFirstName', 'StudentLastName', 'StudentCode', 'Mileage', 'ServiceDate', 'Status', 'CancellationReason', 'StartTime', 'EndTime']
    ]

    et_mileage = et_mileage.sort_values(by=['Provider', 'StudentFirstName', 'ServiceDate'], ascending=True)
    cr_mileage = cr_mileage.sort_values(by=['Provider', 'StudentFirstName', 'ServiceDate'], ascending=True)

    for frame in [et_mileage, cr_mileage]:
        for column in frame.select_dtypes(include=['object']).columns:
            frame[column] = frame[column].astype(str).str.strip()
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
    return mile_diffs, et_virtual, cr_mileage, et_mileage

def find_time_discrepancies(et_data, appointment_match_data, et_virtual):
    aligned_match_data = appointment_match_data[
        ['Provider', 'StudentFirstName', 'StudentLastName', 'StudentCode', 'ServiceDate', 'StartTime', 'EndTime', 'Status', 'CancellationReason']
    ]
    aligned_et_data = et_virtual[
        ['Provider', 'StudentFirstName', 'StudentLastName', 'StudentCode', 'ServiceDate', 'StartTime', 'EndTime']
    ]
    
    for df in [aligned_match_data, aligned_et_data]:
        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].astype('object')

    aligned_match_data.drop_duplicates(inplace=True)
    aligned_et_data.drop_duplicates(inplace=True)
    
    #start_time_diffs = find_start_time_diffs(aligned_match_data, aligned_et_data)
    #end_time_diffs = find_end_time_diffs(aligned_match_data, aligned_et_data)
    #time_differences = pd.concat([start_time_diffs, end_time_diffs], keys=['CR', 'ET'], names=['Source'])

    time_differences = find_time_diffs(aligned_match_data, aligned_et_data)

    missing_from = find_missing_from(aligned_match_data, aligned_et_data, appointment_match_data, time_differences)
    
    return time_differences, missing_from  

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

    # Extracting hour and minute from StartTime and EndTime
    cr_copy["StartTime_Hour"] = pd.to_datetime(cr_copy["StartTime"]).dt.hour
    cr_copy["StartTime_Minute"] = pd.to_datetime(cr_copy["StartTime"]).dt.minute
    cr_copy["EndTime_Hour"] = pd.to_datetime(cr_copy["EndTime"]).dt.hour
    cr_copy["EndTime_Minute"] = pd.to_datetime(cr_copy["EndTime"]).dt.minute
    
    et_copy["StartTime_Hour"] = pd.to_datetime(et_copy["StartTime"]).dt.hour
    et_copy["StartTime_Minute"] = pd.to_datetime(et_copy["StartTime"]).dt.minute
    et_copy["EndTime_Hour"] = pd.to_datetime(et_copy["EndTime"]).dt.hour
    et_copy["EndTime_Minute"] = pd.to_datetime(et_copy["EndTime"]).dt.minute

    # Merge based on matching hour values for StartTime and EndTime
    merged = pd.merge(cr_copy, et_copy, on=["Provider", "StudentFirstName", "StudentLastName", "StudentCode", "ServiceDate", "StartTime_Hour", "EndTime_Hour"], how="outer", suffixes=("_CR", "_ET"))

    merged = merged.dropna(subset=["StartTime_CR", "StartTime_ET", "EndTime_CR", "EndTime_ET"])
    
    # Initialize DiscrepancyType
    merged["DiscrepancyType"] = None

    # Check for StartTime discrepancies: compare minutes when hours match
    start_time_minute_diff = (merged["StartTime_Minute_CR"] != merged["StartTime_Minute_ET"])
    merged.loc[start_time_minute_diff, "DiscrepancyType"] = "Time(Start)"

    # Check for EndTime discrepancies: compare minutes when hours match
    end_time_minute_diff = (merged["EndTime_Minute_CR"] != merged["EndTime_Minute_ET"])
    merged.loc[end_time_minute_diff, "DiscrepancyType"] = "Time(End)"

    # Filter out rows with no discrepancies
    merged["DiscrepancyType"] = merged["DiscrepancyType"].fillna("No Discrepancy")
    
    # Only include rows with discrepancies
    discrepancy_df = merged[["Provider", "StudentFirstName", "StudentLastName", "StudentCode", "ServiceDate", 'Status', 'CancellationReason', 
                             "StartTime_CR", "StartTime_ET", "EndTime_CR", "EndTime_ET", "DiscrepancyType"]]

    # Drop rows with no discrepancies
    discrepancy_df = discrepancy_df[discrepancy_df['DiscrepancyType'] != "No Discrepancy"]
    
    # Ensure ServiceDate is in the correct format
    discrepancy_df['ServiceDate'] = pd.to_datetime(discrepancy_df['ServiceDate']).dt.strftime('%m/%d/%Y').astype(object)

    # Sort and remove duplicates
    discrepancy_df = discrepancy_df.sort_values(by=['Provider', 'StudentFirstName', 'ServiceDate'], ascending=True)
    discrepancy_df.drop_duplicates(inplace=True)
    
    return discrepancy_df

def find_start_time_diffs(cr_copy, et_copy):
    cr_copy = cr_copy.rename(columns={"StartTime": "StartTime_CR"})
    et_copy = et_copy.rename(columns={"StartTime": "StartTime_ET"})

    merged_on_end = pd.merge(cr_copy, et_copy, on=["Provider", "StudentFirstName", "StudentLastName", "StudentCode", "ServiceDate", "EndTime"], how="outer")
    
    start_time_diff = (merged_on_end["StartTime_CR"] != merged_on_end["StartTime_ET"])
    
    start_time_diff["DiscrepancyType"] = None
    merged_on_end.loc[start_time_diff & merged_on_end["StartTime_CR"].notna() & merged_on_end["StartTime_ET"].notna(), "DiscrepancyType"] = "Time(Start)"
    
    discrepancy_df = merged_on_end[["Provider", "StudentFirstName", "StudentLastName", "StudentCode", "ServiceDate", 'Status', 'CancellationReason', 
                                "StartTime_CR", "StartTime_ET", "EndTime", "DiscrepancyType"]]
    
    discrepancy_df = discrepancy_df.dropna(subset=["DiscrepancyType"])
    
    discrepancy_df['ServiceDate'] = pd.to_datetime(discrepancy_df['ServiceDate']).dt.strftime('%m/%d/%Y').astype(object)
    discrepancy_df = discrepancy_df.sort_values(by=['Provider', 'StudentFirstName', 'ServiceDate'], ascending=True)
    discrepancy_df.drop_duplicates(inplace=True)
    
    return discrepancy_df

def find_end_time_diffs(cr_copy, et_copy):
    cr_copy = cr_copy.rename(columns={"EndTime": "EndTime_CR"})
    et_copy = et_copy.rename(columns={"EndTime": "EndTime_ET"})
    
    merged_on_start = pd.merge(cr_copy, et_copy, on=["Provider", "StudentFirstName", "StudentLastName", "StudentCode", "ServiceDate", "StartTime"], how="outer")
    
    end_time_diff = (merged_on_start["EndTime_CR"] != merged_on_start["EndTime_ET"])
    
    end_time_diff["DiscrepancyType"] = None
    merged_on_start.loc[end_time_diff & merged_on_start["EndTime_CR"].notna() & merged_on_start["EndTime_ET"].notna(), "DiscrepancyType"] = "Time(End)"
    
    discrepancy_df = merged_on_start[["Provider", "StudentFirstName", "StudentLastName", "StudentCode", "ServiceDate", 'Status', 'CancellationReason', 
                                "StartTime", "EndTime_CR", "EndTime_ET", "DiscrepancyType"]]
    
    discrepancy_df = discrepancy_df.dropna(subset=["DiscrepancyType"])
    
    discrepancy_df['ServiceDate'] = pd.to_datetime(discrepancy_df['ServiceDate']).dt.strftime('%m/%d/%Y').astype(object)
    discrepancy_df = discrepancy_df.sort_values(by=['Provider', 'StudentFirstName', 'ServiceDate'], ascending=True)
    discrepancy_df.drop_duplicates(inplace=True)
    
    return discrepancy_df

def find_missing_from(aligned_match_data, aligned_et_data, cr_data, time_diffs):
    aligned_match_data = aligned_match_data.rename(columns={"StartTime": "StartTime_CR"})
    aligned_et_data = aligned_et_data.rename(columns={"StartTime": "StartTime_ET"})
    
    merged_df = pd.merge(aligned_match_data, aligned_et_data, on=["Provider", "StudentFirstName", "StudentLastName", "StudentCode", "ServiceDate", "EndTime"], how="outer")
    
    merged_df["DiscrepancyType"] = None
    #merged_df.loc[merged_df["StartTime_CR"].isna(), "DiscrepancyType"] = "Missing from CR"
    merged_df.loc[merged_df["StartTime_ET"].isna(), "DiscrepancyType"] = "Missing from ET"

    

    discrepancy_df = merged_df[["Provider", "StudentFirstName", "StudentLastName", "StudentCode", "ServiceDate", 'Status', 'CancellationReason', 
                                "StartTime_CR", "EndTime", "DiscrepancyType"]]
    
    discrepancy_df = discrepancy_df[discrepancy_df['Status'] != 'Un-Converted']
    
    discrepancy_df = merged_df.dropna(subset=["DiscrepancyType"])
    
    discrepancy_df['ServiceDate'] = pd.to_datetime(discrepancy_df['ServiceDate']).dt.strftime('%m/%d/%Y').astype(object)
    discrepancy_df = discrepancy_df.sort_values(by=['Provider', 'StudentFirstName', 'ServiceDate'], ascending=True)
    discrepancy_df.drop_duplicates(inplace=True)
    
    return discrepancy_df

def calculate_match_percentage(row1, row2):
    # Count the number of matching values between the two rows
    matches = np.sum(row1 == row2)
    
    # Calculate match percentage
    match_percentage = matches / len(row1)  # Number of matches divided by total number of columns
    
    return match_percentage

"""
CODE TO DISPLAY DATATYPES FOR TYPE MATCHING

match_data_types = aligned_match_data.dtypes.to_dict()
print("Data types of aligned_match_data columns:")
for column, dtype in match_data_types.items():
    print(f"{column}: {dtype}")
"""