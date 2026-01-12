import pandas as pd
import numpy as np
from sqlalchemy import create_engine
from pandas import ExcelWriter
from fuzzywuzzy import fuzz, process
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
            FROM InsightSessionComparison
            WHERE CONVERT(DATE, [Service Date], 101) BETWEEN '{range_start_101}' AND '{range_end_101}'
        """

        if 'Employee' in employment_type and not 'Contractor' in employment_type:
            appointment_match_query += f""" AND (Provider IN ({', '.join([f"'{s}'" for s in employee_providers])}))"""
        elif 'Contractor' in employment_type and not 'Employee' in employment_type:
            appointment_match_query += f""" AND (Provider NOT IN ({', '.join([f"'{s}'" for s in employee_providers])}))"""

        appointment_match_data = pd.read_sql_query(appointment_match_query, engine)

        appointment_match_data.drop_duplicates(inplace=True)
        appointment_match_data.drop('School', axis=1, inplace=True)
        
        #if employment_type:
        #    appointment_match_data = appointment_match_data[appointment_match_data['EmploymentType'].isin(employment_type)]

        if rsm_file:
            rsm_file.seek(0)
            rsm_data = pd.read_excel(rsm_file)

            rsm_data.rename(columns={'Student ID number': 'ID Number'}, inplace=True)
            rsm_data.rename(columns={'Service Log Service Name': 'Service Name'}, inplace=True)
            rsm_data.rename(columns={'Service Log Delivery Status': 'Delivery Status'}, inplace=True)


            # Normalize the uploaded RSM file into expected structure
            if "Student First Name" in rsm_data.columns:
                # Build Student Name as "First M Last"
                rsm_data["Student Name"] = (
                    rsm_data["Student First Name"].astype(str).str.strip()
                    + " "
                    #+ rsm_data["Student Middle Name"].fillna("").astype(str).str.strip().str[:1]
                    #+ np.where(rsm_data["Student Middle Name"].notna() & (rsm_data["Student Middle Name"].str.strip() != ""), " ", "")
                    + rsm_data["Student Last Name"].astype(str).str.strip()
                ).str.replace("  ", " ")  # clean double spaces

                # Build Provider as "First Last"
                rsm_data["Provider"] = (
                    rsm_data["Provider First Name"].astype(str).str.strip()
                    + " "
                    + rsm_data["Provider Last Name"].astype(str).str.strip()
                )

                # Map long structure into expected short structure
                column_mapping = {
                    "ID": "ID",
                    "Status": "Status",
                    "Service Date": "Service Date",
                    "ID Number": "ID Number",
                    "Student Name": "Student Name",
                    "Service Name": "Service Name",
                    "Delivery Status": "Delivery Status",
                    "Start Time": "Start Time",
                    "End Time": "End Time",
                    "Duration": "Duration",
                    "Billable Units": "Decimal",  # if this is the same meaning
                    "Provider": "Provider",
                    "Updated At": "Updated At"
                }

                # Keep only what you need
                rsm_data = rsm_data[[col for col in column_mapping.keys() if col in rsm_data.columns]].rename(columns=column_mapping)

            rsm_data = rsm_data.sort_values(by=['Provider', 'Student Name'], ascending=True)
            rsm_data['Status'] = rsm_data['Delivery Status'].apply(lambda x: 'Converted' if 'Administered' in str(x) else 'Cancelled')
            
            appointment_match_data['Service Date'] = pd.to_datetime(appointment_match_data['Service Date']).dt.normalize().astype('object')
            appointment_match_data['Start Time'] = pd.to_datetime(appointment_match_data['Start Time'], errors='coerce').dt.strftime('%I:%M%p').astype('object')
            appointment_match_data['End Time'] = pd.to_datetime(appointment_match_data['End Time'], errors='coerce').dt.strftime('%I:%M%p').astype('object')
            appointment_match_data['ID Number'] = appointment_match_data['ID Number'].astype('object')
            appointment_match_data['Student Name'] = appointment_match_data['Student Name'].astype('object')
            appointment_match_data['Status'] = appointment_match_data['Status'].astype('object')
            
            rsm_data['Student Name'] = rsm_data['Student Name'].astype('object')
            rsm_data['Service Name'] = rsm_data['Service Name'].astype('object')
            rsm_data['Delivery Status'] = rsm_data['Delivery Status'].astype('object')
            rsm_data['Service Date'] = pd.to_datetime(rsm_data['Service Date']).dt.normalize().astype('object')
            rsm_data['ID Number'] = rsm_data['ID Number'].astype('object')
            rsm_data['Start Time'] = pd.to_datetime(rsm_data['Start Time'], format='%H:%M:%S', errors='coerce').dt.strftime('%I:%M%p').astype('object')
            rsm_data['End Time'] = pd.to_datetime(rsm_data['End Time'], format='%H:%M:%S', errors='coerce').dt.strftime('%I:%M%p').astype('object')

            #rsm_data = pd.merge(rsm_data, appointment_match_data[["Provider", "EmploymentType"]], 
            #                    on='Provider', how='left')
            
            #if employment_type:
            #    rsm_data = rsm_data[rsm_data['EmploymentType'].isin(employment_type)]

            for df in [appointment_match_data, rsm_data]:
                for col in df.select_dtypes(include=['object']).columns:
                    df[col] = df[col].astype(str).str.strip()
                    df[col] = df[col].astype('object')
            
                for col in ['Provider', 'Student Name']:
                    if col in df.columns:
                        df[col] = df[col].astype(str).str.upper()
                        df[col] = df[col].str.replace(r'\s*(JR\.|SR\.|III|II|IV)\s*$', '', regex=True)
                        df[col] = df[col].str.replace('-', ' ', regex=False)
                        df[col] = df[col].astype('object')

            rsm_data = normalize_names_with_reference(rsm_data, appointment_match_data)
            
            time_diffs, missing_from, status_diffs, bsc_bcba_diffs = find_time_discrepancies(appointment_match_data, rsm_data)

            appointment_match_data.drop_duplicates(inplace=True)
            #rsm_data.drop_duplicates(inplace=True)
            
            output_file = io.BytesIO()
            with ExcelWriter(output_file, engine='openpyxl') as writer:
                appointment_match_data.to_excel(writer, sheet_name="CR Data", index=False)
                rsm_data.to_excel(writer, sheet_name="Portal Data", index=False)
                missing_from.to_excel(writer, sheet_name="Missing From Portal", index=False)
                time_diffs.to_excel(writer, sheet_name="Time Discrepancies", index=False)
                bsc_bcba_diffs.to_excel(writer, sheet_name="BSC and BCBA Discrepancies", index=False)
                status_diffs.to_excel(writer, sheet_name="Status Discrepancies", index=False)

            output_file.seek(0)
            return output_file
        else:
            appointment_match_data['Start Time'] = appointment_match_data['Start Time'].dt.strftime('%m/%d/%Y %I:%M%p')
            appointment_match_data['End Time'] = appointment_match_data['End Time'].dt.strftime('%m/%d/%Y %I:%M%p')
            
            output_file = io.BytesIO()
            appointment_match_data.to_excel(output_file, index=False)
            output_file.seek(0)
            return output_file

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

    finally:
        engine.dispose()

def normalize_names_with_reference(rsm_data, appointment_match_data, threshold=90):
    ref_names = appointment_match_data["Student Name"].astype(str).unique().tolist()
    
    normalized_names = []
    match_scores = []
    
    for raw_name in rsm_data["Student Name"].astype(str):
        match, score = process.extractOne(raw_name, ref_names, scorer=fuzz.token_set_ratio)
        
        if score >= threshold:
            normalized_names.append(match)  # canonical form from reference
        else:
            normalized_names.append(raw_name)  # fallback: keep original
        
        match_scores.append(score)  # keep score for auditing
    
    rsm_data["Student Name (Normalized)"] = normalized_names
    rsm_data["Name Match Score"] = match_scores
    
    return rsm_data

def categorize_service(text):
    if pd.isna(text):
        return 'Unknown'
    text = text.lower()
    if 'bsc' in text or 'behavior specialist' in text:
        if 'group' in text:
            return 'BSC Group'
        elif 'indiv' in text or 'individual' in text:
            return 'BSC Individual'
        elif 'progress report' in text:
            return 'BSC Progress Report'
        elif 'iep' in text:
            return 'BSC IEP Meeting'
        elif 'consult' in text:
            return 'BSC Individual'
        else:
            return 'BSC Other'
    elif 'bcba' in text or 'behavior analyst' in text:
        if 'group' in text:
            return 'BCBA Group'
        elif 'indiv' in text or 'individual' in text:
            return 'BCBA Individual'
        elif 'progress report' in text:
            return 'BCBA Progress Report'
        elif 'iep' in text:
            return 'BCBA IEP Meeting'
        elif 'consult' in text:
            return 'BCBA Individual'
        else:
            return 'BCBA Other'
    else:
        return 'Other'

def find_time_discrepancies(cr_data, rsm_data):
    aligned_cr_data = cr_data[
        ['Provider', 'ID Number', 'Student Name', 'BillingCode', 'BillingDesc', 'Service Date', 'Start Time', 'End Time', 'Status', 'CancellationReason']    
    ]
    aligned_rsm_data = rsm_data[
        ['Provider', 'ID Number', 'Student Name (Normalized)', 'Service Name', 'Service Date', 'Start Time', 'End Time', 'Status']        
    ]

    aligned_rsm_data.rename(columns={'Student Name (Normalized)': 'Student Name'}, inplace=True)

    bsc_bcba_rsm = aligned_rsm_data[
        aligned_rsm_data['Service Name'].str.contains('BCBA|Behavior Specialist', case=False, na=False)
    ]
    bsc_bcba_cr = aligned_cr_data[
        aligned_cr_data['BillingCode'].str.contains('BSC|BCBA', case=False, na=False)
    ]

    aligned_rsm_data = aligned_rsm_data[
        ~aligned_rsm_data['Service Name'].str.contains('BCBA|Behavior Specialist', case=False, na=False)
    ]
    aligned_cr_data = aligned_cr_data[
        ~aligned_cr_data['BillingCode'].str.contains('BSC|BCBA', case=False, na=False)
    ]

    bsc_bcba_diffs = find_bsc_bcba_discrepancies(bsc_bcba_rsm, bsc_bcba_cr)
    time_diffs = find_time_diffs(aligned_cr_data, aligned_rsm_data)
    missing_from  = find_missing_from(aligned_cr_data, aligned_rsm_data, time_diffs)
    status_diffs = find_status_diffs(aligned_cr_data, aligned_rsm_data, missing_from)

    missing_from.rename(columns={'Start Time_CR': 'Start Time'}, inplace=True)

    missing_from = (
        missing_from
            .merge(status_diffs, on=["Provider", "Student Name", "Start Time"], how="left", indicator=True)
            .query('_merge == "left_only"')
            .drop(columns=['_merge'])
    )

    missing_from.rename(columns={'Start Time': 'Start Time_CR'}, inplace=True)

    time_diffs = time_diffs.style.applymap(highlight_diff_type_cells, subset=['DiscrepancyType'])

    return time_diffs, missing_from, status_diffs, bsc_bcba_diffs

def find_missing_from(aligned_match_data, aligned_rsm_data, time_diffs):
    merged_df = pd.merge(aligned_match_data, aligned_rsm_data, on=["Provider", "Student Name", "Status", "Service Date", "End Time"], how="left", suffixes=("_CR", "_IPP"))
    
    missing_from_rsm_df = merged_df[merged_df["Start Time_IPP"].isna()]

    missing_from_rsm_df["DiscrepancyType"] = "Missing from IPP"
    
    missing_from_rsm_df['Service Date'] = pd.to_datetime(missing_from_rsm_df['Service Date']).dt.strftime('%m/%d/%Y').astype(object)
    
    discrepancy_df = merged_df = pd.merge(missing_from_rsm_df, time_diffs, on=["Provider", "Student Name", "Status", "Service Date", "Start Time_CR"], how="left", indicator=True, suffixes=("_CR", "_IPP"))
    
    discrepancy_df = discrepancy_df[merged_df['_merge'] == 'left_only']
    
    discrepancy_df.drop(columns=['_merge'], inplace=True)

    discrepancy_df = discrepancy_df[["Provider", "Student Name", 
                                         "ID Number_IPP", "BillingCode_CR", "Status", "Service Date",
                                         "Start Time_CR", "End Time", "CancellationReason_CR"]]
    
    discrepancy_df = discrepancy_df[discrepancy_df['Status'] != 'Un-Converted']
    discrepancy_df = discrepancy_df[
        ~discrepancy_df['BillingCode_CR'].str.contains('BSC|BCBA', na=False)
    ]
    
    discrepancy_df.drop_duplicates(inplace=True)
    
    #discrepancy_df['ServiceDate'] = pd.to_datetime(discrepancy_df['ServiceDate']).dt.strftime('%m/%d/%Y').astype(object)
    
    discrepancy_df = discrepancy_df.sort_values(by=['Provider', 'Student Name', 'Service Date'], ascending=True)

    #discrepancy_df = discrepancy_df[discrepancy_df['StudentFirstName'] != 'AGORA CLASSROOM']
    #discrepancy_df = discrepancy_df[discrepancy_df['StudentFirstName'] != 'AGORA SEL GROUP']
    #discrepancy_df = discrepancy_df[discrepancy_df['StudentFirstName'] != 'AGORA SOCIAL SKILLS']
    
    return discrepancy_df

def find_time_diffs(cr_copy, rsm_copy):
    cr_copy = cr_copy.copy()
    rsm_copy = rsm_copy.copy()

    start_merge = find_start_time_diffs(cr_copy, rsm_copy)
    end_merge = find_end_time_diffs(cr_copy, rsm_copy)

    cr_copy["StartTime_Hour"] = pd.to_datetime(cr_copy["Start Time"], errors='coerce').dt.hour
    cr_copy["StartTime_Minute"] = pd.to_datetime(cr_copy["Start Time"], errors='coerce').dt.minute
    cr_copy["EndTime_Hour"] = pd.to_datetime(cr_copy["End Time"], errors='coerce').dt.hour
    cr_copy["EndTime_Minute"] = pd.to_datetime(cr_copy["End Time"], errors='coerce').dt.minute
    
    rsm_copy["StartTime_Hour"] = pd.to_datetime(rsm_copy["Start Time"], errors='coerce').dt.hour
    rsm_copy["StartTime_Minute"] = pd.to_datetime(rsm_copy["Start Time"], errors='coerce').dt.minute
    rsm_copy["EndTime_Hour"] = pd.to_datetime(rsm_copy["End Time"], errors='coerce').dt.hour
    rsm_copy["EndTime_Minute"] = pd.to_datetime(rsm_copy["End Time"], errors='coerce').dt.minute

    merged = pd.merge(cr_copy, rsm_copy, on=["Provider", "Student Name", "ID Number", "Status", "Service Date", "StartTime_Hour", "EndTime_Hour"], how="outer", suffixes=("_CR", "_IPP"))

    minute_match_condition = (
        (merged["StartTime_Minute_CR"] == merged["StartTime_Minute_IPP"]) |
        (merged["EndTime_Minute_CR"] == merged["EndTime_Minute_IPP"])
    )

    merged = merged.dropna(subset=["Start Time_CR", "Start Time_IPP", "End Time_CR", "End Time_IPP"])
    
    merged["DiscrepancyType"] = None

    start_time_minute_diff = (merged["StartTime_Minute_CR"] != merged["StartTime_Minute_IPP"]) & minute_match_condition
    merged.loc[start_time_minute_diff, "DiscrepancyType"] = "Time(Start)"

    end_time_minute_diff = (merged["EndTime_Minute_CR"] != merged["EndTime_Minute_IPP"]) & minute_match_condition
    merged.loc[end_time_minute_diff, "DiscrepancyType"] = "Time(End)"

    merged["DiscrepancyType"] = merged["DiscrepancyType"].fillna("No Discrepancy")
    
    discrepancy_df = merged[["Provider", "Student Name", "ID Number", "BillingCode", "Service Date", 'Status', 'CancellationReason', 
                         "Start Time_CR", "Start Time_IPP", "End Time_CR", "End Time_IPP", "DiscrepancyType"]]

    discrepancy_df = discrepancy_df[discrepancy_df['DiscrepancyType'] != "No Discrepancy"]
    
    #discrepancy_df['ServiceDate'] = pd.to_datetime(discrepancy_df['ServiceDate']).dt.strftime('%m/%d/%Y').astype(object)

    start_merge.rename(columns={
        'Start Time_CR': 'StartTime_CR_start_merge', 
        'Start Time_IPP': 'StartTime_IPP_start_merge',
        'End Time': 'EndTime_start_merge'
    }, inplace=True)

    end_merge.rename(columns={
        'Start Time': 'StartTime_end_merge',
        'End Time_CR': 'EndTime_CR_end_merge',
        'End Time_IPP': 'EndTime_IPP_end_merge'
    }, inplace=True)

    discrepancy_df = pd.merge(
        discrepancy_df, 
        start_merge, 
        on=["Provider", "Student Name", "ID Number", "BillingCode", "Status", "Service Date", "CancellationReason"],
        how='outer'
    )

    discrepancy_df = pd.merge(
        discrepancy_df, 
        end_merge, 
        on=["Provider", "Student Name", "ID Number", "BillingCode", "Status", "Service Date", "CancellationReason"],
        how='outer'
    )
    discrepancy_df.drop_duplicates(subset=['Provider', 'Student Name', 'ID Number', 'BillingCode', 'Service Date', 'Status', 'CancellationReason', 'DiscrepancyType'], inplace=True)

    discrepancy_df['Start Time_CR'] = discrepancy_df['Start Time_CR'].fillna(discrepancy_df['StartTime_CR_start_merge']).fillna(discrepancy_df['StartTime_end_merge'])
    discrepancy_df['Start Time_IPP'] = discrepancy_df['Start Time_IPP'].fillna(discrepancy_df['StartTime_IPP_start_merge']).fillna(discrepancy_df['StartTime_end_merge'])
    discrepancy_df['End Time_CR'] = discrepancy_df['End Time_CR'].fillna(discrepancy_df['EndTime_start_merge']).fillna(discrepancy_df['EndTime_CR_end_merge'])
    discrepancy_df['End Time_IPP'] = discrepancy_df['End Time_IPP'].fillna(discrepancy_df['EndTime_IPP_end_merge']).fillna(discrepancy_df['EndTime_start_merge'])
    discrepancy_df['DiscrepancyType'] = discrepancy_df['DiscrepancyType'].fillna(discrepancy_df['DiscrepancyType_x']).fillna(discrepancy_df['DiscrepancyType_y'])

    discrepancy_df.drop(columns=['DiscrepancyType_x', 'StartTime_CR_start_merge', 'StartTime_IPP_start_merge', 'EndTime_start_merge', 'DiscrepancyType_y', 'StartTime_end_merge', 'EndTime_CR_end_merge', 'EndTime_IPP_end_merge'], inplace=True)

    overlapping_times = find_overlapping_appointments(cr_copy, rsm_copy)

    discrepancy_df = pd.merge(discrepancy_df, overlapping_times, on=['Provider', 'Student Name', 'Service Date', 'Status', 'Start Time_CR', 'End Time_CR', 'Start Time_IPP', 'End Time_IPP'], how='outer')

    discrepancy_df['DiscrepancyType'] = np.where(
        discrepancy_df['Start Time_CR'] == discrepancy_df['Start Time_IPP'], "Time(End)", 
        np.where(
            discrepancy_df['End Time_CR'] == discrepancy_df['End Time_IPP'], "Time(Start)", 
            "Overlapping"
        )
    )

    discrepancy_df['BillingCode'] = discrepancy_df['BillingCode_x'].fillna(discrepancy_df['BillingCode_y'])
    discrepancy_df['ID Number'] = discrepancy_df['ID Number'].fillna(discrepancy_df['ID Number_CR']).fillna(discrepancy_df['ID Number_IPP'])
    discrepancy_df['CancellationReason'] = discrepancy_df['CancellationReason_x'].fillna(discrepancy_df['CancellationReason_y'])

    discrepancy_df.drop(columns=['BillingCode_x', 'BillingCode_y', 'CancellationReason_x', 'ID Number_CR', 'ID Number_IPP', 'CancellationReason_y'], inplace=True)

    discrepancy_df = discrepancy_df[["Provider", "Student Name", "ID Number", "BillingCode", "Service Date", 'Status', 'CancellationReason', 
                     "Start Time_CR", "Start Time_IPP", "End Time_CR", "End Time_IPP", "DiscrepancyType"]]
    
    """
    discrepancy_df["Service Date"] = (
        discrepancy_df["Service Date"]
        .astype(str)                # ensure strings
        .str.strip()                # remove leading/trailing spaces
        .str.replace(r"[-]", "/", regex=True)  # unify separators
    )

    discrepancy_df["Service Date"] = pd.to_datetime(discrepancy_df["Service Date"], errors="coerce")
    discrepancy_df["Service Date"] = discrepancy_df["Service Date"].dt.strftime("%m/%d/%Y")
    """
    
    #discrepancy_df['Service Date'] = pd.to_datetime(discrepancy_df['Service Date']).dt.strftime('%m/%d/%Y')
    discrepancy_df = discrepancy_df.sort_values(by=['Provider', 'Student Name', 'Service Date'], ascending=True)
    discrepancy_df = discrepancy_df[
        ~discrepancy_df['BillingCode'].str.contains('BSC|BCBA', na=False)
    ]
    discrepancy_df.drop_duplicates(inplace=True)
    
    return discrepancy_df

def find_start_time_diffs(cr_copy, rsm_copy):
    merged_on_end = pd.merge(
        cr_copy, rsm_copy,
        on=["Provider", "Student Name", "ID Number", "Service Date", "End Time", "Status"],
        how="outer",
        suffixes=("_CR", "_IPP")
    )

    # Initialize discrepancy column
    merged_on_end["DiscrepancyType"] = pd.NA

    # Condition for mismatched start times
    mask = (
        (merged_on_end["Start Time_CR"] != merged_on_end["Start Time_IPP"])
        & merged_on_end["Start Time_CR"].notna()
        & merged_on_end["Start Time_IPP"].notna()
    )
    merged_on_end.loc[mask, "DiscrepancyType"] = "Time(Start)"

    # Keep only relevant columns
    discrepancy_df = merged_on_end[
        ["Provider", "Student Name", "ID Number", "BillingCode", "Service Date",
         "Status", "CancellationReason", "Start Time_CR", "Start Time_IPP",
         "End Time", "DiscrepancyType"]
    ]

    # Keep only rows where discrepancy exists
    discrepancy_df = discrepancy_df[discrepancy_df["DiscrepancyType"].notna()].reset_index(drop=True)

    # Format date and sort
    discrepancy_df["Service Date"] = pd.to_datetime(discrepancy_df["Service Date"]).dt.strftime("%m/%d/%Y")
    discrepancy_df = discrepancy_df.sort_values(by=["Provider", "Student Name", "Service Date"], ascending=True)
    discrepancy_df.drop_duplicates(inplace=True)

    return discrepancy_df

def find_end_time_diffs(cr_copy, rsm_copy):
    merged_on_start = pd.merge(
        cr_copy, rsm_copy,
        on=["Provider", "Student Name", "ID Number", "Service Date", "Start Time", "Status"],
        how="outer",
        suffixes=("_CR", "_IPP")
    )

    # Initialize discrepancy column
    merged_on_start["DiscrepancyType"] = pd.NA

    # Condition for mismatched end times
    mask = (
        (merged_on_start["End Time_CR"] != merged_on_start["End Time_IPP"])
        & merged_on_start["End Time_CR"].notna()
        & merged_on_start["End Time_IPP"].notna()
    )
    merged_on_start.loc[mask, "DiscrepancyType"] = "Time(End)"

    # Keep only relevant columns
    discrepancy_df = merged_on_start[
        ["Provider", "Student Name", "ID Number", "BillingCode", "Service Date",
         "Status", "CancellationReason", "Start Time", "End Time_CR", "End Time_IPP",
         "DiscrepancyType"]
    ]

    # Keep only rows where discrepancy exists
    discrepancy_df = discrepancy_df[discrepancy_df["DiscrepancyType"].notna()].reset_index(drop=True)

    # Format date and sort
    discrepancy_df["Service Date"] = pd.to_datetime(discrepancy_df["Service Date"]).dt.strftime("%m/%d/%Y")
    discrepancy_df = discrepancy_df.sort_values(by=["Provider", "Student Name", "Service Date"], ascending=True)
    discrepancy_df.drop_duplicates(inplace=True)

    return discrepancy_df

def find_overlapping_appointments(cr_copy, rsm_copy):
    merged = pd.merge(cr_copy, rsm_copy, on=["Provider", "Student Name", "Status", "Service Date"], 
                      suffixes=('_CR', '_IPP'), how='outer')
    
    merged["Start Time_CR"] = pd.to_datetime(merged["Start Time_CR"], format='%I:%M%p').dt.strftime('%H:%M')
    merged["End Time_CR"] = pd.to_datetime(merged["End Time_CR"], format='%I:%M%p').dt.strftime('%H:%M')
    merged["Start Time_IPP"] = pd.to_datetime(merged["Start Time_IPP"], format='%I:%M%p').dt.strftime('%H:%M')
    merged["End Time_IPP"] = pd.to_datetime(merged["End Time_IPP"], format='%I:%M%p').dt.strftime('%H:%M')
    
    overlap_condition = (
        (merged["Start Time_CR"] < merged["End Time_IPP"]) & (merged["Start Time_CR"] > merged["Start Time_IPP"]) |
        (merged["End Time_CR"] < merged["End Time_IPP"]) & (merged["End Time_CR"] > merged["Start Time_IPP"]) | 
        (merged["Start Time_IPP"] < merged["End Time_CR"]) & (merged["Start Time_IPP"] > merged["Start Time_CR"]) |
        (merged["End Time_IPP"] < merged["End Time_CR"]) & (merged["End Time_IPP"] > merged["Start Time_CR"])
    )

    overlap_df = merged[overlap_condition]
    
    overlap_df["Start Time_CR"] = pd.to_datetime(overlap_df["Start Time_CR"], format='%H:%M').dt.strftime('%I:%M%p').astype('object')
    overlap_df["End Time_CR"] = pd.to_datetime(overlap_df["End Time_CR"], format='%H:%M').dt.strftime('%I:%M%p').astype('object')
    overlap_df["Start Time_IPP"] = pd.to_datetime(overlap_df["Start Time_IPP"], format='%H:%M').dt.strftime('%I:%M%p').astype('object')
    overlap_df["End Time_IPP"] = pd.to_datetime(overlap_df["End Time_IPP"], format='%H:%M').dt.strftime('%I:%M%p').astype('object')

    overlap_df['Service Date'] = pd.to_datetime(overlap_df['Service Date']).dt.strftime('%m/%d/%Y')
    
    return overlap_df

def find_bsc_bcba_discrepancies(rsm_df, cr_df):
    # --- Convert times and calculate hours ---
    for df in [rsm_df, cr_df]:
        df['Start Time'] = pd.to_datetime(df['Start Time'], errors='coerce')
        df['End Time'] = pd.to_datetime(df['End Time'], errors='coerce')
        df['Hours'] = (df['End Time'] - df['Start Time']).dt.total_seconds() / 3600

    # --- Categorize services ---
    cr_df['ServiceCategory'] = cr_df['BillingDesc'].apply(categorize_service)
    rsm_df['ServiceCategory'] = rsm_df['Service Name'].apply(categorize_service)

    # --- Group and pivot RSM ---
    rsm_totals = (
        rsm_df
        .groupby(['Provider', 'Student Name', 'ServiceCategory', 'Status'], as_index=False)
        .agg({'Hours': 'sum', 'Service Name': 'count'})
        .rename(columns={'Service Name': 'Appointments'})
        .pivot_table(
            index=['Provider', 'Student Name', 'ServiceCategory'],
            columns='Status',
            values=['Hours', 'Appointments'],
            fill_value=0
        )
        .reset_index()
    )

    # Flatten MultiIndex columns
    rsm_totals.columns = ['_'.join(col).strip('_') for col in rsm_totals.columns.values]

    # Guarantee both status columns exist
    for col in ['Converted', 'Cancelled']:
        for metric in ['Hours', 'Appointments']:
            colname = f"{metric}_{col}"
            if colname not in rsm_totals.columns:
                rsm_totals[colname] = 0

    # Rename columns for consistency
    rsm_totals = rsm_totals.rename(columns={
        'Hours_Converted': 'TotalConverted_IPP',
        'Hours_Cancelled': 'TotalCancelled_IPP',
        'Appointments_Converted': 'AppointmentsConverted_IPP',
        'Appointments_Cancelled': 'AppointmentsCancelled_IPP'
    })

    # --- Group and pivot CR ---
    cr_totals = (
        cr_df
        .groupby(['Provider', 'Student Name', 'ServiceCategory', 'Status'], as_index=False)
        .agg({'Hours': 'sum', 'BillingDesc': 'count'})
        .rename(columns={'BillingDesc': 'Appointments'})
        .pivot_table(
            index=['Provider', 'Student Name', 'ServiceCategory'],
            columns='Status',
            values=['Hours', 'Appointments'],
            fill_value=0
        )
        .reset_index()
    )

    # Flatten MultiIndex columns
    cr_totals.columns = ['_'.join(col).strip('_') for col in cr_totals.columns.values]

    # Guarantee both status columns exist
    for col in ['Converted', 'Cancelled']:
        for metric in ['Hours', 'Appointments']:
            colname = f"{metric}_{col}"
            if colname not in cr_totals.columns:
                cr_totals[colname] = 0

    # Rename columns for consistency
    cr_totals = cr_totals.rename(columns={
        'Hours_Converted': 'TotalConverted_CR',
        'Hours_Cancelled': 'TotalCancelled_CR',
        'Appointments_Converted': 'AppointmentsConverted_CR',
        'Appointments_Cancelled': 'AppointmentsCancelled_CR'
    })

    # --- Merge CR + RSM totals ---
    discrepancy_df = pd.merge(
        cr_totals, rsm_totals,
        on=['Provider', 'Student Name', 'ServiceCategory'],
        how='outer'
    ).fillna(0)

    # --- Add metadata (BillingDesc + Service Name) ---
    cr_meta = (
        cr_df.groupby(['Provider', 'Student Name', 'ServiceCategory'], as_index=False)
        [['BillingDesc']].first()
    )
    rsm_meta = (
        rsm_df.groupby(['Provider', 'Student Name', 'ServiceCategory'], as_index=False)
        [['Service Name']].first()
    )

    discrepancy_df = (
        discrepancy_df
        .merge(cr_meta, on=['Provider', 'Student Name', 'ServiceCategory'], how='left')
        .merge(rsm_meta, on=['Provider', 'Student Name', 'ServiceCategory'], how='left')
    )

    # --- Compute total-hour and appointment columns for comparison ---
    discrepancy_df['TotalHours_CR'] = discrepancy_df['TotalConverted_CR'] + discrepancy_df['TotalCancelled_CR']
    discrepancy_df['TotalHours_IPP'] = discrepancy_df['TotalConverted_IPP'] + discrepancy_df['TotalCancelled_IPP']
    discrepancy_df['TotalAppointments_CR'] = discrepancy_df['AppointmentsConverted_CR'] + discrepancy_df['AppointmentsCancelled_CR']
    discrepancy_df['TotalAppointments_IPP'] = discrepancy_df['AppointmentsConverted_IPP'] + discrepancy_df['AppointmentsCancelled_IPP']

    # --- Filter only mismatched totals (rounded to 2 decimals) ---
    discrepancy_df = discrepancy_df[
        discrepancy_df['TotalHours_CR'].round(2) != discrepancy_df['TotalHours_IPP'].round(2)
    ]

    # --- Reorder columns for clarity ---
    cols = [
        'Provider', 'Student Name', 'ServiceCategory',
        'BillingDesc', 'Service Name',
        'TotalConverted_CR', 'TotalCancelled_CR', 'TotalHours_CR',
        'TotalConverted_IPP', 'TotalCancelled_IPP', 'TotalHours_IPP',
        'TotalAppointments_CR', 'TotalAppointments_IPP'
    ]
    discrepancy_df = discrepancy_df[cols]

    return discrepancy_df

def find_status_diffs(cr_copy, rsm_copy, missing_from):
    #missing_from['StartTime'] = missing_from['StartTime_CR']
    merged = pd.merge(cr_copy, rsm_copy, on=['Provider', 'Student Name', 'ID Number', 'Service Date', 'Start Time', 'End Time'], how='inner', suffixes=('_CR', '_IPP'))
    status_diffs = merged[merged['Status_CR'] != merged['Status_IPP']]
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