import pandas as pd
from sqlalchemy import create_engine
from datetime import datetime
import pymssql
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os

def generate_util_tracker(file_path):
    user_name = os.getlogin()
    documents_path = f"C:/Users/{user_name}/Documents/"
    connection_string = f"mssql+pymssql://MeetCTP\Joshua.Bliven:$Unlock03@CTP-DB/CRDB2"
    engine = create_engine(connection_string)
    
    dfs = pd.read_excel(file_path, sheet_name=None)
    processed_dfs = {}
    all_dates = []
    for sheet_name, df in dfs.items():
        processed_df = df
        processed_dfs[sheet_name] = processed_df
        if 'Date' in df.columns:
            all_dates.extend(df['Date'].dropna().tolist())
            
    if all_dates:
        start_date = min(all_dates)
        end_date = max(all_dates)
    else:
        raise ValueError("No valid dates found in the provided sheets.")
    
    start_date_str = datetime.strftime(start_date, '%Y-%m-%d')
    end_date_str = datetime.strftime(end_date, '%Y-%m-%d')
    
    util_query = f"""
        SELECT
            *
        FROM ClinicalUtilizationTracker
        WHERE CONVERT(DATE, ServiceDate, 101) BETWEEN '{start_date_str}' AND '{end_date_str}'
    """
    util_data = pd.read_sql_query(util_query, engine)
    
    util_data['ServiceDate'] = pd.to_datetime(util_data['ServiceDate']).dt.date
    util_data['AppHours'] = pd.to_numeric(util_data['AppHours'], errors='coerce')
    
    weekly_columns = [
        'Principal1Name', 'Principal2Name', 'Authorized Hours', 'Direct Hours', 'Indirect Time',
        'Progress Reports', 'Make Up Time', 'Meetings/other indirect', 'Completed Hours',
        'Late Cancels', 'Advanced Cancels', 'Teacher Cancels', 'Provider Cancels',
        'Services on Hold', 'School Closed'
    ]
    weekly_df = pd.DataFrame(columns=weekly_columns)

    merged_dfs = {}
    for sheet_name, df in processed_dfs.items():
        # Ensure the date column is in the correct format
        df['Date'] = pd.to_datetime(df['Date']).dt.date
        df['EventHours'] = pd.to_numeric(df['EventHours'], errors='coerce')
        # Perform the merge
        merged_df = pd.merge(df, util_data, how='left', 
                                left_on=['Date', 'Principal1Name', 'Principal2Name', 'EventHours'],
                                right_on=['ServiceDate', 'provider', 'Client', 'AppHours'])
        merged_dfs[sheet_name] = merged_df
        
    weekly_df
        
    output_file_path = documents_path + 'Formatted_Data.xlsx'
        
    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
        for sheet_name, merged_df in merged_dfs.items():
            merged_df.to_excel(writer, sheet_name=sheet_name)
            
    wb = load_workbook(output_file_path)
    for sheet_name in merged_dfs.keys():
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=2, min_col=2, max_col=ws.max_column):
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0.00'
                    
    for sheet in wb.sheetnames:
        wb[sheet].sheet_state = 'visible'
            
    wb.save(output_file_path)