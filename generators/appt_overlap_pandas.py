import pandas as pd
from sqlalchemy import create_engine
import openpyxl
import pymssql
from datetime import datetime as dt
import io
import os

def generate_appt_overlap_report(start_date, end_date, provider, client):
    try:
        user_name = os.getlogin()
        connection_string = f"mssql+pymssql://MeetCTP\Administrator:$Unlock01@CTP-DB/CRDB2"
        engine = create_engine(connection_string)
        
        query = f"""
            SELECT DISTINCT *
            FROM ApptOverlapView
            WHERE (ServiceDate BETWEEN '{start_date}' AND '{end_date}')
        """
        if provider:
            query += f" AND (Provider = '{provider}')"
        if client:
            query += f" AND (Client = '{client}')"
            
        data = pd.read_sql_query(query, engine)

        data = data.sort_values(by=['Provider', 'ServiceDate'], ascending=True)

        data['ServiceDate'] = pd.to_datetime(data['ServiceDate']).dt.strftime('%m/%d/%Y')
        data['StartTime'] = data['StartTime'].dt.strftime('%I:%M%p')
        data['EndTime'] = data['EndTime'].dt.strftime('%I:%M%p')

        data_copy = data.copy()

        data = find_overlapping_appointments(data, data_copy)

        data.drop_duplicates(inplace=True)
        
        output_file = io.BytesIO()
        data.to_excel(output_file, index=False)
        output_file.seek(0)
        return output_file

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

    finally:
        engine.dispose()
        
def find_overlapping_appointments(data, data_copy):
    # Merge the two DataFrames on the appropriate columns
    merged = pd.merge(data, data_copy, on=["Provider", "ServiceDate"], suffixes=('_1', '_2'), how='outer')

    # Convert StartTime and EndTime columns to datetime format for comparison
    merged["StartTime_1"] = pd.to_datetime(merged["StartTime_1"], format='%I:%M%p').dt.strftime('%H:%M')
    merged["EndTime_1"] = pd.to_datetime(merged["EndTime_1"], format='%I:%M%p').dt.strftime('%H:%M')
    merged["StartTime_2"] = pd.to_datetime(merged["StartTime_2"], format='%I:%M%p').dt.strftime('%H:%M')
    merged["EndTime_2"] = pd.to_datetime(merged["EndTime_2"], format='%I:%M%p').dt.strftime('%H:%M')

    # Define the overlap condition
    overlap_condition = (
        (merged["StartTime_1"] < merged["EndTime_2"]) & (merged["StartTime_1"] > merged["StartTime_2"]) |
        (merged["EndTime_1"] > merged["StartTime_2"]) & (merged["EndTime_1"] < merged["EndTime_2"]) |
        (merged["StartTime_2"] < merged["EndTime_1"]) & (merged["StartTime_2"] > merged["StartTime_1"]) |
        (merged["EndTime_2"] > merged["StartTime_1"]) & (merged["EndTime_2"] < merged["EndTime_1"]) |
        (merged["StartTime_1"] == merged["StartTime_2"]) & (merged["EndTime_1"] == merged["EndTime_2"]) & (merged["Client_1"] != merged["Client_2"])
    )

    # Filter the merged DataFrame based on the overlap condition
    overlap_df = merged[overlap_condition]

    # Convert the StartTime and EndTime columns back to the original format
    overlap_df["StartTime_1"] = pd.to_datetime(overlap_df["StartTime_1"], format='%H:%M').dt.strftime('%I:%M%p').astype('object')
    overlap_df["EndTime_1"] = pd.to_datetime(overlap_df["EndTime_1"], format='%H:%M').dt.strftime('%I:%M%p').astype('object')
    overlap_df["StartTime_2"] = pd.to_datetime(overlap_df["StartTime_2"], format='%H:%M').dt.strftime('%I:%M%p').astype('object')
    overlap_df["EndTime_2"] = pd.to_datetime(overlap_df["EndTime_2"], format='%H:%M').dt.strftime('%I:%M%p').astype('object')

    overlap_df['Client_pair'] = overlap_df.apply(lambda row: tuple(sorted([row['Client_1'], row['Client_2']])), axis=1)

    # Drop mirrored duplicates based on the sorted Client pair column
    overlap_df = overlap_df.drop_duplicates(subset=['Client_pair', 'ServiceDate', 'Provider', 'StartTime_1', 'EndTime_1', 'StartTime_2', 'EndTime_2'])

    # Drop the 'Client_pair' column, as it's no longer needed
    overlap_df = overlap_df.drop(columns=['Client_pair'])

    # Optionally, you can drop the extra columns if you don't need them in the final output
    #overlap_df = overlap_df.drop(columns=['StartTime_1', 'EndTime_1', 'StartTime_2', 'EndTime_2'])

    return overlap_df