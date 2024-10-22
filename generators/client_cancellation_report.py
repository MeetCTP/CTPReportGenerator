import pandas as pd
import numpy as np
from sqlalchemy import create_engine
import pymssql
import openpyxl
from flask import jsonify
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import sys
import os
import io

def generate_client_cancel_report(provider, client, cancel_reasons, range_start):
    try:
        user_name = os.getlogin()
        connection_string = f"mssql+pymssql://MeetCTP\Administrator:$Unlock01@CTP-DB/CRDB2"
        engine = create_engine(connection_string)
        now = datetime.now()
        curr_range_start = now - timedelta(days=now.day - 1)
        curr_range_end = now.date()
        past_range_start = datetime.strptime(range_start, '%Y-%m-%d')
        prev_range_start = past_range_start + relativedelta(months=1)
        past_range_end = prev_range_start - timedelta(days=1)
        prev_range_end = curr_range_start - timedelta(days=1)
        query_past_month = f"""
            SELECT DISTINCT * FROM ClientCancellationView
            WHERE (CONVERT(DATE, ServiceDate, 101) BETWEEN '{past_range_start.strftime('%Y-%m-%d')}' AND '{past_range_end.strftime('%Y-%m-%d')}') 
            AND (CancelledReason IN ({', '.join([f"'{s}'" for s in cancel_reasons])}))
        """
        if provider:
            query_past_month += f" AND Provider = '{provider}'"
        if client:
            query_past_month += f" AND Client = '{client}'"

        # Query for the previous month data
        query_prev_month = f"""
            SELECT DISTINCT * FROM ClientCancellationView
            WHERE (CONVERT(DATE, ServiceDate, 101) BETWEEN '{prev_range_start.strftime('%Y-%m-%d')}' AND '{prev_range_end.strftime('%Y-%m-%d')}') 
            AND (CancelledReason IN ({', '.join([f"'{s}'" for s in cancel_reasons])}))
        """
        if provider:
            query_prev_month += f" AND Provider = '{provider}'"
        if client:
            query_prev_month += f" AND Client = '{client}'"

        # Query for the current month data
        query_curr_month = f"""
            SELECT DISTINCT * FROM ClientCancellationView
            WHERE (CONVERT(DATE, ServiceDate, 101) BETWEEN '{curr_range_start.strftime('%Y-%m-%d')}' AND '{curr_range_end.strftime('%Y-%m-%d')}') 
            AND (CancelledReason IN ({', '.join([f"'{s}'" for s in cancel_reasons])}))
        """
        if provider:
            query_curr_month += f" AND Provider = '{provider}'"
        if client:
            query_curr_month += f" AND Client = '{client}'"

        all_past_query = f"""
            SELECT * from ClientCancellationView
            WHERE (CONVERT(DATE, ServiceDate, 101) BETWEEN '{past_range_start.strftime('%Y-%m-%d')}' AND '{past_range_end.strftime('%Y-%m-%d')}');
        """

        all_prev_query = f"""
            SELECT * from ClientCancellationView
            WHERE (CONVERT(DATE, ServiceDate, 101) BETWEEN '{prev_range_start.strftime('%Y-%m-%d')}' AND '{prev_range_end.strftime('%Y-%m-%d')}');
        """

        all_curr_query = f"""
            SELECT * from ClientCancellationView
            WHERE (CONVERT(DATE, ServiceDate, 101) BETWEEN '{curr_range_start.strftime('%Y-%m-%d')}' AND '{curr_range_end.strftime('%Y-%m-%d')}');
        """

        # Fetch data for each period
        data_past_month = pd.read_sql_query(query_past_month, engine)
        data_prev_month = pd.read_sql_query(query_prev_month, engine)
        data_curr_month = pd.read_sql_query(query_curr_month, engine)
        all_past_data = pd.read_sql_query(all_past_query, engine)
        all_prev_data = pd.read_sql_query(all_prev_query, engine)
        all_curr_data = pd.read_sql_query(all_curr_query, engine)

        # Calculate ThreeCancels and CancellationPercentage for each period
        three_cancels_past = check_three_cancels_in_a_row(all_past_data, cancel_reasons)
        cancel_percentage_past = calculate_cancellation_percentage(all_past_data, cancel_reasons)
        
        three_cancels_prev = check_three_cancels_in_a_row(all_prev_data, cancel_reasons)
        cancel_percentage_prev = calculate_cancellation_percentage(all_prev_data, cancel_reasons)
        
        three_cancels_curr = check_three_cancels_in_a_row(all_curr_data, cancel_reasons)
        cancel_percentage_curr = calculate_cancellation_percentage(all_curr_data, cancel_reasons)

        # Merge data for final report
        combined_data = pd.concat([data_past_month, data_prev_month, data_curr_month])

        # Assign the calculated values to the final data
        combined_data['ThreeCancels_FirstMonth'] = combined_data['Client'].map(three_cancels_past)
        combined_data['CancellationPercentage_FirstMonth'] = combined_data['Client'].map(cancel_percentage_past)
        
        combined_data['ThreeCancels_LastMonth'] = combined_data['Client'].map(three_cancels_prev)
        combined_data['CancellationPercentage_LastMonth'] = combined_data['Client'].map(cancel_percentage_prev)
        
        combined_data['ThreeCancels_CurrentMonth'] = combined_data['Client'].map(three_cancels_curr)
        combined_data['CancellationPercentage_CurrentMonth'] = combined_data['Client'].map(cancel_percentage_curr)

        combined_data.drop_duplicates(inplace=True)
        combined_data.sort_values(by=['Client', 'AppStart'])

        # Output the final data to an Excel file
        output_file = io.BytesIO()
        combined_data.to_excel(output_file, index=False)
        output_file.seek(0)

        return output_file
    
    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

    finally:
        engine.dispose()

def check_three_cancels_in_a_row(data, cancel_reasons):
    current_client = None
    cancel_count = 0
    result = {}

    for index, row in data.iterrows():
        client = row['Client']
        reason = row['CancelledReason']

        if client != current_client:
            current_client = client
            cancel_count = 0

        if reason in cancel_reasons:
            cancel_count += 1
            if cancel_count == 3:
                result[client] = True
        else:
            cancel_count = 0

    for client in data['Client'].unique():
        if client not in result:
            result[client] = False

    return result

def calculate_cancellation_percentage(data, cancel_reasons):
    client_sessions = {}
    client_cancellations = {}

    for index, row in data.iterrows():
        client = row['Client']
        reason = row['CancelledReason']

        if client not in client_sessions:
            client_sessions[client] = 0
            client_cancellations[client] = 0

        client_sessions[client] += 1

        if reason in cancel_reasons:
            client_cancellations[client] += 1

    cancellation_percentages = {}

    for client in client_sessions:
        total_sessions = client_sessions[client]
        cancellations = client_cancellations[client]
        cancellation_percentages[client] = (cancellations / total_sessions) * 100

    return cancellation_percentages