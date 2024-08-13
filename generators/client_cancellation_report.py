import pandas as pd
import numpy as np
import dask.dataframe as dd
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
        query_all_periods = f"""
            SELECT *, 
                CASE 
                    WHEN CONVERT(DATE, ServiceDate, 101) BETWEEN '{past_range_start}' AND '{past_range_end}' THEN 'PastMonth' 
                    WHEN CONVERT(DATE, ServiceDate, 101) BETWEEN '{prev_range_start}' AND '{prev_range_end}' THEN 'PrevMonth' 
                    ELSE 'CurrMonth' 
                END AS Period 
            FROM ClientCancellationView 
            WHERE CONVERT(DATE, ServiceDate, 101) BETWEEN '{past_range_start}' AND '{curr_range_end}' AND CancelledReason IN ({', '.join([f"'{c}'" for c in cancel_reasons])})
        """
        if provider:
            query_all_periods += f" AND Provider = '{provider}'"
        if client:
            query_all_periods += f" AND Client = '{client}'"
        query_all_periods += " ORDER BY Client, AppStart, BillingCode;"

        data_all_periods = pd.read_sql_query(query_all_periods, engine)

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
        data_past_month = data_all_periods[data_all_periods['Period'] == 'PastMonth']
        data_prev_month = data_all_periods[data_all_periods['Period'] == 'PrevMonth']
        data_curr_month = data_all_periods[data_all_periods['Period'] == 'CurrMonth']
        all_past_data = pd.read_sql_query(all_past_query, engine)
        all_prev_data = pd.read_sql_query(all_prev_query, engine)
        all_curr_data = pd.read_sql_query(all_curr_query, engine)

        #Convert pandas dataframes to dask dataframes for processing
        all_past_data_dd = dd.from_pandas(all_past_data, npartitions=4)
        all_prev_data_dd = dd.from_pandas(all_prev_data, npartitions=4)
        all_curr_data_dd = dd.from_pandas(all_curr_data, npartitions=4)

        # Merge data for final report
        combined_data = pd.concat([data_past_month, data_prev_month, data_curr_month])

        # Calculate ThreeCancels and CancellationPercentage for each period
        three_cancels_past = check_three_cancels_in_a_row(all_past_data)
        cancel_percentage_past = calculate_cancellation_percentage(all_past_data)
        
        three_cancels_prev = check_three_cancels_in_a_row(all_prev_data)
        cancel_percentage_prev = calculate_cancellation_percentage(all_prev_data)
        
        three_cancels_curr = check_three_cancels_in_a_row(all_curr_data)
        cancel_percentage_curr = calculate_cancellation_percentage(all_curr_data)

        # Assign the calculated values to the final data
        combined_data['ThreeCancels_FirstMonth'] = combined_data['Client'].map(three_cancels_past)
        combined_data['CancellationPercentage_LastMonth'] = combined_data['Client'].map(cancel_percentage_past)
        
        combined_data['ThreeCancels_LastMonth'] = combined_data['Client'].map(three_cancels_prev)
        combined_data['CancellationPercentage_LastMonth'] = combined_data['Client'].map(cancel_percentage_prev)
        
        combined_data['ThreeCancels_CurrentMonth'] = combined_data['Client'].map(three_cancels_curr)
        combined_data['CancellationPercentage_CurrentMonth'] = combined_data['Client'].map(cancel_percentage_curr)

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

@njit
def check_three_cancels_in_a_row(data):
    result = np.zeros(len(data), dtype=np.int32)
    for i in range(2, len(data)):
        if data[i] == 1 and data[i-1] == 1 and data[i-2] == 1:
            result[i] = 1
    return result

@njit
def calculate_cancellation_percentage(data):
    cancel_count = data.sum()
    return (cancel_count / len(data)) * 100