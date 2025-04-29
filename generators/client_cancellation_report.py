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

def generate_client_cancel_report(provider, client, cancel_reasons, start_date, end_date):
    try:
        db_user = os.getenv("DB_USER")
        db_pw = os.getenv("DB_PW")
        connection_string = f"mssql+pymssql://{db_user}:{db_pw}@CTP-DB/CRDB2"
        engine = create_engine(connection_string)

        # Parse start_date and end_date
        start_date = datetime.strptime(start_date, '%Y-%m-%d')
        end_date = datetime.strptime(end_date, '%Y-%m-%d')

        # Generate month ranges
        month_ranges = []
        current_start = start_date.replace(day=1)
        while current_start <= end_date:
            next_month_start = (current_start + relativedelta(months=1)).replace(day=1)
            current_end = min(next_month_start - timedelta(days=1), end_date)
            month_ranges.append((current_start, current_end))
            current_start = next_month_start

        # Data collection for each month
        monthly_data = {}
        for start, end in month_ranges:
            query = f"""
                SELECT DISTINCT * FROM ClientCancellationView
                WHERE (CONVERT(DATE, ServiceDate, 101) BETWEEN '{start.strftime('%Y-%m-%d')}' AND '{end.strftime('%Y-%m-%d')}') 
                AND (CancelledReason IN ({', '.join([f"'{s}'" for s in cancel_reasons])}))
            """
            if provider:
                query += f" AND Provider = '{provider}'"
            if client:
                query += f" AND Client = '{client}'"
            monthly_data[f"{start.strftime('%Y-%m')}"] = pd.read_sql_query(query, engine)

        # Aggregate calculations
        combined_data = pd.DataFrame()
        for month, data in monthly_data.items():
            all_data = pd.read_sql_query(f"""
                SELECT * FROM ClientCancellationView
                WHERE (CONVERT(DATE, ServiceDate, 101) BETWEEN '{month_ranges[0][0].strftime('%Y-%m-%d')}' AND '{month_ranges[-1][1].strftime('%Y-%m-%d')}')
            """, engine)

            all_data = all_data.sort_values(by='ServiceDate', ascending=True)

            three_cancels = check_three_cancels_in_a_row(all_data, cancel_reasons)
            cancel_percentage = calculate_cancellation_percentage(all_data, cancel_reasons)

            data[f'ThreeCancels_{month}'] = data['Client'].map(three_cancels)
            data[f'CancellationPercentage_{month}'] = data['Client'].map(cancel_percentage)

            combined_data = pd.concat([combined_data, data])

        combined_data.drop_duplicates(inplace=True)
        combined_data.sort_values(by=['Client', 'AppStart'], ascending=[True, True])
        combined_data['AppStart'] = combined_data['AppStart'].dt.strftime('%m/%d/%Y %I:%M %p')

        # Output to Excel
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