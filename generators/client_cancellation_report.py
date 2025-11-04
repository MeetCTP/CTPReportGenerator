import pandas as pd
from pandas import ExcelWriter
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
        # Data collection per month
        raw_combined_data = pd.DataFrame()
        three_cancels_combined = pd.DataFrame()
        percentage_combined = pd.DataFrame()

        for start, end in month_ranges:
            month_str = start.strftime('%Y-%m')
            # Pull data for that month
            query = f"""
                SELECT DISTINCT
                    Provider,
                    Client,
                    School,
                    ServiceDate,
                    AppStart,
                    AppEnd,
                    CancelledReason
                FROM ClientCancellationView
                WHERE (CONVERT(DATE, ServiceDate, 101) BETWEEN '{start.strftime('%Y-%m-%d')}' AND '{end.strftime('%Y-%m-%d')}')
                AND (CancelledReason IN ({', '.join([f"'{s}'" for s in cancel_reasons])}))
            """
            if provider:
                query += f" AND Provider = '{provider}'"
            if client:
                query += f" AND Client = '{client}'"

            data = pd.read_sql_query(query, engine)

            # Pull all data within full range for calculations
            all_data = pd.read_sql_query(f"""
                SELECT DISTINCT
                    Provider,
                    Client,
                    School,
                    ServiceDate,
                    AppStart,
                    AppEnd,
                    CancelledReason
                FROM ClientCancellationView
                WHERE (CONVERT(DATE, ServiceDate, 101) BETWEEN '{month_ranges[0][0].strftime('%Y-%m-%d')}' AND '{month_ranges[-1][1].strftime('%Y-%m-%d')}')
            """, engine).sort_values(by='ServiceDate', ascending=True)

            # Compute once for all data (only once outside loop)
            if 'cancel_percentage' not in locals():
                cancel_percentage = calculate_cancellation_percentage(all_data, cancel_reasons)
                three_cancels = check_three_cancels_in_a_row(all_data, cancel_reasons)

            # Monthly data
            raw_combined_data = pd.concat([raw_combined_data, data])

            # Build Three Cancels sheet (subset of clients who met the 3-in-a-row condition)
            rows = []
            for (client_key, provider_key), cancel_rows in three_cancels.items():
                for cancel_row in cancel_rows:
                    rows.append(cancel_row)

            if rows:
                month_three_cancels_df = pd.DataFrame(rows)
                three_cancels_combined = pd.concat([three_cancels_combined, month_three_cancels_df])

            # Build Percentage sheet — include *all* clients from all_data, not just monthly data
            month_percentage_df = all_data.copy()
            month_percentage_df[f'CancellationPercentage_{month_str}'] = month_percentage_df['Client'].map(cancel_percentage)

            percentage_combined = pd.concat([percentage_combined, month_percentage_df])

        # Final formatting
        for df in [raw_combined_data, three_cancels_combined, percentage_combined]:
            df.drop_duplicates(inplace=True)
            df.sort_values(by=['Provider', 'AppStart'], ascending=[True, True], inplace=True)
            if 'AppStart' in df.columns and pd.api.types.is_datetime64_any_dtype(df['AppStart']):
                df['AppStart'] = df['AppStart'].dt.strftime('%m/%d/%Y %I:%M %p')

        # Output to Excel
        output_file = io.BytesIO()
        with ExcelWriter(output_file, engine='openpyxl') as writer:
            raw_combined_data.to_excel(writer, sheet_name='Raw', index=False)
            three_cancels_combined.to_excel(writer, sheet_name='ThreeCancels', index=False)
            percentage_combined.to_excel(writer, sheet_name='Percentage', index=False)
        output_file.seek(0)

        return output_file

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

    finally:
        engine.dispose()

def check_three_cancels_in_a_row(data, cancel_reasons):
    results = {}

    # Sort by Client, Provider, then ServiceDate
    data = data.sort_values(by=['Client', 'Provider', 'AppStart'])

    current_key = None
    sequence = []

    for _, row in data.iterrows():
        client = row['Client']
        provider = row['Provider']
        reason = row['CancelledReason']
        key = (client, provider)

        if key != current_key:
            current_key = key
            sequence = []

        if reason in cancel_reasons:
            sequence.append(row)
            if len(sequence) == 3:
                # We found three in a row — store them and stop for this key
                results[key] = sequence.copy()
        else:
            sequence = []

    return results

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