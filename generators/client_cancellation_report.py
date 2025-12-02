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

def generate_client_cancel_report(provider, client, cancel_reasons, start_date, end_date, overrides):
    try:
        db_user = os.getenv("DB_USER")
        db_pw = os.getenv("DB_PW")
        connection_string = f"mssql+pymssql://{db_user}:{db_pw}@CTP-DB/CRDB2"
        engine = create_engine(connection_string)

        # Parse date inputs
        start_date = datetime.strptime(start_date, '%Y-%m-%d')
        end_date = datetime.strptime(end_date, '%Y-%m-%d')

        # ---------------------------------------------------------------------
        # 1) PULL ALL APPOINTMENTS — NO CANCELLATION FILTERING
        # ---------------------------------------------------------------------
        all_query = f"""
            SELECT DISTINCT
                Provider,
                Client,
                School,
                ServiceDate,
                AppStart,
                AppEnd,
                CancelledReason
            FROM ClientCancellationView
            WHERE CONVERT(DATE, ServiceDate, 101)
                  BETWEEN '{start_date.strftime('%Y-%m-%d')}'
                      AND '{end_date.strftime('%Y-%m-%d')}'
        """
        if provider:
            all_query += f" AND Provider = '{provider}'"
        if client:
            all_query += f" AND Client = '{client}'"

        all_data = pd.read_sql_query(all_query, engine)

        # Parse datetimes
        for col in ['ServiceDate', 'AppStart', 'AppEnd']:
            if col in all_data.columns:
                all_data[col] = pd.to_datetime(all_data[col], errors='coerce')

        # ---------------------------------------------------------------------
        # Apply school overrides
        # ---------------------------------------------------------------------
        all_data = apply_school_overrides(all_data, overrides)

        # ---------------------------------------------------------------------
        # 2) RAW SHEET = cancellations only (based on cancel_reasons)
        # ---------------------------------------------------------------------
        raw_combined_data = all_data[
            all_data['CancelledReason'].isin(cancel_reasons)
        ].copy()

        # ---------------------------------------------------------------------
        # 3) THREE CANCELS IN A ROW (using ALL appointments)
        # ---------------------------------------------------------------------
        three_cancels = check_three_cancels_in_a_row(all_data, cancel_reasons)

        rows = []
        for (_, _), cancel_rows in three_cancels.items():
            rows.extend(cancel_rows)

        three_cancels_combined = pd.DataFrame(rows) if rows else pd.DataFrame()

        # ---------------------------------------------------------------------
        # 4) CANCELLATION PERCENTAGE (using ALL appointments)
        # ---------------------------------------------------------------------
        cancel_percentage = calculate_cancellation_percentage(all_data, cancel_reasons)

        percentage_combined = (
            all_data[['Provider', 'Client', 'School']]
            .drop_duplicates()
            .sort_values(['Provider', 'Client', 'School'])
            .reset_index(drop=True)
        )
        percentage_combined['CancellationPercentage'] = \
            percentage_combined['Client'].map(cancel_percentage)

        # ---------------------------------------------------------------------
        # Final formatting for Raw and ThreeCancels
        # ---------------------------------------------------------------------
        for df in (raw_combined_data, three_cancels_combined):
            if df.empty:
                continue

            df.drop_duplicates(inplace=True)

            # Sort by Provider then Appt start
            if 'AppStart' in df.columns and pd.api.types.is_datetime64_any_dtype(df['AppStart']):
                df.sort_values(['Provider', 'AppStart'], inplace=True)
                df['AppStart'] = df['AppStart'].dt.strftime('%m/%d/%Y %I:%M %p')
            else:
                df.sort_values(['Provider'], inplace=True)

        # ---------------------------------------------------------------------
        # OUTPUT EXCEL
        # ---------------------------------------------------------------------
        output_file = io.BytesIO()
        with ExcelWriter(output_file, engine='openpyxl') as writer:
            raw_combined_data.to_excel(writer, sheet_name='Raw', index=False)
            three_cancels_combined.to_excel(writer, sheet_name='ThreeCancels', index=False)
            percentage_combined.to_excel(writer, sheet_name='Percentage', index=False)

        output_file.seek(0)
        return output_file

    except Exception as e:
        print("Error occurred while generating the report:", e)
        raise e

    finally:
        engine.dispose()

def check_three_cancels_in_a_row(data, cancel_reasons):
    results = {}

    data = data.sort_values(by=['Client', 'Provider', 'ServiceDate'])

    current_key = None
    sequence = []

    for _, row in data.iterrows():
        key = (row['Client'], row['Provider'])
        reason = row['CancelledReason']

        if key != current_key:
            current_key = key
            sequence = []

        if reason in cancel_reasons:
            sequence.append(row)
            if len(sequence) == 3:
                results[key] = sequence.copy()
        else:
            sequence = []

    return results

def calculate_cancellation_percentage(data, cancel_reasons):
    client_sessions = {}
    client_cancellations = {}

    for _, row in data.iterrows():
        client = row['Client']
        reason = row['CancelledReason']

        client_sessions[client] = client_sessions.get(client, 0) + 1

        if reason in cancel_reasons:
            client_cancellations[client] = client_cancellations.get(client, 0) + 1

    percentages = {}
    for client in client_sessions:
        total = client_sessions[client]
        cancels = client_cancellations.get(client, 0)
        percentages[client] = (cancels / total) * 100 if total > 0 else 0

    return percentages

def apply_school_overrides(df, overrides):
    """
    Removes rows where a student's school does NOT match the student's current school
    according to the override dictionary.
    """

    # Normalize old to always be a list
    normalized = {}
    for student, info in overrides.items():
        old_schools = info.get("old", [])
        if not isinstance(old_schools, list):
            old_schools = [old_schools]

        normalized[student] = {
            "current": info["current"],
            "old": old_schools,
        }

    # If a student is not in overrides, keep the row.
    # If they ARE in overrides, keep only the row where School == current school.
    def should_keep(row):
        student = row["Client"]          # <== important: uses the "Client" column from SchoolUtilization
        school = row["School"]

        if student not in normalized:
            return True  # No override, keep

        current = normalized[student]["current"]

        return school == current  # Keep only correct school row

    return df[df.apply(should_keep, axis=1)]