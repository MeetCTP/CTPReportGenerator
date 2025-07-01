import pandas as pd
import numpy as np
from sqlalchemy import create_engine
from datetime import datetime, timedelta
import pymssql
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from pandas import ExcelWriter
import os
import io
import re

service_keywords = {
    'BCBA F2F': [r'Board Certified Behavior Analyst.*Face to Face.*'],
    'BCBA V': [r'Board Certified Behavior Analyst.*Virtual.*'],
    'BSC F2F': [r'Behavior Specialist Consultant.*Face to Face.*'],
    'BSC V': [r'Behavior Specialist Consultant.*Virtual.*'],
    'LC F2F': [r'Learning Coach.*Face to Face.*'],
    'LC V': [r'Learning Coach.*Virtual.*'],
    'Social Skills V': [r'Social Skills.*Virtual.*'],
    'Social Skills F2F': [r'Social Skills.*Face to Face.*'],
    'Para F2F': [r'Personal Care Assistant.*Face to Face.*', r'Instructional Aide.*Face to Face.*'],
    'Para V': [r'Personal Care Assistant.*Virtual.*', r'Instructional Aide.*Virtual.*'],
    'Sped Teacher (CSET) F2F': [r'Certified Special Education Teacher.*Face to Face.*'],
    'Sped Teacher (CSET) V': [r'Certified Special Education Teacher.*Virtual.*'],
    'Counseling F2F': [r'Certified Counselor.*Face to Face.*', r'Licensed Counselor.*Face to Face.*'],
    'Counseling V': [r'Certified Counselor.*Virtual.*', r'Licensed Counselor.*Virtual.*'],
    'RBT F2F': [r'Personal Care Assistant.*Face to Face.*', r'Instructional Aide.*Face to Face.*'],
    'RBT V': [r'Personal Care Assistant.*Virtual.*', r'Instructional Aide.*Virtual.*'],
    'Tutor F2F': [r'Tutor.*Face to Face.*'],
    'Tutor V': [r'Tutor.*Virtual.*'],
    'Print Materials': [r'.*Print.*Materials.*'],
    'Social Work F2F': [r'Social Worker.*Face to Face.*'],
    'Social Work V': [r'Social Worker.*Virtual.*'],
    'Wilson Reading F2F': [r'Wilson Instructor.*Face to Face.*'],
    'Wilson Reading V': [r'Wilson Instructor.*Virtual.*'],
    'Life Skills Group Class': [r'.*Life Skills.*Group.*'],
    'Classroom IA V': [r'Instructional Aide.*Group.*Virtual.*'],
    'Classroom SEL Support V': [r'SEL.*Support.*Virtual.*'],
    'Speech': [r'Speech Therapist.*(Face to Face|Virtual)'],
    'Progress Reports': [r'.*Progress Report.*'],
    'IEP Meetings': [r'.*IEP.*'],
    'Indirect Services': [r'.*Indirect.*'],
}

def get_week_bounds(date):
    start = date - timedelta(days=date.weekday())  # Monday
    end = start + timedelta(days=6)                # Sunday
    return start, end

def label_service(description):
    if pd.isna(description):
        return "Uncategorized"
    for label, patterns in service_keywords.items():
        for pattern in patterns:
            if re.search(pattern, description, re.IGNORECASE):
                return label
    return "Uncategorized"

def generate_school_util_report(start_date, end_date):
    try:
        db_user = os.getenv("DB_USER")
        db_pw = os.getenv("DB_PW")
        connection_string = f"mssql+pymssql://{db_user}:{db_pw}@CTP-DB/CRDB2"
        engine = create_engine(connection_string)

        # Parse start_date and end_date
        start_date = datetime.strptime(start_date, '%Y-%m-%d')
        end_date = datetime.strptime(end_date, '%Y-%m-%d')

        # Snap start_date to Monday of that week
        start_of_current_week = start_date - timedelta(days=start_date.weekday())

        # Define previous weeks
        start_of_previous_week = start_of_current_week - timedelta(weeks=1)
        start_of_two_weeks_ago = start_of_current_week - timedelta(weeks=2)

        # Define ends (Sunday of each week)
        end_of_current_week = start_of_current_week + timedelta(days=6)
        end_of_previous_week = start_of_previous_week + timedelta(days=6)
        end_of_two_weeks_ago = start_of_two_weeks_ago + timedelta(days=6)

        query_start = start_of_two_weeks_ago.strftime('%Y-%m-%d')
        query_end = end_of_current_week.strftime('%Y-%m-%d')

        query = f"""
            SELECT *
            FROM SchoolUtilization
            WHERE ServiceDate BETWEEN '{query_start}' AND '{query_end}'
        """
        data = pd.read_sql_query(query, engine)

        data.drop_duplicates(inplace=True)

        def get_week_label(date):
            if start_of_two_weeks_ago <= date <= end_of_two_weeks_ago:
                return f"Two Weeks Ago ({start_of_two_weeks_ago.strftime('%m/%d')}-{end_of_two_weeks_ago.strftime('%m/%d')})"
            elif start_of_previous_week <= date <= end_of_previous_week:
                return f"Previous Week ({start_of_previous_week.strftime('%m/%d')}-{end_of_previous_week.strftime('%m/%d')})"
            elif start_of_current_week <= date <= end_of_current_week:
                return f"Current Week ({start_of_current_week.strftime('%m/%d')}-{end_of_current_week.strftime('%m/%d')})"
            else:
                return "Outside Range"

        data["ServiceDate"] = pd.to_datetime(data["ServiceDate"])
        data["Week"] = data["ServiceDate"].apply(get_week_label)

        client_counts = (
            data.groupby(['School', 'Week'])['Client']
            .nunique()
            .reset_index()
            .pivot(index='School', columns='Week', values='Client')
            .fillna(0)
            .astype(int)
            .reset_index()
        )

        # Optional: Ensure columns are ordered consistently
        desired_order = [
            f"Current Week ({start_of_current_week.strftime('%m/%d')}-{end_of_current_week.strftime('%m/%d')})",
            f"Previous Week ({start_of_previous_week.strftime('%m/%d')}-{end_of_previous_week.strftime('%m/%d')})",
            f"Two Weeks Ago ({start_of_two_weeks_ago.strftime('%m/%d')}-{end_of_two_weeks_ago.strftime('%m/%d')})"
        ]
        columns_in_order = ['School'] + [col for col in desired_order if col in client_counts.columns]
        client_counts = client_counts[columns_in_order]

        event_hours = (
            data.groupby(['School', 'Week'])['EventHours']
            .sum()
            .reset_index()
            .pivot(index='School', columns='Week', values='EventHours')
            .fillna(0)
            .round(2)
            .reset_index()
        )

        columns_in_order_hours = ['School'] + [col for col in desired_order if col in event_hours.columns]
        event_hours = event_hours[columns_in_order_hours]

        data['ServiceCategory'] = data['ServiceCodeDescription'].apply(label_service)

        service_hours = (
            data.groupby(['ServiceCategory', 'Week'])['EventHours']
            .sum()
            .reset_index()
            .pivot(index='ServiceCategory', columns='Week', values='EventHours')
            .fillna(0)
            .round(2)
            .reset_index()
        )

        columns_in_order_services = ['ServiceCategory'] + [col for col in desired_order if col in service_hours.columns]
        service_hours = service_hours[columns_in_order_services]

        output_file = io.BytesIO()
        with ExcelWriter(output_file, engine='openpyxl') as writer:
            data.to_excel(writer, sheet_name='Raw', index=False)
            client_counts.to_excel(writer, sheet_name='# of Clients', index=False)
            event_hours.to_excel(writer, sheet_name='# of Hours', index=False)
            service_hours.to_excel(writer, sheet_name='Hours by Service', index=False)

        output_file.seek(0)
        return output_file
    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e
    finally:
        engine.dispose()