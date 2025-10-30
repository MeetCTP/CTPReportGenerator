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
    'RBT F2F': [r'Registered Behavior Technician.*Face to Face.*', r'Registered Behavior Technician.*Face to Face.*'],
    'RBT V': [r'Registered Behavior Technician.*Virtual.*', r'Registered Behavior Technician.*Virtual.*'],
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
    'Speech V': [r'Speech Therapist.*(Face to Face|Virtual)'],
    'Assessments': [r'.*Assessment.*'],
    'Indirect Services': [r'.*Indirect.*', r'.*IEP.*', r'.*Progress Report.*'],
    'BCBA-ABA': [r'.*BCBA-ABA.*'],
    'BCaBA-ABA': [r'.*BCaBA-ABA'],
    'BHT-ABA': [r'.*administered by technician.*', r'.*BHT-ABA.*'],
    'BC-ABA': [r'.*BC-ABA.*', r'.*Behavior Consultation.*'],
    'Mobile Therapy-ABA': [r'.*Mobile Therapy.*non ABA.*'],
    'BHT': [r'Behavioral Health Technician.*non.*', r'Behavioral Health Technician'],
    'BC': [r'.*Behavior Consultation.*non ABA.*'],
    'Mobile Therapy': [r'.*Mobile Therapy.*non ABA.*'],
}

aba_pairs = {
    'BHT-ABA': 'BHT',
    'BC-ABA': 'BC',
    'BCBA-ABA': 'BCBA',
    'BCaBA-ABA': 'BCaBA',
    'Mobile Therapy-ABA': 'Mobile Therapy'
}

def get_week_bounds(date):
    start = date - timedelta(days=date.weekday())  # Monday
    end = start + timedelta(days=6)                # Sunday
    return start, end

def label_service(row):
    description = row.get("ServiceCodeDescription", "")
    code = str(row.get("ServiceCode", ""))

    if pd.isna(description):
        return "Uncategorized"

    for label, patterns in service_keywords.items():
        for pattern in patterns:
            if re.search(pattern, description, re.IGNORECASE):

                # If this label has both ABA and non-ABA versions
                if label in aba_pairs:
                    # If 'H' in code, return the non-ABA version
                    if 'H' in code:
                        return aba_pairs[label]
                    else:
                        return label  # Keep ABA version
                elif label in aba_pairs.values():  # It's the non-ABA version
                    if 'H' not in code:
                        # Force ABA version if ServiceCode lacks 'H'
                        return [aba for aba, non_aba in aba_pairs.items() if non_aba == label][0]
                    else:
                        return label

                return label  # Everything else
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
        start_of_current_week = start_date

        # Define previous weeks
        start_of_previous_week = start_of_current_week - timedelta(weeks=1)
        start_of_two_weeks_ago = start_of_current_week - timedelta(weeks=2)

        # Define ends (Sunday of each week)
        end_of_current_week = end_date
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

        ins_query = f"""
            SELECT *
            FROM InsuranceClinicalUtil
            WHERE AppStart BETWEEN '{query_start}' AND '{query_end}'
        """
        ins_data = pd.read_sql_query(ins_query, engine)

        data.drop_duplicates(inplace=True)

        data = data[data['ServiceCodeDescription'] != 'GPAT']

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

        ins_data["AppStart"] = pd.to_datetime(ins_data["AppStart"])
        ins_data["Week"] = ins_data["AppStart"].apply(get_week_label)

        client_counts = (
            data.groupby(['School', 'Week'])['Client']
            .nunique()
            .reset_index()
            .pivot(index='School', columns='Week', values='Client')
            .fillna(0)
            .astype(int)
            .reset_index()
        )

        county_client_counts = (
            ins_data.groupby(['County', 'Week'])['Client']
            .nunique()
            .reset_index()
            .pivot(index='County', columns='Week', values='Client')
            .fillna(0)
            .astype(int)
            .reset_index()
        )

        county_client_counts = county_client_counts.rename(columns={'County': 'School'})


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

        ins_event_hours = (
            ins_data.groupby(['County', 'Week'])['EventHours']
            .sum()
            .reset_index()
            .pivot(index='County', columns='Week', values='EventHours')
            .fillna(0)
            .round(2)
            .reset_index()
        )

        ins_event_hours = ins_event_hours.rename(columns={'County': 'School'})

        columns_in_order_hours = ['School'] + [col for col in desired_order if col in event_hours.columns]
        event_hours = event_hours[columns_in_order_hours]

        data['ServiceCategory'] = data.apply(label_service, axis=1)

        service_hours = (
            data.groupby(['ServiceCategory', 'Week'])['Client']
            .nunique()
            .reset_index()
            .pivot(index='ServiceCategory', columns='Week', values='Client')
            .fillna(0)
            .astype(int)
            .reset_index()
        )

        columns_in_order_services = ['ServiceCategory'] + [col for col in desired_order if col in service_hours.columns]
        service_hours = service_hours[columns_in_order_services]

        ins_data['ServiceCategory'] = ins_data.apply(label_service, axis=1)

        ins_service_hours = (
            ins_data.groupby(['ServiceCategory', 'Week'])['Client']
            .nunique()
            .reset_index()
            .pivot(index='ServiceCategory', columns='Week', values='Client')
            .fillna(0)
            .astype(int)
            .reset_index()
        )

        ins_service_hours = ins_service_hours[['ServiceCategory'] + [col for col in desired_order if col in ins_service_hours.columns]]

        client_counts = pd.concat([client_counts, county_client_counts], ignore_index=True)
        event_hours = pd.concat([event_hours, ins_event_hours], ignore_index=True)
        service_hours = pd.concat([service_hours, ins_service_hours], ignore_index=True)

        output_file = io.BytesIO()
        with ExcelWriter(output_file, engine='openpyxl') as writer:
            data.to_excel(writer, sheet_name='Raw', index=False)
            ins_data.to_excel(writer, sheet_name='Raw Ins Data', index=False)
            client_counts.to_excel(writer, sheet_name='# of Clients', index=False)
            event_hours.to_excel(writer, sheet_name='# of Hours', index=False)
            service_hours.to_excel(writer, sheet_name='# of Clients by Service', index=False)

        output_file.seek(0)
        return output_file
    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e
    finally:
        engine.dispose()