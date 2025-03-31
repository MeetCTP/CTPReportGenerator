import pandas as pd
from pandas import ExcelWriter
from sqlalchemy import create_engine
from urllib.parse import quote_plus
from datetime import datetime
from io import BytesIO
import pymssql
import openpyxl
import io
import os

def generate_no_show_late_cancel_report(app_start, app_end, provider, client, school):
    try:
        user_name = os.getlogin()
        documents_path = f"C:/Users/{user_name}/Documents/"
        connection_string = f"mssql+pymssql://MeetCTP\Administrator:$Unlock01@CTP-DB/CRDB2"
        engine = create_engine(connection_string)
        app_start_dt = pd.to_datetime(app_start)
        app_end_dt = pd.to_datetime(app_end)
        app_start_101 = datetime.strftime(app_start, '%Y-%m-%d')
        app_end_101 = datetime.strftime(app_end, '%Y-%m-%d')

        provider_sessions_query = f"""
            SELECT
                AppStart,
                AppEnd,
                AppMinutes,
                AppHours,
                ClientId,
                Client,
                ProviderId,
                Provider,
                ProviderEmail,
                School,
                BillingCode,
                BillingDesc,
                Mileage,
                Status,
                CancellationReason
            FROM LateCancelNoShowView
            WHERE Status = 'Cancelled'
                AND (CancellationReason = 'No Show'
                OR CancellationReason LIKE '%Late Cancel%')
                AND CONVERT(DATE, AppStart, 101) BETWEEN '{app_start_101}' AND '{app_end_101}'
        """
        if provider:
            provider_sessions_query += f" AND Provider = '{provider}'"
        if client:
            provider_sessions_query += f" AND Client = '{client}'"
        if school:
            provider_sessions_query += f" AND School = '{school}'"
            
        provider_session_data = pd.read_sql_query(provider_sessions_query, engine)

        scheduling_query = f"""
            SELECT
                SchedulingPrincipal1Name,
                SchedulingPrincipal1Id,
                SchedulingPrincipal2Name,
                SchedulingPrincipal2Id,
                SchedulingChangeNote,
                SchedulingCancelledReason,
                SchedulingSegmentStartDateTime
            FROM Scheduling
            WHERE (SchedulingCancelledReason = 'No Show'
                OR SchedulingCancelledReason LIKE '%Late Cancel%')
                AND SchedulingSegmentStartDateTime BETWEEN '{app_start_dt}' AND '{app_end_dt}'
        """
        if provider:
            scheduling_query += f" AND SchedulingPrincipal1Name = '{provider}'"
        if client:
            scheduling_query += f" AND SchedulingPrincipal2Name = '{client}'"

        scheduling_data = pd.read_sql_query(scheduling_query, engine)

        provider_session_data['AppStart'] = pd.to_datetime(provider_session_data['AppStart'])
        scheduling_data['SchedulingSegmentStartDateTime'] = pd.to_datetime(scheduling_data['SchedulingSegmentStartDateTime'])

        report_data = pd.merge(provider_session_data, scheduling_data,
                            left_on=['AppStart', 'Provider', 'Client'],
                            right_on=['SchedulingSegmentStartDateTime', 'SchedulingPrincipal1Name', 'SchedulingPrincipal2Name'],
                            how='inner')

        report_data.drop_duplicates(inplace=True)

        columns_to_drop = ['SchedulingPrincipal1Name', 'SchedulingPrincipal1Id', 
                        'SchedulingPrincipal2Id', 'SchedulingPrincipal2Name', 
                        'SchedulingCancelledReason', 'SchedulingSegmentStartDateTime', 'Status']

        report_data.drop(columns=columns_to_drop, inplace=True)
        report_data['AppStart'] = report_data['AppStart'].dt.strftime('%m/%d/%Y %I:%M%p')

        output_file = io.BytesIO()
        report_data.to_excel(output_file, index=False)
        output_file.seek(0)
        return output_file

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

    finally:
        engine.dispose()

def generate_no_show_late_cancel_report_single(app_start, app_end, provider, client, school):
    try:
        user_name = os.getlogin()
        documents_path = f"C:/Users/{user_name}/Documents/"
        connection_string = f"mssql+pymssql://MeetCTP\Administrator:$Unlock01@CTP-DB/CRDB2"
        engine = create_engine(connection_string)
        app_start_dt = pd.to_datetime(app_start)
        app_end_dt = pd.to_datetime(app_end)
        app_start_101 = datetime.strftime(app_start, '%Y-%m-%d')
        app_end_101 = datetime.strftime(app_end, '%Y-%m-%d')
        schools = list(school.split(', '))

        output_file = BytesIO()

        with ExcelWriter(output_file, engine='openpyxl') as writer:
            for school_name in schools:
                clean_school_name = school_name.replace("School: ", "").strip()
                print(f"Generating report for {school_name}")
                
                # Query for provider session data
                provider_sessions_query = f"""
                    SELECT
                        AppStart,
                        AppEnd,
                        AppMinutes,
                        AppHours,
                        ClientId,
                        Client,
                        ProviderId,
                        Provider,
                        ProviderEmail,
                        School,
                        BillingCode,
                        BillingDesc,
                        Mileage,
                        Status,
                        CancellationReason
                    FROM LateCancelNoShowView
                    WHERE Status = 'Cancelled'
                        AND (CancellationReason = 'No Show'
                        OR CancellationReason LIKE '%Late Cancel%')
                        AND CONVERT(DATE, AppStart, 101) BETWEEN '{app_start_101}' AND '{app_end_101}'
                        AND School = '{school_name}'
                """
                if provider:
                    provider_sessions_query += f" AND Provider = '{provider}'"
                if client:
                    provider_sessions_query += f" AND Client = '{client}'"
                
                provider_session_data = pd.read_sql_query(provider_sessions_query, engine)

                # Query for scheduling data
                scheduling_query = f"""
                    SELECT
                        SchedulingPrincipal1Name,
                        SchedulingPrincipal1Id,
                        SchedulingPrincipal2Name,
                        SchedulingPrincipal2Id,
                        SchedulingChangeNote,
                        SchedulingCancelledReason,
                        SchedulingSegmentStartDateTime
                    FROM Scheduling
                    WHERE (SchedulingCancelledReason = 'No Show'
                        OR SchedulingCancelledReason LIKE '%Late Cancel%')
                        AND SchedulingSegmentStartDateTime BETWEEN '{app_start_dt}' AND '{app_end_dt}'
                """
                if provider:
                    scheduling_query += f" AND SchedulingPrincipal1Name = '{provider}'"
                if client:
                    scheduling_query += f" AND SchedulingPrincipal2Name = '{client}'"
                
                scheduling_data = pd.read_sql_query(scheduling_query, engine)
                
                print(f"Scheduling data for {school_name}: {len(scheduling_data)} rows")
                
                provider_session_data['AppStart'] = pd.to_datetime(provider_session_data['AppStart'])
                scheduling_data['SchedulingSegmentStartDateTime'] = pd.to_datetime(scheduling_data['SchedulingSegmentStartDateTime'])

                # Merging the two data sets
                report_data = pd.merge(provider_session_data, scheduling_data,
                                    left_on=['AppStart', 'Provider', 'Client'],
                                    right_on=['SchedulingSegmentStartDateTime', 'SchedulingPrincipal1Name', 'SchedulingPrincipal2Name'],
                                    how='inner')
                
                report_data.drop_duplicates(inplace=True)

                # Dropping unnecessary columns
                columns_to_drop = ['SchedulingPrincipal1Name', 'SchedulingPrincipal1Id', 
                                'SchedulingPrincipal2Id', 'SchedulingPrincipal2Name', 
                                'SchedulingCancelledReason', 'SchedulingSegmentStartDateTime', 'Status']
                report_data.drop(columns=columns_to_drop, inplace=True)
                
                report_data['AppStart'] = report_data['AppStart'].dt.strftime('%m/%d/%Y %I:%M%p')

                # Only write the sheet if data exists
                if not report_data.empty:
                    sheet_name = clean_school_name[:31]  # Ensure sheet name is under 31 characters
                    report_data.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # Make sure the sheet is visible
                    workbook = writer.book
                    sheet = workbook[sheet_name]
                    sheet.sheet_state = 'visible'

                else:
                    print(f"No data for {school_name}, skipping sheet.")

        output_file.seek(0)
        return output_file

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

    finally:
        engine.dispose()