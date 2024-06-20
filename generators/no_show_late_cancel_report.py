import pandas as pd
from sqlalchemy import create_engine
from datetime import datetime
import pymssql
import openpyxl
import io
import os

def generate_no_show_late_cancel_report(app_start, app_end, provider, client, school):
    try:
        user_name = os.getlogin()
        documents_path = f"C:/Users/{user_name}/Documents/"
        connection_string = f"mssql+pymssql://MeetCTP\Joshua.Bliven:$Unlock03@CTP-DB/CRDB"
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
                ProviderId,
                Provider,
                ProviderEmail,
                ClientId,
                Client,
                School,
                BillingCode,
                BillingDesc,
                Mileage,
                Status,
                CancellationReason
            FROM ProviderSessions_New
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
            FROM Scheduling_Raw_New
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

        report_data = pd.merge(provider_session_data, scheduling_data,
                            left_on=['AppStart', 'Provider', 'Client'],
                            right_on=['SchedulingSegmentStartDateTime', 'SchedulingPrincipal1Name', 'SchedulingPrincipal2Name'],
                            how='inner')

        columns_to_drop = ['SchedulingPrincipal1Name', 'SchedulingPrincipal1Id', 
                        'SchedulingPrincipal2Id', 'SchedulingPrincipal2Name', 
                        'SchedulingCancelledReason', 'SchedulingSegmentStartDateTime', 'Status']

        report_data.drop(columns=columns_to_drop, inplace=True)

        output_file = io.BytesIO()
        report_data.to_excel(output_file, index=False)
        output_file.seek(0)  # Reset the file pointer to the beginning of the BytesIO object
        return output_file

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

    finally:
        engine.dispose()