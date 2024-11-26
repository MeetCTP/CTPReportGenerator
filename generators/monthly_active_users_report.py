import pandas as pd
from sqlalchemy import create_engine
from datetime import datetime
import pymssql
import io
import os
import re
import openpyxl

def generate_monthly_active_users(start_date, end_date):
    try:
        user_name = os.getlogin()
        connection_string = f"mssql+pymssql://MeetCTP\Administrator:$Unlock01@CTP-DB:1433/CRDB2"
        engine = create_engine(connection_string)

        start_date = datetime.strptime(start_date, '%Y-%m-%d')
        end_date = datetime.strptime(end_date, '%Y-%m-%d')

        active_contacts_query = f"""
            SELECT
                FirstName,
                LastName,
                IsActive,
                IsEmployee,
                CreationDate,
                IsDeleted,
                LastDeactivatedOn,
                LastReactivatedOn
            FROM MonthlyActiveContacts
            WHERE 
                CreationDate <= '{end_date}' AND 
                (
                    LastDeactivatedOn IS NULL OR 
                    (LastDeactivatedOn IS NOT NULL AND LastReactivatedOn <= '{end_date}')
                )
        """
        report_data = pd.read_sql_query(active_contacts_query, engine)

        report_data['ActiveStart'] = report_data[['CreationDate', 'LastReactivatedOn']].max(axis=1)
        report_data['ActiveEnd'] = report_data['LastDeactivatedOn'].fillna(end_date)

        report_data = report_data[
            (report_data['ActiveStart'] <= end_date) &
            (report_data['ActiveEnd'] >= start_date)
        ]

        """def expand_active_period(row):
            return pd.date_range(
                max(row['ActiveStart'], start_date), 
                min(row['ActiveEnd'], end_date), 
                freq='D'
            )

        expanded_data = report_data.apply(expand_active_period, axis=1)
        active_days = pd.concat([pd.Series(dates) for dates in expanded_data], ignore_index=True)
        active_days = active_days.value_counts()
        max_active_users = active_days.max()

        report_data['MaxActiveUsers'] = max_active_users"""
        report_data.drop_duplicates(inplace=True)

        output_file = io.BytesIO()
        report_data.to_excel(output_file, index=False)
        output_file.seek(0)
        return output_file

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

    finally:
        engine.dispose()