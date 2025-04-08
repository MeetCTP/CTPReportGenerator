import pandas as pd
from sqlalchemy import create_engine
import pymssql
import openpyxl
from datetime import datetime, timedelta
import io
import os

def generate_cert_exp_report(status, timeframe, provider):
    try:
        user_name = os.getlogin()
        connection_string = f"mssql+pymssql://MeetCTP\Administrator:$Unlock01@CTP-DB/CRDB2"
        engine = create_engine(connection_string)

        now = datetime.now()
        target_date = now + timedelta(days=int(timeframe))
        target_date = datetime.strftime(target_date, '%Y-%m-%d')
        
        query = f"""
            SELECT
                FullName,
                DocumentName,
                DocumentType,
                DocumentExpirationDate,
                ExpirationStatus,
                Status,
                EmailAddress
            FROM CertificationExpiration
            WHERE (DocumentExpirationDate < '{target_date}') AND (Status IN ({', '.join([f"'{s}'" for s in status])}))
        """
        if provider:
            query += f""" AND (FullName = '{provider}')"""
        data = pd.read_sql_query(query, engine)

        data['DocumentExpirationDate'] = pd.to_datetime(data['DocumentExpirationDate'], format='%m/%d/%Y')
        data = data.sort_values(by='DocumentExpirationDate', ascending=True)
        data['DocumentExpirationDate'] = data['DocumentExpirationDate'].dt.strftime('%m/%d/%Y')
        data.drop_duplicates(inplace=True)

        output_file = io.BytesIO()
        data.to_excel(output_file, index=False)
        output_file.seek(0)
        return output_file
    
    except Exception as e:
        print("Error occurred while generating report: ", e)
        raise e
    finally:
        engine.dispose()