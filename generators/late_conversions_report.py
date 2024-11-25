import pandas as pd
from sqlalchemy import create_engine
import pymssql
import openpyxl
from datetime import datetime
import io
import os

def generate_late_conversions_report(app_start, app_end, converted_after):
    try:
        user_name = os.getlogin()
        connection_string = f"mssql+pymssql://MeetCTP\Administrator:$Unlock01@CTP-DB/CRDB2"
        engine = create_engine(connection_string)
        
        #app_start_str = datetime.strftime(app_start, '%Y-%m-%d %H:%M:%S.%f')
        #app_end_str = datetime.strftime(app_end, '%Y-%m-%d %H:%M:%S.%f')
        #converted_after_str = datetime.strftime(converted_after, '%Y-%m-%d %H:%M:%S.%f')

        query = f"""
            SELECT ProviderName,
                ProviderEmail,
                ClientName,
                School,
                AppStart,
                AppEnd,
                ConvertedDT,
                LengthHr,
                RateProvider,
                PayForHours,
                Mileage,
                MilesRate,
                PayForMileage,
                TotalPay
            FROM LateConversions
            WHERE AppStart >= '{app_start}' AND AppEnd <= DATEADD(day, 1, '{app_end}') AND ConvertedDT >= '{converted_after}'
        """
        data = pd.read_sql_query(query, engine)

        data.drop_duplicates(inplace=True)

        data = data.sort_values(by='ConvertedDT', ascending=True)

        data['AppStart'] = data['AppStart'].dt.strftime('%m/%d/%Y %I:%M%p')
        data['AppEnd'] = data['AppEnd'].dt.strftime('%m/%d/%Y %I:%M%p')
        data['ConvertedDT'] = data['ConvertedDT'].dt.strftime('%m/%d/%Y %I:%M%p')

        output_file = io.BytesIO()
        data.to_excel(output_file, index=False)
        output_file.seek(0)
        return output_file
    
    except Exception as e:
        print("Error occurred while generating report: ", e)
        raise e
    finally:
        engine.dispose()