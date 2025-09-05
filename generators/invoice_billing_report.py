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

def generate_invoice_billing_report(start_date, end_date, school):
    try:
        user_name = os.getlogin()
        db_user = os.getenv("DB_USER")
        db_pw = os.getenv("DB_PW")
        connection_string = f"mssql+pymssql://{db_user}:{db_pw}@CTP-DB:1433/CRDB2"
        engine = create_engine(connection_string)

        match school:
            case "AH":
                print('Coming Soon')
            case "Agora":
                print('Coming soon')
            case "CCA":
                data = cca_report(start_date, end_date, engine)
            case "PAD":
                print('Coming soon')
            case _:
                pass

        output_file = io.BytesIO()
        data.to_excel(output_file, index=False)
        output_file.seek(0)
        return output_file
    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

    finally:
        engine.dispose()

def cca_report(start_date, end_date, engine):
    query = f"""
        SELECT *
        FROM Invoice_CCA
        WHERE ServiceDate BETWEEN '{start_date}' AND '{end_date}'
    """
    data = pd.read_sql_query(query, engine)
    data.drop_duplicates(inplace=True)



    return data