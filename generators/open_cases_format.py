import pandas as pd
from datetime import datetime
import numpy as np
from pandas import ExcelWriter
from flask import jsonify
from sqlalchemy import create_engine
import pymssql
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import re
import io
import os

def generate_open_cases_report(uploaded_file):

    if not uploaded_file:
        return jsonify({'error': 'Could not read file'}), 415

    try:
        filename = uploaded_file.filename.lower()
        if filename.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        df['Student ID'] = df['Student ID'].astype('object')

        df['Frequency/Duration'] = np.where(
            df['Minutes'] < 60,
            df['Minutes'].astype(str) + "min/" + df['Frequency'].astype(str),
            (df['Minutes'] / 60).round(2).astype(str) + "hrs/" + df['Frequency'].astype(str)
        )

        note_cols = [col for col in df.columns if re.search(r'note', col, re.IGNORECASE)]
        if note_cols:
            # If multiple note-like columns exist, merge them into one (taking first non-null value)
            df['Notes'] = df[note_cols].bfill(axis=1).iloc[:, 0]

        df = df[['Student ID', 'County', 'Zip Code', 'Service', 'Frequency/Duration', 'Location', 'Notes']]

        output_file = io.BytesIO()
        df.to_excel(output_file, index=False)
        output_file.seek(0)
        return output_file

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e