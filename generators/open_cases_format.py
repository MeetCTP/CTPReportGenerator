import pandas as pd
from datetime import datetime
import numpy as np
from pandas import ExcelWriter
from flask import jsonify
from sqlalchemy import create_engine
from dotenv import load_dotenv
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from smartsheet import models
import pymssql
import smartsheet
import re
import io
import os

service_keywords = {
    'BCBA': [r'.*Board Certified Behavior Analyst.*', r'.*BCBA.*'],
    'BSC': [r'.*Behavior Specialist Consultant.*', r'.*BSC.*'],
    'CSET': [r'.*Speech/Language.*', r'.*Speech.*', r'.*SLP.*', r'.*Language.*'],
    'WRI': [r'.*Wilson.*', r'.*Wilson Reading.*'],
    'LSW': [r'.*Licensed Social Worker.*', r'.*Social Worker.*', r'.*Social Work.*'],
    'Counselor': [r'.*Counselor.*', r'.*Counseling.*', r'.*Counsel.*'],
    'IA/PCA': [r'.*Personal Care Assistant.*', r'.*Instructional Aide.*', r'.*Instructional Assistant.*', r'.*PCA.*', r'.*IA.*'],
    'BHT': [r'.*Behavioral Health Technician.*', r'.*Behavior Health Technician.*', r'.*BHT.*'],
    'CT Tutor': [r'.*Certified Teacher.*', r'.*Tutor.*', r'.*Tutoring.*'],
    'SLP': [r'.*Speech.*', r'.*Speech Therapist.*', r'.*Speech Therapy.*', r'.*SLP.*'],
    'Classroom IA': [r'.*Classroom Assistant.*', r'.*Classroom IA.*'],
    'RBT': [r'.*RBT.*', r'.*Registered Behavior Technician.*'],
    'Social Skills': [r'.*Social Skills.*'],
}

def generate_open_cases_report(uploaded_file, school):
    if not uploaded_file:
        return jsonify({'error': 'Could not read file'}), 415

    try:
        filename = uploaded_file.filename.lower()
        if filename.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        match school:
            case 'cca':
                df = cca_referrals(df, school)
            case 'agora':
                df = agora_referrals(df, school)
            case 'insight':
                df = insight_referrals(df, school)
            case _:
                return jsonify({'message': f'Report for {school} could not be uploaded to Smartsheet. Please ensure you have uploaded the correct file.'}), 500
            
        if 'Location' in df.columns:
            df['Location'] = df['Location'].apply(lambda x: normalize_location(x))

        df = normalize_group_individual(df, school)
        df = normalize_school_column(df, school)
        df = apply_service_normalization(df)

        try:
            df[['Service']].to_csv('C:/temp/service_debug.csv', index=False)
        except Exception as e:
            pass  # don't fail if the directory doesn't exist

        push_to_smartsheet(df, "Open Case Referral Applications")

        # --- return success message ---
        return jsonify({'message': f'Report for {school} uploaded to Smartsheet successfully.'}), 200

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e
    
def cca_referrals(df, school):
    df.rename(columns={
        'Key': 'School Referral ID',
        'Location of Service': 'Location',
        'Related Service': 'Service',
        'County of residence (enter NA for virtual services)': 'County',
        'ZIP Code of Residence': 'Zip Code',
        'Additional Portal Information:': 'Notes',
        'Grade Level': 'Grade'
    }, inplace=True)

    if 'County' in df.columns:
        df['County'] = df['County'].apply(clean_county)

    return df

def agora_referrals(df, school):
    df.rename(columns={
        'Posting ID': 'School Referral ID',
        'Comments': 'Notes'
    }, inplace=True)

    if 'County' in df.columns:
        df['County'] = df['County'].apply(lambda x: clean_county(x, remove_word=True))

    return df

def insight_referrals(df, school):
    df.rename(columns={
        'Referral ID': 'School Referral ID',
        'Student First Name': 'First Name',
        'Student Last Name': 'Last Name',
        'Student Grade Level': 'Grade',
        'ISA Service Name': 'Service',
        'ISA Delivery Method': 'Location'
    }, inplace=True)

    return df

def normalize_location(value):
    """Normalize location values to Smartsheet dropdown options (F2F, Virtual)."""
    if pd.isna(value) or not str(value).strip():
        return ""

    val = str(value).strip().lower()

    # Multi-location match
    if "and" in val and "virtual" in val and ("face" in val or "f2f" in val or "in-person" in val):
        # Means something like "Face to Face and/or Virtual"
        return ["F2F", "Virtual"]

    # Singular matches
    if any(x in val for x in ["virtual", "online", "remote"]):
        return "Virtual"
    if any(x in val for x in ["f2f", "face", "in-person", "in person", "onsite", "on-site"]):
        return "F2F"

    # Fallback — return the original so you can spot it in review
    return value

def clean_county(value, remove_word=False):
    """Cleans county names depending on school logic."""
    if pd.isna(value):
        return ""

    val = str(value).strip()

    # Handle n/a or similar
    if val.lower() in ["n/a", "na", "none"]:
        return ""

    # Remove "County" if requested
    if remove_word:
        val = val.replace("County", "").strip()
    
    return val

def normalize_group_individual(df, school):
    """
    Normalizes or infers the 'Group/Individual' column values depending on school logic.
    Smartsheet accepts only: 'Group', 'Individual', 'Classroom'.
    """

    # --- CCA ---
    # Does not have an indicator; skip adding Group/Individual column
    if school.lower() == "cca":
        return df

    # --- AGORA ---
    # Has a column with exact matching values; no normalization needed
    if school.lower() == "agora":
        if 'Group/Individual' in df.columns:
            df['Group/Individual'] = df['Group/Individual'].str.strip().replace('', None)
        return df

    # --- INSIGHT ---
    # No Group/Individual column; infer from Service column if "indiv" appears
    if school.lower() == "insight":
        if 'Service' in df.columns:
            df['Group/Individual'] = df['Service'].apply(
                lambda x: 'Individual' if isinstance(x, str) and 'indiv' in x.lower() else ''
            )
        return df

    return df

def normalize_school_column(df, school):
    """
    Ensures the 'School' column exists and is standardized to match Smartsheet formatting.
    Converts the passed-in school argument to the correct display case.
    """
    school_name_map = {
        'agora': 'Agora',
        'cca': 'CCA',
        'insight': 'Insight'
    }

    school_display = school_name_map.get(school.lower(), school.capitalize())
    df['School'] = school_display

    valid_values = {'Agora', 'CCA', 'Insight'}
    if school_display not in valid_values:
        raise ValueError(f"'{school_display}' is not an accepted School value in Smartsheet")
    
    return df

def normalize_service(service_value: str) -> str:
    """Normalize service name to match Smartsheet dropdown values."""
    if not isinstance(service_value, str) or not service_value.strip():
        return "Uncategorized"
    
    # force to single-line clean string
    service_value = service_value.strip().replace("\n", " ")

    for service_name, patterns in service_keywords.items():
        for pattern in patterns:
            if re.search(pattern, service_value, re.IGNORECASE):
                return service_name

    # Fallback default
    return "Uncategorized"


def apply_service_normalization(df: pd.DataFrame) -> pd.DataFrame:
    """Apply normalize_service() to the Service column if present."""
    if 'Service' in df.columns:
        df['Service'] = df['Service'].apply(normalize_service)
    return df

def get_smartsheet():
    load_dotenv()
    TOKEN = os.getenv("SMARTSHEET_TOKEN")

    if not TOKEN:
        raise ValueError("SMARTSHEET_TOKEN not found in environment variables.")

    smartsheet_client = smartsheet.Smartsheet(TOKEN)
    smartsheet_client.errors_as_exceptions(True)
    return smartsheet_client

def push_to_smartsheet(df, sheet_name):
    client = get_smartsheet()

    # 1️⃣ Get the correct sheet
    sheet_list = client.Sheets.list_sheets(include_all=True)
    target_sheet = next((s for s in sheet_list.data if s.name == sheet_name), None)
    if not target_sheet:
        raise ValueError(f"Sheet '{sheet_name}' not found in your Smartsheet workspace.")

    SHEET_ID = target_sheet.id

    # 2️⃣ Fetch sheet columns and make name → ID map
    sheet = client.Sheets.get_sheet(SHEET_ID)
    column_map = {col.title: col.id for col in sheet.columns}

    # 4️⃣ Prepare rows for insertion
    rows_to_add = []
    for _, record in df.iterrows():
        cells = []
        for col in df.columns:
            if col in column_map:
                value = "" if pd.isna(record[col]) else str(record[col])
                
                # Get Smartsheet column type (so we can handle Multi-Picklists correctly)
                col_type = next((c.type for c in sheet.columns if c.title == col), None)

                # --- Handle MULTI_PICKLIST columns ---
                if col_type == "MULTI_PICKLIST":
                    # Convert things like "F2F Virtual" → "F2F, Virtual"
                    if isinstance(value, str):
                        # Normalize separators (split on spaces or semicolons)
                        parts = re.split(r'[;,]\s*|\s+', value.strip())
                        # Remove any empty strings and deduplicate
                        parts = [p for p in parts if p]
                        # Rejoin into comma-separated string
                        value = ", ".join(parts)

                # --- Build the cell ---
                cell = models.Cell()
                cell.column_id = column_map[col]
                cell.value = value
                cells.append(cell)
        if cells:
            row = models.Row()
            row.to_top = False
            row.cells = cells
            rows_to_add.append(row)

    # 5️⃣ Add rows (batch to handle Smartsheet limits)
    BATCH_SIZE = 400
    for i in range(0, len(rows_to_add), BATCH_SIZE):
        batch = rows_to_add[i:i + BATCH_SIZE]
        response = client.Sheets.add_rows(SHEET_ID, batch)
        print(f"Added {len(response.data)} rows to '{sheet_name}'")

    print(f"✅ Smartsheet '{sheet_name}' updated successfully.")