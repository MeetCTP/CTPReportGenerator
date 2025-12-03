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
    'CSET': [r'.*Special Education.*', r'.*SPED.*', r'.*Certified Special Education.*'],
    'WRI': [r'.*Wilson.*', r'.*Wilson Reading.*'],
    'LSW': [r'.*Licensed Social Worker.*', r'.*Social Worker.*', r'.*Social Work.*'],
    'Counselor': [r'.*Counselor.*', r'.*Counseling.*', r'.*Counsel.*'],
    'IA/PCA': [r'.*Personal Care Assistant.*', r'.*Instructional Aide.*', r'.*Instructional Assistant.*', r'.*PCA.*', r'.*IA.*'],
    'BHT': [r'.*Behavioral Health Technician.*', r'.*Behavior Health Technician.*', r'.*BHT.*'],
    'CT Tutor': [r'.*Certified Teacher.*', r'.*Tutor.*', r'.*Tutoring.*'],
    'SLP': [r'.*Speech.*', r'.*Speech/Language.*', r'.*Speech Therapist.*', r'.*Speech Therapy.*', r'.*Language.*', r'.*SLP.*'],
    'Classroom IA': [r'.*Classroom Assistant.*', r'.*Classroom IA.*'],
    'RBT': [r'.*RBT.*', r'.*Registered Behavior Technician.*'],
    'Social Skills': [r'.*Social Skills.*'],
}

def generate_open_cases_report(cca_file, agora_file, insight_file, other_file):
    try:
        # --- Load each file if present ---
        dfs = []

        if cca_file:
            df = load_referral_file(cca_file)
            df = cca_referrals(df, "cca")
            if 'School' not in df.columns:
                df['School'] = 'CCA'
            dfs.append(df)

        if agora_file:
            df = load_referral_file(agora_file)
            df = agora_referrals(df, "agora")
            dfs.append(df)

        if insight_file:
            df = load_referral_file(insight_file)
            df = insight_referrals(df, "insight")
            dfs.append(df)

        if other_file:
            df = load_referral_file(other_file)
            df = cca_referrals(df, "cca")
            if 'School' not in df.columns:
                df['School'] = 'Other'
            dfs.append(df)

        # --- Merge all input files into one dataframe ---
        if not dfs:
            return jsonify({'error': 'No valid files uploaded.'}), 400

        full_df = pd.concat(dfs, ignore_index=True)

        # Normalizations
        if 'Location' in full_df.columns:
            full_df['Location'] = full_df['Location'].apply(normalize_location)

        full_df = normalize_group_individual(full_df, "")
        full_df = apply_service_normalization(full_df)

        full_df = full_df[full_df['Service'] != 'Uncategorized'].copy()

        # Push to Smartsheet
        push_to_smartsheet(full_df, "Open Case Referral Applications")

        return jsonify({'message': 'All referrals processed and uploaded successfully.'}), 200

    except Exception as e:
        print("Error:", e)
        return jsonify({'error': str(e)}), 500
    
def load_referral_file(upload):
    filename = upload.filename.lower()
    if filename.endswith(".csv"):
        return pd.read_csv(upload)
    return pd.read_excel(upload)
    
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

    df['School'] = 'Agora'

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

    if 'Service' in df.columns:
        # parse_insight_service returns (group_ind, location)
        parsed = df['Service'].apply(lambda s: parse_insight_service(s) if isinstance(s, str) else (None, None))

        # Turn into two columns
        df['Group/Individual'] = parsed.apply(lambda t: t[0] or "")
        df['Location'] = parsed.apply(lambda t: t[1] or "")

    df['School'] = 'Insight'

    return df

def normalize_location(value):
    if pd.isna(value) or not str(value).strip():
        return ""

    val = str(value).strip().lower()

    # Multi-location match
    if "and" in val and "virtual" in val and ("face" in val or "f2f" in val or "in-person" in val):
        return "F2F, Virtual"  # <- FIXED
    
    # Single-location matches
    if any(x in val for x in ["virtual", "online", "remote"]):
        return "Virtual"
    if any(x in val for x in ["f2f", "face", "in-person", "in person", "onsite", "on-site"]):
        return "F2F"

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

def parse_insight_service(service_raw):
    """
    Extract service, group/individual, and location from Insight's combined Service field.
    Example: 'BCBA - Indiv(online)' → Individual, Virtual
             'BSC - Group(In-Person)' → Group, F2F
    """

    if not isinstance(service_raw, str):
        return None, None, None

    text = service_raw.strip().lower()

    # ---------- GROUP / INDIVIDUAL ----------
    if "indiv" in text:
        group_ind = "Individual"
    elif "group" in text:
        group_ind = "Group"
    else:
        group_ind = ""

    # ---------- LOCATION ----------
    if "online" in text or "virtual" in text or "remote" in text:
        location = "Virtual"
    elif "person" in text or "in-person" in text or "in person" in text:
        location = "F2F"
    else:
        location = ""

    return group_ind, location

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

def clear_sheet(client, sheet_id):
    """Delete all existing rows in the sheet before upload."""
    sheet = client.Sheets.get_sheet(sheet_id)
    row_ids = [row.id for row in sheet.rows]

    if not row_ids:
        print("Sheet already empty.")
        return

    # Smartsheet API allows batch delete up to 500 at a time
    BATCH = 400
    for i in range(0, len(row_ids), BATCH):
        batch = row_ids[i:i+BATCH]
        client.Sheets.delete_rows(sheet_id, batch)
        print(f"Deleted {len(batch)} rows.")

def push_to_smartsheet(df, sheet_name):
    client = get_smartsheet()

    # 🔹 Get the correct sheet
    sheet_list = client.Sheets.list_sheets(include_all=True)
    target_sheet = next((s for s in sheet_list.data if s.name == sheet_name), None)
    if not target_sheet:
        raise ValueError(f"Sheet '{sheet_name}' not found in your Smartsheet workspace.")

    SHEET_ID = target_sheet.id

    # 🔹 CLEAR the sheet first
    print(f"Clearing Smartsheet '{sheet_name}'...")
    clear_sheet(client, SHEET_ID)
    print("Sheet cleared.")

    # 🔹 Fetch sheet columns and mapping
    sheet = client.Sheets.get_sheet(SHEET_ID)
    column_map = {col.title: col.id for col in sheet.columns}

    # 🔹 Prepare rows
    rows_to_add = []

    for _, record in df.iterrows():
        cells = []
        for col in df.columns:
            if col in column_map:
                cell = models.Cell()
                cell.column_id = column_map[col]

                # Handle Hyperlink column
                if col == "Apply Here":
                    cell.value = "Apply Here"
                    cell.hyperlink = models.Hyperlink(
                        url="https://forms.office.com/pages/responsepage.aspx?id=lz_WkfQhpkOUXq_ZK7JpXxb-2YMqoyRCsildbVBwsaFURE0yTFhMT0dWOUM5TVJPREZFWk5EQldXSS4u&route=shorturl"
                    )
                    cells.append(cell)
                    continue
                
                value = "" if pd.isna(record[col]) else str(record[col])

                # Handle Multi-Picklist
                col_type = next((c.type for c in sheet.columns if c.title == col), None)
                if col_type == "MULTI_PICKLIST":
                    if isinstance(value, str):
                        parts = re.split(r'[;,]\s*|\s+', value.strip())
                        parts = [p for p in parts if p]
                        value = ", ".join(parts)

                cell.value = value
                cells.append(cell)

        if cells:
            row = models.Row()
            row.to_top = False
            row.cells = cells
            rows_to_add.append(row)

    # 🔹 Upload rows in batches
    BATCH_SIZE = 400
    for i in range(0, len(rows_to_add), BATCH_SIZE):
        batch = rows_to_add[i:i+BATCH_SIZE]
        client.Sheets.add_rows(SHEET_ID, batch)
        print(f"Added {len(batch)} new rows.")

    print(f"✅ Smartsheet '{sheet_name}' fully refreshed.")