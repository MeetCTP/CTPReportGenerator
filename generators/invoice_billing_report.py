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

def generate_invoice_billing_report(input_file, lcns_file, school):
    try:
        user_name = os.getlogin()
        if input_file:
            input_file.seek(0)
            if input_file.filename.endswith('.csv'):
                data = pd.read_csv(input_file)
            else:
                data = pd.read_excel(input_file)

            match school:
                case "AH":
                    print('Coming Soon')
                case "PACyber":
                    data = pa_cyber_report(data, lcns_file)
                case "CCA":
                    data = cca_report(data)
                case "PAD":
                    data = pad_report(data, lcns_file)
                case _:
                    pass

            output_file = io.BytesIO()
            data.to_excel(output_file, index=False)
            output_file.seek(0)
            return output_file
    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e
    
SCHOOL_RULES = {
    'CCA': {
        'base': {
            'BCBA': r'Board Certified Behavior Analyst.*',
            'BSC': r'Behavior Specialist Consultant.*',
            'LC': r'Learning Coach.*',
            'Social Skills': r'Social Skills.*',
            'IA': r'Instructional Aide.*',
            'PCA': r'Personal Care Assistant.*',
            'CSET': r'Certified Special Education Teacher.*',
            'Counselor': r'Certified Counselor.*',
            'RBT': r'Registered Behavior Technician.*',
            'CT': r'(Tutor.*|Certified Teacher.*)',
            'Social Worker': r'Social Worker.*',
            'WRI': r'Wilson Instructor',
            "Print Materials": r"Print.*Materials",
        },
        'modifiers': {
            '1': r'.*1.*',
            '2': r'.*2.*',
            'Make Up': r'.*Make Up.*',
            'IEP Meeting': r'.*IEP.*',
            'PR': r'.*Progress Report.*',
            'Group': r'.*Group.*',
            'Virtual': r'.*Virtual.*',
            'F2F': r'.*Face To Face.*',
        },
    },
    'PAD': {
        'occupation': {
            'BCBA': r'Board Certified Behavior Analyst.*',
            'BSC': r'Behavior Specialist Consultant.*',
            'LC': r'Learning Coach.*',
            'Social Skills': r'Social Skills.*',
            'IA': r'Instructional Aide.*',
            'PCA': r'Personal Care Assistant.*',
            'CSET': r'Certified Special Education Teacher.*',
            'Counselor': r'Certified Counselor.*',
            'RBT': r'Registered Behavior Technician.*',
            'CT': r'(Tutor.*|Certified Teacher.*)',
            'Social Worker': r'Social Worker.*',
            'WRI': r'Wilson Instructor',
            "Print Materials": r"Print.*Materials",
        },
        'service': {
            'Regular': r'.*Regular.*',
            'Make Up': r'.*Make Up.*',
            'No Show': r'.*No Show.*',
            'Late Cancel': r'.*Late Cancel.*',
        },
    },
    'PACyber': {
        'base': {
            'Behavior Support': r'',
            'Counseling': r'',
            'Hearing': r'',
            'Instructional Aide': r'.*Instructional Aide.*',
            'Personal Care Assistant': r'.*Personal Care Assistant.*',
            'Social Skills': r'.*Social Skills.*',
            'Speech Therapy': r'.*Speech.*',
            'Tutoring': r'(Tutor.*|Certified Teacher.*)',
        },
        'cancels': {
            'No Show': r'.*No Show.*',
        },
    }
}

def apply_service_rules(description, rules):
    desc = str(description)

    # 1. Find base
    base_found = None
    for label, pattern in rules['base'].items():
        if re.search(pattern, desc, re.IGNORECASE):
            base_found = label
            break

    if not base_found:
        return "Other"

    # 2. Find modifiers
    modifiers = []
    for label, pattern in rules['modifiers'].items():
        if re.search(pattern, desc, re.IGNORECASE):
            modifiers.append(label)

    return " ".join([base_found] + modifiers)

def apply_pad_service_rules(description, rules):
    desc = str(description)

    # 1. Find base
    base_found = None
    for label, pattern in rules['occupation'].items():
        if re.search(pattern, desc, re.IGNORECASE):
            base_found = label
            break

    if not base_found:
        return "Other"

    # 2. Find modifiers
    modifier = None
    for label, pattern in rules['service'].items():
        if re.search(pattern, desc, re.IGNORECASE):
            modifier = label
            break

    return base_found, modifier

def apply_lcns_service_rules(description, rules):
    desc = str(description)

    # 1. Find base
    base_found = None
    for label, pattern in rules['occupation'].items():
        if re.search(pattern, desc, re.IGNORECASE):
            base_found = label
            break

    if not base_found:
        return "Other"
    
    return base_found

def apply_pac_service_rules(description, rules):
    desc = str(description)

    # 1. Find base
    base_found = None
    for label, pattern in rules['base'].items():
        if re.search(pattern, desc, re.IGNORECASE):
            base_found = label
            break

    if not base_found:
        return "Other"
    
    return base_found

def cca_report(data):
    rules = SCHOOL_RULES['CCA']
    data['Student Name'] = data['ClientLastName'].astype(str) + ", " + data['ClientFirstName'].astype(str)
    data['Therapist Name'] = data['ProviderFirstName'].astype(str) + " " + data['ProviderLastName'].astype(str)
    data['Service Type'] = data["ProcedureCodeDescription"].apply(lambda desc: apply_service_rules(desc, rules))
    data['Service Date'] = pd.to_datetime(data['DateOfService'])
    data['Service Date'] = data['Service Date'].dt.strftime('%m/%d/%Y')
    data['Rate'] = data['RateClient'].apply(lambda n: n * 4)
    
    data = data[['Student Name',
                 'Therapist Name',
                 'Service Type',
                 'Service Date',
                 'TimeWorkedInMins',
                 'Rate',
                 'Mileage']]
    
    data = data.sort_values(by=['Student Name', 'Therapist Name', 'Service Date'], ascending=True)
    return data

def pad_report(data, lcns_file):
    rules = SCHOOL_RULES['PAD']
    lcns_file.seek(0)
    lcns = pd.read_excel(lcns_file)

    lcns['Client'] = lcns['Client'].apply(
        lambda x: f"{x.split()[1]}, {x.split()[0]}"
    )
    lcns['Provider'] = lcns['Provider'].apply(
        lambda x: f"{x.split()[1]}, {x.split()[0]}"
    )

    lcns.rename(columns={'BillingDesc': 'ProcedureCodeDescription'}, inplace=True)
    lcns.rename(columns={'AppMinutes': 'TimeWorkedInMins'}, inplace=True)
    lcns.rename(columns={'AppHours': 'TimeWorkedInHours'}, inplace=True)
    lcns.rename(columns={'Client': 'Student Name'}, inplace=True)
    lcns.rename(columns={'Provider': 'Provider Name'}, inplace=True)
    lcns.rename(columns={'CancellationReason': 'Service Type'}, inplace=True)
    lcns.rename(columns={'AppStart': 'DateOfService'}, inplace=True)

    data['Student Name'] = data['ClientLastName'].astype(str) + ", " + data['ClientFirstName'].astype(str)
    data['Provider Name'] = data['ProviderLastName'].astype(str) + ", " + data['ProviderFirstName'].astype(str)
    data[['Occupation Type', 'Service Type']] = pd.DataFrame(data["ProcedureCodeDescription"].apply(lambda desc: apply_pad_service_rules(desc, rules)).tolist(), index=data.index)
    lcns['Occupation Type'] = lcns["ProcedureCodeDescription"].apply(lambda desc: apply_lcns_service_rules(desc, rules))

    data = pd.merge(data, lcns, how='outer')

    data['Service Date'] = pd.to_datetime(data['DateOfService'], errors='coerce', format='mixed')
    data['Service Date'] = data['Service Date'].dt.strftime('%m/%d/%Y')
    data['Session Type'] = data['Occupation Type'] + " " + data['Service Type']
    data['Rate'] = data['RateClient'].apply(lambda n: n * 4)

    data = data[['Student Name',
                 'Provider Name',
                 'Session Type',
                 'Service Date',
                 'Occupation Type',
                 'Service Type',
                 'TimeWorkedInHours',
                 'TimeWorkedInMins',
                 'Rate',
                 'Mileage']]
    return data

def pa_cyber_report(data, lcns_file):
    rules = SCHOOL_RULES['PACyber']
    lcns_file.seek(0)
    lcns = pd.read_excel(lcns_file)

    lcns.rename(columns={'BillingDesc': 'ProcedureCodeDescription'}, inplace=True)
    lcns.rename(columns={'AppMinutes': 'TimeWorkedInMins'}, inplace=True)
    lcns.rename(columns={'AppHours': 'TimeWorkedInHours'}, inplace=True)
    lcns.rename(columns={'Client': 'Student Name'}, inplace=True)
    lcns.rename(columns={'Provider': 'Therapist Name'}, inplace=True)
    lcns.rename(columns={'CancellationReason': 'Service Type'}, inplace=True)
    lcns.rename(columns={'AppStart': 'DateOfService'}, inplace=True)

    data['Student Name'] = data['ClientFirstName'].astype(str) + " " + data['ClientLastName'].astype(str)
    data['Therapist Name'] = data['ProviderFirstName'].astype(str) + " " + data['ProviderLastName'].astype(str)
    data['Service Type'] = data["ProcedureCodeDescription"].apply(lambda desc: apply_pac_service_rules(desc, rules))

    data = pd.merge(data, lcns, how='outer')

    data['Service Date'] = pd.to_datetime(data['DateOfService'], errors='coerce', format='mixed')
    data['Service Date'] = data['Service Date'].dt.strftime('%m/%d/%Y')
    data['Rate'] = data['RateClient'].apply(lambda n: n * 4)
    data['Mileage'] = np.where(
        data['Mileage'].isna() | (data['Mileage'] == "") | (data['Mileage'] == 0),
        "N",
        "Y"
    )

    data = data[['Therapist Name',
                 'Service Date',
                 'Student Name',
                 'Service Type',
                 'TimeWorkedInMins',
                 'Rate',
                 'Mileage']]
    return data