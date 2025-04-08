import pandas as pd
from sqlalchemy import create_engine
from datetime import datetime
from pyairtable import Api
import io
import os
import re

def get_inactive_employee_list():
    try:
        connection_string = f"mssql+pymssql://MeetCTP\Administrator:$Unlock01@CTP-DB:1433/CRDB2"
        engine = create_engine(connection_string)

        query = f"""
            SELECT
                *
            FROM ActiveContacts
            WHERE Status = 'Inactive' AND ServiceType = 'employee'
        """
        data = pd.read_sql_query(query, engine)

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e

def get_no_contact_list():
    try:

        api = Api('patpaS7kXYs546WpG.cc10e36e0d622e8e5b8d1be51a6b27eaabb16b2ce3cd8009157bc4cef04c7783')

        counselors_social_table = api.table('applyILT6MqcpyHWU', 'tblcISPJ1KskmFJ3V')
        bcba_lbs_table = api.table('app9O5xkhfInyGoip', 'tbl0YfBacdKvvNqpq')
        wilson_table = api.table('appACRGeTxgqokzXT', 'tblfk2P4ZiZsy2ZAV')
        speech_table = api.table('appFwul5GLBW3XXkA', 'tblkZq8PRZCGPykBi')
        sped_table = api.table('appGj6OWRMqrdcydL', 'tblDWEdcCkYnXGNcb')
        paras_table = api.table('app5obuWU6q9BKfiL', 'tblWylyMP4shyRhpM')
        mobile_table = api.table('app27nPo3s0RmlPyW', 'tbl42Un3FVBkJGXpe')
        archived_para_21_22 = api.table('appJMe2I9C9NMSu9d', 'tblwpjA57QX8h8xj6')
        archived_para_19_21 = api.table('appkZep4g2h0AGfR9', 'tblwDYMALGzp1Gfbl')
        archived_para_22 = api.table('appCsoodShQ4P4JrV', 'tbltCys3NfScMbLyW')

        tables = [
            (counselors_social_table, "Counselors and Social Workers"),
            (bcba_lbs_table, "BCBA and LBS"),
            (wilson_table, "Wilson Reading Instructors"),
            (speech_table, "Speech Therapists"),
            (sped_table, "SPED Teachers and Tutors"),
            (paras_table, "Paraprofessional"),
            (mobile_table, "Mobile Therapist"),
            (archived_para_21_22, "Archived Para Apps 2021-2022"),
            (archived_para_19_21, "Archived Para Apps 2019-2021"),
            (archived_para_22, "Archived Para Apps 08.15.2022")
        ]

    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e