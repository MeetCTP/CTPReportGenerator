import pandas as pd
from pyairtable import Api
from pandas import ExcelWriter
import io


def get_all_at_tables():
    try:
        # Setup Airtable API client
        api = Api('patpaS7kXYs546WpG.cc10e36e0d622e8e5b8d1be51a6b27eaabb16b2ce3cd8009157bc4cef04c7783')

        counselors_social_table = api.table('applyILT6MqcpyHWU', 'tblcISPJ1KskmFJ3V')
        bcba_lbs_table = api.table('app9O5xkhfInyGoip', 'tbl0YfBacdKvvNqpq')
        wilson_table = api.table('appACRGeTxgqokzXT', 'tblfk2P4ZiZsy2ZAV')
        speech_table = api.table('appFwul5GLBW3XXkA', 'tblkZq8PRZCGPykBi')
        sped_table = api.table('appGj6OWRMqrdcydL', 'tblDWEdcCkYnXGNcb')
        paras_table = api.table('app5obuWU6q9BKfiL', 'tblWylyMP4shyRhpM')
        mobile_table = api.table('app27nPo3s0RmlPyW', 'tbl42Un3FVBkJGXpe')

        # Connect to Airtable and get all records
        counselors_social_records = counselors_social_table.all()
        bcba_lbs_records = bcba_lbs_table.all()
        wilson_records = wilson_table.all()
        speech_records = speech_table.all()
        sped_records = sped_table.all()
        paras_records = paras_table.all()
        mobile_records = mobile_table.all()

        # Convert the Airtable data (list of dictionaries) to a pandas DataFrame
        # Extract the 'fields' part of each record to get the actual data
        counselors_social_data = [record['fields'] for record in counselors_social_records]
        bcba_lbs_data = [record['fields'] for record in bcba_lbs_records]
        wilson_data = [record['fields'] for record in wilson_records]
        speech_data = [record['fields'] for record in speech_records]
        sped_data = [record['fields'] for record in sped_records]
        paras_data = [record['fields'] for record in paras_records]
        mobile_data = [record['fields'] for record in mobile_records]

        # Create a DataFrame from the data
        counselors_social = pd.DataFrame(counselors_social_data)
        bcba_lbs = pd.DataFrame(bcba_lbs_data)
        wilson = pd.DataFrame(wilson_data)
        speech = pd.DataFrame(speech_data)
        sped = pd.DataFrame(sped_data)
        paras = pd.DataFrame(paras_data)
        mobile = pd.DataFrame(mobile_data)

        counselors_social = counselors_social.astype('object')
        bcba_lbs = bcba_lbs.astype('object')
        wilson = wilson.astype('object')
        speech = speech.astype('object')
        sped = sped.astype('object')
        paras = paras.astype('object')
        mobile = mobile.astype('object')

        output_file = io.BytesIO()
        with ExcelWriter(output_file, engine='openpyxl') as writer:
            counselors_social.to_excel(writer, sheet_name="Counselors and Social Workers", index=False)
            bcba_lbs.to_excel(writer, sheet_name="BCBA and LBS", index=False)
            wilson.to_excel(writer, sheet_name="Wilson Reading Instructors", index=False)
            speech.to_excel(writer, sheet_name="Speech Therapists", index=False)
            sped.to_excel(writer, sheet_name="SPED Teachers and Tutors", index=False)
            paras.to_excel(writer, sheet_name="Paraprofessional", index=False)
            mobile.to_excel(writer, sheet_name="Mobile Therapist", index=False)

        output_file.seek(0)
        return output_file
    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e
    
def get_all_records(table):
    records = []
    page_size = 100  # Airtable's API default limit per request
    offset = None

    while True:
        if offset:
            records_page = table.all(offset=offset)
        else:
            records_page = table.all()

        records.extend(records_page)
        offset = records_page.get('offset', None)

        if not offset:
            break
    return records