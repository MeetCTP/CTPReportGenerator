import pandas as pd
from pyairtable import Api
from pandas import ExcelWriter
import io


def get_all_at_tables(start_date, end_date):
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

        tables = [
            (counselors_social_table, "Counselors and Social Workers"),
            (bcba_lbs_table, "BCBA and LBS"),
            (wilson_table, "Wilson Reading Instructors"),
            (speech_table, "Speech Therapists"),
            (sped_table, "SPED Teachers and Tutors"),
            (paras_table, "Paraprofessional"),
            (mobile_table, "Mobile Therapist")
        ]
        total_ncns = 0
        total_interviews = 0

        output_file = io.BytesIO()

        # Create a new ExcelWriter object to write data to the output_file
        with ExcelWriter(output_file, engine='openpyxl') as writer:
            for table, sheet_name in tables:
                # Get all records from Airtable
                records = table.all()

                # Convert records to a DataFrame
                data = [record['fields'] for record in records]
                df = pd.DataFrame(data)

                # Count interviews and NCNS
                row_count, ncns_count = count_ncns_in_interviews(df)

                # Add to the totals
                total_ncns += ncns_count
                total_interviews += row_count

                # Write the DataFrame to the Excel sheet with the corresponding sheet name
                df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Optional: You could also add a summary sheet with totals if you want
            summary_df = pd.DataFrame({
                'Total NCNS': [total_ncns],
                'Total Interviews': [total_interviews]
            })
            summary_df.to_excel(writer, sheet_name="Summary", index=False)

        # Ensure the file pointer is at the beginning of the file before returning
        output_file.seek(0)

        # Return the output file as an attachment
        return output_file
    except Exception as e:
        print('Error occurred while generating the report: ', e)
        raise e
    
def count_ncns_in_interviews(table):
    # Step 1: Extract rows where "Interview Scheduled" has a value
    interviews = table[table['Interview Scheduled'].notna()]

    row_count = len(interviews)
    
    # Step 2: Find the column name containing "Status" (whether it's 'Status' or 'Hiring Status')
    status_column = [col for col in table.columns if 'Status' in col][0]  # Find the column with "Status"
    
    # Step 3: Filter interviews where the status contains "NCNS"
    ncns_count = interviews[interviews[status_column].str.contains('NCNS', case=False, na=False)].shape[0]
    
    return ncns_count, row_count