import pandas as pd
from pyairtable import Api
from pandas import ExcelWriter
import io


def get_all_at_tables(start_date, end_date):
    try:
        start_date = pd.to_datetime(start_date)
        end_date = pd.to_datetime(end_date)
        
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
        
        total_paras_ncns = 0
        total_paras_interviews = 0
        
        completed_interviews = 0

        output_file = io.BytesIO()

        # Create a new ExcelWriter object to write data to the output_file
        with ExcelWriter(output_file, engine='openpyxl') as writer:
            for table, sheet_name in tables:
                # Get all records from Airtable
                records = table.all()

                # Convert records to a DataFrame
                data = [record['fields'] for record in records]
                df = pd.DataFrame(data)
                
                df['Interview Scheduled'] = pd.to_datetime(df['Interview Scheduled'], errors='coerce')
                
                filtered_df = df[(df['Interview Scheduled'] >= pd.to_datetime(start_date)) & 
                                 (df['Interview Scheduled'] <= pd.to_datetime(end_date))]

                # If no records match the date range, skip this table
                if filtered_df.empty:
                    print(f"Skipping {sheet_name} as no data matches the date range.")
                    continue

                # Count interviews and NCNS
                ncns_count, row_count = count_ncns_in_interviews(filtered_df)
                
                if sheet_name == "Paraprofessional":
                    total_paras_ncns += ncns_count
                    total_paras_interviews += row_count

                # Add to the totals
                total_ncns += ncns_count
                total_interviews += row_count

                # Write the DataFrame to the Excel sheet with the corresponding sheet name
                filtered_df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Optional: You could also add a summary sheet with totals if you want
            summary_df = pd.DataFrame({
                'Total # of NCNS': [total_ncns],
                'Total # of Interviews': [total_interviews],
                '% of NCNS': ["{:.2f}%".format(total_ncns / total_interviews * 100)],
                'Total # of NCNS (Para)': [total_paras_ncns],
                'Total # of Interviews (Para)': [total_paras_interviews],
                '% of NCNS (Para)': ["{:.2f}%".format(total_paras_ncns / total_paras_interviews * 100)]
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
    
    # Step 2: Filter interviews where the status contains "NCNS"
    interviews['Status'] = interviews['Status'].astype(str)
    ncns_count = interviews[interviews['Status'].str.contains('NCNS', case=False, na=False)].shape[0]
    
    return ncns_count, row_count