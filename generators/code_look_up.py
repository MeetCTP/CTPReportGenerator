import openpyxl
from flask import Flask, url_for, jsonify

def code_search(search_query):
    path = url_for('static', filename='servicecodetable.xlsx')
    query = search_query.strip().lower()

    found = False
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
        
    def search_table(sheet, query):
        results = []

        for row in sheet.iter_rows(min_row=2, values_only=True):
            code_name, desc, role, session_type, service_type, interaction, *details = row
            
            if (code_name and query in str(code_name).lower()) or (desc and query in str(desc).lower()):
                results.append(row)
        
        return results

    search_results = search_table(sheet, query)

    if search_results:
        return jsonify({'results': search_results})
    else:
        return jsonify({'message': 'No matches found.', 'results': []}), 404