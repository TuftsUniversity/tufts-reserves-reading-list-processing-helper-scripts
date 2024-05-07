import pandas as pd
# from tkinter import Tk
# from tkinter.filedialog import askopenfilename
import requests
import sys
sys.path.append('config/')
import secrets_local
import glob
import json

import sys
# # Initialize Tkinter and hide the main window
# Tk().withdraw()


bnf_files = glob.glob('Barnes and Noble/*', recursive = True)

bAndN_filename = bnf_files[0]
barnes_and_noble_df_input = pd.read_excel(bAndN_filename, header=1, dtype={"EAN-13": "str"}, engine='openpyxl')
courses_url = "https://api-na.hosted.exlibrisgroup.com/almaws/v1/courses?"

api_key = secrets_local.prod_courses_api_key


barnes_and_noble_df_input['course_code'] = ""
barnes_and_noble_df_input['section'] = ""
barnes_and_noble_df_input['course_name'] = ""
barnes_and_noble_df_input['processing_department'] = ""

barnes_and_noble_df = barnes_and_noble_df_input.copy()
barnes_and_noble_df = barnes_and_noble_df[~barnes_and_noble_df['EAN-13'].isna()]
x = 0
for index, row in barnes_and_noble_df.iterrows():
    # if x == 100:
    #     break
    semester = row['Term']

    if 'F' in semester:
        semester = semester.replace('F', 'Fa')

    elif 'W' in semester:
        semester = semester.replace('W', 'Sp')

    request_url = courses_url + "apikey=" + api_key + "&q=name~" + semester + "*" + row['Dept'] + "*" + row['Course'] + "*" + row['Sec'] + "&format=json"

    response = requests.get(request_url).json()




    print(str(index) + "\t-" + request_url)



    try:
        course_code = response['course'][0]['code']
    except:
        course_code = "Error finding course" + json.dumps(response)

    try:
        section = response['course'][0]['section']

    except:
        section = "Error finding course" + json.dumps(response)

    try:
        course_name = response['course'][0]['name']

    except:
        course_name = "Error finding course" + json.dumps(response)

    try:
        course_processing_department = response['course'][0]['processing_department']['desc']

    except:
        course_processing_department = "Error finding processing department: " + json.dumps(response)
    barnes_and_noble_df.loc[index, 'course_code'] = course_code
    barnes_and_noble_df.loc[index, 'section'] = section
    barnes_and_noble_df.loc[index, 'course_name'] = course_name
    barnes_and_noble_df.loc[index, 'processing_department'] = course_processing_department


    x += 1

barnes_and_noble_df.to_excel('Barnes and Noble Parsed/Updated Barnes and Noble.xlsx', index=False)
sys.exit()
# Read the Excel file into a pandas DataFrame
lookup_titles_df = pd.read_excel(excel_file_path, dtype={'MMS ID': str}, engine='openpyxl')

# Read the structure of "create reading list input.txt" to get the column names
template_df = pd.read_csv("./column_headers_for_course_loader.txt", sep='\t', nrows=0)
output_df = pd.DataFrame(columns=template_df.columns)

# Map and fill the columns based on the Excel file

x = 0
for index, row in lookup_titles_df.iterrows():
    if x > 100:
        break
    # Extract and process the necessary fields
    title = row['Title']
    mms_id = str(row['MMS ID'])
    isbn = str(int(row['ISBN(Matching Identifier)'])) if not pd.isna(row['ISBN(Matching Identifier)']) else ''

    # Hardcode the invariant values
    coursecode = str(row['Course Code'])

    response = requests.get("https://api-na.hosted.exlibrisgroup.com/almaws/v1/courses?apikey=" + secrets_local.prod_courses_api_key + "&format=json&q=code~" + coursecode).json()
    section_id = response['course'][0]['section']

    citation_secondary_type = 'TEXT_RESOURCE'
    citation_status = 'ReadyForProcessing'
    reading_list_status = 'ReadyForProcessing'
    reading_list_name = str(row['Course Name'])
    reading_list_code = str(row['Course Name']) + "-v2"
    section_name = 'Resources'

    # Prepare a new row with all the column values
    new_row = {col: '' for col in template_df.columns}
    new_row.update({
        'coursecode': coursecode,
        'section_id': section_id,
        'citation_secondary_type': citation_secondary_type,
        'citation_status': citation_status,
        'citation_title': title,
        'citation_isbn': isbn,
        'citation_mms_id': mms_id,
        'reading_list_code': reading_list_code,
        'readin_list_name': reading_list_name,
        'reading_list_status': reading_list_status,
        'section_name': section_name
    })

    # Append to the output DataFrame
    output_df = output_df.append(new_row, ignore_index=True)

    x += 1
output_df = output_df.dropna()
# Save the DataFrame to a new TXT file
output_file_path = "Barnes and Noble Parsed/2Data-Parsed Barnes and Noble.txt"
output_df.to_csv(output_file_path, sep='\t', index=False)
