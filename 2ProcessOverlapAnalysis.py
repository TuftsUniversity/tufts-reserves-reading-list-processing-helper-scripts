import pandas as pd
# from tkinter import Tk
# from tkinter.filedialog import askopenfilename
import requests
import sys
sys.path.append('config/')
import secrets_local
import glob
import json
import re
import os

import sys
# # Initialize Tkinter and hide the main window
# Tk().withdraw()


bnf_files = glob.glob("Overlap Analysis Input Directory/*", recursive = True)

bAndN_filename = bnf_files[0]
barnes_and_noble_df_input = pd.read_excel(bAndN_filename, dtype={"ISBN": "str", "ISBN(Matching Identifier)": "str", "ISBN(13)": "str", "MMS ID": "str"}, engine='openpyxl')
courses_url = "https://api-na.hosted.exlibrisgroup.com/almaws/v1/courses?"

api_key = secrets_local.prod_courses_api_key


barnes_and_noble_df_input['course_code'] = ""
barnes_and_noble_df_input['section'] = ""
barnes_and_noble_df_input['course_name'] = ""
barnes_and_noble_df_input['processing_department'] = ""

barnes_and_noble_df = barnes_and_noble_df_input.copy()

for column in barnes_and_noble_df.columns:
    barnes_and_noble_df[column] = barnes_and_noble_df[column].astype(str)
    #if barnes_and_noble_df[column].dtype == "object":
    barnes_and_noble_df[column] = barnes_and_noble_df[column].apply(lambda x: x.replace('"', ''))
#print(barnes_and_noble_df)
x = 0
for index, row in barnes_and_noble_df.iterrows():
    # if x == 100:
    #     break

    # print(row)
    #
    # sys.exit()
    semester = row['Term']

    if 'F' in semester:
        semester = semester.replace('F', 'Fa')

    elif 'W' in semester:
        semester = semester.replace('W', 'Sp')


    request_url = courses_url + "apikey=" + api_key + "&q=name~" + semester + "*" + row['Dept'] + "*" + row['Course'] + "*" + row['Sec'] + "&format=json"
    #print(request_url)
    response = requests.get(request_url).json()




    #print(str(index) + "\t-" + request_url)



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

oDir = "Barnes and Noble Parsed"

if not os.path.isdir(oDir) or not os.path.exists(oDir):
	os.makedirs(oDir)

barnes_and_noble_df.to_excel('Barnes and Noble Parsed/Updated Barnes and Noble.xlsx', index=False)
sys.exit()
