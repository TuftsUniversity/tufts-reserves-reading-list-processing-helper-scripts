import pandas as pd
import json
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import numbers
import sys
sys.path.append('config/')
import secrets_local
#from tkinter.filedialog import askopenfilename
import glob
import pandas as pd
import json
import re
import requests
import xml.etree.ElementTree as ET
import time
import os
import socket


def search(row, df, api_key_courses, output_df):

    # print(initial_row)
    # print(type(initial_row))

    course_url = f"https://api-na.hosted.exlibrisgroup.com/almaws/v1/courses?apikey={api_key_courses}&q=code~" + str(row['Course Code'] + "%20AND%20section~" + str(row['Course Section'])) + "&format=json"
    #course_url_without_instructor = f"https://api-na.hosted.exlibrisgroup.com/almaws/v1/courses?apikey={api_key_courses}&q=name~{course_semester}-{course_number}&format=json"


    response_course = requests.get(course_url)

    # print(response_course.status_code)
    # print(response_course.text)
    json_course = response_course.json() if response_course.status_code == 200 else {}

    if response_course.status_code == 200:
        json_course = response_course.json()

        course_id = json_course['course'][0]['id']


        reading_lists_url = f"https://api-na.hosted.exlibrisgroup.com/almaws/v1/courses/{course_id}/reading-lists?apikey={api_key_courses}&format=json"


        response_reading_lists = requests.get(reading_lists_url)

        if response_reading_lists.status_code == 200 and 'reading_list' in response_reading_lists.text:

            reading_list_instance = response_reading_lists.json()
            #for reading_list_instance in response_reading_lists['reading_list']:
            json_reading_list = reading_list_instance['reading_list'][0]

            # print(json.dumps(json_reading_list))




            reading_list_id = json_reading_list['id']
            citations_url = f"https://api-na.hosted.exlibrisgroup.com/almaws/v1/courses/{course_id}/reading-lists/{reading_list_id}/citations?apikey={api_key_courses}&format=json"


            response_citations = requests.get(citations_url)


            if response_citations.status_code == 200:


                reading_list_mms_id_list = []



                # try:
                #     reading_list_mms_id_list = [record['metadata']['mms_id'] for record in response_citations.json()['citation']]
                #
                # except:
                try:
                    for citation in response_citations.json()['citation']:

                        reading_list_mms_id_list.append(citation['metadata']['mms_id'])
                except:
                    for mms_id in row['MMS ID']:

                        new_row = initial_row.copy()


                        #new_row = new_row.loc[0,
                        output_df = pd.concat([output_df, df[(df['MMS ID'] == mms_id)].to_frame().T], ignore_index =True)

                        new_row.loc['reading_list_mms_id'] =  ""
                        new_row.loc['citation_id'] =  ""
                        new_row.loc['On Reading List Already?'] =  "No citations on reading list"


                        output_df = pd.concat([output_df, new_row.to_frame().T], ignore_index =True)


                        return output_df



                                #reading_list_citation_id_list = [record['citation_id'] for record in response_citations.json()['citation']]

                y = 0
                for input_mms_id in row['MMS ID']:
                    print("Course: " + json_course['course'][0]['code'] + ", MMS ID: " + input_mms_id)
                    # print(initial_row)
                    if input_mms_id in reading_list_mms_id_list:
                        initial_row.loc['On Reading List Already?'] = "Yes"
                        initial_row.loc['reading_list_mms_id'] =  input_mms_id
                        # print(json_dumps)
                        matching_citation_id_list = [citation['id'] for citation in response_citations.json()['citation'] if 'metadata' in citation and 'mms_id' in citation['metadata'] and citation['metadata']['mms_id'] == input_mms_id]
                        initial_row.loc['citation_id'] = ";".join(matching_citation_id_list)

                        output_df = pd.concat([output_df, initial_row.to_frame().T], ignore_index =True)

                    else:
                        initial_row.loc['On Reading List Already?'] = "No"
                        initial_row.loc['reading_list_mms_id'] = ""
                        initial_row.loc['citation_id']  = ""
                        output_df = pd.concat([output_df, initial_row.to_frame().T], ignore_index =True)
                    pd.set_option('display.max_columns', None)
                    pd.set_option('display.max_colwidth', None)
                    # print(input_mms_id)
                    # print(row)
                    y += 1




            else:

                #print("Error retrieving citations Course: " + json_course['course'][0]['code'] + ", MMS ID: " + row['MMS ID'])
                initial_row.loc['On Reading List Already?'] = "Error in retrieving citation list for reading list"
                initial_row.loc['reading_list_mms_id'] = ""
                initial_row.loc['citation_id'] = ""
                output_df = pd.concat([output_df, new_row.to_frame().T], ignore_index =True)
                print("ERROR Citation")
                print(response_citations.text)



        else:

            for mms_id in row['MMS ID']:

                new_row = initial_row.copy()


                #new_row = new_row.loc[0,
                new_row.loc['MMS ID'] = mms_id
                new_row.loc['reading_list_mms_id'] =  ""
                new_row.loc['citation_id'] =  ""
                new_row.loc['On Reading List Already?'] =  "No reading list for course"


                output_df = pd.concat([output_df, new_row.to_frame().T], ignore_index =True)




            print("No Reading List for " + row['Course Code'] + "-" + row['Course Section'])



    else:
        for mms_id in row['MMS ID']:

            new_row = initial_row.copy()
            new_row.loc['MMS ID'] =  mms_id
            new_row.loc['reading_list_mms_id'] =  ""
            new_row.loc['citation_id'] =  ""
            new_row.loc['On Reading List Already?'] =  "Course not found"


            output_df = pd.concat([output_df, new_row.to_frame().T], ignore_index =True)


            print("Error retrieving course for " + row['Course Code'] + "-" + row['Course Section'])
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_colwidth', None)
    # print(output_df)
    return output_df

bnf_files = glob.glob('Barnes and Noble Parsed/*', recursive = True)

file_path = bnf_files[0]

x = 0
pd.options.mode.chained_assignment = None
df = pd.read_excel(file_path, engine="openpyxl", dtype={'MMS ID': 'str', 'Barcode': 'str', 'Section': 'str', 'ISBN': 'str'})

output_df = pd.DataFrame(columns=df.columns)

output_columns = df.columns


output_df['On Reading List Already?'] = ""
output_df['reading_list_mms_id'] = ""
output_df['citation_id'] = ""


iteration_dataframe =  df.groupby(["Course Code", "Course Section"])["MMS ID"].apply(lambda x: ';'.join(x).split(";")).reset_index()
for _, row in iteration_dataframe.iterrows():
    if x != 0 and x % 100 == 0:
        time.sleep(30)


    api_key_courses = secrets_local.prod_courses_api_key
    output_df = search(row, df, api_key_courses, output_df)
    x += 1
pd.set_option('display.max_columns', None)
# print(output_df)

output_df.to_excel("Barnes and Noble Parsed/Matching Barnes and Noble Inventory - Reading List and Citation Creation Status.xlsx", index=False)
# course_url = f"https://api-na.hosted.exlibrisgroup.com/almaws/v1/courses?apikey={api_key_courses}&q=name~{course_semester}-{course_number}%20AND%20instructors~{instructor}&format=json"
#       course_url_without_instructor = f"https://api-na.hosted.exlibrisgroup.com/almaws/v1/courses?apikey={api_key_courses}&q=name~{course_semester}-{course_number}&format=json"
#
#       response_course = ""
#       if instructor != "":
#
#           response_course = requests.get(course_url)
#       else:
#           response_course = requests.get(course_url_without_instructor)
#
#
#
#       json_course = response_course.json() if response_course.status_code == 200 else {}
