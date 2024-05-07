#!/usr/bin/env python3
#######################################################################
#######################################################################
####    Author: Henry Steele, LTS, Tufts University
####    Title: createReadingListTable.py
####    Purpose:
####      - query extant reading lists for a given semester and produce
####        desired Outputs
####
####    Architecture:
####      - query from the Primo API using the semester prefix
####        part of the course code as a search term
####
####
####
####    Inputs:
####        - env vars to get semester prefix, and test prod, and maybe
####          a variable STOP|GO to indicate whether this should be running
####        - Barnes and Noble list from either FTP server or Drupal
####    Outputs:
####        - DataFrame --> CSV of list of reading lists that SIS can use
####          to construct links in the SIS catalog, with the following get_fields
####            - course_code
####            - link to citations of course in Primo
####        - DataFrame --> CSV for Drupal to consume to go on websites:
####           - title
####           - ISBN
####           - MMS ID (maybe)
####           - cost


import requests
import json
import os
import time
import csv
import re
import datetime
import sys
import shutil
#from bs4 import BeautifulSoup
import glob
from bs4 import BeautifulSoup
from lxml import etree
import xml.etree.ElementTree as et
# Initialize Tkinter and hide the main window

#from tkinter.filedialog import askopenfilename

import pandas as pd
import numpy as np

sys.path.append('config/')
import secrets_local
semester = sys.argv[1]

print(semester)
primo_prod_api = secrets_local.primo_prod_api

primo_sandbox_api = secrets_local.primo_sandbox_api

offset = 0



# # print(response)
# # sys.exit()
# file = open("Primo Results.json", "w+")
# print(response)
# print(response1)
# file.write(json.dumps(response))
# file.write("\n\n")
# file.write(json.dumps(response1))
# file.close()






# this is where I'll pull the daily Barnes and noble file from a web endpoint.
# one question is whether this needs to be secure. I would hope not.
# but regardless this script will either live on the SFTP server or on the
# libapps server, in which case the FTP server will just need to send it
# or libapps will retrieve it


sru_url = secrets_local.sru_url
namespaces = {'ns1': 'http://www.loc.gov/MARC21/slim'}







primo_column_list = ['processing_department', 'course_code', 'course_name', 'course_section', 'primo_course_link', 'primo_citation_link', 'mms_id', 'format', 'title', 'isbn', 'usage_restriction']
primo_df = pd.DataFrame(columns=primo_column_list)



y = 0

primo_url = "https://api-na.hosted.exlibrisgroup.com/primo/v1/search?q=course_code,contains," + str(semester) + "-&apikey=" + primo_prod_api + "&vid=" + secrets_local.primo_inst + "&tab=CourseReserves&scope=CourseReserves&limit=10&offset=" + str(y)

response_1 = requests.get(primo_url).json()


limit = int(response_1['info']['total'])



last = 0
while (int(response_1['info']['last']) == 10 or response_1['info']['last'] != last):

    last = int(response_1['info']['last'])


    primo_url = "https://api-na.hosted.exlibrisgroup.com/primo/v1/search?q=course_code,contains," + str(semester) + ",-&apikey=" + primo_prod_api + "&vid=" + secrets_local.primo_inst + "&tab=CourseReserves&scope=CourseReserves&limit=10&offset=" + str(y)

    #primo_test_url = "https://api-na.hosted.exlibrisgroup.com/primo/v1/search?q=mms_id,equals,991018066046303851&apikey=" + primo_prod_api + "&vid=01TUN_INST:01TUN&tab=CourseReserves&scope=CourseReserves"
    #response = requests.get(primo_url + "&offset" + str(0)).json()
    #response1 = requests.get(primo_url + "&offset" + str(10)).json()


    # while True:
    # try:
    response_initial = requests.get(primo_url).json()
    print(response_initial['info']['total'])
    # except:
    #     continue
    #     print("error")
        # finally:
        #     break




    #
    # print(response)
    # sys.exit

    if len(response_initial['docs']) > 0:
        for x in range (0, len(response_initial['docs'])):


            mms_id = response_initial['docs'][x]['pnx']['display']['mms'][0]

            primo_full_record_url = "https://api-na.hosted.exlibrisgroup.com/primo/v1/search?q=mms_id,equals," + mms_id + "&apikey=" + primo_prod_api + "&vid=" + secrets_local.primo_inst + "&lang=en&tab=Everything&scope=MyInst_and_CI"
            response = requests.get(primo_full_record_url).json()

            title = response['docs'][0]['pnx']['display']['title'][0]

            # print(response['docs'][x])

            print(str(x) + "\t" + title + "-" + mms_id)
            result = requests.get(sru_url + mms_id)
            # print(result.content)
            tree_bib_record = et.ElementTree(et.fromstring(result.content.decode('utf-8')))
            root_bib_record = tree_bib_record.getroot()

            try:
                usage_restriction = root_bib_record.findall(".//ns1:datafield[@tag='AVE']/ns1:subfield[@code='n']", namespaces)[0]
                usage_restriction = usage_restriction.text
            except:
                usage_restriction = ""

            # for this section, parse the course code, section, and course name out of the crsinfo Primo results.
            # the course code will be used in SIS output
            # the course name and section will be used in Drupal/library output



            try:
                isbns = response["docs"][0]['pnx']['display']['identifier'][0]

            except:
                isbns = []
            # if "Cyrus" in title:
            #     print(response['docs'][x])
            #
            #     sys.exit()
            try:
                raw_identifier_list = isbns.split(";")
            except:
                raw_identifier_list = isbns
            #print(raw_identifier_list)
            isbn_list = []
            for identifier in raw_identifier_list:

                if 'ISBN' in identifier:

                    if len(re.sub(r'.*\$\$V(\d+).*', r'\1', identifier)) == 13:

                        isbn_list.append(re.sub(r'.*\$\$V(\d+).*', r'\1', identifier))

            z = 0

            delivery_category = response['docs'][0]['delivery']['deliveryCategory']
            if "Alma-E" in delivery_category or "Alma-D" in delivery_category:
                physical_or_electronic = "Electronic"
            else:
                physical_or_electronic = "Physical"
            # physical_or_electronic = ""
            # if 'bestlocation' in response['docs'][x]['delivery'] and response['docs'][x]['delivery']['bestlocation'] != None:
            #     physical_or_electronic = "Physical"
            # else:
            #     physical_or_electronic = "Electronic"
            # print(response['docs'][x]['pnx']['display']['crsinfo'])
            # sys.exit()
            while z < len(response_initial['docs'][x]['pnx']['display']['crsinfo']):
                course_info = response_initial['docs'][x]['pnx']['display']['crsinfo'][z]# + ": " + response["docs"][x]['pnx']['display']['mms'][0] + ": " + response['docs'][x]['pnx']['display']['title'][0])


                course_code = re.sub(r'.*(\d{4}-\d{5}).*', r'\1', course_info)

                primo_citation_link = "https://tufts.primo.exlibrisgroup.com/discovery/search?query=any,contains," + str(mms_id) + "&tab=Everything&search_scope=MyInst_and_CI&vid=01TUN_INST:01TUN&lang=en&offset=0"

                primo_course_link = "https://tufts.primo.exlibrisgroup.com/discovery/search?query=course_code,contains," + course_code + "&tab=CourseReserves&search_scope=CourseReserves&vid=01TUN_INST:01TUN&lang=en&mode=advanced&offset=0"

                api_key = secrets_local.prod_courses_api_key

                courses_url = "https://api-na.hosted.exlibrisgroup.com/almaws/v1/courses?"

                request_url = courses_url + "apikey=" + api_key + "&q=code~" + course_code + "&format=json"

                print(request_url)
                try:
                    response_course = requests.get(request_url).json()

                except:
                    continue
                # finally:
                #     break


                processing_department = response_course['course'][0]['processing_department']

                processing_department = processing_department['desc']

                course_name = re.sub(r'.*([A-Z][a-z]\d{2}[^;]+).*', r'\1', course_info)

                course_semester = re.sub(r'([A-Z][a-z]\d{2}).+', r'\1', course_name)

                #print(course_semester)
                if "Fa" in course_semester:
                    course_semester = course_semester.replace("Fa", "F")
                #print(course_semester)
                # sys.exit()
                if "Sp" in course_semester:
                    course_semester = course_semester.replace("Sp", "W")
                course_department = re.sub(r'^[A-Z][a-z]\d{2}-([A-Z]+).+', r'\1', course_name)
                course_course = re.sub(r'^[A-Z][a-z]\d{2}-([A-Z]+)[\-\s]+([A-Z0-9]+).+', r'\2', course_name)
                course_section = re.sub(r'.*\$\$N([0-9A-Z]+).*', r'\1', course_info)
                # print(course_code)
                # print(course_name)
                # print(course_section)
                # print(mms_id)
                # print(title)
                primo_df = pd.concat([primo_df, pd.DataFrame({'processing_department': [processing_department], 'course_code': [course_code], 'course_name': [course_name], 'course_section': [course_section], 'primo_course_link': [primo_course_link], 'primo_citation_link': [primo_citation_link], 'mms_id':[mms_id], 'format': [physical_or_electronic], 'title': [title], 'isbn': [isbn_list], 'usage_restriction': [usage_restriction]}, index=[0])])
                z += 1


            # for isbn in isbn_list:
            #     print(isbn)
            #
            # print("------")

        y += 10
    else:
        break
#barnes_and_noble_df = barnes_and_noble_df[barnes_and_noble_df[['EAN-13', 'Term', 'Dept', 'Course', 'Sec']].notnull().all(axis=1)]
#df =                  df                 [df                 [cols]                                      .notnull().all(axis=1)]
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
# print(barnes_and_noble_df)


#primo_df = primo_df.explode('isbn').reset_index()

primo_df.to_excel("Output/6Data-primo output.xlsx", index=False)



print(primo_df)
