import numpy as np
import pandas as pd
import streamlit as st

import sys
import os
from os import listdir
from datetime import datetime
import pandas as pd
import re
import matplotlib.pyplot as plt
import numpy as np
import re
from openpyxl import load_workbook

from io import BytesIO


def clear_sheet(sheet):
    for row in sheet:
        sheet.delete_rows(1, sheet.max_row+1)
    return sheet

def copy_paste_cleaned_data(wb, target_sheet_name, source_sheet, reporting_year):
    clear_sheet(wb[target_sheet_name])
    target_sheet = wb[target_sheet_name]

    # calculate total number of rows and 
    # columns in source excel file
    mr = source_sheet.max_row
    mc = source_sheet.max_column
      
    # copying the cell values from source 
    # excel file to destination excel file
    for i in range (1, mr + 1):
        for j in range (1, mc + 1):
            # reading cell value from source excel file
            c = source_sheet.cell(row = i, column = j)
      
            # writing the read value to destination excel file
            target_sheet.cell(row = i, column = j).value = c.value
    target_sheet.insert_cols(0)
    target_sheet.insert_cols(0)
    target_sheet['B1'] = "Year"
    target_sheet['A1'] = "Helper"

    # create helper column and year column in dataset
    for i in range (2, mr + 1):
        if target_sheet.cell(row = i, column = 7).value == None:
            continue
        else:
            formula = "=G" + str(i) + "&B" + str(i) 
            target_sheet.cell(row = i, column = 1).value = formula
            target_sheet.cell(row = i, column = 2).value = reporting_year
    return wb

def convert_empty_to_zero(sheet, col):
    print(sheet.cell(row = 1, column = col).value)
    mr = sheet.max_row
    #start at row 2 to skip headers
    for i in range(2, mr):
        if isinstance(sheet.cell(row = i, column = col).value, str) == True: 
            if sheet.cell(row = i, column = col).value== "-":
                sheet.cell(row = i, column = col).value = 0
            if str(sheet.cell(row = i, column = col).value).lower().strip() == "n/a":
                sheet.cell(row = i, column = col).value = 0
        if sheet.cell(row = i, column = col).value == None:
            continue
            

    return sheet

# this function is designed to clean numerical columns that erroneously have strings in them
# for example: "Approximately 150" -> should be replaced with 150
#               "Zero (0)" -> should be replaced with 0 
#               "828 - DENTAL Imaging Only" -> replaced with 828
def find_numbers(sheet, col):

    mr = sheet.max_row
    #start at row 2 to skip headers
    for i in range(2, mr):
        if sheet.cell(row = i, column = col).value == None:
            continue
        else:
            digits_found = re.findall(r'\d+', str(sheet.cell(row = i, column = col).value))
            if len(digits_found) >0:
                sheet.cell(row = i, column = col).value = int(digits_found[0])
            else:
                sheet.cell(row = i, column = col).value = 0
    return sheet

def clean_data(wb, data_sheet_name):

    numerical_col_list = [  "What was the total number of 30-day on-site prescriptions filled or medications dispensed by the clinic in the past year. Note that this is different from the number of medications prescribed by the clinic. Please provide your best estimate.",
                            "What was your total cash-operating expenditure in the past year? ",
                            "What is the approximate total revenue received from patient fees and reimbursements of services in the past year?",
                            "What was your total cash-operating expenditure in the past year?",
                            "What is the approximate total revenue received from patient fees and reimbursements of services in the past year?",   
                            "Approximately how many volunteer-hours were provided at your clinic within the past year? Please provide your best estimate.",
                            "Total number of NEW Patients in Past Year",
                            "Total number of patients served in past year (new and established combined)",
                            "Number of Dental VISITS",
                            "Total Number of all Medical VISITS (both primary care AND specialty visits for both new and established patients)",
                            "Number of Mental Health/Behavioral Health VISITS",
                            "Total Patients VISITS (sum of above visits)",
                            "Of the total MEDICAL visits, approximately what number of MEDICAL visits would have occurred at the ED if the clinic was not in operation? If an estimation cannot be provided, consider surveying patients this question during the visit.",
                            "Of the total DENTAL visits, approximately what number of DENTAL visits would have occurred at the ED if the clinic was not in operation? If an estimation cannot be provided, consider surveying patients this question during the visit.", 
                            "Number of In-House Imaging Tests in past year",    
                            "Number of In-House Lab Tests in past year",
                            "Number of In-House COVID Tests in past year",
                            "Number of COVID Vaccinations in past year provided  in your clinic by an outside organization (e.g. the state or local public health department or a private provider)",
                            "Number of COVID Vaccinations in past year provided independently by your clinic (vaccines administered by clinic personnel, not by an outside organization)",
                            "% Diabetes Screening/Management",
                            "% HTN Screening/Management",   
                            "% Cancer Screening/Management",
                            "% Obesity Screening/Management",   
                            "% Dental Care",
                            "% Sexual Health Screening/Management",
                            "% Dyslipidemia/Hypercholesterolemia Screening/Management", 
                            "% Mental Health Screening/management",
                            "% Influenza Immunization",  
                            "% Other Immunizations (ex: shingles, pneumonia, COVID, etc)",
                            "% Asthma/COPD Management", 
                            "% Dermatology Screening & Management", 
                            "% Heart Disease Screening & Management",
                            "% Vision Screenings and Exams",    
                            "% Arthritis/Musculoskeletal Screening/Management", 
                            "% Physicals (school, sport, or general)",
                            "% Acute Injury Management",    
                            "% Hearing Screening and Exam",
                            "% Transgender FTM (female-to-male) Patients",
                            "% Transgender MTF (male-to-female) Patients", 
                            "% Gender Non-Conforming Patients",
                            "% 0-17 years old",    
                            "% 18-64 years old",
                            "% 65+ years old",
                            "% Latino or Hispanic",
                            "% White",
                            "% Black or African American",
                            "% Asian",
                            "% American Indian or Alaskan Native",
                            "%  Native Hawaiian or Pacific Islander",
                            "% Multi-racial or Bi-Racial",
                            "% Other",
                            "% Unknown",   
                            "% : Below 100% of FPL",
                            "% : Between 100% and 200% of FPL",
                            "% : Over 200% of FPL"]    
    ws = wb[data_sheet_name]
    column_headers_obs = ws[1]
    column_headers = []
    for col in column_headers_obs:
        if col.value != None:
            column_headers = column_headers + [col.value.strip()]
            print(col.value)
    mc = ws.max_column

    for col in numerical_col_list:
        col_index = column_headers.index(col.strip()) + 1
        convert_empty_to_zero(ws,col_index)
        find_numbers(ws, col_index)
    return wb

def data_quality_check():
    return True


#sheet is the openpyxl worksheet with the raw data
def pull_clinic_list(raw_input_sheet, reporting_year):

    clinic_list = []
    for i, row in enumerate(raw_input_sheet):
        #skip if row is not the appropriate year
        if row[1].value != reporting_year:
            continue
        # Skip the first row (the row with the column names)
        if i == 0:
            continue
        if row[0].value == None:
            continue

        else:
        # Get the value of the first cell in the row
            clinic_name = row[6].value
        # Add the value to the list
            clinic_list.append(clinic_name)
            print(clinic_name)

    return clinic_list

def update_clinic_dashbaords(dest,updated_dest, clinic_name, reporting_year):
    #Open an xlsx for reading
    wb = load_workbook(filename = dest)
    #Get the current Active Sheet
    ws = wb['0. Dashboard']
    #You can also select a particular sheet
    #based on sheet name
    #ws = wb.get_sheet_by_name("Sheet1")
    #Open the csv file
    ws['I12'] = reporting_year 
    ws['R12'] = clinic_name
    wb.save(updated_dest)




output = BytesIO()
dest = "./static/2022-IAFCC-Master-Dashboard copy.xlsx"
updated_dest = "./static/IAFCC-FCC-Template-Dashboard-test.xlsx"
spectra = st.file_uploader("upload file", type={"csv", "xlsx"})
reporting_year = 2022
if spectra is not None:
	
	clean_source_wb = load_workbook(spectra)
	clean_source_ws = clean_source_wb.active
	st.write("Completed upload")
	wb = load_workbook(filename = dest)
	wb = copy_paste_cleaned_data(wb, 'Cleaned Responses',clean_source_ws, reporting_year = reporting_year)
	wb = clean_data(wb,"Cleaned Responses")
	clinic_list = pull_clinic_list(wb['Cleaned Responses'], reporting_year)
	# wb['Raw Model Inputs'].delete_columns(1,1)

	mr = wb['Raw Model Inputs'].max_row
	for i in range(0, mr-2):
	    if i <len(clinic_list):
	        wb['Raw Model Inputs'].cell(row = i+2, column = 1).value = clinic_list[i]
	    else: 
	        wb['Raw Model Inputs'].cell(row = i+2, column = 1).value = None

	wb['0. Dashboard']["I12"].value = "2022"
	wb.save(updated_dest)

	# for clinic in clinic_list:
		# final_path = dashboard_folder_path + str(reporting_year) + "-" + clinic + " dashboard.xlsx"
		# update_clinic_dashbaords(updated_dest, final_path, clinic, reporting_year)

	# with open(dest, 'rb') as f:
	with open(updated_dest, 'rb') as f:
   		st.download_button('Download File', f, file_name='test.xlsx')  # Defaults to 'application/octet-stream'




# st.write(
#     f'<span style="font-size: 78px; line-height: 1">🐱</span>',
#     unsafe_allow_html=True,
# )

# """
# # Static file serving
# """

# st.caption(
#     "[Code for this demo](https://github.com/streamlit/static-file-serving-demo/blob/main/streamlit_app.py)"
# )

# """
# Streamlit 1.18 allows you to serve small, static media files via URL. 

# ## Instructions

# - Create a folder `static` in your app's root directory.
# - Place your files in the `static` folder.
# - Add the following to your `config.toml` file:

# ```toml
# [server]
# enableStaticServing = true
# ```

# You can then access the files on `<your-app-url>/app/static/<filename>`. Read more in our 
# [docs](https://docs.streamlit.io/library/advanced-features/static-file-serving).

# ## Examples

# You can use this feature with `st.markdown` to put a link on an image:
# """

# with st.echo():
st.markdown("[![Click me](./app/static/cat.jpg)](https://streamlit.io)")

# """
# Or you can use images in HTML or SVG:
# """

# with st.echo():
#     st.markdown(
#         '<img src="./app/static/dog.jpg" height="333" style="border: 5px solid orange">',
#         unsafe_allow_html=True,
#     )