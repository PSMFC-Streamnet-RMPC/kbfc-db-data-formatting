import pandas as pd
from datetime import datetime as dt
import os
from pathlib import Path

# READ ME: The purpose of this script is to format remote detection data into an upload-ready excel sheet. 

#CHANGE
# Path to KBFC Detection template. Adjust this file path according to your file structure. 
detection_template = Path(r"C:\Users\tpeterschmidt\Desktop\KBFC Data Conversion Project\KBFC_Templates\KBFC_Detection.xlsm")

# CHANGE
# Path to where the final product will be located. Adjust file path to wherever you would like the final output to be located.
output_path = Path(r"C:\Users\tpeterschmidt\Desktop\KBFC Data Conversion Project\Yurok\Cleaned")

# CHANGE 
# The file path to the source data file. Adjust this file path according to your file structure.
source_file = Path(r"C:\Users\tpeterschmidt\Desktop\KBFC Data Conversion Project\Yurok\2020-24_PITarrays_NotInDB.xlsx")

# Dataframe that will contain the formatted data.
detection_df = pd.read_excel(detection_template, sheet_name='Detection')

source_df = pd.read_excel(source_file, sheet_name='Downloaded Tag IDs')

# Save the output files with today's date in file name
mask = '%Y%m%d'
dte = dt.now().strftime(mask)

# Set Date Time format
datetime_format = 'YYYY-MM-DD HH:MM:SS'


# Formatting the date and time columns in the source data to match the required detection file format.
source_df['Scan Date'] = pd.to_datetime(source_df['Scan Date']).dt.strftime("%Y-%m-%d")

source_df['Scan Time'] = pd.to_datetime(source_df['Scan Time']).dt.strftime("%H:%M:%S")

detection_df['DetectionDateTime'] = source_df['Scan Date'] + ' ' + source_df['Scan Time']

# Adding the hexidecimal PIT tag codes from the source file to the detection dataframe. 
detection_df['PITtag'] = source_df['HEX Tag ID']




