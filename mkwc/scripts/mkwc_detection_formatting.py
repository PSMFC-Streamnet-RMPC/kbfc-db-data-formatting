import pandas as pd
from datetime import datetime as dt
import os
from pathlib import Path

# READ ME: The purpose of this script is to format remote detection data into an upload-ready excel sheet. 

#CHANGE
# Path to KBFC Detection template. Adjust this file path according to your file structure. 
detection_template = Path(r"C:\Users\tpeterschmidt\OneDrive - PSMFC\Documents\KBFC\kbfc-db-data-formatting\mkwc\templates\KBFC_Detection.xlsm")

# CHANGE
# Path to where the final product will be located. Adjust file path to wherever you would like the final output to be located.
output_path = Path(r"C:\Users\tpeterschmidt\OneDrive - PSMFC\Documents\KBFC\kbfc-db-data-formatting\mkwc\formatted_data")

# CHANGE 
# The file path to the source data file. Adjust this file path according to your file structure.
source_file = Path(r"C:\Users\tpeterschmidt\OneDrive - PSMFC\Documents\KBFC\kbfc-db-data-formatting\mkwc\source_data\antenna2-7-2-2026.xlsx")

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

source_df['Scan Time'] = pd.to_datetime(source_df['Scan Time'], format='%H:%M:%S.%f').dt.strftime("%H:%M:%S")

detection_df['DetectionDateTime'] = source_df['Scan Date'] + ' ' + source_df['Scan Time']

print("Date/Time formatting successful.")

# Adding the hexidecimal PIT tag codes from the source file to the detection dataframe. 
detection_df['PITtag'] = source_df['HEX Tag ID']

print("PIT tags added successfully.")

# Adding the correct remote deployment and Site ID to to the detection dataframe.
reader_deployment_map = {
    "01": "ImportID1",
    "02": "ImportID2",
}

detection_df['ImportID'] = source_df['Reader ID'].map(reader_deployment_map)

print("ImportID assignment successful.")

# Adding the correct Site ID to the detection dataframe.
importid_siteid_map = {
    "ImportID1": "SiteID1",
    "ImportID2": "SiteID2",
}

detection_df['SiteID'] = detection_df['ImportID'].map(importid_siteid_map)

print("SiteID assignment successful.")

detection_df['AntennaID'] = 2

print("AntennaID assignment successful.")

# Export the detection dataframe to an excel file. 

with pd.ExcelWriter(output_path / f"mkwc_detection_formatted_{dte}.xlsx") as writer:
    detection_df.to_excel(writer, sheet_name='Detection', index=False)

print("Data formatted successfully and exported to excel file.")