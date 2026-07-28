import pandas as pd
from datetime import datetime as dt
from pathlib import Path

# READ ME: The purpose of this script is to format remote detection data into an upload-ready excel sheet.

# ==========================================================================================
# IMPORTANT: For this script to work properly, you will need to change
# the file paths in this section to match your file structure.
# ==========================================================================================

# Your file path to KBFC Detection template. 
# Adjust this file path according to your file structure. 
# detection_template = Path(r"C:\Your\File\Path\To\The\Deteciton\Template\KBFC_Detection.xlsm")

detection_template = Path(r"C:\Users\tpeterschmidt\OneDrive - PSMFC\Documents\KBFC\kbfc-db-data-formatting\mkwc\templates\KBFC_Detection.xlsm")

# Path to where the final product will be located. 
# Adjust file path to the folder that you would like the output saved to.
# output_path = Path(r"C:\Your\File\Path\To\Your\Output\Folder")
output_path = Path(r"C:\Users\tpeterschmidt\OneDrive - PSMFC\Documents\KBFC\KBFC_DB_data\MKWC\formatted_data")
 
# The file path to the source data file. 
# Adjust this file path according to your file structure.
# source_file = Path(r"C:\Your\File\Path\To\The\source_data.xlsx")
source_file = Path(r"C:\Users\tpeterschmidt\OneDrive - PSMFC\Documents\KBFC\KBFC_DB_data\MKWC\source_data\antenna2-7-2-2026.xlsx")

# ==========================================================================================
# ==========================================================================================

# Dataframe that will contain the formatted data.
detection_df = pd.read_excel(detection_template, sheet_name='Detection')

# ==========================================================================================
# IMPORTANT: Make sure the sheet_name matches the sheet name for the source data file.
# ==========================================================================================

source_df = pd.read_excel(source_file, sheet_name='Downloaded Tag IDs')

# ==========================================================================================
# ==========================================================================================

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

# ============================================================================================
# IMPORTANT: This section will need to be adjusted according to 
# the specific remote deployment associated with the data.
# ============================================================================================ 

detection_df['ImportID'] = 'XLRAC'

print("ImportID assignment successful.")



# ============================================================================================
# IMPORTANT: This section will need to be adjusted according to the specific site associated with the data.
# ============================================================================================

detection_df['SiteID'] = '1000R'

print("SiteID assignment successful.")



# ============================================================================================
# IMPORTANT: This section will need to be adjusted according to 
# the specific remote deployment associated with the data.
# ============================================================================================

# Map each Reader ID from the source data to the appropriate AntennaID.
#
# Format:
#     ReaderID: "AntennaID"
#
# The value on the left must match a value found in the 'Reader ID' column
# of the source data file.
# The value on the right will be the corresponding AntennaID
# that should be written to the output file.
# 
# If there is a leading 0 in the Reader ID values, don't add it when setting up the below antenna_map.  

antenna_map = {
    1: "001",  
    2: "002",
}

detection_df['AntennaID'] = source_df['Reader ID'].map(antenna_map)

print("AntennaID assignment successful.")

# ============================================================================================
# ============================================================================================

# Export the detection dataframe to an excel file. 

with pd.ExcelWriter(output_path / f"mkwc_detection_formatted_{dte}.xlsx") as writer:
    detection_df.to_excel(writer, sheet_name='Detection', index=False)

print("Data formatted successfully and exported to excel file.")