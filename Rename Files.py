## This script allows you to rename a directory of files using an excel table containing both the original filenames and desired filenames.
## Simply enter the directory below, accompanies by an excel file containing headers 'Input Filename' and Output Filename' and execute.

import pandas as pd
import os

dir = r''                                                                                               ## Enter directory where files are

filenameTable = pd.read_excel(r'C:\Users\om11\Documents\Rename Table.xlsx',                             ## Rename excel table
    dtype = {'Input Filename': str, 'Output Filename': str})   

for index, row in filenameTable.iterrows():                 # For each row  
    input_name = str(row["Input Filename"])                 # Input/original filename
    output_name = str(row["Output Filename"])               # Output/new filename

    input_name = dir + "\\" + input_name                    # 
    output_name = dir + "\\" + output_name                  # Append filename to directory

    os.rename(input_name, output_name)                      # Rename file from input_name to output_name
