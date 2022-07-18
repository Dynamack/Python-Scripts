## This file transforms tables from Azure JSON files into Excel

import json
import pprint as pp
import pandas as pd
import numpy as np
import os

array = np.chararray(shape=(100000, 20), itemsize=1000, unicode = True)                     # array where values are stored
docNumber = 0                                                                               # variable keeps track of document number
index = 0                                                                                   # need this index to add new page rows to bottom of array

rootDir = r'C:\Users\om11\Documents\Work Tasks\(11) Jun 2022\Bank Statement Analysis - David Merritt\HSBC Type 1 That Failed\JSONs'

dirTree = os.walk(rootDir)                                                                  # creates tree of directory and subpaths

for dirPaths, subPaths, files in dirTree:                                                   # for each subpath
    
    for currentFile in files:                                                               # for each file

        docNumber += 1

        filePath = os.path.join(dirPaths, currentFile)                                      # get file path
        
        JSON_file = open(filePath)                                                          # open JSON file

        JSON_data = json.load(JSON_file)                                                    # convert JSON to dict

        for i in range(len(JSON_data['tables'])):                                           # iterate through tables/pages
    
            for j in range(len(JSON_data['tables'][str(i+1)][0]['cells'])):                 # iterate through cells in table
        
                row = JSON_data['tables'][str(i+1)][0]['cells'][j]['row']
                column = JSON_data['tables'][str(i+1)][0]['cells'][j]['column']
                value = JSON_data['tables'][str(i+1)][0]['cells'][j]['text']
                array[index + row, column] = value

            index += JSON_data['tables'][str(i+1)][0]['rows']

print('Number of documents: ' + str(docNumber))

df = pd.DataFrame(data=array, dtype = str, index=None, columns=None)                        # write array to pandas dataframe

storage_location = 'C:\\Users\\om11\\Documents\\Work Tasks\\(11) Jun 2022\\Bank Statement Analysis - David Merritt\\Results\\'

filename = input("Please enter filename: ")                                                 # enter name
    
df.to_excel(storage_location + filename + '.xlsx', header=False, index=False)               # write to excel

print("COMPLETE")