## This script deletes a series of files defined in an excel spreadsheet ##

import pandas as pd
import os

dir = r'C:\Users\om11\Documents\Grosvenor Liverpool\Indexed and Flattened Files - Leases Only'

filesToDelete = pd.read_excel(r'C:\Users\om11\Documents\Grosvenor Liverpool\Files to delete - leases only.xlsx', dtype = {'Filename': str})

count = 0

for index, row in filesToDelete.iterrows():
    for dirPaths, subDirs, files in os.walk(dir):
        for currentFile in files:
            if currentFile == row['Filename']:
                print("Deleting: " + currentFile)
                os.remove(os.path.join(dir, currentFile))
                count += 1
        
print("\n" + str(count) + " files deleted")