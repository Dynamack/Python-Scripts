from docx import Document # The docx module creates, reads and writes Microsoft Office Word 2007 docx files
import pandas as pd
import re


## REMAINING REQUIREMENTS:
#  Change log 
#  Highlight new elements in word doc as red? - May not be necessary if we have change log
#  Resilience to different cases (all caps, lowercase, camelcase etc)


## Open directory ?


changes = pd.read_excel(r'C:\Users\om11\\Documents\\DocRegExFindAndReplace\\Changes.xlsx')                                  ## Import excel with input-output transforms

document = Document(r'C:\Users\om11\\Documents\\DocRegExFindAndReplace\\Test Document.docx')                                ## Open Word document

paragraphs = [p.text for p in document.paragraphs]                                                                          ## Creates list of paragraphs in Word doc

new_document = open(r'C:\Users\om11\\Documents\\DocRegExFindAndReplace\\Test Document NEW.txt', 'w')                        ## Save new document
document_name = 'Test Document.docx'


## Scan for input 
i = 0
for line in paragraphs:                                                                                                     ## Iterate through paragraphs
    for index, row in changes.iterrows():                                                                                   ## Iterate through input clauses ## SWAP ABOVE AND BELOW? Test speed
        if row['Input'] in line:
            print('FOUND: ' + row['Input'])
            line = line.replace(row['Input'], row['Output'])                                                                ## Replace substring with output from excel
            print(line)
    new_document.write(line + "\n")                                                                                         ## Write new line to output document
            #paragraphs[i] = paragraphs[i].replace(row['Input'], row['Output'])
            #print(paragraphs[i])
#i += 1


#new_document.close()