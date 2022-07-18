from pdfminer.high_level import extract_text
import re

text = extract_text(r'C:\\Users\\om11\\Documents\\VCT Novation Automation\\HSTN1001SC26 Supply of Technical Services.pdf')
print(text)

#companyNameRegEx = 'and.VINCI\sCONSTRUCTION\sGRANDS\sPROJETS' #'(-\s)?(AND|and)(\s-)?\n(\(2\)\s)?[A-Za-z \t]{5,40}'
companyName = re.findall('(AND|and)( -)?[\s]{1,5}(\(2\)){0,1}[\w \t]{5,40}', text)

print(companyName)
print(companyName[0])

if companyName != []:
    print('company name: ' + companyName.string)
else:
    print('No results found')



#(-\s)?(AND|and)(\s-)?\n(\(2\)\s)?[A-Za-z \t]{5,40}