## This tool is used to extract standardised data from CoT documents using RegEx
import tkinter as tk
from tkinter import filedialog
import numpy as np
import pandas as pd
import docx2txt
import re
import os

# regex extractors
regex_tenant = 'Company\smeans\s([^\n|^;|^:]+)'
regex_landlord = 'Name\sand\saddress\sof\s[the\s]*present\slandlord,\sprovided\sby\sthe\sCompany:\s*([\w\W]+)Name\s[andres\s]*of\sany\spresent\sguarantor\sof\sthe\stenant'
regex_guarantor = 'Name\s[andres\s]*of\sany\spresent\sguarantor\sof\sthe\stenant:\s*([^\n]+)'
regex_property = 'Brief\s[Dd]escription\:?\s([\w\W]*?)Tenure\:?'
regex_lease_title_num = '(Registered\s[Tt]itle\s|Folio\s)[Nn]umber:*?\s([\w\W]*?)(Conveyancing|Root\sof\stitle)'
regex_lease_expiry_date = 'Contractual\sterm\sexpiry\sdate:\s*([^\n]+)'
regex_tenant_termination_rights = 'Options\sand\srights\sof\sfirst\srefusal\s[\w\W]*?Disclosures:?\s+(Tenants\sRight\sto\sTerminate\s)?([\w\W]*?)(1995\sAct|Collateral\sassurances|Landlord[\W]s\sRight\sto\sTerminate)'
regex_subject_to_charge = 'Charges\n[\w\W]*?Disclosures:?([\w\W]*?)(Agreements|Encumbrances)'
regex_current_annual_rent = 'Current\sannual\srent\:?([\w\W]*?)Rent\sreview\sfrequency\:?'
regex_rent_calculation = 'Original\sannual\srent\sincluding\sdetails\sof\sany\spremium\spaid\:?([\w\W]*?)Current\sannual\srent\:?'
regex_index_linked_original_rent = 'Original\sannual\srent\sincluding\sdetails\sof\sany\spremium\spaid\:?([\w\W]*?)Current\sannual\srent\:?'
regex_index_linked_rent_review_frequency = 'Rent\sreview\sfrequency\:?([\w\W]*?)Remaining\srent\sreview\sdates\:?'
regex_tenant_indemnity = 'No\sother\smaterial\smatters\n[\w\W]*?Disclosures:?([\w\W]*?)(SCHEDULE\s5|PART\s5|THE\sLETTING\sDOCUMENTS|[\W]\sNOT\sUSED)'
regex_change_of_control = 'Alienation([\w\W]*?Disclosures([\w\W]*?))Insurance\n'
regex_landlord_consent = 'Alienation([\w\W]*?Disclosures([\w\W]*?))Insurance\n'
regex_req_for_auth_guarantee = 'No\sother\smaterial\smatters\n[\w\W]*?Disclosures:?([\w\W]*?)(SCHEDULE\s5|PART\s5|THE\sLETTING\sDOCUMENTS|[\W]\sNOT\sUSED)'
regex_lease_rights = 'Summary\sof\sthe\srights\sgranted\sto\sthe\stenant\:?\s([\w\W]*?)Summary\sof\sthe\srights\sreserved\sto\sthe\slandlord'
regex_landlord_reservations = 'Summary\sof\sthe\srights\sreserved\sto\sthe\slandlord:?\s*([\w\W]*?)(Specified\sinsured\srisks|Name\sand\saddress\sof\s(the\s)?present\slandlord)'
regex_access_public_road = 'Access\n[\w\W]*?Disclosures:?([\w\W]*?)Benefits\n'
regex_repowering_application = 'Pending\sapplications\n([\w\W]*?Disclosures[\w\W]*?)(([\d]{1,2}\.)?Planning\sagreements)'
regex_COT_liability = 'LIMITATION\sOF\sLIABILITY\n([\w\W]*?)(Schedule\s1|SCHEDULE\s1|Date:)'

# regex list
regex_list = [regex_tenant, regex_landlord, regex_guarantor, regex_property, regex_lease_title_num, regex_lease_expiry_date, \
    regex_tenant_termination_rights, regex_subject_to_charge, regex_current_annual_rent, regex_rent_calculation, regex_index_linked_original_rent, \
        regex_index_linked_rent_review_frequency, regex_tenant_indemnity, regex_change_of_control, regex_landlord_consent, \
            regex_req_for_auth_guarantee, regex_lease_rights, regex_landlord_reservations, regex_access_public_road, regex_repowering_application, \
                regex_COT_liability]

# import word document(s)
root = tk.Tk()
root.withdraw()
wordDocs = filedialog.askopenfilenames()

# get filenames
filenames = [[""] for x in range(len(wordDocs))]
i = 0

while (i < len(wordDocs)):
    filenames[i] = os.path.basename(wordDocs[i])
    i += 1

## array for all results
results = np.empty((len(wordDocs), 21), dtype=object)

i = 0

while (i < len(wordDocs)):

    # get text from current word document
    docText = docx2txt.process(wordDocs[i]) ## works for .docx, not .doc

    # array that stores current word docs regex results
    docResults = np.empty((21), dtype=object)

    # iterates over list of regexes
    j = 0

    # following code extracts regex results from current word doc
    while (j < len(regex_list)):
        if (re.search(regex_list[j], docText) is None):
            docResults[j] = '[Blank]'
            j += 1
        elif (j == 4 or j == 6):
            docResults[j] = re.search(regex_list[j], docText).group(2).strip() # due to some awkwardness with lease_title_num regex, need to take group 2
            j += 1
        else:
            docResults[j] = re.search(regex_list[j], docText).group(1).strip()
            j += 1

    # add current doc results to all results
    results[i] = docResults

    # iterate counter to get next document results
    i += 1

# pandas dataframe
df = pd.DataFrame(results.T, index = ['Tenant', 'Landlord', 'Guarantor', 'Property', 'Lease Title Num', 'Lease Expiry Date', \
    'Tenant Termnination Rights', 'Subject to Charge', 'Current Annual Rent', 'Rent Calculation', \
        'Index Linked (Original Rent)', 'Index Linked (Rent Review Frequency)', 'Tenant Indemnity', \
            'Change of Control', 'Landlord Consent to Charge Req', 'Req for Auth Guarantee', \
                'Lease Rights', 'Landlord Reservations', 'Access to Public Road', \
                    'Repowering Application Submittee', 'COT Liability'], columns = [filenames])

# write to excel
df.to_excel("C:\\Users\\om11\\Documents\\Python\\Testing\\CoT_Test_Output.xlsx")