#!/usr/bin/python3

import csv
import sys
import os
import openpyxl
from datetime import datetime

month = datetime.now()
totals = []
dir = '<csv download destination path here>'
dest_dir = '<parsed csv destination path here>'
blank_audit = str(dest_dir + '/Audit Template.xlsx')
final_audit = str(dest_dir + '/2021/' + month.strftime('%B') + '.xlsx')

def counter(csv_download):
    with open(dir + csv_download) as file:
        csv_reader = csv.reader(file, delimiter=',')
        exclusions = ['expertek.com', 'acumen.com', 'mits.com', 'centraldata.com', 'infor.com', 'birst.com']
        accounts = []
        statuses = []
        disabled = []

        #compiles all accounts and statuses into respective lists
        for row in csv_reader:
            accounts.append(row[3])
            statuses.append(row[5])

        #excludes third party domains from the overall count
        excluded = [item for item in accounts if any(sub in item.lower() for sub in exclusions)]
        
        #finds disabled accounts that match the excluded domains above
        for i in range(len(statuses)):
            if statuses[i] == 'Disabled':
                disabled.append(accounts[i])

        #subtracts total disabled accounts from disabled excluded accounts is they exist
        excluded_disabled = [acc for acc in disabled if any(sub in acc for sub in excluded)]

        if len(excluded_disabled) != 0:
            adjusted_excluded = len(excluded) - len(excluded_disabled)
        else:
            adjusted_excluded = len(excluded)
        
        #subtracts excluded and disabled users from total accounts
        total = len(accounts) - 1 - adjusted_excluded - len(disabled)
        totals.append(total)

#iterates through downloaded user export files and runs the above function against each
for files in os.listdir(dir):
    exports = []

    if files.startswith('UsersExport'):
        exports.append(files)
        for items in exports:
            counter(items)

#prepares the blank audit sheet for writing
audit_sheet = openpyxl.load_workbook(blank_audit)
sheet = audit_sheet['Audit']

#interates through enumerated totals list, starts on row 2
for row, val in enumerate(totals, start=2):
    sheet.cell(row=row, column=3).value = val
    
audit_sheet.save(final_audit)
