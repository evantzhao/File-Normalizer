#!/usr/bin/env python3

"""
Rectifier
Written by Evan Zhao

This program is designed to take any non-Excel problem file (TSV, CSV, etc) and attempt to correct the problems in its formatting
to make it compatible with this program's complimentary program: Converter.

This program will first scan the document to see how many spaces there are between each piece of data. What it will
then try to do is it will replace all "Multi-spaced" modules, with one tabbed delimiter. Then it will check to see if
the dimensions of the resulting array fit within the parameters of the wanted header.

If not, it will ruthlessly replace all spaces with tabbed delimiters, and combine the spaces before the column
identified as "Vendor Name" until it does fit the dimensions wanted.

This methodology is in no way guarenteed to have a 100% success rate, and so check the completed files and use at
your own risk.

Built in Windows 7
"""

import os
import datetime
from fuzzywuzzy import fuzz
import time
import csv
import shutil

def main():
    # Uncomment this line later
    print("Please put the files that will be reformatted into a folder on your local Windows Desktop named 'Problem'.")
    source = "Problem"

    header = ['Name', 'Vendor_Id', 'Invoice Id', 'Gross Amt', 'Txn Currency Cd', 'Invoice Dt', 'Accounting Dt',
              'Entered Date']  # get rid of this later

    print("\nYour new files will be saved in a folder on your Desktop named 'Source'.")
    name = "Source"

    cycle(source, header, name)

# Cycles through every single folder in the path, converting each file to an excel file.
def cycle(source, header, name):
    home = os.path.expanduser('~')
    folder_path = home + "/Desktop/" + source + "/"
    filename = seek(folder_path)
    pathway = direct(name)

    for x in range(0, len(seek(folder_path))):
       try:
            file = read_txt_file("%s%s" % (folder_path, (seek(folder_path))[x]))
            trialOne = spacify(file)
            if tsa_checkpoint(trialOne, header) and equalizer(trialOne):
                excelling(trialOne, filename[x], pathway)
            else:
                roundThree = []
                for index in range(len(brutalizer(file))):
                    roundThree.append(get_row(brutalizer(file)[index]))
                for i in range(len(roundThree[0])):
                    roundThree[0][i] = file_recognition(roundThree[0][i])
                final = normalize(roundThree)
                if tsa_checkpoint(final, header) and equalizer(final):
                    excelling(final, filename[x], pathway)

            os.remove("%s%s" % (folder_path, seek(folder_path)[x]))
       except:
            pass

def normalize(rows):
    normal = len(rows[0])
    index = None
    for i in range(0, normal):
        if rows[0][i] == "Name":
            index = i

    holder = list()
    holder.append(rows[0])

    for irow in range(1, len(rows)):
        temp = rows[irow]
        while len(temp) > normal:
            temp = merge(temp, index)
        holder.append(temp)

    return holder

def merge(arr, index):
    temp = []
    for i in range(0, index):
        temp.append(arr[i])
    str = arr[index] + ' ' + arr[index + 1]
    temp.append(str)
    for j in range(index + 2, len(arr)):
        temp.append(arr[j])
    return temp

def brutalizer(rows):
    temp = rows
    for irow in range(1, len(rows)):
        temp[irow] = '\t'.join(rows[irow].split())
    return temp

def headify(rows):
    temp = rows
    for index in range(len(rows[0])):
        temp[0][index] = file_recognition(rows[0][index])
    return temp

def equalizer(rows):
    precedent = len(rows[0])
    for irow in range(len(rows)):
        if len(rows[irow]) != precedent:
            return False
    return True

def spacify(rows):
    temp = rows
    counter = 2
    for row in range(len(rows)):
        while rows[row].find(" " * counter) != -1:
            counter += 1
        while counter > 1:
            temp[row] = rows[row].replace(" " * counter, "\t")
            counter -= 1
    return temp

def tsa_checkpoint(rows, header):
    if len(rows[0]) > len(header):
        return False
    return True

# File recognition program. This automates the process of inputting the headers. Because there are only a set amount
# of header formats, I can map the known formats to a hashmap, with the keys being the source file header type, and
# the output being the target output format. Easy peezy lemon squeezy.

# Mode determines if the file_recognition function will output a boolean showing success or failure, or if it will
# output a new header.
def file_recognition(unit, mode="rewrite"):
    header_hash = {'Name': ['Vendor Name', 'Name1', 'Name', 'Vname', 'Vendor_name'],
                   'Vendor_Id': ['Vendor_Id', 'Vendor ID', 'Vendor Number', 'Vendor', 'Duns_no', 'Vendor_number',
                                 'Vendor #', 'Vend_no'],
                   'Invoice_Id': ['Invoice Id', 'AP Invoice Number', 'Invoice', 'Doc_Number', 'Credit Memo',
                                  'Reference Number', 'Invoice #', 'Inv_no', 'Vendor Credit memo/reference',
                                  'Credit Memo Number', 'Invoice num', 'Invoice_Id'],
                   'Gross_Amt': ['Gross Amount', 'Gross Amt', 'TranInvAmt', 'cost_amt', 'Invoice Amt',
                                 'Invoice Amount', 'Amount', 'Amt', 'Gross_amt', 'Inv Amt', 'Credit memo Amount',
                                 'Invoice Amount SUM'],
                   'Txn_Currency_Cd': ['Currency', 'Curr', 'Inv Currency', 'InvCurrency', 'Currency USD',
                                       'Invoice Currency Code'],
                   'Invoice_Dt': ['AP Invoice Date', 'Invoice Date', 'doc_date', 'Invoice Dte', 'Document Date',
                                  'CreditMemo_Date', 'Credit Memo Date', 'Invoice Dt', "InvoiceDte", 'Inv_date',
                                  'Inv Date'],
                   'Accounting_Dt': ['AP Check Date', 'Payment Date', 'pay_due_date', 'AP Payment Due Date',
                                     'Clear Date', 'Clearing_date', 'Date Processed', 'Clearing Date', 'PaymentDate',
                                     'Check_date', 'Pymt Date'],
                   'Entered_Date': ['Date_added', 'Post Date', 'Entered Date', 'Posting_Date', 'Posting Date',
                                    'Create Date', 'ReconDate', 'entry_date', 'Create Dt', 'Payment Entry Date']}

    # If unit has a match within the table, this automatically cuts out of the for loop and returns the key.
    # Else, it sets the errorCounter value to 1.
    for key in header_hash:
        counter = 0
        errorCounter = 0
        while not fuzzy(unit, header_hash[key][counter]) and counter < len(header_hash[key]) - 1:
            counter += 1
        if fuzzy(unit, header_hash[key][counter]):
            return key
        else:
            errorCounter += 1
    # program would never run out here if the unit was located in the hash map.

    if mode == "rewrite":
        return unit
    else:
        return False

def fuzzy(input, thata, benchmark=90):
    input = input.lower()
    thata = thata.lower()
    r = (fuzz.ratio(input, thata))
    if r >= benchmark:
        return True
    else:
        return False

# Moves the columns of data into a TXT file, delimited with '|'
def excelling(rows, save, pathway):
    now = datetime.date.today().strftime("%m.%d.%y").rstrip('.txt.xls.xlsx.tsv')

    with open("%s/%s (%s).txt" % (pathway, save, now), 'w', newline='') as f:
        writer = csv.writer(f, delimiter='\t')
        writer.writerows(rows)

    print("Excelled")

# Creates a new path to store the created excel files in.
def direct(name):
    userhome = os.path.expanduser('~')
    pathway = "%s/Desktop/%s" % (userhome, name)
    if not os.path.isdir(pathway):
        os.mkdir(pathway)
    return pathway

# Finds all files within a folder. Argument takes a directory path.
def seek(path):
    files = os.listdir(path)
    return files

# Reads the TXT file in its entirety. Not really formatted well, outputs as a list.
def read_txt_file(name):
    filename = name
    # Uses the newer "with" construct to close the file automatically. Works with everything not Excel
    with open(filename) as f:
        data = f.readlines()

    return data

# Gets the row of each TXT file. Takes the content of txt files as an input, as an array.
def get_row(data, expected=None):
    newstring = data
    try:
        # Removes the random spacing characters
        if newstring.count('\t') > newstring.count(','):
            newstring = (data.split('\t'))
        else:
            newstring = data.split(',')

        newstring[len(newstring) - 1] = (newstring[len(newstring) - 1]).replace('\n', '')
        counter = 0
        for i in range(0, len(newstring)):
            newstring[i] = newstring[i].strip(' ')
            newstring[i] = newstring[i].replace('"', "")
            if newstring[i] == '':
                counter += 1
    except:
        counter = 0
        for i in range(0, len(newstring)):
            try:
                newstring[i] = newstring[i].strip('"')
                newstring[i] = newstring[i].strip(' ')
            except:
                pass
            if newstring[i] == '':
                counter += 1

    if counter > 3 * len(newstring)/4:
        return None
    else:
        return newstring

# Runs the main, after establishing that this is not a library.
if __name__ == "__main__":
    main()
