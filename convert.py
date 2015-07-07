#!/usr/bin/env python3
"""
File Conversion Program
Written by Evan Zhao

This program is designed to convert TXT and Excel files into formats compatible with Lavante, while placing each
"column" of data from the source file into a preferred order in the target file. Also does minor formatting,
such as getting rid of white space, omitting leading zeros, and formatting dates.

This project is written in Python 3.4, and requires the xlrd module so that the program can read Excel files. This
capability extends to older Excel formats, like .xls.

The new files will be saved in a new directory, or in an auxiliary directory that will be created on demand.

Update 1:
Now includes a bit of data validation. This program will check to see if the rows are completely filled out. In the
case that the rows are faulty, or omit several sections, this program will weed out those rows and replace them with an
empty row. This should leave incomplete rows, since in the environment that this program is run in, some txt spaces are
filled with ' ', which counts as a unit in this code, and will be treated as a character in this program.

Also saves all created excel files within a specific folder on the user's local desktop.

Update 2:
No longer removes the leading zeros from the Invoice Id. The column must be specifically named, Invoice Id for this
effect to be expressed.

Better whitespace processing. New ability to retain the original file name, and place a datestamp on the created file
name. Also, another step towards automation, as this program now has the ability to store multiple files into an
existing folder, or create a new folder.

Update 3:
Final update. Fuzzy string reader implemented into the program, which allows it to recognize headers of similar style
and then normalize the results, such that the computer can easily recognize which column is which. Also, Date parsing
is also complete, will recognize several formats and covnert them to the appropriate style. Optimized for speed now,
and the final file format is now a .txt file.


Required Downloads:
    xlrd: Excel File Reader
    Fuzzywuzzy: Fuzzy string reader
    CSV: TXT file writer
"""

import os
import datetime
import xlrd
from fuzzywuzzy import process
import time
import csv
import shutil

def main():
    source = "Source"
    new_header = ['Name', 'Number', 'Reference', 'Amount', 'Currency', 'Invoice Date', 'Payment Date', 'Entered Date']
    print("\nYour new files will be saved in a folder on your Desktop called 'Target'")
    new_folder = "Target"

    cycle(source, new_header, new_folder)

# Cycles through every single folder in the path, converting each file to an excel file.
def cycle(source, header, new_folder):
    folder_path = os.path.expanduser('~') + "/Desktop/" + source + "/"
    file_names = seek(folder_path)
    save_pathway = direct(new_folder)
    problem_pathway = direct("Problem/")

    for x in range(0, len(file_names)):
        # Catches errors with the try, except structure.
        # try:
        raw_data, date_mode = read_txt_file("%s%s" % (folder_path, file_names[x]))
        print("Cycling")
        col = order(general_parse(simplify(header, colify(raw_data)), header, date_mode), header)
        row = rowify(col)
        convert(row, file_names[x], save_pathway)
        # except:
        #     # Move file to the problem folder.
        #     shutil.copy("%s%s" % (folder_path, file_names[x]), "%s%s" % (problem_pathway, file_names[x]))

    for file in file_names:
        os.remove("%s%s" % (folder_path, file))

# Assumes that simplify has been run before this function. In other words, each column exists within the final header.
def order(col, header):
    ordered_data = []
    # Cycles through each header
    for header_index in range(len(header)):
        col_index = 0
        # Cycles through each column, while checking to see if the column header is equal to the header above.
        # If a match is found, the while loop is broken, and the code continues on.
        while col_index < len(col) and not col[col_index][0] == header[header_index]:
            col_index += 1
        ordered_data.append(col[col_index])
    if len(col) < len(header):
        raise ValueError
    return ordered_data

# File recognition program. This automates the process of inputting the headers. Because there are only a set amount
# of header formats, I can map the known formats to a hash map, with the keys being the source file header type, and
# the output being the target output format. EACH CATEGORY MUST CONTAIN ROUGHLY THE SAME AMOUNT OF ELEMENTS, ELSE
# THE ACCURACY OF THE MATCHER WILL DECREASE SIGNIFICANTLY IN SOME CASES

def file_recognition(unit, mode="rewrite"):
    header_hash = {'Name': ['Vendor Name', 'Name1', 'Name', 'Vname', 'Vendor', 'Vend name'],
                   'Number': ['Vendor Id', 'Vendor Number', 'Duns no', 'Vendor #', 'Vend no', 'Vend ID', 'Vendor num',
                              'Vend num', 'Vendor #', 'VendorID'],
                   'Reference': ['Invoice Id', 'AP Invoice Number', 'Invoice', 'Doc Number', 'Credit Memo',
                                 'Reference Number', 'Invoice #', 'Inv no', 'Vendor Credit memo/reference',
                                 'Credit Memo Number', 'Invoice num', 'Invoice Number'],
                   'Amount': ['Gross Amount', 'Gross Amt', 'TranInvAmt', 'cost amt', 'Invoice Amt',
                              'Invoice Amount', 'Amount', 'Amt', 'Inv Amt', 'Credit memo Amount',
                              'Invoice Amount SUM'],
                   'Currency': ['Currency', 'Curr', 'Inv Currency', 'InvCurrency', 'Currency USD',
                                'Invoice Currency Code', 'Txn Currency Cd'],
                   'Invoice Date': ['AP Invoice Date', 'Invoice Date', 'doc date', 'Document Date',
                                    'CreditMemo Date', 'Credit Memo Date', 'Invoice Dt', "InvoiceDte", 'Inv date',
                                    'Inv Date'],
                   'Payment Date': ['AP Check Date', 'Payment Date', 'pay due date', 'AP Payment Due Date',
                                    'Clear Date', 'Clearing date', 'Date Processed', 'Clearing Date', 'PaymentDate',
                                    'Check date', 'Pymt Date', 'Accounting Dt'],
                   'Entered Date': ['Date added', 'Post Date', 'Entered Date', 'Posting Date', 'Posting Date',
                                    'Create Date', 'ReconDate', 'entry date', 'Create Dt', 'Payment Entry Date',
                                    'Entered Date']}
    unit = unit.replace('_', ' ')

    # If unit has a match within the table, this automatically cuts out of the for loop and returns the key.
    # Else, it sets the errorCounter value to 1.
    master = {}
    for key in header_hash:
        recon = process.extractBests(unit, header_hash[key], score_cutoff=93)
        if len(recon) > 0:
            master["%s" % key] = len(recon) / len(header_hash[key])

    # Master should now hold an array of all the best values for each key. I'm just going to count up which ones are the
    # most heavily populated, and take it from there.
    if len(master) > 0:
        return max(master, key=lambda i: master[i])
    elif mode == "rewrite":
        return unit
    else:
        return False


# Moves the columns of data into a TXT file, delimited with '|'
def convert(rows, save, pathway):
    now = datetime.date.today().strftime("%m.%d.%y")

    with open("%s/%s (%s).txt" % (pathway, save, now), 'w', newline='') as f:
        writer = csv.writer(f, delimiter='\t')
        writer.writerows(rows)

    print("Excelled")

# Creates a new path to store the created excel files in.
def direct(new_folder):
    home = os.path.expanduser('~')
    pathway = "%s/Desktop/%s" % (home, new_folder)
    if not os.path.isdir(pathway):
        os.mkdir(pathway)
    return pathway

def simplify(header, col):
    new_data = []

    for index in range(len(col)):
        if col[index][0] in header:
            new_data.append(col[index])
    if len(new_data) == 0:
        print("Simplification fail")
        raise SyntaxError
    return new_data

def rowify(col):
    new_row = []
    for row_index in range(1, len(col[0])):
        temp = []
        for col_index in range(0, len(col)):
            temp.append(col[col_index][row_index])
        new_row.append(temp)
    return new_row

# Returns the columns of a TXT file, as a nested array. Automatically modifies the heading.
def colify(data_array):
    header_start = 0
    boolean = False
    while not boolean:
        ticker = 0
        for header_pos in range(len(get_row(data_array[header_start], "override"))):
            if not file_recognition(get_row(data_array[header_start], "override")[header_pos], "bool"):
                ticker += 1

        if ticker > len(data_array[header_start]) * 3 / 4:
            header_start += 1
        else:
            boolean = True

    # Nested array to hold the values outputted. Returned at the end of the function
    new_data = create_nested_array(len(get_row(data_array[header_start], "override")))

    for row in range(header_start, len(data_array)):
        checkout = get_row(data_array[row])
        if row == header_start:
            for header_index in range(len(checkout)):
                checkout[header_index] = file_recognition(checkout[header_index])

        if checkout is not None:
            for i in range(0, len(checkout)):
                new_data[i].append(checkout[i])
    return new_data

# Creates a nested array
def create_nested_array(size):
    array = [None] * size
    for x in range(0, size):
        array[x] = []
    return array

# Finds all files within a folder. Argument takes a directory path.
def seek(path):
    files = os.listdir(path)
    return files

# Reads the TXT file in its entirety. Not really formatted well, outputs as a list.
def read_txt_file(name):
    date_mode = 3
    try:
        # Uses the newer "with" construct to close the file automatically. Works with everything not Excel
        with open(name) as f:
            data = f.readlines()
    except UnicodeDecodeError:
        # Excel file parsing.
        data = []
        book = xlrd.open_workbook(name)
        date_mode = book.datemode
        for wsnum in range(0, book.nsheets):
            ws = book.sheet_by_index(wsnum)
            if wsnum == 0:
                start = 0
            else:
                start = 1
            for rows in range(start, ws.nrows):
                temp = ws.row_values(rows)
                try:
                    for x in range(len(temp)):
                        temp[x] = temp[x].strip('"')
                except AttributeError:
                    pass
                data.append(temp)

    return data, date_mode

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

    if counter > 3 * len(newstring)/4 and expected != "override":
        return None
    else:
        return newstring

def general_parse(column, header, date_mode):
    col = column
    first = []
    for z in range(len(col)):
        first.append(col[z][0])

    for q in range(len(header)):
        if header[q] not in first:
            temp = ["NULL"] * len(col[0])
            temp[0] = header[q]
            if q != len(col) - 1:
                col.insert(q, temp)
            else:
                col.append(temp)

    for index in range(0, len(col)):
        if col[index][0] == 'Invoice Date' or col[index][0] == 'Payment Date' or col[index][0] == 'Entered Date':
            #Runs through all of the headers in the document, then cycles through all of the. Break out of it after.
            for i in range(len(col[index])):
                col[index][i] = timemachine(col[index][i], date_mode)
        # Strips out all of the leading zeros in the Vendor ID
        elif col[index][0] == 'Number':
            for i in range(len(col[index])):
                if type(col[index][i]) == str:
                    col[index][i] = col[index][i].lstrip("0")

    return col

# Mainly hit and miss. I wish I could make this a bit smarter, but for now, its just going to try each date format
# If it works, it works. If there's an error, it tries the next date format. I don't like this because there is the
# slim possibility that a date will be able to work for MM/DD/YYYY and YYYY/MM/DD.
def timemachine(date, date_mode):
    # For properly formatted dates, not the weird crazy stuff that most companies standardize. Only works for floats,
    # so this function must be the first to be run.
    try:
        temp = xlrd.xldate_as_tuple(date, date_mode)
        timeholder = datetime.date(temp[0], temp[1], temp[2])
        return timeholder.strftime("%m/%d/%Y")
    except:
        pass
    try:
        # Converts dates from floats to strings. They won't ever be needed as floats again.
        dates = str(int(date))
    except ValueError:
        dates = date
    try:
        timeholder = time.strptime(dates, "%Y%m%d")
        return time.strftime("%m/%d/%Y", timeholder)
    except ValueError:
        pass
    try:
        timeholder = time.strptime(dates, "%m%d%Y")
        return time.strftime("%m/%d/%Y", timeholder)
    except ValueError:
        pass
    try:
        timeholder = time.strptime(dates, "%Y/%m/%d")
        return time.strftime("%m/%d/%Y", timeholder)
    except ValueError:
        pass
    return date

# Runs the main, after establishing that this is not a library.
if __name__ == "__main__":
    main()
