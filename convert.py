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
import cProfile

import os
import datetime
import xlrd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import time
import csv
import shutil

def main():
    new_header = ['Supplier Name', 'Supplier Number', 'Reference', 'Amount', 'Currency', 'Invoice Date', 'Payment Date',
                  'Entered Date']
    print("\nYour new files will be saved in a folder on your Desktop called 'Target'")

    home = os.path.expanduser('~') + "/Desktop/"
    fo = Folder(home, "Source/", "Target/", "Problem/")

    cycle(new_header, fo)

# Cycles through every single folder in the path, converting each file to an excel file.
def cycle(header, folder):
    file_names = folder.seek()

    for x in range(0, len(file_names)):
        print('Cycle start')
        try:
            case = None
            check = process.extractOne(os.path.splitext(file_names[x])[0], ['SummaHealth'])
            if check[1] > 85:
                case = "Summa"

            fi = File(folder.source, file_names[x], header, case)

            convert(fi.convert(), fi.filename, folder.target)
        except:
            # Move file to the problem folder.
            shutil.copy("%s%s" % (folder.source, file_names[x]), "%s%s" % (folder.problem, file_names[x]))

    for file in file_names:
        os.remove("%s%s" % (folder.source, file))

# Moves the columns of data into a TXT file, delimited with '\t'
def convert(rows, save, pathway):
    now = datetime.date.today().strftime("%m.%d.%y")

    with open("%s/%s (%s).txt" % (pathway, save, now), 'w', newline='') as f:
        writer = csv.writer(f, delimiter='\t')
        writer.writerows(rows)

    print("Excelled")

# File class. Deals with everything that manipulates a file. Converter bulk.
class File:
    def __init__(self, source, name, header, case):
        self.raw_name = name
        self.location = "%s%s" % (source, name)
        self.filename = os.path.splitext(name)[0]
        self.raw_data, self.date_mode = self.read_txt_file()
        self.header = header
        self.case = case

    def convert(self):
        return self.colify()

    # Assumes that simplify has been run before this function. In other words, each column exists within the final header.

    # File recognition program. This automates the process of inputting the headers. Because there are only a set amount
    # of header formats, I can map the known formats to a hash map, with the keys being the source file header type, and
    # the output being the target output format.

    def order(self, col):
        ordered_data = []
        # Cycles through each header
        for header_index in range(len(self.header)):
            col_index = 0
            # Cycles through each column, while checking to see if the column header is equal to the header above.
            # If a match is found, the while loop is broken, and the code continues on.
            while col_index < len(col) and not col[col_index][0] == self.header[header_index]:
                col_index += 1
            ordered_data.append(col[col_index])
        if len(col) < len(self.header):
            raise ValueError
        return ordered_data

    # Mode determines if the file_recognition function will output a boolean showing success or failure, or if it will
    # output a new header.

    def file_recognition(self, unit, case, mode="rewrite"):
        header_hash = {'Supplier Name': ['Vendor Name', 'Name1', 'Name', 'Vname', 'Vendor'],
                       'Supplier Number': ['Vendor Id', 'Vendor ID', 'Vendor Number', 'Duns no', 'Vendor number',
                                           'Vendor #', 'Vend no'],
                       'Reference': ['Invoice Id', 'AP Invoice Number', 'Invoice', 'Doc Number', 'Credit Memo',
                                     'Reference Number', 'Invoice #', 'Inv no', 'Vendor Credit memo/reference',
                                     'Credit Memo Number', 'Invoice num', 'Invoice Number', 'Invoice/Reference'],
                       'Amount': ['Gross Amount', 'Gross Amt', 'TranInvAmt', 'cost amt', 'Invoice Amt',
                                  'Invoice Amount', 'Amount', 'Amt', 'Inv Amt', 'Credit memo Amount',
                                  'Invoice Amount SUM'],
                       'Currency': ['Currency', 'Curr', 'Inv Currency', 'InvCurrency', 'Currency USD',
                                    'Invoice Currency Code', 'Txn Currency Cd'],
                       'Invoice Date': ['AP Invoice Date', 'Invoice Date', 'doc date', 'Invoice Dte', 'Document Date',
                                        'CreditMemo Date', 'Credit Memo Date', 'Invoice Dt', "InvoiceDte", 'Inv date'],
                       'Payment Date': ['AP Check Date', 'Payment Date', 'pay due date', 'AP Payment Due Date',
                                        'Clear Date', 'Clearing date', 'Date Processed', 'PaymentDate',
                                        'Check date', 'Pymt Date', 'Accounting Dt'],
                       'Entered Date': ['Date added', 'Post Date', 'Entered Date', 'Posting Date', 'Posting Date',
                                        'Create Date', 'ReconDate', 'entry date', 'Create Dt', 'Payment Entry Date',
                                        'Entered Date', 'Clearing Date']}
        if case == "Summa":
            header_hash['Supplier Name'] = ['Vendor Name']
            header_hash['Supplier Number'] = ['Vendor']

        unit = unit.replace('_', ' ')
        # If unit has a match within the table, this automatically cuts out of the for loop and returns the key.
        # Else, it sets the errorCounter value to 1.
        for key in header_hash:
            counter = 0
            while not self.fuzzy(unit, header_hash[key][counter]) and counter < len(header_hash[key]) - 1:
                counter += 1
            if self.fuzzy(unit, header_hash[key][counter]):
                return key
        # program would never run out here if the unit was located in the hash map.

        if mode == "rewrite":
            return unit
        else:
            return False

    def fuzzy(self, text, aux, benchmark=93):
        text = text.lower()
        aux = aux.lower()
        r = (fuzz.ratio(text, aux))
        if r >= benchmark:
            return True
        else:
            return False

    def simplify(self, data):
        new_data = []

        for index in range(len(data)):
            if data[index][0] in self.header:
                new_data.append(data[index])
        if len(new_data) == 0:
            print("Simplification fail")
            raise SyntaxError
        return self.rowify(self.order(self.general_parse(new_data)))

    def rowify(self, col):
        new_row = []
        for row_index in range(1, len(col[0])):
            temp = []
            for col_index in range(0, len(col)):
                temp.append(col[col_index][row_index])
            new_row.append(temp)
        return new_row
    # Returns the columns of a TXT file, as a nested array. Automatically modifies the heading.

    def colify(self):
        header_start = 0
        boolean = False
        while not boolean:
            ticker = 0
            for header_pos in range(len(self.get_row(self.raw_data[header_start], "override"))):
                if not self.file_recognition(self.get_row(self.raw_data[header_start], "override")[header_pos], None,
                                             "bool"):
                    ticker += 1

            if ticker > len(self.raw_data[header_start]) * 3 / 4:
                header_start += 1
            else:
                boolean = True

        # Nested array to hold the values outputted. Returned at the end of the function
        new_data = self.create_nested_array(len(self.get_row(self.raw_data[header_start], "override")))

        for row in range(header_start, len(self.raw_data)):
            checkout = self.get_row(self.raw_data[row])
            if row == header_start:
                for header_index in range(len(checkout)):
                    checkout[header_index] = self.file_recognition(checkout[header_index], self.case)

            if checkout is not None:
                for i in range(0, len(checkout)):
                    new_data[i].append(checkout[i])
        return self.simplify(new_data)

    # Creates a nested array
    def create_nested_array(self, size):
        array = [None] * size
        for x in range(0, size):
            array[x] = []
        return array
    # Finds all files within a folder. Argument takes a directory path.

    # Reads the TXT file in its entirety. Not really formatted well, outputs as a list.
    def read_txt_file(self):
        date_mode = 3
        try:
            # Uses the newer "with" construct to close the file automatically. Works with everything not Excel
            with open(self.location) as f:
                data = f.readlines()
        except UnicodeDecodeError:
            # Excel file parsing.
            data = []
            book = xlrd.open_workbook(self.location)
            date_mode = book.datemode
            for wsnum in range(0, book.nsheets):
                ws = book.sheet_by_index(wsnum)
                if wsnum == 0:
                    start = 0
                else:
                    # Possible issue that happens if the header has multiple rows
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
    def get_row(self, data, expected=None):
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

    def general_parse(self, column):
        col = column
        first = []
        for z in range(len(col)):
            first.append(col[z][0])

        for q in range(len(self.header)):
            if self.header[q] not in first:
                if self.header[q] == 'Supplier Number':
                    print("Issue with the headers. Program did not find a Vendor ID.")
                    raise ValueError
                temp = [""] * len(col[0])
                temp[0] = self.header[q]
                if q != len(col) - 1:
                    col.insert(q, temp)
                else:
                    col.append(temp)

        for index in range(0, len(col)):
            if col[index][0] == 'Invoice Date' or col[index][0] == 'Payment Date' or col[index][0] == 'Entered Date':
                # Runs through all of the headers in the document, then cycles through all of the. Break out of it after
                for i in range(len(col[index])):
                    col[index][i] = self.timemachine(col[index][i], self.date_mode)
            if col[index][0] == 'Currency':
                for i in range(len(col[index])):
                    if str(col[index][i]).isspace() or not col[index][i]:
                        col[index][i] = "USD"
            if col[index][0] == 'Supplier Number' or col[index][0] == 'Reference':
                for i in range(len(col[index])):
                    try:
                        col[index][i] = str(col[index][i]).rstrip('.0')
                    except:
                        pass
        return col

    # Mainly hit and miss. I wish I could make this a bit smarter, but for now, its just going to try each date format
    # If it works, it works. If there's an error, it tries the next date format. I don't like this because there is the
    # slim possibility that a date will be able to work for MM/DD/YYYY and YYYY/MM/DD.
    def timemachine(self, date, date_mode):
        dates = date
        # For properly formatted dates, not the weird crazy stuff that most companies standardize. Only works for floats
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
            dates = dates.replace('-', '')
        except:
            pass
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

# Folder class. Deals with everything that has to do with the file locations.
class Folder:
    def __init__(self, home, source, target, problem):
        self.home = home
        self.source = "%s%s" % (home, source)
        self.target = self.direct(target)
        self.problem = self.direct(problem)

    # Creates a new path to store the created excel files in.
    @staticmethod
    def direct(new_folder):
        home = os.path.expanduser('~')
        pathway = "%s/Desktop/%s" % (home, new_folder)
        if not os.path.isdir(pathway):
            os.mkdir(pathway)
        return pathway

    def seek(self):
        files = os.listdir(self.source)
        return files

# Runs the main, after establishing that this is not a library.
if __name__ == "__main__":
    # main()
    cProfile.run('main()')
