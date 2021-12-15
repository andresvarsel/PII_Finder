# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import spacy
import en_core_web_sm

import csv
import docx
import plum.exceptions
from exif import Image
import magic
from openpyxl import load_workbook
import os
import PyPDF2
import re
import time

localtime = time.asctime(time.localtime(time.time()))
utctime = time.asctime(time.gmtime(time.time()))
# Class to hold all hits (matches for search criteria).
class Hits:
    def __init__(self):
        self.Hits_li_key = []
        self.Hits_li_email = []
        self.Hits_li_idNum = []
        self.Hits_li_cardNum = []
        self.Hits_li_gps = []
        self.Hits_li_names = []
        self.Hits_li_num = ''

# Variable for class Hits
Hits_ = Hits()

# Must be re-written for all instance variables.
# Write hits to csv file
def hits_to_csv():
    time_desc = "UTC Time: "
    time = utctime
    with open('E:\hits.csv', 'w+', newline='') as file:
        columns = ['Match', 'Full_path']
        writer = csv.DictWriter(file, fieldnames=columns)

        writer.writeheader()
        writer.writerow({'Match': '', 'Full_path': ''})
        writer.writerow({f'Match': {time_desc}, 'Full_path': {time}})
        writer.writerow({'Match': '', 'Full_path': ''})
        for hits in set(Hits_.Hits_li):
            hits = str(hits)
            hits = hits.split(',')
            writer.writerow({f'Match': {hits[0]}, 'Full_path': {hits[1]}})
    file.close()


# Check if string has digit. Used to minimize false positives when finding names (PERSON labels).
def has_digit(inp_str):
    return any(char.isdigit() for char in inp_str)


# Reduce false positives when searching for names (PERSON labels) with casy (nltk).
def only_letter_and_hyphen():
    to_match = r'^[æøåÆØÅa-zA-Z\s.-]+$'
    return to_match

# Encode string to utf-8
def convert_to_bytes(x):
    x = x.encode('utf-8')
    return x


# Lists of keywords to search for.
def key_matcher():
    keyword_li = ['Horse', 'exception', 'andre sele', 'problem', 'OLaV', 'eTTeRnavneNe', 'johansen', 'PNg', 'ÅDne']
    lower_li = []
    for i in keyword_li:
        lower_li.append(i.casefold())
    return lower_li


# Lists of regex to search for.
def re_mail_matcher():
    # Email address regex with æøå.
    re_mail = [r'[æøåÆØÅa-zA-Z0-9+._-]+@[æøåÆØÅa-zA-Z0-9._-]+\.[æøåÆØÅa-zA-Z0-9_-]+']
    return re_mail


def re_idNum_matcher():
    # ID num regex (Norway, Poland, UK, US, Iceland, Denmark, Sweden, Finland).
    re_idNum = [r'\b\d{11}\b',
                r'\b[a-ceghj-npr-tw-zA-CEGHJ-PR-TW-Z]{2}(?:\d){6}[a-dA-D]?\b',
                r'\b\d{3}\-\d{2}\-\d{4}\b', r'\b\d{11}\d', r'\b\d{6}\-\d{4}\b', r'\b\d{6}\-\d{3}[a-zA-Z]\b']
    return re_idNum

def re_cardNum_matcher():
    # Regex for standard monetary card number format.
    re_cardNum = [r'\b\d{4}\-\d{4}\-\d{4}\-\d{4}\b']
    return re_cardNum

# Find names with casy (nltk).
def name_finder(text):
    nlp = en_core_web_sm.load()
    doc = nlp(text)
    for x in doc.ents:
        if x.label_ == 'PERSON':
            if bool(re.match(only_letter_and_hyphen(), str(x))) == True:
                hit = str(x)
                Hits_.Hits_li_names.append(hit)



        else:
            pass

def gps_coord(File_Name):
    file_name = File_Name
    pathpath = os.path.normpath(file_name)
    lat = 'gps_latitude'
    lng = 'gps_longitude'
    # Errors are raised when using exif.Image on png files. See exception list.
    try:
        with open(file_name, 'rb') as img_file:
            img = Image(img_file)

            if img.has_exif:
                if lat and lng in img.list_all():
                    h_lat = img.gps_latitude
                    h_lng = img.gps_longitude
                    hit = "Lat:" + str(h_lat), "Long:" + str(h_lng), pathpath
                    Hits_.Hits_li_gps.append(str(hit))
            else:
                pass
    except (OSError, ValueError, plum.exceptions.UnpackError):
        pass


# Extract text etc from xlsx file to search for given values.
def xlsx_reader(File_Name):
    info_li = []
    file_name = File_Name
    pathpath = os.path.normpath(file_name)
    # Open xlsx file
    wb = load_workbook(file_name)
    # Read sheet
    ws = wb.active
    # Extract values from cells
    cells = (list(ws.rows))
    for cell in cells:
        for info in cell:
            if info.value != None:
                i = str(info.value)
                info_li.append(i)
    text = ' '.join(info_li)


    for i in key_matcher():
        if i in info_li:
            hit = i + ', ' + pathpath
            Hits_.Hits_li_key.append(hit)

    for i in re_mail_matcher():
        res = re.findall(i, text)
        if res:
            for i in res:
                hit = i + ', ' + pathpath
                Hits_.Hits_li_email.append(hit)

    for i in re_idNum_matcher():
        res = re.findall(i, text)
        if res:
            for i in res:
                hit = i + ', ' + pathpath
                Hits_.Hits_li_idNum.append(hit)

    for i in re_cardNum_matcher():
        res = re.findall(i, text)
        if res:
            for i in res:
                hit = i + ', ' + pathpath
                Hits_.Hits_li_cardNum.append(hit)

    name_finder(text)

# Extract hits from files of ftype: application/pdf
def pdf_reader(File_Name):

    file_name = File_Name

    # open the pdf file
    object = PyPDF2.PdfFileReader(file_name)
    pathpath = os.path.normpath(file_name)
    # get number of pages
    NumPages = object.getNumPages()

    # extract text and do the search
    for i in range(0, NumPages):
        PageObj = object.getPage(i)
        # Extract text from pdf
        Text = PageObj.extractText()
        name_finder(Text)
        Text = Text.casefold()
        #name_finder(Text)
        # Call on function "matcher" which contain list of search items
        for i in key_matcher():
            # Use re to search for items in "matcher"
            ResSearch = re.findall(i.casefold(), Text)# case insensitive match!
            #print(ResSearch)

            # If matches are found
            if ResSearch:
                # Insert matches into match_li
                hit = i + ', ' + pathpath
                #print("This is hit", hit)
                #print(f'Match for string:"{i}", Path = {pathpath}')
                Hits_.Hits_li_key.append(hit)

            else:
                continue

        for i in re_mail_matcher():
            #

            ResSearch = re.findall(i.casefold(), Text)  # make case insensitive

            if ResSearch:

                re_hit = ''.join(ResSearch), pathpath
                #.group(0)
                #print(re_hit)
                #print(ResSearch.group(0))
                #match_li.append(i + '---' + pathpath)
                Hits_.Hits_li_email.append(re_hit)
                # print(type(i))
                # print(f'Match for string:"{i}", Path = {pathpath}')


            else:
                # print(f'No match for the word: {i}')
                continue

        for i in re_idNum_matcher():
            res = re.findall(i, Text)
            if res:
                for i in res:
                    hit = i + ', ' + pathpath
                    Hits_.Hits_li_idNum.append(hit)

        for i in re_cardNum_matcher():
            res = re.findall(i, Text)
            if res:
                for i in res:
                    hit = i + ', ' + pathpath
                    Hits_.Hits_li_cardNum.append(hit)

# Extract hits from files of ftype: application/vnd.openxmlformats-officedocument.wordprocessingml.document
def docx_reader(File_Name):
    file_name = File_Name
    pathpath = os.path.normpath(file_name)
    doc = docx.Document(file_name)
    Text = []
    for para in doc.paragraphs:
        Text.append(para.text)
    Text = '\n'.join(Text)
    Text = Text.casefold()
    #print(Text)
    for i in key_matcher():

        ResSearch = re.findall(i, Text)
        # match_li = []
        if ResSearch:

            # Insert matches into match_li
            Hits_.Hits_li_key.append(str(i) + ", " + pathpath)
            continue

        else:
            # print(f'No match for the word: {i}')
            continue

    for i in re_mail_matcher():

        #print(i)
        #print(type(i))
        res = re.findall(i, Text)
        if res:
            for i in res:



                Hits_.Hits_li_email.append(i + ", " + pathpath)
            # Insert matches into match_li
            #Hits_.Hits_li.append(str(i) + ", " + pathpath)
            #continue

        else:
            # print(f'No match for the word: {i}')
            continue

    for i in re_idNum_matcher():
        res = re.findall(i, Text)
        if res:
            for i in res:
                hit = i + ', ' + pathpath
                Hits_.Hits_li_idNum.append(hit)

    for i in re_cardNum_matcher():
        res = re.findall(i, Text)
        if res:
            for i in res:
                hit = i + ', ' + pathpath
                Hits_.Hits_li_cardNum.append(hit)

# Read files and adds matches to match_li
def read_file(File_Name):
    file_name = File_Name
    pathpath = os.path.normpath(file_name)
    match_li = []

    f = open(file_name, mode='rb')
    t = f.read()

    # Search for keyword matches in t
    for i in key_matcher():
        # Add to match_li if match is found
        if convert_to_bytes(i) in t.lower():  # case sensitivity!!!
            Hits_.Hits_li_key.append(str(i) + ", " + pathpath)
            # print(i + " --- " + pathpath)

        else:
            continue

    # Search for regex matches in t
    for i in re_mail_matcher():
        i = i.encode()
        # Add to match_li if match is found
        res = re.findall(i, t, re.IGNORECASE)  # case sensitivity!!!
        # print(i, "is match for:", str(res), pathpath)
        if res:
            # print(res)

            for i in res:

                Hits_.Hits_li_email.append(str(i) + ", " + pathpath)
                # add_hit_to_li(str(i) + " !!! " + pathpath)
        else:
            continue

    for i in re_idNum_matcher():
        i = i.encode()
        res = re.findall(i, t, re.IGNORECASE)

        if res:
            for i in res:
                hit = str(i) + ', ' + pathpath
                Hits_.Hits_li_idNum.append(hit)

    for i in re_cardNum_matcher():
        i = i.encode()
        res = re.findall(i, t, re.IGNORECASE)
        if res:
            for i in res:
                hit = i + ', ' + pathpath
                Hits_.Hits_li_cardNum.append(hit)

    f.close()


# Iterates through directories
def walker(Directory):
    directory = Directory
    count = 0
    print("List of matching strings, and absolute file-path:")
    print()
    #print("Matches found for:")
    for subdir, dirs, files in os.walk(directory):

        for file in files:
            # print(file)
            # File-path from os
            paths = os.path.join(subdir, file)
            ftype = magic.from_file(paths, mime=True)
            count += 1
            # print(paths, ftype)
            # if "pdf" in ftype:
            #    pdf_reader(paths)

            # else:
            #    continue
            # Counts number of files processed

            if "pdf" in ftype:
                pdf_reader(paths)
            elif ftype == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
                docx_reader(paths)
            elif ftype == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
                xlsx_reader(paths)
            elif ftype[:5] == 'image':
                gps_coord(paths)

            else:
                read_file(paths)





    print()
    file_num = count
    Hits_.Hits_li_num = file_num



start = time.time()

if __name__ == "__main__":
    walker('E:\Iter_open_test')  # Enter drive/directory to search here!!!


#hits_to_csv()

print("Number of files searched:", Hits_.Hits_li_num)
#print("Number of Hits: ", len(set(Hits_.Hits_li_email))
print()
print("Hit List: ")
print()

print('Keywords:')
print()
for hit in set(Hits_.Hits_li_key):
    print(hit)
print()
print('Emails:')

for hit in set(Hits_.Hits_li_email):
    print(hit)

print()
print('ID Numbers:')
for hit in set(Hits_.Hits_li_idNum):
    print(hit)

print()
print("Card Numbers:")
for hit in set(Hits_.Hits_li_cardNum):
    print(hit)

print()
print("GPS Coordinates from Image files:")
for hit in Hits_.Hits_li_gps:
    print(hit)

print()
print("Names and plenty of false positives")
print(len(Hits_.Hits_li_names))
for hit in set(Hits_.Hits_li_names):
    print(hit)

print()
stop = time.time()
process_time = round(stop - start, 2)

# Print time used in seconds or minutes
if process_time > 60:
    process_time /= 60
    print("Minutes used: ", process_time)
else:
    print("Seconds used: ", process_time)


print("Local current time :", localtime)

print("UTC current time   :", utctime)





