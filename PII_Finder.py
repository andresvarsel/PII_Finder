
"""
This program is created with the purpose of searching for PII (Personal Identifiable Information).
Specifically email addresses, names of persons, personal id numbers, monetary card numbers.
"""

__author__ = "Andre Sele"

# --- IMPORT SECTION ---

import csv
import sqlite3

import docx
import magic
import os
import plum.exceptions
import re
import spacy
import time

from exif import Image
from openpyxl import load_workbook
from pdfminer.high_level import extract_text
from spacy.language import Language
from spacy_language_detection import LanguageDetector

import en_core_web_md
import nb_core_news_lg

#import PyPDF2 # remove after modification for pdfminer!


localtime = time.asctime(time.localtime(time.time()))
utctime = time.asctime(time.gmtime(time.time()))
# Class to hold all hits (matches for search criteria).
class Hits:
    """
    | Class variables for hits.
    """
    def __init__(self):
        self.Hits_li_key = []
        self.Hits_li_email = []
        self.Hits_li_idNum = []
        self.Hits_li_cardNum = []
        self.Hits_li_gps = []
        self.Hits_li_names = []
        self.Hits_li_num = ''

# Variable for accessing class Hits
Hits_ = Hits()

# Must be re-written for all instance variables.
# Write hits to csv file
def hits_to_csv():
    """
    | Write search results to csv file.
    """
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


def get_lang_detector(nlp, name):
    """
    | Spacy LanguageDetector class
    """
    return LanguageDetector(seed=42)


def state_language(text: str) -> str:
    """
    | Detects language of text strings.
    | Return appropriate spaCy model.
    """
    nlp_model = spacy.load("en_core_web_sm")
    Language.factory("language_detector", func=get_lang_detector)
    nlp_model.add_pipe('language_detector', last=True)

    doc = nlp_model(text)
    language = doc._.language
    id_lang = language.get('language')
    if id_lang == 'no': # Norwegian
        mod = 'nb_core_news_lg'
    else:
        mod = 'en_core_web_md'
    print(mod)
    return mod


def has_digit(inp_str) -> bool:
    """
    | Return True if string has digit.
    """
    return any(char.isdigit() for char in inp_str)


def only_letter_and_hyphen() -> str:
    """
    | Check that string contain letters and hyphen only.
    | Allow string of length 2-26.
    """
    to_match = r'^[æøåÆØÅa-zA-Z\s.-]{2,26}$'
    return to_match


def decrease_false_pos() -> list:
    """
    | Attempt at reducing false positives by
    | specifying patterns for human names.
    """
    patterns = [r'^[æøåÆØÅa-zA-Z]+$', r'^[æøåÆØÅa-zA-Z.]+[\s-][æøåÆØÅa-zA-Z.]+$', r'^[æøåÆØÅa-zA-Z]{2,10}[\s-]{1}[æøåÆØÅa-zA-Z.]{1,10}[\s-]{1}[æøåÆØÅa-zA-Z.]{1,10}$']
    return patterns


def convert_to_bytes(x: str) -> bytes:
    """
    | Convert input string to bytes
    """
    x = x.encode('utf-8')
    return x


def key_matcher() -> list:
    """
    | Keywords for search.
    | List of keywords is converted to
    | lower-case before return.
    """
    keyword_li = ['Horse', 'exception', 'andre sele', 'problem', 'OLaV', 'eTTeRnavneNe', 'johansen', 'PNg', 'ÅDne']
    lower_li = []
    for i in keyword_li:
        lower_li.append(i.casefold())
    return lower_li


def re_mail_matcher() -> str:
    """
    | For email address search.
    | Regular expression for email addresses.
    """
    re_mail = [r'[æøåÆØÅa-zA-Z0-9+._-]+@[æøåÆØÅa-zA-Z0-9._-]+\.[æøåÆØÅa-zA-Z0-9_-]+']
    return re_mail


def re_idNum_matcher() -> list:
    """
    | List of regular expressions for personal id numbers.
    | Match standard format for Nordic countries, Poland, UK, US.
    """
    re_idNum = [r'\b\d{11}\b',
                r'\b[a-ceghj-npr-tw-zA-CEGHJ-PR-TW-Z]{2}(?:\d){6}[a-dA-D]?\b',
                r'\b\d{3}\-\d{2}\-\d{4}\b', r'\b\d{11}\d', r'\b\d{6}\-\d{4}\b', r'\b\d{6}\-\d{3}[a-zA-Z]\b']
    return re_idNum

def re_cardNum_matcher() -> list:
    """
    | Regex for standard monetary card number format.
    """
    re_cardNum = [r'\b\d{4}\-\d{4}\-\d{4}\-\d{4}\b'] # include more!
    return re_cardNum


def name_finder(text, path): # INCLUDE PATH FOR MATCHES!
    """
    | Use spaCy to find human names.
    """

    nlp = spacy.load(state_language(text))
    doc = nlp(text)
    #nlp = spacy.load("nb_core_news_sm")

    #name_li = []

    #with open("data/NameList.txt", "r") as f:
    #    lines = f.readlines()
    #    for line in lines:
    #        name_li.append(line)

    #ruler = nlp.add_pipe("entity_ruler", after="ner")

    #name_li = [item.strip() for item in name_li]

    #patterns = []
    #for name in name_li:
    #    pattern = {"label": "PERSON", "pattern": name}
    #    patterns.append(pattern)

    per_li = []
    #ruler.add_patterns(patterns)
    #doc = nlp(text)

    for ent in doc.ents:
        if ent.label_ == "PERSON" or ent.label_== 'PER':#and bool(re.search(only_letter_and_hyphen(), str(ent))) == True\
                #and len(str(ent)) > 3:
            per_li.append(ent)

    # print(len(set(per_li)))
    per_li = [str(item) for item in per_li]
    per_li = list(set(per_li))
    per_li = sorted(per_li)
    for i in per_li:
        Hits_.Hits_li_names.append(i + ', ' + path)
    #print(per_li)
    # print(len(set(per_li)))


'''
# Find names with spacy (nltk).
def name_finder(text):
    names = []
    #one_name = r'^[æøåÆØÅa-zA-Z]{2,14}$'
    #two_names = r'^[æøåÆØÅa-zA-Z.]+[\s-][æøåÆØÅa-zA-Z.]+$'
    #three_names = r'^[æøåÆØÅa-zA-Z]{2,10}[\s-]{1}[æøåÆØÅa-zA-Z.]{1,10}[\s-]{1}[æøåÆØÅa-zA-Z.]{1,10}$'
    #regex_list = [one_name, two_names, three_names]
    nlp = spacy.load("en_core_web_sm")
    with open("data/NameList.txt", "r") as f:
        lines = f.readlines()
        for line in lines:
            names.append(line)
    names = [item.strip() for item in names]
    ruler = nlp.add_pipe("entity_ruler", after="ner")
    patterns = []
    for name in names:
        pattern = {"label": "PERSON", "pattern": name}
        patterns.append(pattern)
    ruler.add_patterns(patterns)
    doc = nlp(text)
    #per_li = []
    for ent in doc.ents:
        if ent.label == "PERSON":
            Hits_.Hits_li_names.append(ent)
    #per_li = [str(item) for item in per_li]
    #per_li = list(set(per_li))
    #per_li = sorted(per_li)
    #for i in per_li:

        #Hits_.Hits_li_names.append(i)

    #doc = nlp(text)
    #for x in doc.ents:
    #    #if x.label_ == 'PERSON' and bool(re.search(only_letter_and_hyphen(), str(x))) == True:
    #    for y in decrease_false_pos():
    #        if x.label_ == 'PERSON':
    #            hit = str(x)
    #            for regex in regex_list:
    #                res = re.findall(regex, hit)
    #                if res:
    #                    for r in res:
    #                        Hits_.Hits_li_names.append(r)

            #if bool(re.match(only_letter_and_hyphen(), str(x))) == True:
            #hit = str(x)
            #Hits_.Hits_li_names.append(hit)



        else:
            pass
'''
# EXPAND TO INCLUDE TIME AND DATE OF CREATION!
def gps_coord(File_Name):
    """
    | Check for gps coordinates in image files.
    """
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
    """
    | Extract text from excel files for search/match process.
    """
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
    name_finder(text, pathpath)

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


# USE PDF MINER INSTEAD!!!!!!!!!!!!!
# Extract hits from files of ftype: application/pdf
'''
def pdf_reader(File_Name):

    file_name = File_Name
    Text = et(file_name)

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
        #name_finder(str(Text))
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
'''

def pdf_reader(file_name):
    pathpath = os.path.normpath(file_name)
    pdf = file_name
    Text = extract_text(pdf)
    name_finder(Text, pathpath)

    for i in key_matcher():
        # Use re to search for items in "matcher"
        ResSearch = re.findall(i.casefold(), Text.casefold())  # case insensitive match!
        # print(ResSearch)

        # If matches are found
        if ResSearch:
            # Insert matches into match_li
            hit = i + ', ' + pathpath
            # print("This is hit", hit)
            # print(f'Match for string:"{i}", Path = {pathpath}')
            Hits_.Hits_li_key.append(hit)

        else:
            continue

    for i in re_mail_matcher():
        #

        ResSearch = re.findall(i.casefold(), Text.casefold())  # make case insensitive

        if ResSearch:
            for i in ResSearch:
                hit = i + ', ' + pathpath
                Hits_.Hits_li_email.append(hit)

            #re_hit = ''.join(ResSearch), pathpath

            # .group(0)
            # print(re_hit)
            # print(ResSearch.group(0))
            # match_li.append(i + '---' + pathpath)

            #Hits_.Hits_li_email.append(re_hit)

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
    """
    | Extract text from docx files for search/match process.
    """
    file_name = File_Name
    pathpath = os.path.normpath(file_name)
    doc = docx.Document(file_name)
    Text = []
    for para in doc.paragraphs:
        Text.append(para.text)
    Text = '\n'.join(Text)
    name_finder(Text, pathpath)
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

#THIS WILL CHECK ALL TABLE NAMES AND READ ROWS FROM ALL TABLES!!!
def db_reader(File_Name):
    """
    | Connect to and read rows of database tables.

    """
    file_name = File_Name
    pathpath = os.path.normpath(file_name)

    # connect to db
    with sqlite3.connect(file_name) as connection:
        c = connection.cursor()
        for tables in c.execute("SELECT name FROM sqlite_master WHERE type='table';"):
            for table in tables:
                c.execute(f"SELECT * FROM {table}")
                Text = c.fetchall()
                Text = str(Text)
                name_finder(Text, pathpath)
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

                    # print(i)
                    # print(type(i))
                    res = re.findall(i, Text)
                    if res:
                        for i in res:
                            Hits_.Hits_li_email.append(i + ", " + pathpath)
                        # Insert matches into match_li
                        # Hits_.Hits_li.append(str(i) + ", " + pathpath)
                        # continue

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
    """
    | Standard file opener and reader.
    | Open and read files in byte-form (mode=rb).
    """
    file_name = File_Name
    pathpath = os.path.normpath(file_name)
    match_li = []
    fn = open(file_name, mode='r')
    tn = fn.read()
    name_finder(tn, pathpath)
    fn.close()
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
                hit = i.decode() + ', ' + pathpath
                Hits_.Hits_li_cardNum.append(hit)

    f.close()


# Iterates through directories
def walker(Directory):
    """
    | Walks directories and subdirectories to execute search.
    """
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
            elif ftype == 'application/x-sqlite3':
                db_reader(paths)

            else:
                read_file(paths)





    print()
    file_num = count
    Hits_.Hits_li_num = file_num



start = time.time()

if __name__ == "__main__":
    walker('E:\Iter_open_test\Atesting')  # Enter drive/directory to search here!!!


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
print("Names")
print(len(Hits_.Hits_li_names))
for hit in Hits_.Hits_li_names:
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





