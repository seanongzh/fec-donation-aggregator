import requests
import re
from bs4 import BeautifulSoup
from openpyxl import load_workbook

DIRECTORY_URL = "https://directory.tufts.edu/searchresults.cgi"
WORKBOOK_NAME = "Working_Copy_MASTER.xlsx"
NAME_SHEET = "DirectoryResults"


# This script works on Excel Sheets with a single column in the A column of
# a unique listing of names. It searches for every name in the tufts directory,
# then strips out a department if it is found and places it in the B column next
# to the search query.

def getDirectoryPage(name):
    rawPage = requests.post(DIRECTORY_URL, data={"type": "Faculty", "search": name})
    soup = BeautifulSoup(rawPage.text, "html.parser")
    for child in soup.find_all(href=re.compile("department.cgi")):
        return child.contents[0].strip()

dataBook = load_workbook(WORKBOOK_NAME)
nameSheet = dataBook[NAME_SHEET]

for index in range(2, nameSheet.max_row + 1):
    currentName = nameSheet["A{0}".format(index)].value
    affiliation = getDirectoryPage(currentName)
    nameSheet["B{0}".format(index)] = affiliation

    # Rudimentary backup every 100 entries
    if index % 100 == 0:
        print("Progress: {:.2%}".format(index / (nameSheet.max_row + 1)))
        dataBook.save(WORKBOOK_NAME)

dataBook.save(WORKBOOK_NAME)
