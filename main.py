from hashlib import new
from xlwt import Workbook
from time import sleep
from os import system

# Intro Animation
for _ in range(0, 3):
    system("clear")
    print("//// Mom's Program ////")
    sleep(.2)
    system("clear")
    print("~~~~ Mom's Program ~~~~")
    sleep(.2)
    system("clear")
    print("\\\\\\\\ Mom's Program \\\\\\\\")
    sleep(.2)
    system("clear")
    print("|||| Mom's Program ||||")
    sleep(.2)

# Try to import file
found = False
while not found:
    filename = "ShoutbombMarch2022.txt"
    try:
        with open(filename, "r") as f:
            email = f.read()
        found = True
    except FileNotFoundError:
        print("File not found...")

# Variables
newFilename = filename.replace(".txt", "")
newFilename = newFilename.replace('Shoutbomb', '')

queries = [
    "Hold notices sent for the month",
    "Hold cancel notices sent for the month",
    "Overdue notices sent for the month",
    "Overdue items eligible for renewal, notices sent for the month",
    "Overdue items ineligible for renewal, notices sent for the month",
    "Overdue items renewed successfully by patrons for the month",
    "Overdue items unsuccessfully renewed by patrons for the month",
    "Renewal notices sent for the month",
    "Items eligible for renewal notices sent for the month",
    "Items ineligible for renewal notices sent for the month",
    "Items renewed successfully by patrons for the month",
    "Items unsuccessfully renewed by patrons for the month",
    "Totals?"
    ]

libraries = [
    "Atkinson", "Bay View", "Villard", "Wash Park", "Capitol",
    "Mitchell St.", "Zablocki", "Center St.",
    "Hales Corners", "Whitefish Bay", "Shorewood", "Cudahy",
    "North Shore", "Brown Deer", "Tippecanoe", "St. Francis",
    "Good Hope", "West Allis", "Wauwatosa", "Oak Creek",
    "West Milwaukee", "King", "Greendale", "Greenfield",
    "East", "South Milwaukee", "Franklin", "Central"
    ]

workbook = Workbook()
splittedEmail = email.split("=TOTALS BY BRANCH=")[0]


def parse(data, query):
    valuesList = []
    for line in data.splitlines():
        for key in query:
            if key in line:
                newLine = line.replace(key, "")
                newLine = newLine.replace(" = ", "")
                valuesList.append(int(newLine))
    return valuesList


# First Sheet
totalsByBranch = workbook.add_sheet(f"Totals {newFilename}")

emailText = splittedEmail.split("=TOTALS=")[0]
for query in queries:
    totalsByBranch.write(int(queries.index(query)+1), 0, query)
row = 0
column = 0
for branch in emailText.split("Branch:: "):
    for library in libraries:
        row = 0
        if library in branch:
            column += 1
            totalsByBranch.write(0, column, library)
            libQueries = parse(branch, queries.copy())
            for query in libQueries:
                row += 1
                totalsByBranch.write(row, column, query)

row = 0 
column += 1
totals = parse(splittedEmail.split("=TOTALS=")[1], queries.copy())
totalsByBranch.write(row, column, f"Totals")
for query in totals:
    row += 1
    totalsByBranch.write(row, column, query)

# Second Sheet
#make that shit parse
totalTexts = workbook.add_sheet(f"Text Notices {newFilename}")
emailText = splittedEmail.split("=TOTALS OF REGISTERED PATRON BY BRANCH=")[1]
print(emailText)


# Third Sheet
totalUsers = workbook.add_sheet(f"Patrons Registered for Text Notices {newFilename}")




# Save workbook
#workbook.save(filename.replace(".txt", ".xls"))
print("Saved Successfully...")
