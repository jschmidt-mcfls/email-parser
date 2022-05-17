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

totalsByBranch = workbook.add_sheet(f"Totals by Branch {newFilename.replace('Shoutbomb', '')}")
totalUsers = workbook.add_sheet("Total Registered Users")


# Import dictionary data to Totals by Branch sheet
def parse(what):
    valuesList = []
    for line in what.splitlines():
        for key in queries.copy():
            if key in line:
                newLine = line.replace(key, "")
                newLine = newLine.replace(" = ", "")
                valuesList.append(int(newLine))
    return valuesList
    
# First Sheet
email_text = splittedEmail.split("=TOTALS=")[0]
for query in queries:
    totalsByBranch.write(int(queries.index(query)+1), 0, query)
row = 0
column = 0
for branch in email_text.split("Branch:: "):
    for library in libraries:
        row = 0
        if library in branch:
            column += 1
            totalsByBranch.write(0, column, library)
            libQueries = parse(branch)
            for query in libQueries:
                row += 1
                totalsByBranch.write(row, column, query)

row = 0 
column += 1
totals = parse(splittedEmail.split("=TOTALS=")[1])
totalsByBranch.write(row, column, "Totals")
for query in totals:
    row += 1
    totalsByBranch.write(row, column, query)

# Save workbook
workbook.save(filename.replace(".txt", ".xls"))
print("Saved Successfully...")
