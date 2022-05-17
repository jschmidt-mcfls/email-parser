# Imports
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
    filename = "ShoutbombApril2022.txt"
    try:
        with open(filename, 'r') as f:
            email = f.read()
        found = True
    except FileNotFoundError:
        print("File not found...")

# Variables
new_filename = filename.replace(".txt", "")
queries = {
    'Hold notices sent for the month': 0,
    'Hold cancel notices sent for the month': 0,
    'Overdue notices sent for the month': 0,
    'Overdue items eligible for renewal, notices sent for the month': 0,
    'Overdue items ineligible for renewal, notices sent for the month': 0,
    'Overdue items renewed successfully by patrons for the month': 0,
    'Overdue items unsuccessfully renewed by patrons for the month': 0,
    'Renewal notices sent for the month': 0,
    'Items eligible for renewal notices sent for the month': 0,
    'Items ineligible for renewal notices sent for the month': 0,
    'Items renewed successfully by patrons for the month': 0,
    'Items unsuccessfully renewed by patrons for the month': 0,
    }

libraries = [
    'Atkinson', 'Bay View', 'Villard', 'Wash Park', 'Capitol',
    'Mitchell St.', 'Zablocki', 'Center St.',
    'Hales Corners', 'Whitefish Bay', 'Shorewood', 'Cudahy',
    'North Shore', 'Brown Deer', 'Tippecanoe', 'St. Francis',
    'Good Hope', 'West Allis', 'Wauwatosa', 'Oak Creek',
    'West Milwaukee', 'King', 'Greendale', 'Greenfield',
    'East', 'South Milwaukee', 'Franklin', 'Central'
    ]

workbook = Workbook()

# rename
totalsByBranch = workbook.add_sheet(f"Totals by Branch {new_filename.replace('Shoutbomb', '')}")
totals = workbook.add_sheet(f"Totals {new_filename.replace('Shoutbomb', '')}")


# Import dictionary data to Totals by Branch sheet
def firstSheet():
    global email
    email = email.split("=TOTALS=")[0]
    queryKeys = list(queries.keys())
    for key in queries.keys():
        totalsByBranch.write(queryKeys.index(key)+1, 0, key)
    for branch in email.split('Branch:: '):
        for library in libraries:
            if library in branch:
                y_pos = libraries.index(library)+1
                totalsByBranch.write(0, y_pos, library)
                new_queries = queries.copy()
                for line in branch.splitlines():
                    for key in new_queries:
                        if key in line:
                            new_line = line.replace(key, '')
                            new_line = new_line.replace(' = ', '')
                            line = new_line
                            new_queries[key] = int(line)

                for query in list(new_queries.keys()):
                    x_pos = list(new_queries.keys()).index(query)
                    totalsByBranch.write(x_pos+1, y_pos, new_queries[query])


def secondSheet():
    pass


# Run functions
firstSheet()
secondSheet()

# Save workbook
workbook.save(filename.replace(".txt", ".xls"))
print("Saved Successfully...")
