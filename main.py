from xlwt import Workbook

# potential queries for totals
queries = {
    # old queries

    # "Hold text notices sent for the month": 0,
    # "Hold cancel notices sent for the month": 0,
    # "Overdue notices sent for the month": 0,
    # "Overdue items eligible for renewal, notices sent for the month": 0,
    # "Overdue items ineligible for renewal, notices sent for the month": 0,
    # "Overdue items renewed successfully by patrons for the month": 0,
    # "Overdue items unsuccessfully renewed by patrons for the month": 0,
    # "Renewal notices sent for the month": 0,
    # "Items eligible for renewal notices sent for the month": 0,
    # "Items ineligible for renewal notices sent for the month": 0,
    # "Items renewed successfully by patrons for the month": 0,
    # "Items unsuccessfully renewed by patrons for the month": 0,

    # new queries
    "Hold text notices sent for the month": 0,
    "Hold cancel notices sent for the month": 0,
    "Overdue text notices sent for the month": 0,
    "Overdue items eligible for renewal, text notices sent for the month": 0,
    "Overdue items ineligible for renewal, text notices sent for the month": 0,
    "Overdue (text) items renewed successfully by patrons for the month": 0,
    "Overdue (text) items unsuccessfully renewed by patrons for the month": 0,
    "Renewal text notices sent for the month": 0,
    "Items eligible for renewal text notices sent for the month": 0,
    "Items ineligible for renewal text notices sent for the month": 0,
    "Items (text) renewed successfully by patrons for the month": 0,
    "Items (text) unsuccessfully renewed by patrons for the month": 0
}

# dictionary of possible libraries in file
libraries = {
    "Atkinson": 0, "Bay View": 0, "Villard": 0, "Wash Park": 0, "Capitol": 0,
    "Mitchell St.": 0, "Zablocki": 0, "Center St.": 0,
    "Hales Corners": 0, "Whitefish Bay": 0, "Shorewood": 0, "Cudahy": 0,
    "North Shore": 0, "Brown Deer": 0, "Tippecanoe": 0, "St. Francis": 0,
    "Good Hope": 0, "West Allis": 0, "Wauwatosa": 0, "Oak Creek": 0,
    "West Milwaukee": 0, "King": 0, "Greendale": 0, "Greenfield": 0,
    "East": 0, "South Milwaukee": 0, "Franklin": 0, "Central": 0,
}

# create spreadsheet
workbook = Workbook()

# welcome statement
print('Noah Dinan | MCFLS email parser | input "exit" to exit')

# Try to import text file
while True:
    found = False
    filename = None
    while not found:
        filename = input('input file [XXXX-XX.txt]: ')

        if filename.lower() == "exit":
            exit("exiting...")
        else:
            try:
                with open(f"Input/{filename}", "r") as f:
                    email = f.read()
                found = True
            except FileNotFoundError:
                print("File not found...")

    splitEmail = email.split("=TOTALS BY BRANCH=")[0]

    # Parser
    def parse(data, query):
        for line in data.splitlines():
            for key in query.keys():
                if key in line:
                    new_line = line.split(" = ")
                    new_line = int(new_line[1])
                    query[key] = new_line
        return query


    # Totals Sheet
    totalsByBranch = workbook.add_sheet("Totals")

    emailText = splitEmail.split("=TOTALS=")[0]
    queriesList = list(queries.keys())
    for query in queriesList:
        totalsByBranch.write(int(queriesList.index(query)+1), 0, query)
    row = 0
    column = 0
    libraryCopy1 = libraries.copy()
    for branch in emailText.split("Branch:: "):
        for library in libraryCopy1:
            row = 0
            if library in branch:
                column += 1
                totalsByBranch.write(0, column, library)
                libQueries = parse(branch, queries.copy())
                for query in libQueries.values():
                    row += 1
                    totalsByBranch.write(row, column, query)

    row = 0
    column += 1
    totals = parse(splitEmail.split("=TOTALS=")[1], queries.copy())
    totalsByBranch.write(row, column, f"Totals")
    for query in totals.values():
        row += 1
        totalsByBranch.write(row, column, query)

    # Text Notices Sent sheet
    row = 1
    libraryCopy2 = libraries.copy()
    textNotices = workbook.add_sheet("Text Notices Sent")
    textNotices.write(0, 1, "Total Text Notices")
    splitEmail = email.split("=TOTALS BY BRANCH=")[1]
    emailText = splitEmail.split("=TOTALS OF REGISTERED PATRON BY BRANCH=")[0]
    values = parse(emailText, libraryCopy2)
    for line in emailText.splitlines():
        for library in libraryCopy2:
            if library in line:
                textNotices.write(row, 0, library)
                textNotices.write(row, 1, values[library])
                row += 1

    # Third Part
    row = 1
    registeredUsers = workbook.add_sheet("Registered Patrons")
    registeredUsers.write(0, 1, "Total Registered Patrons")
    emailText = splitEmail.split("=TOTALS OF REGISTERED PATRON BY BRANCH=")[1]
    emailText = emailText.split("These patron phone numbers")[0]

    # setup custom parsing
    libraryCopy3 = libraries.copy()
    for line in emailText.splitlines():
        for library in libraries.keys():
            if library in line:
                if "has" in line:
                    newLine = line.split(" has ")[1]
                    newLine = newLine.replace(" registered patrons for text notices", "")
                    newLine = int(newLine)
                    libraryCopy3[library] += newLine

    # write to sheet and remove duplication
    used = []
    for line in emailText.splitlines():
        for library in libraryCopy3.keys():
            if library in line and library not in used:
                registeredUsers.write(row, 0, library)
                registeredUsers.write(row, 1, libraryCopy3[library])
                used.append(library)
                row += 1

    # Save workbook
    workbook.save(f"Output/{filename.replace('.txt', '.xls')}")
    print("saved successfully.")
