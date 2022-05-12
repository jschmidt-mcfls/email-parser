from xlwt import Workbook

workbook = Workbook()

filename = 'ShoutbombMarch2022.txt'

with open(filename, 'r') as f:
    email = f.read()

email = email.split("=TOTALS=")[0]

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

queries2 = {
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

libraries = {'Hales Corners': queries,
             'Whitefish Bay': queries,
             'Shorewood': queries,
             'Cudahy': queries,
             'North Shore': queries,
             'Brown Deer': queries,
             'Tippecanoe': queries,
             'St. Francis': queries,
             'West Allis': queries,
             'Wauwatosa': queries,
             'Oak Creek': queries,
             'West Milwaukee': queries,
             'King': queries,
             'Greendale': queries,
             'Greenfield': queries,
             'East': queries,
             'South Milwaukee': queries,
             'Franklin': queries,
             'Central': queries,
             'Center St.': queries,
            }


def get_data(data):
    for line in data:
        for key in queries:
            if key in line:
                new_line = line.replace(key, '')
                new_line = new_line.replace(' = ', '')
                line = new_line
                queries[key] = int(line)
    return queries


for branch in email.split('Branch:: '):
    for library in libraries:
        if library in branch:
            libraries[library] = get_data(branch.splitlines())
            queries = queries2.copy()

# Import dictionary data to sheet
sheet1 = workbook.add_sheet('Totals by Branch')
libraryNames = list(libraries.keys())

for lib in libraries:
    sheet1.write(libraryNames.index(lib)+1, 0, lib)
  
workbook.save(filename.replace(".txt", ".xls"))
