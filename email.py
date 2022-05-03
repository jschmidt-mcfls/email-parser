with open('ShoutbombMarch2022.txt', 'r') as f:
    email = f.readlines()
    
libraries = []
dictionary = {
    'Hold notices sent for the month' : 0,
    'Hold cancel notices sent for the month' : 0,
    'Overdue notices sent for the month' : 0,
    'Overdue items eligible for renewal, notices sent for the month' : 0,
    'Overdue items ineligible for renewal, notices sent for the month' : 0,
    'Overdue items renewed successfully by patrons for the month' : 0,
    'Overdue items unsuccessfully renewed by patrons for the month' : 0,
    'Renewal notices sent for the month' : 0,
    'Items eligible for renewal notices sent for the month' : 0,
    'Items ineligible for renewal notices sent for the month' : 0,
    'Items renewed successfully by patrons for the month' : 0,
    'Items unsuccessfully renewed by patrons for the month' : 0,
}

for line in email:
    if 'Branch:: ' in line:
        print(line)

# https://www.geeksforgeeks.org/working-with-excel-spreadsheets-in-python/