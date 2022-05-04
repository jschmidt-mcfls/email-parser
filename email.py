with open('text.txt', 'r') as f:
    email = f.readlines()
    
libraries = []
queries = {
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
    for key in queries:
        if key in line:
            new_line = line.replace(key, '')
            new_line = new_line.replace(' = ', '')
            line = new_line
            queries[key] = int(line)

print(queries)

#use .replace()
# https://www.geeksforgeeks.org/working-with-excel-spreadsheets-in-python/