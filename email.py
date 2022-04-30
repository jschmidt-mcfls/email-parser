# open email text file and put it into a variable
with open('ShoutbombMarch2022.txt', 'r') as f:
    msg_lines = f.readlines()

for l in msg_lines:
    if "::" in l:
        print(l) 
# https://datatofish.com/convert-text-file-to-csv-using-python-tool-included/  