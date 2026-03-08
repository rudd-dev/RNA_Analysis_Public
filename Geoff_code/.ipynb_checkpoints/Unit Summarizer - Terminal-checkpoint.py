import csv
import sys
import os
import xlsxwriter

def getChains(input_string):
    result_list = input_string.split(',')

    # Strip any whitespace around the elements (optional, if necessary)
    result_list = [element.strip() for element in result_list]

    return result_list

def getPDB(input_file):
    with open(input_file, 'r') as infile:
        first_line = infile.readline().strip()
    return first_line[0:4]

def isChain(str,list):
    pos = []
    for i, char in enumerate(str):
        if str[i] == "|":
            pos.append(i)
    for a in range(len(list)):
        if(str[pos[1] + 1:pos[2]] == list[a]):
            return True
    return False

script = sys.argv[1]
pdb = getPDB(sys.argv[2])
file_name = script[0:1] + "-" + pdb + "_basepair.txt"
desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
file_path = os.path.join(desktop_path, file_name)

list = getChains(script[2:])
print(list)

output = []
with open(file_path, 'w') as file:
    for i in range(2,len(sys.argv)):
        with open(sys.argv[i], 'r') as text:
            for line in text:
                if(isChain(line,list)):
                    file.write(line)

print("file created on desktop")
file.close()