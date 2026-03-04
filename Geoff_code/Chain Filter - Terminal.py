import openpyxl
import xlsxwriter
import numpy as np
import csv
import sys
import os
from datetime import date

def isChain(line, chains):
    chain1 = getID(line[0])
    chain2 = getID(line[2])

    for a in range(0,len(chains)):
        for b in range(0,len(chains)):
            if(chain1 == chains[a] and chain2 == chains[b]):
                return True
    return False

def getID(str):
    pos = []
    for i, char in enumerate(str):
        if str[i] == "|":
            pos.append(i)
    return str[pos[1]+1:pos[2]]
def divideLine(line):
    pos = []
    for i in range(len(line) - 1):
        if (line[i:i + 1] == ","):
            pos.append(i)
        print(pos)
    return [line[1:pos[0]-1],line[pos[0]+2:pos[1]-1],line[pos[1]+2:len(line)-2]]

small = [sys.argv[-1]]
large = []
for i in range(4,len(sys.argv)-1):
    large.append(sys.argv[i])

for a in range(1,4):
    name = os.path.basename(sys.argv[a])
    file_name = str(name[0:-4])+"(L)"+str(name[-4:])
    desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
    file_path = os.path.join(desktop_path, file_name)

    matrix = []
    with open(sys.argv[a], newline='') as infile:
        reader = csv.reader(infile)
        for row in reader:
            if(isChain(row, large)):
                matrix.append(row)

    with open(file_path, 'w', newline='') as outfile:
        writer = csv.writer(outfile)
        writer.writerows(matrix)
    print(f"Matrix saved to: {file_path}")

for b in range(1,4):
    name = os.path.basename(sys.argv[b])
    file_name = str(name[0:-4]) + "(S)" + str(name[-4:])
    desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
    file_path = os.path.join(desktop_path, file_name)

    matrix = []
    with open(sys.argv[b], newline='') as infile:
        reader = csv.reader(infile)
        for row in reader:
            if(isChain(row, small)):
                matrix.append(row)

    with open(file_path, 'w', newline='') as outfile:
        writer = csv.writer(outfile)
        writer.writerows(matrix)
    print(f"Matrix saved to: {file_path}")

print("file complete")