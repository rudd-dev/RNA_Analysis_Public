import openpyxl
import xlsxwriter
import numpy as np
import csv
import sys
import os
from datetime import date

class Interaction():
    base1 = -1
    pos1 = -1
    base2 = -1
    pos2 = -1
    int = ""
    chain1 = ""
    chain2 = ""
    stacking = False

    def __init__(self, b1, p1, b2, p2, inter, chan1, chan2):
        self.base1 = b1
        self.pos1 = p1
        self.base2 = b2
        self.pos2 = p2
        self.int = inter
        self.chain1 = chan1
        self.chain2 = chan2

    def getBase1(self):
        return self.base1

    def getPos1(self):
        return self.pos1

    def getInt(self):
        return self.int

    def getBase2(self):
        return self.base2

    def getPos2(self):
        return self.pos2

    def getChain1(self):
        return self.chain1

    def getChain2(self):
        return self.chain2

    def getChains(self):
        return str(self.chain1) + "/" + str(self.chain2)

    def isStacking(self):
        return self.stacking

    def addStacking(self):
        self.stacking = True

    def getPair(self):
        return str(self.base1) + "-" + str(self.base2)

    def __str__(self):
        return self.chain1 + " (" + str(self.pos1) + ") " + self.base1 + " - " + self.base2 + " (" + str(
            self.pos2) + ") " + self.chain2 + " " + self.int

def getName(name):
    if(len(name) < 27):
        return name[0:4]
    return name[0:4]+ "-"+ name[-20:-13]
def createInteraction(line):
    parts = divideLine(line)
    base1 = getNucleo(parts[0])
    pos1 = getPos(parts[0])
    base2 = getNucleo(parts[2])
    pos2 = getPos(parts[2])
    int = simplifyInt(parts[1])
    chain1 = getID(parts[0])
    chain2 = getID(parts[2])
    return Interaction(base1,pos1,base2,pos2,int,chain1,chain2)
# 1S72|1|0|U|12","tHS","1S72|1|0|G|531

def divideLine(line):
    pos = []
    for i in range(len(line) - 1):
        if (line[i:i + 1] == "\t"):
            pos.append(i)
    return [line[0:pos[0]],line[pos[0]+1:pos[1]],line[pos[1]+1:pos[2]]]
def getDate():
    today = date.today()
    return str(today.month) + "_" + str(today.day) + "_" + str(today.year)
def getPos(str):
    pos = []
    for i, char in enumerate(str):
        if str[i] == "|":
            pos.append(i)
    if(len(pos) > 4):
        return int(str[pos[3]+1:pos[4]])
    return int(str[pos[3]+1:])

def percentage(part, whole):
  per = 100 * float(part)/float(whole)
  return round(per*100)

def getNucleo(str):
    pos = []
    for i, char in enumerate(str):
        if str[i] == "|":
            pos.append(i)
    seq = str[pos[2]+1:pos[3]]
    if(len(seq) > 1):
        for a, char in enumerate(seq):
            if (seq[a] == 'A'or seq[a] == 'C'or seq[a] == 'G'or seq[a] == 'U'):
                return seq[a]
    return seq
def getID(str):
    pos = []
    for i, char in enumerate(str):
        if str[i] == "|":
            pos.append(i)
    return str[pos[1]+1:pos[2]]
def getInt(str):
    pos = []
    for i in range(len(str) - 1):
        if (str[i:i + 1] == "\t"):
            pos.append(i)
    return simplifyInt(str[pos[0] + 1:pos[1]])

#list - > intreaction
#matrix -> [23S/5S, Position 1, 23S/5S, Position 2]
def findModules(list,matrix):
    modules = []
    for a in range(len(matrix)):
        for b in range(len(list)):
            if(matrix[a][0] == list[b].getChain1() and matrix[a][1] == list[b].getPos1()):
                for b2 in range(len(list)):
                    if(matrix[a][2] == list[b2].getChain1() and matrix[a][3] == list[b2].getPos1()):
                        modules.append([list[b],list[b2]])
    return modules

def simplifyInt(str):
    if(str[0:1] == "n"):
        return str[1:4]
    return str[0:3]
def getReverse(str):
  for i in range(len(str)):
    if str[i] == "/":
      return getReverse(str[0:i]) + "/" + getReverse(str[i+1:])
  if(len(str) == 4):
    return str[0:2] + str[3:4] + str[2:3]
  if(str[0:2] == "nt" or str[0:2] == "nc"):
    return str[0:2] + str[3:4] + str[2:3] + str[4:]
  if(len(str)>3):
    return str[0:1] + str[2:3] + str[1:2] + str[3:]
  return str[0:1] + str[2:3] + str[1:2]

def isStacking(c1,p1,c2,p2, stackPos):
    for i in range(len(stackPos)):
        if(stackPos[i][0] == c1 and stackPos[i][1] == p1):
            if(stackPos[i][2] == c2 and stackPos[i][3] == p2):
                return True
    return False

def removeRedundancy(modules):
    matrix = []
    double = []
    for a in range(len(modules)):
        isRed = True
        for b in range(len(double)):
            if(modules[a][0].getChain1() == double[b][0].getChain2() and modules[a][0].getPos1() == double[b][0].getPos2()):
                if (modules[a][1].getChain1() == double[b][1].getChain2() and modules[a][1].getPos1() == double[b][
                    1].getPos2()):
                    isRed = False
            if (modules[a][1].getChain1() == double[b][0].getChain2() and modules[a][1].getPos1() == double[b][
                0].getPos2()):
                if (modules[a][0].getChain1() == double[b][1].getChain2() and modules[a][0].getPos1() == double[b][
                    1].getPos2()):
                    isRed = False
        if(isRed):
            matrix.append(modules[a])
            double.append(modules[a])
    return matrix
def removeBprime(module,Bprime):
    matrix = []
    for a in range(len(module)):
        x = True
        for b in range(len(Bprime)):
            if(module[a] == Bprime[b]):
                x = False
        if(x):
            matrix.append(module[a])
    return matrix

def findModuleAB(modules):
    matrix = []
    for a in range(len(modules)):
        if (abs(modules[a][0].getPos2() - modules[a][1].getPos2()) > 1):
            matrix.append(modules[a])
    return matrix
def findModuleA(modules, stackPos):
    matrix = []
    for a in range(len(modules)):
        if(isStacking(modules[a][0].getChain2(),modules[a][0].getPos2(),modules[a][1].getChain2(),modules[a][1].getPos2(),stackPos) == False):
            if(abs(modules[a][0].getPos2()-modules[a][1].getPos2()) > 1):
                matrix.append(modules[a])
    return matrix
def findModuleB(modules, stackPos):
    matrix = []
    for a in range(len(modules)):
        if (abs(modules[a][0].getPos2() - modules[a][1].getPos2()) > 1):
            if(isStacking(modules[a][0].getChain2(), modules[a][0].getPos2(), modules[a][1].getChain2(), modules[a][1].getPos2(), stackPos)):
                matrix.append(modules[a])
    return matrix

def findModuleBprime(modules,stackPos):
    matrix = []
    for a in range(len(modules)):

            if (abs(modules[a][0].getPos2() - modules[a][1].getPos2()) > 1):
                    if(isStacking(modules[a][0].getChain1(),modules[a][0].getPos1(),modules[a][1].getChain2(),modules[a][1].getPos2(),stackPos)):
                        matrix.append(modules[a])
                    elif (isStacking(modules[a][0].getChain2(), modules[a][0].getPos2(), modules[a][1].getChain1(),
                                   modules[a][1].getPos1(), stackPos)):
                        matrix.append(modules[a])
    return matrix
def findModuleC(modules, stackPosCon):
    matrix = []
    for a in range(len(modules)):
        if (modules[a][0].getPos2() - modules[a][1].getPos2() == -1):
            matrix.append(modules[a])
    return removeRedundancy(matrix)
def findModuleD(modules, stackPosCon):
    matrix = []
    for a in range(len(modules)):
        if (modules[a][0].getPos2() - modules[a][1].getPos2() == 1):
            matrix.append(modules[a])
    return removeRedundancy(matrix)
def getAbundance(modules):
    matrix = []
    for a in range(len(modules)):
        check = True
        for b in range(len(matrix)):
            if(modules[a][0].getInt() == matrix[b][0] and modules[a][1].getInt() == matrix[b][1]):
                matrix[b][2] += 1
                check = False
        if(check):
            matrix.append([modules[a][0].getInt(),modules[a][1].getInt(),1])
    matrix.sort(key=lambda x: x[2], reverse=True)
    return matrix

def writeSheet(modules, abund, name, pdb):
    output = workbook.add_worksheet(name)
    output.write(0, 0, "PDB")
    output.write(0, 1, "Chain res i (1-4)")
    output.write(0, 2, "Chain res j (2-3)")
    output.write(0, 3, "Pos 1")
    output.write(0, 4, "Pos 4")
    output.write(0, 5, "Pos 2")
    output.write(0, 6, "Pos 3")
    output.write(0, 7, "1:4")
    output.write(0, 8, "2:3")
    output.write(0, 9, "1:4 int")
    output.write(0, 10, "2:3 int")
    output.write(0, 12, "Interactions Observed")
    output.write(0, 13, "Interaction Abundance")
    output.write(0, 14, "Interaction Frequency")
    output.set_column(12, 14, 20)

    for x in range(len(modules)):
        output.write(x + 1, 0, pdb)
        output.write(x + 1, 1, str(modules[x][0].getChains()))
        output.write(x + 1, 2, str(modules[x][1].getChains()))
        output.write_number(x + 1, 3, modules[x][0].getPos1())
        output.write_number(x + 1, 4, modules[x][0].getPos2())
        output.write_number(x + 1, 5, modules[x][1].getPos1())
        output.write_number(x + 1, 6, modules[x][1].getPos2())
        output.write(x + 1, 7, str(modules[x][0].getPair()))
        output.write(x + 1, 8, str(modules[x][1].getPair()))
        output.write(x + 1, 9, str(modules[x][0].getInt()))
        output.write(x + 1, 10, str(modules[x][1].getInt()))
    for y in range(len(abund)):
        freq = abund[y][2] / len(modules)
        output.write(y + 1, 12, str(abund[y][0] + "/" + abund[y][1]))
        output.write_number(y + 1, 13, abund[y][2])
        output.write_number(y + 1, 14, freq)

def sumTotal(matrix):
    for a in range(len(matrix)):
        sum = 0
        for b in range(len(matrix[a])):
            sum += matrix[a][b]
        matrix[a].append(sum)
    list = []
    for y in range(len(matrix[0])):
        sum = 0
        for x in range(len(matrix)):
            sum += matrix[x][y]
        list.append(sum)
    matrix.append(list)
    return matrix

def makeTable(modules, abund, name):
    output = workbook.add_worksheet(name)
    x = 0
    for q in range(len(abund)):
        output.write(x,0, abund[q][0] + "/" + abund[q][1])
        output.write(x, 20, abund[q][0] + "/" + abund[q][1] + "Percentages")
        matrix = []
        row = ["A-A", "A-C", "A-G", "A-U", "C-A", "C-C", "C-G", "C-U","G-A", "G-C", "G-G", "G-U","U-A", "U-C", "U-G", "U-U"]
        for line in range(len(row)):
            matrix.append([0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0])
        count = 0
        for i in range(len(modules)):
            if(modules[i][0].getInt() == abund[q][0] and modules[i][1].getInt() == abund[q][1]):
                for a in range(len(row)):
                    for b in range(len(row)):
                        if(modules[i][0].getPair() == row[b] and modules[i][1].getPair() == row[a]):
                            matrix[a][b] += 1
                            count += 1
        matrix = sumTotal(matrix)

        row.append("Total")
        writeTables(matrix, x+1, row, output,count)
        x += 21
def writeTables(matrix, x, row, output,count):
    output.add_table(x, 0, x + 17, 17, {'header_row': True, 'autofilter': False, 'first_column': True })
    output.write(x, 0, "n = " + str(count))
    output.write_row(x, 1, row)
    output.write_column(x + 1, 0, row)

    yellow = workbook.add_format({'bg_color': '#FFFF00'})
    for a in range(len(matrix)):
        for b in range(len(matrix[0])):
            if(matrix[a][b]>0):
                output.write(x+1+a,b+1,matrix[a][b],yellow)
            else:
                output.write(x+1+a,b+1,matrix[a][b])
    cell_format = workbook.add_format()
    cell_format.set_num_format(10)
    cell_format1 = workbook.add_format({'bg_color': 'yellow', 'num_format': '0.0%'})

    for w in range(len(matrix)):
        for z in range(len(matrix[0])):
            if(count > 0):
                matrix[w][z] = matrix[w][z] / count

    output.add_table(x, 20, x + 17, 37, {'header_row': True, 'autofilter': False, 'first_column': True})
    output.write(x, 20, "n = " + str(count))
    output.write_row(x, 21, row)
    output.write_column(x + 1, 20, row)

    for a1 in range(len(matrix)):
        for b1 in range(len(matrix[0])):
            if (matrix[a1][b1] > 0):
                output.write(x + 1 + a1, b1 + 21, matrix[a1][b1], cell_format1)
            else:
                output.write(x + 1 + a1, b1 + 21, matrix[a1][b1],cell_format)
def removeRed(abund):
    list = []
    for a in range(len(abund)):
        isNew = True
        for b in range(len(list)):
            if(abund[a][0] == list[b][0] and abund[a][1] == list[b][1]):
                list[b][2] += abund[a][2]
                isNew = False
        if(isNew):
            list.append(abund[a])
    list.sort(key=lambda x: x[2], reverse=True)
    return list

date = getDate()
#specify filepath here
for file in range(1,len(sys.argv)-1):
    name = os.path.basename(sys.argv[file])
    pdb = name[0:4]
    if(file%2 == 0):
        print("next PDB")
    else:
        if(name[-12] == "stacking.csv"):
            stackname = sys.argv[file]
            interactions = sys.argv[file+1]
            file_name = "Modules_fr3d_" + getName(name)+"_"+ date +".xlsx"
        else:
            stackname = sys.argv[file+1]
            interactions = sys.argv[file]
            file_name = "Modules_fr3d_" + getName(name)+"_"+ date +".xlsx"


        #creates the array of all interactions
        basepairs = []
        with open(interactions,'r') as tfile:
            for line in tfile:
                basepairs.append(createInteraction(line))

        #stackPosCon row -> [23S/5S, Position 1, 23S/5S, Position 2]. Contains no redundancy, contingent stacks
        #stackPos row -> [23S/5S, Position 1, 23S/5S, Position 2]. Contains redundancy
        stackPosCon = []
        stackPos = []
        with open(stackname,'r') as tfile:
            for line in tfile:
                part = divideLine(line)
                stackPos.append([getID(part[0]),getPos(part[0]),getID(part[2]),getPos(part[2])])
                if(getPos(part[0])-getPos(part[2]) == -1):
                    stackPosCon.append([getID(part[0]),getPos(part[0]),getID(part[2]),getPos(part[2])])

        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        file_path = os.path.join(desktop_path, file_name)
        workbook = xlsxwriter.Workbook(file_path)

        #outputs all the modules
        modules = findModules(basepairs,stackPosCon)
        print(len(modules))
        #outputs module A and B
        moduleAB = findModuleAB(modules)
        moduleA = findModuleA(modules, stackPos)
        moduleB = findModuleB(modules, stackPos)
        moduleBprime = findModuleBprime(moduleA, stackPos)
        moduleC = findModuleC(modules, stackPosCon)
        moduleD = findModuleD(modules, stackPosCon)

        moduleA = removeBprime(moduleA,moduleBprime)

        abAbun = getAbundance(moduleAB)
        aAbun = getAbundance(moduleA)
        bAbun = getAbundance(moduleB)
        bpAbun = getAbundance(moduleBprime)
        cAbun = getAbundance(moduleC)
        dAbun = getAbundance(moduleD)

        writeSheet(moduleAB, abAbun, "Output A+B+Bprime", getName(name))
        makeTable(moduleAB, abAbun, "Tables Output A+B+Bprime")
        writeSheet(moduleA, aAbun, "OutputA", getName(name))
        makeTable(moduleA, aAbun, "Tables Output A")
        writeSheet(moduleB, bAbun, "OutputB", getName(name))
        makeTable(moduleB, bAbun, "Tables Output B")
        writeSheet(moduleBprime, bpAbun, "Output B-prime", getName(name))
        makeTable(moduleBprime, bpAbun, "Tables Output B-prime")
        writeSheet(moduleC, cAbun, "OutputC", getName(name))
        makeTable(moduleC, cAbun, "Tables Output C")
        writeSheet(moduleD, dAbun, "OutputD", getName(name))
        makeTable(moduleD, dAbun, "Tables Output D")

        workbook.close()

        print(f"Excel file '{file_name}' created on your desktop.")
print("Program Complete")