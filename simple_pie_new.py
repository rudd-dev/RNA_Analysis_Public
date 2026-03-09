# -*- coding: utf-8 -*-

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import pandas as pd

#this function creates the list that appears on the top
#Bad name. change later. Probably can just use list(String)
def extractData(line, length):
    lis = []
    for i in range(length):
        lis.append(line.iat[i+1])
    return lis


datapath = ".\\test.xlsx"

#AKA Stack agnostic
outA = pd.read_excel(datapath, "Tables Output A")


























DIM = 16
def WGNew(dataframe):
    # Column headers for the top row (pairing labels)
    header = ['C-G\nX-X','G-C\nX-X','A-U\nX-X','U-A\nX-X','U-G\nX-X','G-U\nX-X','A-G\nX-X','G-A\nX-X','A-A\nX-X','A-C\nX-X','C-A\nX-X','C-C\nX-X','C-U\nX-X','G-G\nX-X','U-C\nX-X','U-U\nX-X',]

    # Row headers for the left column (pairing labels)
    side = ['X-X\nC-G','X-X\nG-C','X-X\nA-U','X-X\nU-A','X-X\nU-G','X-X\nG-U','X-X\nA-G','X-X\nG-A','X-X\nA-A','X-X\nA-C','X-X\nC-A','X-X\nC-C','X-X\nC-U','X-X\nG-G','X-X\nU-C','X-X\nU-U',]

    fig, axis = plt.subplots(DIM+1, DIM+1)
    idx = 0
    for i in range(1, DIM+1):
        for j in range(1, DIM + 1):
            data = dataframe.iloc[idx]
            print(data)
            exdata = extractData(data, 15)
            print(exdata)
            if (data.iat[3] == 0):
                axis[j, i].pie([1,0], colors=["#FFFFFF","#FFFFFF"])
            else:
                
                axis[j, i].pie(exdata, colors=["#FFFFFF","#FFFFFF"])
            idx += 1
            
WGNew(outA)