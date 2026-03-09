# -*- coding: utf-8 -*-

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import pandas as pd
import numpy as np

#this function creates the list that appears on the top
#Bad name. change later. Probably can just use list(String)
def extractData(line, length):
    lis = []
    for i in range(length):
        lis.append(line.iat[i+1])
    return lis


datapath = ".\\test.xlsx"

#AKA Stack agnostic
testDF = pd.read_excel(datapath, "Tables Output A")

def GrabSquare(df, pos, sz):
    return df.iloc[ pos[0] : pos[0] + sz, pos[1] : pos[1] + sz]

SZ = 16
init_pos = [1,1] #1,1 offset because of labels
#cursor_jump = 
cjumpvec = [5 + SZ, 0]

def MakeArraysFromSquares(df, init_pos, sz, cursor_jump, n_sq, normalized = True):
    structure = None
    structure = GrabSquare(df, init_pos, sz)

    #working w a new sz x sz x njumps+1 array
    vecarr = np.zeros([sz,sz,n_sq])
    
    for j in range(n_sq):
        print("J ITER ", j)
        pos = [ init_pos[0] + j * cjumpvec[0],
               init_pos[1] + j * cjumpvec[1] ]
        #print("probing pos ", pos[0], " ", pos[1])
        newstru = GrabSquare(df, pos, sz)
        newstrunp = newstru.to_numpy()
        print(newstrunp)
        #print(newstru)
        
        #pasting new struct stuff over
        #very ugly way
        for k in range(sz):
            for l in range(sz):
                #print(newstru.iloc[k, l])
                newval = newstru.iloc[k, l]
                if type(newval) != int:
                    newval = 0
                vecarr[k,l,j] = newval#newstru.iloc[k, l]
        #print(vecarr[:,:,j])
        #structure = pd.concat([ structure, newstru ])#, axis=2)
        #structure.concat(newstrucomp)
    return vecarr#structure








DIM = 16
def WGNew(vecarr):
    # Column headers for the top row (pairing labels)
    header = ['C-G\nX-X','G-C\nX-X','A-U\nX-X','U-A\nX-X','U-G\nX-X','G-U\nX-X','A-G\nX-X','G-A\nX-X','A-A\nX-X','A-C\nX-X','C-A\nX-X','C-C\nX-X','C-U\nX-X','G-G\nX-X','U-C\nX-X','U-U\nX-X',]

    # Row headers for the left column (pairing labels)
    side = ['X-X\nC-G','X-X\nG-C','X-X\nA-U','X-X\nU-A','X-X\nU-G','X-X\nG-U','X-X\nA-G','X-X\nG-A','X-X\nA-A','X-X\nA-C','X-X\nC-A','X-X\nC-C','X-X\nC-U','X-X\nG-G','X-X\nU-C','X-X\nU-U',]

    fig, axis = plt.subplots(DIM+1, DIM+1)
    #idx = 0
    for i in range(vecarr.shape[0]):
        for j in range(vecarr.shape[1]):
            axis[j, i].pie(vecarr[i, j, :])

            #data = dataframe.iloc[idx]
            #print(data)
            #exdata = extractData(data, 15)
            #print(exdata)
            #if (data.iat[3] == 0):
            #    axis[j, i].pie([1,0], colors=["#FFFFFF","#FFFFFF"])
            #else:
                
            #    axis[j, i].pie(exdata, colors=["#FFFFFF","#FFFFFF"])
            #idx += 1
    plt.show()
#WGNew(outA)

vecarr_global = MakeArraysFromSquares(testDF, init_pos, SZ, cjumpvec, 4)
WGNew(vecarr_global)