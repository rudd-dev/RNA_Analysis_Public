import openpyxl
import pandas as pd
from PIL import Image
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
import math
import matplotlib.patches as mpatches

def round_up(number):
    if(number < 50):
        number += 100
    return math.ceil(number / 100) * 100

def autopct_format(pct):
    return ('%1.1f%%' % pct) if pct > 0 else ''

def getFontSize(radius):
    if (radius < 0.5):
        return 6
    if (radius < 0.7):
        return 8
    if (radius < 0.9):
        return 10
    return 11

def extractData(line, len):
    list = []
    for i in range(len):
        list.append(line.iat[i+1])
    return list

def writeGraph(out,struct,name,colors):
    img = Image.open(name+'!.png')
    img = img.resize((700, 700))

    fig, axs = plt.subplots(17,17, figsize=(6, 6))
    p = len(colors)
    max = out.iloc[:, p+1].max()
    total = out.iloc[:, p+1].sum()
    index = 0
    header = ['C-G\nX-X','G-C\nX-X','A-U\nX-X','U-A\nX-X','U-G\nX-X','G-U\nX-X','A-G\nX-X','G-A\nX-X','A-A\nX-X','A-C\nX-X','C-A\nX-X','C-C\nX-X','C-U\nX-X','G-G\nX-X','U-C\nX-X','U-U\nX-X',]
    side = ['X-X\nC-G','X-X\nG-C','X-X\nA-U','X-X\nU-A','X-X\nU-G','X-X\nG-U','X-X\nA-G','X-X\nG-A','X-X\nA-A','X-X\nA-C','X-X\nC-A','X-X\nC-C','X-X\nC-U','X-X\nG-G','X-X\nU-C','X-X\nU-U',]
    for a in range(1,17):
        for b in range(1,17):
            data = out.iloc[index]
            if (data.iat[p+1] == 0):
                axs[b,a].pie([1, 0], colors=['#FFFFFF', '#FFFFFF'])
            else:
                axs[b,a].pie(extractData(data,p),colors=colors,radius=1.5,startangle = 90,counterclock=False)
            index += 1
    legend_patches = [mpatches.Patch(color=color, label=label) for color, label in zip(colors, out.columns.tolist()[1:p+1])]
    fig.legend(handles=legend_patches, bbox_to_anchor=(1.1, 0.6, 0.1, 0.1), fancybox=False, fontsize='small')

    if(name == "Output 2H + NC"):
        fig.figimage(img, xo=5500, yo=4500, resize=False)
    else:
        fig.figimage(img, xo=5234, yo=4000,resize=False)

    axs[0, 0].text(0, .5, 'n='+str(total), ha='center', va='center', fontsize=6)
    axs[0, 0].axis('off')

    for i in range(1,17):
        axs[0,i].text(0.5, 0.5, header[i-1], ha='center', va='center', fontsize=6)
        axs[0, i].axis('off')
        axs[i, 0].text(0.5, 0.5, side[i - 1], ha='center', va='center', fontsize=6)
        axs[i, 0].axis('off')
    plt.savefig(struct+name+'.svg', format='svg', bbox_inches='tight')

    fig, axs = plt.subplots(17, 17, figsize=(5, 5))
    index = 0
    for a in range(1,17):
        for b in range(1,17):
            data = out.iloc[index]
            if (data.iat[p+1] == 0):
                axs[b, a].pie([1, 0], colors=['#FFFFFF', '#FFFFFF'])
            else:
                radius = math.sqrt(1.5*(data.iat[p+1]/max))
                axs[b, a].pie(extractData(data,p), colors=colors, radius=radius,startangle = 90,counterclock=False)
            index += 1
    fig.legend(handles=legend_patches, bbox_to_anchor=(1.2, 0.6, 0.1, 0.1), fancybox=False, fontsize='small')

    img = img.resize((1000, 1000))
    if (name == "Output 2H + NC"):
        fig.figimage(img, xo=9000, yo=8700, resize=False)
    else:
        fig.figimage(img, xo=9000, yo=6500, resize=False)

    axs[0, 0].text(0, .5, 'n='+str(total), ha='center', va='center', fontsize=6)
    axs[0, 0].axis('off')
    for i in range(1, 17):
        axs[0, i].text(0.5, 0.5, header[i - 1], ha='center', va='center', fontsize=6)
        axs[0, i].axis('off')
        axs[i, 0].text(0.5, 0.5, side[i - 1], ha='center', va='center', fontsize=6)
        axs[i, 0].axis('off')
    plt.savefig(struct+name+'-Sized.svg', format='svg', bbox_inches='tight')
    plt.close()

file_path = 'Stack_Abundance_Ribozyme.xlsx'
structure = "Ribozyme_"

outA = pd.read_excel(file_path, sheet_name='Output 2H')
outAA = pd.read_excel(file_path, sheet_name='Output 2H Split')
outAA2H = pd.read_excel(file_path, sheet_name='Output Only 2H Split')
outAAn2H = pd.read_excel(file_path, sheet_name='Output non 2H Split')
outB = pd.read_excel(file_path, sheet_name='Output cWW-x')
outC = pd.read_excel(file_path, sheet_name='Output x-cWW')
outE = pd.read_excel(file_path, sheet_name='Output SS')
outAE = pd.read_excel(file_path, sheet_name='Output 2H + SS')
outNC = pd.read_excel(file_path, sheet_name='Output 2H + NC')
outDub = pd.read_excel(file_path, sheet_name='Output Doublets')
outDubNC = pd.read_excel(file_path, sheet_name='Output Doublets(no cWC)')

#writeGraph(outA,structure,"Output 2H",['#0000ff','#ff0000'])
writeGraph(outAA,structure,"Output 2H Split",['#1edeff','#0000ff','#3d85c6','#ff0000','#e06666','#f4cccc'])
#writeGraph(outAA2H,structure,"Output Only 2H Split",['#1edeff','#0000ff','#3d85c6'])
#writeGraph(outAAn2H,structure,"Output non 2H Split",['#ff0000','#e06666','#f4cccc'])
#writeGraph(outB,structure,"Output cWW-x",['#000080', '#46F0F0', '#CFCFC4', '#D2F53C', '#F58231', '#808080', '#F032E6', '#FF6961', '#008080', '#0082C8', '#3CB44B', '#BFEF45', '#FFD8B1', '#AA6E28', '#808000', '#FABEBE', '#FF0000'])
#writeGraph(outC,structure,"Output x-cWW",['#000080', '#CFCFC4', '#FFD8B1', '#FF6961', '#AA6E28', '#808080', '#F58231', '#008080', '#0082C8', '#F032E6', '#FABEBE', '#3CB44B', '#BFEF45', '#D2F53C', '#46F0F0', '#808000', '#FF0000'])
#writeGraph(outE,structure,"Output SS",['#1e90ff','#5396c4','#3d85c6','#46f0f0','#1edeff','#911eb4','#f032e6','#e6beff','#ff69b4','#ff1493','#cd5c5c','#fffac8','#ffd700','#9a6324','#b8860b','#2f4f4f','#3cb44b','#32cd32','#adff2f','#aaffc3','#e6194b','#ff7f50','#f6b26b','#EBC3C0','#ffc1cb','#8b6747','#808080','#AA9D90','#C69155','#CFAF7A','#558098','#6CA5A5','#AABBB8','#00ced1','#A1DAD2','#ff0000'])
#writeGraph(outAE,structure,"Output 2H + SS", ['#0000ff','#1e90ff','#5396c4','#3d85c6','#46f0f0','#1edeff','#911eb4','#f032e6','#e6beff','#ff69b4','#ff1493','#cd5c5c','#fffac8','#ffd700','#9a6324','#b8860b','#2f4f4f','#3cb44b','#32cd32','#adff2f','#aaffc3','#e6194b','#ff7f50','#f6b26b','#EBC3C0','#ffc1cb','#8b6747','#808080','#AA9D90','#C69155','#CFAF7A','#558098','#6CA5A5','#AABBB8','#00ced1','#A1DAD2','#ff0000'])
#writeGraph(outNC,structure,"Output 2H + NC", ['#0000ff','#1edeff','#3d85c6','#7798cf','#8b6747','#95A56C','#03ff00','#f3cd7d','#ffc1cb','#ff7f50','#6CA5A5','#A1DAD2','#558098','#8E7A62','#AABBB8','#CFAF7A','#5067D1','#C69155','#A566A2','#459A96','#AA9D90','#ff0000'])
#writeGraph(outDub, structure,"Output Doublets",['#0000ff','#81cd9b','#53f500','#a64d79','#2A2A80','#ccc1ab','#A021C3','#bee38b','#5010A2','#454596','#3F4F6C','#7F9F58','#b4a7d6','#ff0000'])
#writeGraph(outDubNC, structure,"Output Doublets(no cWC)",['#81cd9b','#53f500','#a64d79','#2A2A80','#ccc1ab','#A021C3','#bee38b','#5010A2','#454596','#3F4F6C','#7F9F58','#b4a7d6','#c12790','#67c6ed','#3c78d8','#69916E','#387A62','#ff0000'])
