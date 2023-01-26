# 1/9/2022 File has been created

import pandas as pd

initPubIDList = []
initNumList = []
initNumCatList = []
initFigList = []
initFigCatList = []
initCatList = []

dataDict = {}

def importExcel():
    df = pd.read_excel(r'D:\coding\parsedCatalog.xlsx', sheet_name = 'Sheet1' )
    for i in range(len(df.index)):
        initPubIDList.append(df['Pub ID'][i])
        initNumList.append(df['CatalogNum'][i])
        initNumCatList.append(df['CategoryNum'][i])
        initFigList.append(df['CatalogFig'][i])
        initFigCatList.append(df['CategoryFig'][i])
        initCatList.append(df['InheritedCategory'][i])
    return 0

def replacer(list1, list2, a, b):
    for i in range(a, b):
        for j in range(a, b):
            if list1[i] == list2[j]:
                list1[i] = ''
    return list1, list2

def remover(initNumList, initFigList):
    tempID = ''
    a = 0
    b = 0
    for i in range(len(initPubIDList)):
        if tempID == '':
            tempID = initPubIDList[i]
            a = i
            if a + 1 == len(initPubIDList):
                break

            if tempID == initPubIDList[i + 1]:
                continue
            else:
                b = i + 1
                initNumList, initFigList = replacer(initNumList, initFigList, a, b)
                tempID = ''
        elif tempID == initPubIDList[i]:
            continue
        elif tempID != initPubIDList[i]:
            b = i + 1
            initNumList, initFigList = replacer(initNumList, initFigList, a, b)
            tempID = ''
    return 0
        
def exportExcel():
    data = {'Pub ID': initPubIDList, 
    'CatalogNum': initNumList, 
    'CategoryNum': initNumCatList,   
    'CatalogFig': initFigList,
    'CategoryFig': initFigCatList,
    'InheritedCategory': initCatList}
    df = pd.DataFrame(data)

    df.to_excel(r'D:\\coding\\parsedCatalogFinal.xlsx', index = False)

    return 0

importExcel()
remover(initNumList, initFigList)
exportExcel()