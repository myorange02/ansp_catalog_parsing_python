# Date: 1/3/2022

# 1/9/2022
# Made some modification. Make a separate file for removing duplicated catalogs

import pandas as pd

initialPubIDList = []
initialNumberList = []
initialFiguredList = []
initialCategoryList = []

finalPubIDList = []
finalNumberList = []
finalFiguredList = []
finalNumberCategoryList = []
finalFiguredCategoryList = []
finalInheritedCategoryList = []

def importExcel():
    df = pd.read_excel(r'D:\coding\malpub_catalog.xlsx', sheet_name = 'Sheet1' )
    for i in range(len(df.index)):
        initialPubIDList.append(df['Pub_ID'][i])
        initialNumberList.append(df['ANSP lots cited'][i])
        initialFiguredList.append(df['ANSP lots figured'][i])
        initialCategoryList.append(df['ANSP Collection'][i])
    return 0

def splitter():
    for i in range(len(initialPubIDList)):
        pubID = initialPubIDList[i]
        tempNum = initialNumberList[i].split(';')  # ['112345, 132456', 'IP 12345, 15437']
        tempFig = initialFiguredList[i].split(';') # ['x']
        Category = initialCategoryList[i]
        tempCat = ''

        tempNumCount = 0
        tempFigCount = 0

        # Splitting them into each Number
        for item in range(len(tempNum)): # [['112345', '132456'], ['IP 12345', ['15437']]
            tempNum[item] = tempNum[item].split(',')
            for piece in range(len(tempNum[item])):
                tempNum[item][piece] = tempNum[item][piece].strip()
                tempNumCount += 1

        for item in range(len(tempFig)): # [['x']]
            tempFig[item] = tempFig[item].split(',')
            for piece in range(len(tempFig[item])):
                tempFig[item][piece] = tempFig[item][piece].strip()
                tempFigCount += 1

        # appending
        if Category == 'Malacology' or Category == 'Malacology and Invertebrate Paleontology': # add ANSP and WM as worm
            for j in range(len(tempNum)):
                if tempNum[j][0] == '':
                    tempCat = ''
                elif tempNum[j][0] == 'x':
                    tempCat = ''
                elif tempNum[j][0][0:2] == 'IP':
                    tempCat = 'Invertebrate Paleontology'
                elif tempNum[j][0][0] == 'A' or tempNum[j][0][0] == '1' or tempNum[j][0][0] == '2' or tempNum[j][0][0] == '3' or tempNum[j][0][0] == '4' or tempNum[j][0][0] == '5' or tempNum[j][0][0] == '6' or tempNum[j][0][0] == '7' or tempNum[j][0][0] == '8' or tempNum[j][0][0] == '9':
                    tempCat = 'Malacology'
                else:
                    tempCat = 'Exception'
                
                for k in range(len(tempNum[j])):
                    if tempNum[j][k] == '':
                        finalNumberList.append('')
                    elif tempNum[j][k] == 'x':
                        finalNumberList.append('')
                    elif tempNum[j][k][0] == 'A' and tempNum[j][k][1] != ' ':
                        finalNumberList.append(tempNum[j][k][0] + ' ' + tempNum[j][k][1:])
                    elif tempNum[j][k][0:2] == 'IP' and tempNum[j][k][2] != ' ':
                        finalNumberList.append(tempNum[j][k][0:2] + ' ' + tempNum[j][k][2:])
                    else:
                        finalNumberList.append(tempNum[j][k])
                    finalNumberCategoryList.append(tempCat)

            for j in range(len(tempFig)):
                if tempFig[j][0] == '':
                    tempCat = ''
                elif tempFig[j][0] == 'x':
                    tempCat = ''
                elif tempFig[j][0][0:2] == 'IP':
                    tempCat = 'Invertebrate Paleontology'
                elif tempFig[j][0][0] == 'A' or tempFig[j][0][0] == '1' or tempFig[j][0][0] == '2' or tempFig[j][0][0] == '3' or tempFig[j][0][0] == '4' or tempFig[j][0][0] == '5' or tempFig[j][0][0] == '6' or tempFig[j][0][0] == '7' or tempFig[j][0][0] == '8' or tempFig[j][0][0] == '9':
                    tempCat = 'Malacology'
                else:
                    tempCat = 'Exception'
                
                for k in range(len(tempFig[j])):
                    if tempFig[j][k] == '':
                        finalFiguredList.append('')
                    elif tempFig[j][k] == 'x':
                        finalFiguredList.append('')
                    elif tempFig[j][k][0] == 'A' and tempFig[j][k][1] != ' ':
                        finalFiguredList.append(tempFig[j][k][0] + ' ' + tempFig[j][k][1:])
                    elif tempFig[j][k][0:2] == 'IP' and tempFig[j][k][2] != ' ':
                        finalFiguredList.append(tempFig[j][k][0:2] + ' ' + tempFig[j][k][2:])
                    else:
                        finalFiguredList.append(tempFig[j][k])
                    finalFiguredCategoryList.append(tempCat)

        else:
            for j in range(len(tempNum)):
                if tempNum[j][0] == '':
                    tempCat = ''
                elif tempNum[j][0] == 'x':
                    tempCat = ''
                elif tempNum[j][0][0:2] == 'IP':
                    tempCat = 'Invertebrate Paleontology'
                elif tempNum[j][0][0:2] == 'CA':
                    tempCat = 'Crustacea'
                elif tempNum[j][0][0:2] == 'PO':
                    tempCat = 'Porifera'
                elif tempNum[j][0][0:2] == 'WO' or tempNum[j][0][0:2] == 'WM':
                    tempCat = 'Worms'
                elif tempNum[j][0][0] == 'I':
                    tempCat = 'General Invertebrates'
                elif tempNum[j][0][0] == 'R' or tempNum[j][0][0] == 'r':
                    tempCat = 'Rotifers'
                elif tempNum[j][0][0] == '1' or tempNum[j][0][0] == '2' or tempNum[j][0][0] == '3' or tempNum[j][0][0] == '4' or tempNum[j][0][0] == '5' or tempNum[j][0][0] == '6' or tempNum[j][0][0] == '7' or tempNum[j][0][0] == '8' or tempNum[j][0][0] == '9':
                    tempCat = '(refer to inherited category)'
                else:
                    tempCat = 'Exception'

                for k in range(len(tempNum[j])):
                    if tempNum[j][k] == '':
                        finalNumberList.append('')
                    elif tempNum[j][k] == 'x':
                        finalNumberList.append('')
                    elif tempNum[j][k][0] == 'A' and tempNum[j][k][1] != ' ':
                        finalNumberList.append(tempNum[j][k][0] + ' ' + tempNum[j][k][1:])
                    elif tempNum[j][k][0:2] == 'IP' and tempNum[j][k][2] != ' ':
                        finalNumberList.append(tempNum[j][k][0:2] + ' ' + tempNum[j][k][2:])
                    elif tempNum[j][k][0:2] == 'IP' and tempNum[j][k][2] == ' ':
                        finalNumberList.append(tempNum[j][k])
                    elif tempNum[j][k][0] == 'I' and tempNum[j][k][1] != ' ':
                        finalNumberList.append(tempNum[j][k][0] + ' ' + tempNum[j][k][1:])
                    elif tempNum[j][k][0:2] == 'PO' and tempNum[j][k][2] != ' ':
                        finalNumberList.append(tempNum[j][k][0:2] + ' ' + tempNum[j][k][2:])
                    elif tempNum[j][k][0] == 'R' and tempNum[j][k][1] != ' ':
                        finalNumberList.append(tempNum[j][k][0] + ' ' + tempNum[j][k][1:])
                    elif tempNum[j][k][0:2] == 'WO' and tempNum[j][k][2] != ' ':
                        finalNumberList.append(tempNum[j][k][0:2] + ' ' + tempNum[j][k][2:])
                    elif tempNum[j][k][0:2] == 'WM' and tempNum[j][k][2] != ' ':
                        finalNumberList.append(tempNum[j][k][0:2] + ' ' + tempNum[j][k][2:])
                    elif tempNum[j][k][0:2] == 'CA' and tempNum[j][k][2] != ' ':
                        finalNumberList.append(tempNum[j][k][0:2] + ' ' + tempNum[j][k][2:])
                    else:
                        finalNumberList.append(tempNum[j][k])
                    finalNumberCategoryList.append(tempCat)
            
            for j in range(len(tempFig)):
                if tempFig[j][0] == '':
                    tempCat = ''
                elif tempFig[j][0] == 'x':
                    tempCat = ''
                elif tempFig[j][0][0:2] == 'IP':
                    tempCat = 'Invertebrate Paleontology'
                elif tempFig[j][0][0:2] == 'CA':
                    tempCat = 'Crustacea'
                elif tempFig[j][0][0:2] == 'PO':
                    tempCat = 'Porifera'
                elif tempFig[j][0][0:2] == 'WO' or tempFig[j][0][0:2] == 'WM':
                    tempCat = 'Worms'
                elif tempFig[j][0][0] == 'I':
                    tempCat = 'General Invertebrates'
                elif tempFig[j][0][0] == 'R' or tempFig[j][0][0] == 'r':
                    tempCat = 'Rotifers'
                elif tempFig[j][0][0] == '1' or tempFig[j][0][0] == '2' or tempFig[j][0][0] == '3' or tempFig[j][0][0] == '4' or tempFig[j][0][0] == '5' or tempFig[j][0][0] == '6' or tempFig[j][0][0] == '7' or tempFig[j][0][0] == '8' or tempFig[j][0][0] == '9':
                    tempCat = '(refer to inherited category)'
                else:
                    tempCat = 'Exception'
                
                for k in range(len(tempFig[j])):
                    if tempFig[j][k] == '':
                        finalFiguredList.append('')
                    elif tempFig[j][k] == 'x':
                        finalFiguredList.append('')
                    elif tempFig[j][k][0] == 'A' and tempFig[j][k][1] != ' ':
                        finalFiguredList.append(tempFig[j][k][0] + ' ' + tempFig[j][k][1:])
                    elif tempFig[j][k][0:2] == 'IP' and tempFig[j][k][2] != ' ':
                        finalFiguredList.append(tempFig[j][k][0:2] + ' ' + tempFig[j][k][2:])
                    elif tempFig[j][k][0:2] == 'IP' and tempFig[j][k][2] == ' ':
                        finalFiguredList.append(tempFig[j][k])
                    elif tempFig[j][k][0] == 'I' and tempFig[j][k][1] != ' ':
                        finalFiguredList.append(tempFig[j][k][0] + ' ' + tempFig[j][k][1:])
                    elif tempFig[j][k][0:2] == 'PO' and tempFig[j][k][2] != ' ':
                        finalFiguredList.append(tempFig[j][k][0:2] + ' ' + tempFig[j][k][2:])
                    elif tempFig[j][k][0] == 'R' and tempFig[j][k][1] != ' ':
                        finalFiguredList.append(tempFig[j][k][0] + ' ' + tempFig[j][k][1:])
                    elif tempFig[j][k][0:2] == 'WO' and tempFig[j][k][2] != ' ':
                        finalFiguredList.append(tempFig[j][k][0:2] + ' ' + tempFig[j][k][2:])
                    elif tempFig[j][k][0:2] == 'WM' and tempFig[j][k][2] != ' ':
                        finalFiguredList.append(tempFig[j][k][0:2] + ' ' + tempFig[j][k][2:])
                    elif tempFig[j][k][0:2] == 'CA' and tempFig[j][k][2] != ' ':
                        finalFiguredList.append(tempFig[j][k][0:2] + ' ' + tempFig[j][k][2:])
                    else:
                        finalFiguredList.append(tempFig[j][k])
                    finalFiguredCategoryList.append(tempCat)

        # appending pubID after comparing numbers of item from tempNum and tempFig
        if tempNumCount >= tempFigCount: # if there are more or equal items in tempNum, append pubID as the number of items in tempNum.
            for a in range(tempNumCount):
                finalPubIDList.append(pubID)
                if Category == 'x':
                    finalInheritedCategoryList.append('')
                else:
                    finalInheritedCategoryList.append(Category)
            for a in range(tempNumCount - tempFigCount): # For any difference in number of items of tempNum and tempFig, append ''.
                finalFiguredList.append('')
                finalFiguredCategoryList.append('')
        else:
            for a in range(tempFigCount): # vice versa
                finalPubIDList.append(pubID)
                finalInheritedCategoryList.append(Category)
            for a in range(tempFigCount - tempNumCount):
                finalNumberList.append('')
                finalNumberCategoryList.append('')

    print(len(finalPubIDList))
    print(len(finalNumberList))
    print(len(finalFiguredList))
    print(len(finalNumberCategoryList))
    print(len(finalFiguredCategoryList))
    print(len(finalInheritedCategoryList))

def exportExcel():
    data = {'Pub ID': finalPubIDList, 
    'CatalogNum': finalNumberList, 
    'CategoryNum': finalNumberCategoryList,   
    'CatalogFig': finalFiguredList,
    'CategoryFig': finalFiguredCategoryList,
    'InheritedCategory': finalInheritedCategoryList}
    df = pd.DataFrame(data)

    df.to_excel(r'D:\\coding\\parsedCatalog.xlsx', index = False)

    return 0

if __name__ == "__main__":
    importExcel()
    splitter()
    exportExcel()