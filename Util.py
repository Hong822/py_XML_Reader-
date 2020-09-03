import os
import easygui

class ComponentInfo:
    def __init__(self):
        super().__init__()
    
    SWVersion = None
    HWVersion = None
    PartNo = None
    KeyType = None
    DBINFO = None

def PrintFolderPath(dirname):
    try:
        filenames = os.listdir(dirname)
        for filename in filenames:
            full_filename = os.path.join(dirname, filename)
            if os.path.isdir(full_filename):
                PrintFolderPath(full_filename)
            else:
                ext = os.path.splitext(full_filename)[-1]
                if ext == '.xml': 
                   print(full_filename)
    except PermissionError:
        pass

def CollectFiles(dirname, FileList, SearchKeyword):
    try:
        filenames = os.listdir(dirname)
        for filename in filenames:
            full_filename = os.path.join(dirname, filename)
            if os.path.isdir(full_filename):
                CollectFiles(full_filename, FileList, SearchKeyword)
            else:
                ext = os.path.splitext(full_filename)[-1]
                if ext == SearchKeyword: 
                    FileList.append(full_filename)
    except PermissionError:
        pass

def GetIterator (pListName, TagName):
    Result = None
    for iterator in pListName:
        if iterator.tag == TagName:
            Result = iterator
            break
    return Result

def GetInfoData (ChildIter, cInfo):
    SWVersionIter = GetIterator(ChildIter, 'SWVersion')
    if SWVersionIter != None:
        cInfo.SWVersion = SWVersionIter.text
    else:
        cInfo.SWVersion = '-'
    
    HWVersionIter = GetIterator(ChildIter, 'HWVersion')
    if HWVersionIter != None:
        cInfo.HWVersion = HWVersionIter.text
    else:
        cInfo.HWVersion = '-'

    PartNoIter = GetIterator(ChildIter, 'HWTeilenummer')
    if PartNoIter != None:
        cInfo.PartNo = Insert_Dot(PartNoIter.text, 3)
    else:
        cInfo.PartNo = '-'

    KeyTypeIter = GetIterator(ChildIter, 'Schluesseltyp')
    if KeyTypeIter != None:
        cInfo.KeyType = KeyTypeIter.text
    else:
        cInfo.KeyType = '-'

    DBIter = GetIterator(ChildIter, 'SWVersion')
    if DBIter != None:
        cInfo.DBINFO = DBIter.text
    else:
        cInfo.DBINFO = '-'

def DictionarySetting(sSheetName, dDict, nRowOrColIdx, bCol):
    for i in range(1,sSheetName.max_row+1):
        if sSheetName.cell(row=i,column=1).value:
            MaxRow = i + 1
    
    for i in range(1,sSheetName.max_column+1):
        if sSheetName.cell(row=3,column=i).value:
            MaxCol = i + 1

    if bCol == True:
        for ColIdx in range(1, MaxCol):
            if sSheetName.cell(nRowOrColIdx, ColIdx).value != None:
                if sSheetName.cell(nRowOrColIdx, ColIdx).value in dDict:
                    sErrorString = "Sheet Name: " + sSheetName.title + " / Duplicated Key: " + sSheetName.cell(nRowOrColIdx, ColIdx).value
                    easygui.msgbox(sErrorString, "Duplicated Key")
                else:
                    dDict[sSheetName.cell(nRowOrColIdx, ColIdx).value] = ColIdx
    else:
        for RowIdx in range(1, MaxRow):
             if sSheetName.cell(RowIdx, nRowOrColIdx).value != None:
                if sSheetName.cell(RowIdx, nRowOrColIdx).value in dDict:
                    sErrorString = "Sheet Name: " + sSheetName.title + " / Duplicated Key: " + sSheetName.cell(RowIdx, nRowOrColIdx).value
                    easygui.msgbox(sErrorString, "Duplicated Key")
                else:
                    dDict[sSheetName.cell(RowIdx, nRowOrColIdx).value] = RowIdx
   
def GetVINRow (sSheetName, VIN_No, VINColumn, nVINRow):
    bResult = False
    for i in range(4,sSheetName.max_row+2):
        nVINRow[0] = i
        VINValue = sSheetName.cell(row=i,column=VINColumn).value
        if VINValue == VIN_No:
            bResult = True
            break
        elif VINValue == None and (i >= sSheetName.max_row) :
            bResult = False
            break
        else:
            bResult = False
    return bResult

def Insert_Dot(InputString, index):
    if InputString == None:
        ResultStr = '-'
    else:
        ListString = list(InputString)
        End = len(ListString)-1
        for i in range(End, 0, -1):
            if i % 3 == 0:
                ListString.insert(i, '.')
        ResultStr = "".join(ListString)
    return ResultStr
