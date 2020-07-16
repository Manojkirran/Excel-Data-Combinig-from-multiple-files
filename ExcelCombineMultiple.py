# combine Multiple Sheets into One Sheet for Multiple file

import pandas as pd
import os.path
from os import path
import platform
import os
from openpyxl import load_workbook

data = {}
data2 = []

def Mergeallexcel(inputxl,outputxl,*filenames,sheet1 ="sheet"):
    filenae = os.path.split(inputxl)
    tmp = filenae[-1].replace(".xlsx", "")
    xls = pd.ExcelFile(inputxl)
    name1 = xls.sheet_names
    for sheet in name1:
        df = pd.read_excel(inputxl, sheet_name=sheet)
        if sheet in data2:
            df = pd.read_excel(inputxl, sheet_name=sheet)
            sheet = sheet + "_" + tmp
            data[sheet] = df
        data2.append(str(sheet))
        data[sheet] = df
#Output file Function
    ostype = platform.system()
    fu = outputxl
    filenae1 = os.path.split(fu)
    if not filenae1[0] == "":
        if ostype == "Windows":
            ful = fu.split("\\")
        else:
            ful = fu.split("/")
    foldercreat = ""
    if not filenae1[0] == "":
        for i in range(0, len(ful) - 1):
             if ostype == "Windows":
                  foldercreat = foldercreat + "/" + ful[i]
        path1 = os.getcwd()
        validpath = path.exists(path1+foldercreat)
        if validpath == False:
            os.makedirs(path1+foldercreat)


    writer = pd.ExcelWriter(outputxl)
    gt = 0
    for sheet in data:
        df = data[sheet]
        tr3 = data.get(sheet)
        lenthofdf3 = len(tr3)
        df.to_excel(writer, sheet_name=sheet1, index=False, startrow=gt)
        gt = gt + lenthofdf3 + 1
    writer.save()
    writer.close()
    # Working on multiple files
    if not len(filenames) == 0:
        for k in filenames:
            validfile = k.endswith(".xls")
            validfile1 = k.endswith(".xlsx")
            validfile2 = k.endswith(".xlsm")
            if validfile == True or validfile1 == True or validfile2 == True:
                filenae2 = os.path.split(k)
                tmp = filenae2[-1].replace(".xlsx", "")
                xls2 = pd.ExcelFile(k)
                name2 = xls2.sheet_names
                data3 = {}
                data4 = []
                for sheet2 in name2:
                    df2 = pd.read_excel(k, sheet_name=sheet2)
                    if sheet2 in data4:
                        df2 = pd.read_excel(k, sheet_name=sheet2)
                        sheet2 = sheet2 + "_" + tmp
                        data3[sheet2] = df2
                    data4.append(str(sheet2))
                    data3[sheet2] = df2
                Book = load_workbook(outputxl)
                writer2 = pd.ExcelWriter(outputxl,engine='openpyxl')
                writer2.book = Book
                writer2.sheets = dict((ws.title, ws) for ws in Book.worksheets)
                for sheet2 in data3:
                    df2 = data3[sheet2]
                    tr3 = data3.get(sheet2)
                    lenthofdf3 = len(tr3)
                    df2.to_excel(writer2, sheet_name=sheet1, index=False, startrow=gt)
                    gt = gt + lenthofdf3 + 1
                writer2.save()
            else:
                 print("pls give the valid excel file name")
    datadu =pd.read_excel(outputxl)
    datadu2=datadu.drop_duplicates()
    writer = pd.ExcelWriter(outputxl)
    datadu2.to_excel(writer, sheet_name=sheet1, index=False)
    writer.save()

Mergeallexcel("First has 3 sheet.xlsx","manoj\lovers\world\outfile.xlsx","Second has 2 sheet.xlsx","Third has 2 sheet.xlsx",sheet1 ="Manoj")
