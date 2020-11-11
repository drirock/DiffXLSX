#!/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Oct 29 09:46:40 2020
 
@author: Rusconi Sandro
"""
import sys
import os
import zipfile
from xml.etree.ElementTree import iterparse
import tkinter as tk
from tkinter import filedialog
 
def diff_xlsx(fname,fname2,newfile):
    ### GESTISCO FILE EXCEL
    rows = []
    row = {}
    row_copy = {}
    sheetfile={}
    SheetName1={}
    SheetID1={}
    value = ''
 
 
    rows2 = []
    row2 = {}
    row2_copy = {}
    sheetfile2={}
    SheetName2={}
    SheetID2={}
    value2 = ''
    #CREO IL FILE PER OUTPUT
    try:
     OutputFile=open(newfile,'w')
     OutputFile.write("FileName;FileName(Confronto);SheetName;Cella;Valore;Confronto\n")
    except OSError:
     OutputFile=open("ErrorDiffXLS.log",'w')
     OutputFile.write("Impossibile creare il file {0}".format(newfile))
     OutputFile.close()
     sys.exit()
 
    z = zipfile.ZipFile(fname)
    if 'xl/sharedStrings.xml' in z.namelist():
        # Get shared strings
     strings = [element.text for event, element in iterparse(z.open('xl/sharedStrings.xml')) if element.tag.endswith('}t')]
 
    z2 = zipfile.ZipFile(fname2)
    if 'xl/sharedStrings.xml' in z2.namelist():
     # Get shared strings
     strings2 = [element.text for event, element in iterparse(z2.open('xl/sharedStrings.xml')) if element.tag.endswith('}t')]
 
    i=0
    for info in z.infolist():
     if info.filename.endswith('.xml') and 'sheet' in info.filename :
         sheetfile[i]=info.filename
         i+=1
     if info.filename == 'xl/workbook.xml':
        for event, child in iterparse(z.open('xl/workbook.xml'),events=('end',)):
         if child.attrib.get('name'):
          SheetID1=child.attrib.get('sheetId')
          if SheetID1 is not None:
           SheetName1[SheetID1]=child.attrib.get('name')
    i=0
    for info2 in z2.infolist():
     if info2.filename.endswith('.xml') and 'sheet' in info2.filename :
         sheetfile2[i]=info2.filename
         i+=1
     if info2.filename == 'xl/workbook.xml':
        for event2, child2 in iterparse(z2.open('xl/workbook.xml'),events=('end',)):
         if child2.attrib.get('name'):
          SheetID2=child2.attrib.get('sheetId')
          if SheetID2 is not None:
           SheetName2[SheetID2]=child2.attrib.get('name')
    i=0
    XMLName=fname.rsplit('/')[::-1]
    SheetName_1=""
    SheetName_2=""
    CountSheet1=1
    while i < len(sheetfile):
     find=0
     try:
      SheetName_1=SheetName1[str(CountSheet1)]
      print("SheetName_1: {0}".format(SheetName_1))
     except:
      print("Error on index1 i: {0} - CountSheet1: {1}".format(i,CountSheet1))
     CountSheet1+=1
     for event, element in iterparse(z.open(sheetfile[i])):
        # get value or index to shared strings
        if element.tag.endswith('}v') or element.tag.endswith('}t'):
            value = element.text
        # If value is a shared string, use value as an index
        if element.tag.endswith('}c'):
            if element.attrib.get('t') == 's':
                value = strings[int(value)]
            letter = element.attrib['r']
            #row[letter] = value
            row[SheetName_1] = {"FileName":XMLName[0],
                   "SheetName":SheetName_1,
                   "Cella":letter,
                   "Value":value}
            row_copy=row.copy()
            rows.append(row_copy)
            value = ''
     XMLName2=fname2.rsplit('/')[::-1]
     j=0
     CountSheet2=1
     while j < len(sheetfile2):
      try:
       SheetName_2=SheetName2[str(CountSheet2)]
      except:
       print("Error on index2 j: {0}".format(j))
      CountSheet2+=1
      if SheetName_2 == SheetName_1:
        find=1
        for event2, element2 in iterparse(z2.open(sheetfile2[i])):
         # get value or index to shared strings
         if element2.tag.endswith('}v') or element2.tag.endswith('}t'):
             value2 = element2.text
         # If value is a shared string, use value as an index
         if element2.tag.endswith('}c'):
             if element2.attrib.get('t') == 's':
                 value2 = strings2[int(value2)]
             letter2 = element2.attrib['r']
             row2[SheetName_2]={"FileName":XMLName2[0],
                                "SheetName":SheetName_2,
                                "Cella":letter2,
                                "Value":value2}
             row2_copy=row2.copy()
             rows2.append(row2_copy)
             value2 = ''
        for list_idx in rows:
         findCell=0
         for list_idx2 in rows2:
          try:
           if (list_idx[SheetName_1]["Cella"] == list_idx2[SheetName_2]["Cella"]):
            findCell = 1
           if ((list_idx != {}  and list_idx2 != {} ) and (list_idx[SheetName_1]["Cella"] == list_idx2[SheetName_2]["Cella"]) and (list_idx[SheetName_1]["Value"] != list_idx2[SheetName_2]["Value"])):
             print("{0};{1};{2};{3};{4};{5}\n".format(XMLName2[0],XMLName[0],list_idx2[SheetName_2]['SheetName'],list_idx2[SheetName_2]['Cella'],list_idx2[SheetName_1]["Value"],list_idx[SheetName_2]["Value"]))
             OutputFile.write("{0};{1};{2};{3};{4};{5}\n".format(XMLName2[0],XMLName[0],list_idx2[SheetName_2]['SheetName'],list_idx2[SheetName_2]['Cella'],list_idx2[SheetName_1]["Value"],list_idx[SheetName_2]["Value"]))
             break
          except:
           print("Error on write output")
           print(rows)
           OutputFile.write("Error on write output\nlist_idx: {0} \nlist_idx2: {1}\n".format(list_idx,list_idx2))
           os._exit(0)
         if findCell == 0:
           print(("{0};{1};{2};{3};;{4}".format(XMLName2[0],XMLName[0],list_idx[SheetName_1]['SheetName'],list_idx[SheetName_1]['Cella'],list_idx[SheetName_1]["Value"])))
           OutputFile.write(("{0};{1};{2};{3};;{4}\n".format(XMLName2[0],XMLName[0],list_idx[SheetName_1]['SheetName'],list_idx[SheetName_1]['Cella'],list_idx[SheetName_1]["Value"])))
        j+=1
        break
      rows2.clear()
      j+=1
     rows.clear()
     i+=1
     if find == 0:
      OutputFile.write("{0};{1};{2};;FOGLIO ASSENTE\n".format(XMLName2[0],XMLName[0], SheetName_1))
      continue
    OutputFile.close()
 
### Rimuovere commenti per utilizzo linea di comando
#print(len(sys.argv))
#for i in range(1,len(sys.argv)):
# print(sys.argv[i])
 
#if len(sys.argv) > 1 and len(sys.argv) <= 3:
# path_new = sys.argv[1]
# path_old = sys.argv[2]
#else:
# path_new = input("Enter first file name with absolute path: ")
# path_old = input("Enter second file name with absolute path: ")

### Commentare per utilizzo linea di comando
root = tk.Tk()
root.withdraw()

path1 = filedialog.askopenfilename()
path2 = filedialog.askopenfilename()
 
new_file="DiffXLSX.csv"
diff_xlsx(path2,path1,new_file)
 

### Rimuovere commenti per abilitare confronto reverse file
#new_file="DiffXLSX_reverse.csv"
#print("Eseguo reverse")
#diff_xlsx(path1,path2,new_file)
 
sys.exit(0)
os._exit(0)
