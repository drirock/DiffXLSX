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
    sheetfile={}
    SheetName1={}
    SheetID1={}
    value = ''
    

    rows2 = []
    row2 = {}
    sheetfile2={}
    SheetName2={}
    SheetID2={}
    value2 = ''
    #CREO IL FILE PER OUTPUT
    try:
     OutputFile=open(newfile,'w')   
     OutputFile.write("FileName;SheetName;Cella;Valore\n")
    except OSError:
     OutputFile=open("ErrorDiffXLS.log",'w')
     OutputFile.write("Impossibile creare il file {0}".format(newfile))
     OutputFile.close()
     os._exit(0)
    
    z = zipfile.ZipFile(fname)
    if 'xl/sharedStrings.xml' in z.namelist():
        # Get shared strings
     strings = [element.text for event, element in iterparse(z.open('xl/sharedStrings.xml')) if element.tag.endswith('}t')]
     #sheetdict = { element.attrib['name']:element.attrib['sheetId'] for event,element in iterparse(z.open('xl/workbook.xml')) if element.tag.endswith('}sheet') }
     #print(strings)
     
    z2 = zipfile.ZipFile(fname2)
    if 'xl/sharedStrings.xml' in z2.namelist():
     # Get shared strings
     strings2 = [element.text for event, element in iterparse(z2.open('xl/sharedStrings.xml')) if element.tag.endswith('}t')]
     
    i=0
    for info in z.infolist():
    #print(info.filename)
     if info.filename.endswith('.xml') and 'sheet' in info.filename :         
         sheetfile[i]=info.filename
         i+=1
     if info.filename == 'xl/workbook.xml':
        for event, child in iterparse(z.open('xl/workbook.xml'),events=('end',)):
         if child.attrib.get('name'):
          SheetID1=child.attrib.get('sheetId')
          SheetName1[SheetID1]=child.attrib.get('name')
          #print(SheetID1," - ", SheetName1[SheetID1])
    i=0
    for info2 in z2.infolist():
    #print(info.filename)
     if info2.filename.endswith('.xml') and 'sheet' in info2.filename :         
         sheetfile2[i]=info2.filename
         i+=1
     if info2.filename == 'xl/workbook.xml':
        for event2, child2 in iterparse(z2.open('xl/workbook.xml'),events=('end',)):
         if child2.attrib.get('name'):
          SheetID2=child2.attrib.get('sheetId')
          SheetName2[SheetID2]=child2.attrib.get('name')
          #print(SheetID2," - ", SheetName2[SheetID2])
    i=0     
    XMLName=fname.rsplit('/')[::-1]
    
    while i < len(sheetfile):  
     find=0
     NSheet1=sheetfile[i].rsplit('/')[::-1]
     NSheet1[0]=NSheet1[0].replace('.xml',"")
     NSheet1[0]=NSheet1[0][-1]
     SheetName_1=SheetName1[NSheet1[0]]   
     #TotRow[CountTOT]="{0};{1}".format(XMLName[0],SheetName_1)
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
            row[SheetName_1]={"FileName":XMLName[0],
                   "SheetName":SheetName_1,
                   "Cella":letter,
                   "Value":value}
            value = ''
        if element.tag.endswith('}row'):
             #rows2.append(SheetName_2)
             rows.append(row)
             row = {}
     XMLName2=fname2.rsplit('/')[::-1]            
     j=0
     while j < len(sheetfile2):
      NSheet2=sheetfile2[j].rsplit('/')[::-1]
      NSheet2[0]=NSheet2[0].replace('.xml',"")
      NSheet2[0]=NSheet2[0][-1]
      SheetName_2=SheetName2[NSheet2[0]]
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
             #row2[letter2] = value2
             row2[SheetName_2]={"FileName":XMLName2[0],
                   "SheetName":SheetName_2,
                   "Cella":letter2,
                   "Value":value2}
             value2 = ''
         if element2.tag.endswith('}row'):
             rows2.append(row2)
             row2 = {}             
        for list_idx in rows:
         for list_idx2 in rows2:
          if (list_idx[SheetName_1]["Cella"] == list_idx2[SheetName_2]["Cella"]) and (list_idx[SheetName_1]["Value"] != list_idx2[SheetName_2]["Value"]):
           print("{0};{1};{2};{3}/{4}".format(list_idx2[SheetName_1]['FileName'],list_idx2[SheetName_2]['SheetName'],list_idx2[SheetName_2]['Cella'],list_idx[SheetName_1]["Value"],list_idx2[SheetName_2]["Value"]))
           OutputFile.write("{0};{1};{2};{3} - {4}\n".format(list_idx2[SheetName_1]['FileName'],list_idx2[SheetName_2]['SheetName'],list_idx2[SheetName_2]['Cella'],list_idx[SheetName_1]["Value"],list_idx2[SheetName_2]["Value"]))
      rows2.clear()
      j+=1
     rows.clear() 
     i+=1 
     print(rows2)
     if find == 0:
      print(XMLName2[0],";",SheetName_1,"; ASSENTE")
      OutputFile.write("{0};{1};ASSENTE\n".format(XMLName2[0], SheetName_1))
    OutputFile.close()
root = tk.Tk()
root.withdraw()

path_new = filedialog.askopenfilename()
path_old = filedialog.askopenfilename()

new_file="DiffXLSX.csv"
diff_xlsx(path_old,path_new,new_file)

new_file="DiffXLSX_reverse.csv"
print("Eseguo reverse")
diff_xlsx(path_new,path_old,new_file)

sys.exit(0)
os._exit(0)
