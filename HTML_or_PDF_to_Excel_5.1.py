import os
from html.parser import HTMLParser
import PyPDF2                       #need to run 'pip install PyPDF2' at Terminal before using
import pandas as pd                 #need to run 'pip install pandas' at Terminal before using
'''
Also may need to run 'pip install XlsxWriter' at Terminal before using 
'''

#↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
def GetFileNames(path: str) -> list:
    htmls_and_pdfs = [[],[]]          #the list of 2 sets: html and pdf file names
    htmls = []                        #the list for html-file names
    pdfs = []                         #the list for pdf-file names
    os.chdir(path+'\FCT_files')
    dir_list = os.listdir(path+'\FCT_files')
    # print (dir_list)
    for item in dir_list:
        if item.endswith('.html'):    #search for html files
            htmls.append(item)       
        elif item.endswith('.pdf'):   #search for pdf files
            pdfs.append(item)
    print ('This folder has',len(htmls),'html files and',len(pdfs),'pdf files')
    print ('')
    htmls_and_pdfs[0] = htmls
    htmls_and_pdfs[1] = pdfs
    os.chdir('..')
    return htmls_and_pdfs
#↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑#

#class/function for convert html to txt:
#//**
class HTMLFilter(HTMLParser):
    text = ""
    def handle_data(self, data):
        self.text += data
#htmlLines = '"shoud be a line in format of HTML coding"'
# f = HTMLFilter()
# f.feed(htmlLines)
# print(f.text)
#**//

#↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
def html2txt(FileName: str, CreateTXT=False) -> list[str]:

    os.chdir(path+'\FCT_files')
    TheFile = open (FileName,"r")
    htmlLineList = TheFile.readlines()
    TheFile.close()

    txtLineList = []                           #list of lines converted to txt
    txtLineList.append('*** File Name ***')
    txtLineList.append(FileName)

    if CreateTXT == True:
        x = 0                                      #how many lines in that file
        ResultFileName = FileName[:-4]+"txt"           #name a result/new txt file
        with open(ResultFileName,"w") as ResultTxt:        #create a txt with that result/new name
            for L in htmlLineList:
        #   print ("No.",x+1,"line is:")
                f = HTMLFilter()              #use HTMLParser class/function
                f.feed(L.strip())             #put html code line into the above class/function
                txtLineList.append(f.text)    #do converting and store result in a list of string
                f.text = str(x)+'.'+f.text    #add index for debug
                ResultTxt.write(f.text+'\n')  #write the converted result into that result/new txt file
        #       print(f.text)
        #       print (L.strip())
                x = x + 1
        ResultTxt.close()
        print (FileName[:-5],'txt (from HTML)','has',x,'lines')
    else:
        for L in htmlLineList:
            f = HTMLFilter()              #use HTMLParser class/function
            f.feed(L.strip())             #put html code line into the above class/function
            txtLineList.append(f.text)    #do converting and store result in a list of string

    os.chdir('..')
    return txtLineList
#↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑#

#↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
def pdf2txt(FileName: str, CreateTXT=False) -> list[str]:

    os.chdir(path+'\FCT_files')
    PdfFile = open(FileName,'rb')               #open the pdf file
    PdfReader = PyPDF2.PdfFileReader(PdfFile)   #create a variable reading the pdf file
    p = PdfReader.numPages                      #p shows how many pages that pdf file has 

    txtCharacterStr = ''
    for i in range(p):                                        #go through all pages
        PagesInfo = PdfReader.getPage(i)                      #get the info of all pages
        txtCharacterStr += PagesInfo.extract_text() + '\n'   #load all characters in a list
    PdfFile.close()

    txtLineList = []                               #list of lines converted to txt
    txtLineList.append('*** File Name ***')
    txtLineList.append(FileName)

    aline = ''                                     #for store the string of a line
    if CreateTXT == True:
        i = 0
        ResultFileName = FileName[:-3]+"txt"           #name a result/new txt file
        with open(ResultFileName,"w") as ResultTxt:    #create a txt with that result/new name
            for chara in txtCharacterStr:
                if chara != '\n':
                    aline += chara
                elif chara == '\n':
                    # print (aline)
                    txtLineList.append(aline)
                    ResultTxt.write(str(i)+'.'+aline+'\n')
                    i = i + 1
                    aline = ''
        ############# ↓for debug↓ #############
        # x = 0
        # last10char = ''
        # for x in range(10):
        #     last10char += txtCharacterStr[-(10-x)]
        #     print(last10char)
        # print ('the second last element in txtCharaterList is {',txtCharacterStr[-2],'}')
        # print ('the last element in txtCharaterList is {',txtCharacterStr[-1],'}')   
        ############# ↑for debug↑ #############
        ResultTxt.close()
        print (FileName[:-4],'txt (from PDF)','has',i,'lines')
    else:
        for chara in txtCharacterStr:
            if chara != '\n':
                aline += chara
            elif chara == '\n':
                # print (aline)
                txtLineList.append(aline)
                aline = ''

    os.chdir('..')
    return txtLineList
#↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑#

#↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓ group lines into different Sections ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
def SortLines(Lines : list[str]) -> dict[str:list[str]]:
    HeadTail = 'un'         #'H'or'T'or'TH'or'un': Head, Tail, Tail&Head, or Unknown
    # lines_with_marks = []   #lines with mark about whether they're Head, Tail or Tail&Head 
    SectionsData = {}       #seperate data in a file into different Sections  
    record = 'N'            #'Y' or 'N': the determinent for recording or not
    tag = ''     #to store the current Section-Name

    for L in Lines:
        #L = L.lower().strip('\n')      #the result dict would be lower case 
        L = L.strip('\n')       
        if L[0:3] == '***':             #when reach to those star-lines:
            landmark = L.strip(' *')    #store info of those star-lines
            if HeadTail=='un' or HeadTail=='T':     #check the previous star-line attribute 
                HeadTail = 'H'                      #based on previous attribute, it's a Head
                record = 'Y'                        #turn on recording mod
                # L = L + '.' + HeadTail                    #for debug use
                tag = landmark                      #info in the current star-line is a Section-Name
                SectionsData[tag] = ['info']        #create an dict-object with that Section-name
                ### The first element ↑↑↑ represets the attribute of that Section

            elif HeadTail=='H' or HeadTail=='TH':   #check the previous star-line attribute 
                #if landmark=='passed' or landmark=='failed':   #for lower case mod
                if landmark=='PASSED' or landmark=='FAILED':
                    HeadTail = 'T'                  #based on the current star-line's info, it's a Tail
                    # L = L + '.' + HeadTail                #for debug use
                else:
                    HeadTail = 'TH'     #based on the current star-line's info and the previous attribute
                    record = 'Y'                    #turn on recording mod
                    # L = L + '.' + HeadTail                #for debug use
                    tag = landmark                  #info in the current star-line is a Section-Name
                    SectionsData[tag] = ['info']    #create an dict-object with that Section-name
                    ### The first element ↑↑↑ a Section represets the attribute of that Section

        elif record == 'Y':         #when reach to unstar-lines and recording is on, do:
            SectionsData[tag].append(L)
        
        if HeadTail == 'T':                     #when the current status is Tail:
            SectionsData[tag][0] = landmark     #load Pass or Fail info to the current Section 
            record = 'N'                        #trun off recording mod

        # print (L)                 #for debug use
    #print ('--------------------------------------------')
    return SectionsData
#↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑#

#↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
class KeyData:
    def __init__(self) -> str:              #What they called /mean in FCT files:
        self.fn = 'File Name'               #'File Name'
        self.sa = 'Started at'              #'Started at'
        self.did = 'DeviceID'               #'DeviceID'
        self.sn = 'Serial number'           #'Serial number'
        self.pr = 'Part Replaced'           #'Part Replaced'
        self.dmc = 'DataMatrix Code'        #'DataMatrix Code'
        self.oid = 'Operator Id'            #'Operator Id'
        self.tm = 'Test mode'               #'Test mode'
        self.hwv = 'RTZ Hardware-Version'   #'Hardware-Version'
        self.fwv = 'RTZ Firmware-Version'   #'Firmware-Version'
        self.swv = 'RTZ Software-Version'   #'Software-Version'
        # self.tas = 'Temperature at Start'   #'Temperature at Start'
        # self.tae = 'Temperature at end'     #'Temperature at end'
        self.ftp = 'FTP port'               #'FTP port'
        self.pnp = 'P/NP?'                  #'Pass or Failed'
        self.fa = 'test fails at:'
        #↓↓↓↓↓↓↓↓↓↓↓ the 4 below are Section-Names ↓↓↓↓↓↓↓↓↓↓↓
        self.ts = 'Test Started'            #'Test Started'
        self.rtzv = 'RTZ Reader Version'    #'RTZ Reader Version'
        self.btf = 'BOARD TEST FAILED'      #'BOARD TEST FAILED'
        self.te = 'Test Ended'              #'Test Ended'
   
    def filtrate(self, DataDict: dict[str:list[str]]) -> dict[str:str]:   
        VarsDict = {}
        for i in CheckOrder:
                VarsDict[i] = None

        for L in DataDict[K.fn]:
            if 'html' in L:
                VarsDict[K.fn] = L
            elif 'pdf' in L:
                VarsDict[K.fn] = L

        for L in DataDict[K.ts]:
            #print (L)
            if K.sa in L:
                VarsDict[K.sa] = L.lstrip(K.sa + ' :')
            elif K.did in L:
                DID1 = L.lstrip(K.did + ' :') 
            elif K.sn in L:
                VarsDict[K.sn] = L.lstrip(K.sn + ' :')
            elif K.pr in L:
                VarsDict[K.pr] = L.lstrip(K.pr + ' :')
            elif K.dmc in L:
                VarsDict[K.dmc] = L.lstrip(K.dmc + ' :')
            elif K.oid in L:
                VarsDict[K.oid] = L.lstrip(K.oid + ' :')
            elif K.tm in L:
                VarsDict[K.tm] = L.lstrip(K.tm + ' :')

            if 'ATU ' in L:         
                DID2 = L.lstrip(K.did + ' :')
                if DID1 == DID2:
                    VarsDict[K.did] = DID1
                else:
                    VarsDict[K.did] = DID1 +' '+ DID2
        
        for L in DataDict[K.rtzv]:
            #print (L)
            if K.hwv[4::] in L:
                VarsDict[K.hwv] = L.lstrip(K.hwv[4::] + ' :')
            elif K.fwv[4::] in L:
                VarsDict[K.fwv] = L.lstrip(K.fwv[4::] + ' :')
            elif K.swv[4::] in L:
                VarsDict[K.swv] = L.lstrip(K.swv[4::] + ' :')

        if K.te in DataDict:
            ftp1 = ftp2 = ''
            for L in DataDict[K.te]:
                if 'enabling ftp' in L.lower():
                    ftp1 = 'Enabling'
                if 'fail' in L:
                    ftp2 = ', upload Fail'
                if 'success' in L:
                    ftp2 = ', upload Success'
            VarsDict[K.ftp] = ftp1 + ftp2
        else:
            VarsDict[K.ftp] = 'NO "Test Ended" Section'

        if K.btf in DataDict:
            VarsDict[K.pnp] = '✖'
            for D in DataDict:
                if DataDict[D][0] == 'FAILED':
                    VarsDict[K.fa] = D
        else:
            VarsDict[K.pnp] = '✔'

        return VarsDict
#↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑#           

'''===================================================================================================='''
#↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
def CheckRep(aDF:object, col:str, mutate=lambda x:x) -> dict[str:list[int]]:
    repetition = {}     #ends like:{'item1':[list of which row it shows],...}
    idx = 0             #that column's items index
    for i in aDF[col]:
        i_mut = mutate(i)
        if i_mut not in repetition:
            repetition[i_mut] = [idx]
        else:
            repetition[i_mut].append(idx)
        idx += 1
    return repetition   #{item1/mutant1_in_that_col:[list of which row it shows],...}
#↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑#

#↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
def Sort_Split_DF(theDF:object, col:str, mutate=lambda x:x, newDF=False) -> dict or tuple[dict,object]:

    similarity = CheckRep(theDF,col=col, mutate=mutate)
    #print (similarity)

    # writing = pd.ExcelWriter('test1.xlsx')                      #just for debug
    DFdict = {}
    new_DF = pd.DataFrame()
    for s in sorted(similarity):
        DFdict[s] = pd.DataFrame()
        for i in similarity[s]:
            tempDF = theDF.iloc[[i]]
            DFdict[s] = pd.concat([DFdict[s],tempDF], ignore_index=True)
        if newDF == True:
            new_DF = pd.concat([new_DF,DFdict[s]], ignore_index=True)
            #↑↑↑ "new_DF" will be a new sorted (by input col) DataFrame 

    #     DFdict[s].to_excel(writing, sheet_name=s,index=True)    #just for debug
    # writing.save()                                              #just for debug
    if newDF == False:
        return DFdict   #keys-sorted {item1/mutant1_in_that_col:<a DF of rows having it>,...}
    elif newDF == True:
        return DFdict, new_DF
#↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑#

#↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
def SortDF_by_cols(DF:object, bys:list[str], Locs:list[list]=[], keys:list[tuple]=[]) -> tuple[object,list]:
#'bys': a list of cols you want to sort by, first in fisrt priarity/layer           
#'Locs': a list of where the most-recent/last-layer/last-col sorted rows locate     Ex:[[0,1,2],[3],[4,5]]
#'keys': a list of tuple: (which No. of 'by' the 'key' applys to, key's function), starts at No.1  
         #Ex: bys=[col1, col2, col3], keys=[(1,function1),(3,function3)]
            
    key = lambda a: a               #create a default 'key' return orginal input 
    for k in keys:
        if k[0] == len(bys):
            key = k[1]              #find and use the correspunding key-function
    
    if len(bys) == 1:
        SS = Sort_Split_DF(DF, bys[0],mutate=key, newDF=True)
        DF = SS[-1]                             #get the new DateFrame
        Locs = CheckRep(DF,bys[0],mutate=key).values()      #get the new locations
        return DF,Locs
    else:
        pre_DF,pre_Locs = SortDF_by_cols(DF=DF, bys=bys[:-1], keys=keys, Locs=Locs)

        #***↓ Slice input DF to sub-DFss basded on previous sorting ↓***
        DFs = []
        for i in pre_Locs:
            tempdf = pre_DF.iloc[i[0]:i[-1]+1]
            DFs.append(tempdf)
        #******************************↑↑↑******************************

        DF = pd.DataFrame()
        idxRecord = 0
        Locs = []
        for x in DFs:
            SS = Sort_Split_DF(x, bys[-1], mutate=key, newDF=True)  #sort each sub-DF alone
            DF = pd.concat([DF,SS[-1]], ignore_index=True)      #get the new DF by adding up sorted sub-DFs    
            tempL = CheckRep(SS[-1], bys[-1], mutate=key).values()  #get local Locations of each sub-DF itself
            #******↓ Get the new global Locations related to the whole new DF ↓******
            for y in tempL:
                Locs.append([y[0]+idxRecord,y[-1]+idxRecord])   #get the new Locations
            idxRecord = Locs[-1][-1]+1
            #***********************************↑↑↑***********************************

        return DF,Locs
#↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑#

#↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
def SortTable(pandasDF:object, bys:list[str], keys:list[tuple]=[], space=False,IndexCol=None) -> tuple[object,int]:

    pandasDF = SortDF_by_cols(pandasDF, bys=bys, keys=keys)[0]

    if space == True:
        #*****↓ add an empty row between different groups, only for First sorted_col ↓******
        trans = lambda a: a             #create a default key-function return orginal input 
        for k in keys:
            if k[0] == 1:
                trans = k[1]            #find and use the key-function for First sorted_col
        similarity = CheckRep(pandasDF, bys[0], mutate=trans)

        resultDF = pd.DataFrame()
        empltyDF = pd.DataFrame([' '],columns=[bys[0]],index=[''])
        for y in similarity:
            tempDF = pandasDF.iloc[similarity[y][0]:similarity[y][-1]+1]
            tempDF = pd.concat([tempDF,empltyDF], ignore_index=False)
            resultDF = pd.concat([resultDF,tempDF], ignore_index=False)
        #***************************************↑↑↑↑↑***************************************
    else:
        resultDF = pandasDF

    #***↓ What and how many special index needed or just 1 default ↓***
    if IndexCol != None and IndexCol !=[]:
        if type(IndexCol) == str:
            newIdxCols = IndexCol
            IdxCount = 1
        elif type(IndexCol) == list:
            newIdxCols = []
            for i in IndexCol:
                if i not in newIdxCols:
                    newIdxCols.append(i)
            IdxCount = len(newIdxCols)
        resultDF.set_index(newIdxCols, inplace=True)
    else:
        IdxCount = 1
    #******************************↑↑↑******************************
    
    return resultDF,IdxCount
#↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑#

# def Last8SN(series):
#     return series.apply(lambda s: s[-8::])

#↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
def BetterWidth (string: str) -> int:
    HalfWidthGroup = [" ", ":", ";", ".", ",", "/", "i", "l", "j"]
    adjust = 0.0
    for c in string:
        if c in HalfWidthGroup:
            adjust += 0.5
    CellWidth = len(string) - int(adjust)
    return CellWidth
#↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑#

#↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
def toExcel (FinalDataList: list[dict[str:str]], Cols_to_Sort: list[str] = ['Serial number']) -> str:
    ExcelName = 'FCT test files to excel.xlsx'
    MyData = pd.DataFrame(FinalDataList)
    writing = pd.ExcelWriter(ExcelName, engine='xlsxwriter')

    #********↓ Specify what exact info you need from Serial Number or File Name↓********
    ProductDate = lambda x: x[-4]+x[-7:-4]
    Revision = lambda x: x[8:10]
    PlantCode = lambda x: x[10]
    ProductNum = lambda x: x[0:8]
    File_Type = lambda x: x.split('.')[-1]

    # Cols_to_Sort = [K.sn, K.sn, K.sn]                             #for test
    # FuncsList = [(1,ProductDate),(2,ProductNum),(3,Revision)]     #for test

    # Cols_to_Sort = [K.did,K.tm,K.sn,K.pnp]        #sequence decided after 2022 Oct 9th
    FuncsList = []
    my_Indexs = [K.sn,K.pnp,K.fa,K.tm]
    #**********************************↑↑↑**********************************

    MyData,IdxCount = SortTable(MyData,bys=Cols_to_Sort,keys=FuncsList,space=True, IndexCol=Cols_to_Sort)
    ''' parameter 'bys' list needs at least one element, 'keys' and 'Indexcol' can be an empty list 
        bys: Sort the table by which Columns' content
        keys: Sort some specific part you need from Columns to be sorted, such as Revision num from Seriral Num
            eg. [(3,func3)] means applying 'func3' to the 'third' Column in 'bys' list'''

    print (MyData,'\n---------------------------------------------------------------')

    MyData.to_excel(writing, sheet_name='Summary',index=True)
    
    #*********↓ give index column enough width ↓*********
    if IdxCount == 1:
        IdxLens = []
        if MyData.index.name != None:
            IdxLens = [BetterWidth(MyData.index.name)]
            print ("*****************\n",type(MyData.index.name),"\n**********************")
  
        for i in MyData.index:
            IdxLens.append(BetterWidth(str(i))) 
        writing.sheets['Summary'].set_column(0,0,max(IdxLens)+2)
    else:
        IdxLens = {}
        x = 0
        for n in MyData.index.names:
            IdxLens[x] = [BetterWidth(str(n))]
            x = x+1
        for I in MyData.index:
            x = 0
            for i in I:
                IdxLens[x].append(BetterWidth(str(i)))
                x = x+1
        for x in range(IdxCount):
            IdxLens[x] = max(IdxLens[x]) + 2
            writing.sheets['Summary'].set_column(x,x,IdxLens[x])
    #***********************↑↑↑***********************

    for col in MyData:
        toStr = MyData[col].astype(str)    #asign each item in a column to string type
        MaxLen = toStr.map(BetterWidth).max()      #get the length of each item and find Max
        MaxLen = max(MaxLen, BetterWidth(str(col)))      #include the title length into comparison
        # if MaxLen>20 and MaxLen<41:
        #     MaxLen = int(MaxLen*0.9)        #cell width decorate, not necessary
        idx = MyData.columns.get_loc(col) + IdxCount      #index of the current column
        writing.sheets['Summary'].set_column(idx, idx, MaxLen+1)
    
    #print (MyData)
    #print (type(MyData))
    #print (MyData.dtypes)
    #print (MyData.iloc[6:23])
    writing.save()
    return ExcelName
#↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑#

################################ ↓↓↓ <Global Varibales and Main Function> ↓↓↓ ##################################
K = KeyData()       #contain info's titles which the output excel needs
path = os.getcwd()    #get the path to the folder where this .py file is 
#↓↓↓ The default order of columns presenting in the excel ↓↓↓
CheckOrder = [K.fn,K.did,K.sn,K.tm,K.sa,K.pr,K.dmc,K.oid,K.hwv,K.fwv,K.swv,K.ftp,K.pnp,K.fa] 
 
def main():
    FilesList = []      #list of files (a list of a list of string)
    htmlsCount = 0      #how many html files it converts
    pdfsCount = 0       #how many pdf files it converts   
    htmls_and_pdfs = GetFileNames(path)   #the list of 2 sets: html and pdf file names

    while (True):
        print ("Please make sure you put right files into 'FCT_files' folder")
        print ('Which type of files you want to convert?')
        print ('(Enter: 1 for HTML, 2 for PDF, 3 for both, 0 for exit)')
        choose = input()

        if choose == '1':
            for fn in htmls_and_pdfs[0]:
                print ('FilesList[',htmlsCount,'] is:',fn)
                FilesList.append(html2txt(fn))
                htmlsCount += 1
            print ('\nTotally converted',htmlsCount,'HTML files to txt')
            break

        elif choose == '2':
            for fn in htmls_and_pdfs[1]:
                print ('FilesList[',pdfsCount,'] is:',fn)
                FilesList.append(pdf2txt(fn))
                pdfsCount += 1
            print ('\nTotally converted',pdfsCount,'PDF files to txt')
            break

        elif choose == '3':
            for fn in htmls_and_pdfs[0]:
                print ('FilesList[',htmlsCount,'] is:',fn)
                FilesList.append(html2txt(fn))
                htmlsCount += 1
            for fn in htmls_and_pdfs[1]:
                print ('FilesList[',htmlsCount+pdfsCount,'] is:',fn)
                FilesList.append(pdf2txt(fn))
                pdfsCount += 1
            print ('\nTotally loaded',htmlsCount,'HTML and',pdfsCount,'PDF files')
            break

        elif choose == '0':
            print ('Bye, Later!\n')
            break
        else:
            # print (type(choose))
            # print (choose == '')
            print ('I can not understand your command, please try again\n')

    if choose != '0':
        print ('\nHow would you like to sort your table? Options are:')
        print (" 1.'FN': File Name\n 2.'SA': Start at\n 3.'DID': DeviceID\n 4.'SN': Serial number")
        print (" 5.'PR': Part Replaced\n 6.'DMC': DataMatrix Code\n 7.'OID': Operator Id\n 8.'TM': Test mode")
        print (" 9.'HWV': Hardware-Version\n 10.'FWV': Firmware-Version\n 11.'SWV': Software-Version")
        print (" 12.'FTP': FTP port\n 13.'PNP': Pass/NO pass\n 14.'FA': test fails at")
        print ('Please type in your choices(in the form of provided abbreviations) and then press Enter\n'
                "For more than one, seperate them by 'Space'. Orders matter like first priority, second priority..."
                "\neg:DID TM SN FTP")
        
        understood = False
        while (understood == False):
            understood = True
            AimCols:list[str] = []
            InputStr = input()
            inputs = InputStr.split()
            AllowedInputs = ['FN','SA','DID','SN','PR','DMC','OID','TM','HWV','FWV','SWV','FTP','PNP','FA']
            for i in inputs:
                if i not in AllowedInputs:
                    print ('Can Not understand:{',i,'}Please re-type your choices')
                    understood = False
                else:
                    if i == 'FN':
                        AimCols.append(K.fn)
                    elif i == 'SA':
                        AimCols.append(K.sa)
                    elif i == 'DID':
                        AimCols.append(K.did)
                    elif i == 'SN':
                        AimCols.append(K.sn)
                    elif i == 'PR':
                        AimCols.append(K.pr)
                    elif i == 'DMC':
                        AimCols.append(K.dmc)
                    elif i == 'OID':
                        AimCols.append(K.oid)
                    elif i == 'TM':
                        AimCols.append(K.tm)
                    elif i == 'HWV':
                        AimCols.append(K.hwv)
                    elif i == 'FWV':
                        AimCols.append(K.fwv)
                    elif i == 'SWV':
                        AimCols.append(K.swv)
                    elif i == 'FTP':
                        AimCols.append(K.ftp)
                    elif i == 'PNP':
                        AimCols.append(K.pnp)
                    elif i == 'FA':
                        AimCols.append(K.fa)
            # print (AimCols)

        Welldone = []       #each element will be a dict of data we want from a file
        for F in FilesList:
            Sections = SortLines(F)
            Welldone.append(K.filtrate(Sections))
        print ('An excel has been created -> {',toExcel(Welldone,AimCols),'}\n')

if __name__ == '__main__':
    main()









 