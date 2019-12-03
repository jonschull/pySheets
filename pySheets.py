#!/usr/bin/env python
# coding: utf-8

# ## pySheets:  
# ### read andwrite pandas data to and from Excel and google Worksheets 
# #### Introduction

# pySheets uses python to transfer data between spreadsheets (XLSs or Google) and easy-to-manipulate pandas Dataframes.  
# 
# DataFrames are powerful objects, with a complicated.  PySheets provides some utilities for dealing with a more pythonic data object called (here) a DictArray.  (It's the equivalent of __df.to_excel('records')__
# 
# ### Reading and Writing 
# 
# * __getDFB(file_or_gsURL)__  retrieves data from a filepath or the URL of a googleSheet.
# * __writeDFB(DFB, file_or_gsURL)__ (writes data to a filepath or the URL of a googleSheet.
# 
# Thus, to read or write an Excel file, you need its path (e.g., 'myExcelFile.xlsx')
# 
# To read or write a Google Spreadsheet, you need 
# 
# * a credentials file, and the URL of the sheet.__
#     e.g., __gs = getGoogleSheet(credentialsFile, URL)__
#     ([Here](https://gspread.readthedocs.io/en/latest/oauth2.html) is how to get a credentials file.)
# 
# * and the spreadsheet must be shared with the "user" identified in the credentials file.
# 
# We call the resulting data structure a "DataFrameBook" or DBF)
# * An Excel or Google workbook may comprise multiple worksheets, each with it's own name. 
# * __Note: when reading a google sheet, we assume the first row is column headings and we modify the resulting dataframe to preserve the spreadsheet's column ordering*__  
# 
# 
# * Similarly a DataFrameBook (a DFB) comprises one or more dataframes, as returned by *pd.read_excel*.
# 
# 
#     OrderedDict('nameOfFirstDF': firstDF,
#                 'nameofSecondDF: secondDF')
# 
# Then, you can read/write a file or spreadsheet to or from a DFB using __the only two functions you should have to use__ (assuming you have your credentials file, as for GoogleSheets.)
# * __getDFB(file_or_gsURL)__
# * __writeDFB(DFB, file_or_gsURL)__
# 
# The Utilties we provide are 
# * log
# * show
# * showDBF

# ### utilities: log
# Write to pySheets.log.txt, newest first

# In[1]:


#!pip3 install -U -r requirements.txt
from IPython.display import display, HTML
TESTING=False
VERBOSE_LOGGING=False
EXPLAINING = False

#!pip3 install arrow
import arrow
def now():
    return arrow.utcnow().to('US/Eastern').format('MM/DD hh:mm a')
    
def log(*args):
    msg = [str(arg) for arg in args]
    msg=' '.join(msg)
    import arrow
    log=open('pySheets.log.txt', 'a')
    output=str(now() + ' pySheets.log.txt: '   + msg)
    log.write( output + '\n' )
    if VERBOSE_LOGGING: print( output )
        
        
import pandas
from pandas import DataFrame

if TESTING: 
    log('abc','def', now())
    get_ipython().system('head pySheets.log.txt')


# ## introducing DataFrames and DictArrays
# pandas is all about DataFrames.  
# They display nicely in notebooks.
# 
# One way to populate a DataFrame is by using an array of dicts or OrderedDicts.  
# I call an array of dicts a "DictArray"

# In[2]:


if EXPLAINING:
    import pandas
    from pandas import DataFrame
    dictArray = [dict(a=1,b=2),
                 dict(a=3,b=4),
                 dict(a=5,b=6)]

    DataFrame(dictArray)  


# ###: showDF()
# 
# As you see above, DataFrames display nicely in notebooks when they are the last item in a cell.
# 
# But as you see below, they don't print so as nicely.   showDF addresses this issue.

# In[ ]:


def showDF(df):
    """showDF renders dataframes even they are not the last item in a cell' )"""
    display(HTML(df.to_html()))

if EXPLAINING:
    df = DataFrame(dictArray)
    print(df)
    print("\nPrinting dataframes is not pretty.  \n\nSo here's showDF:")
    
    showDF(df)

    print('See? showDF renders dataframes nicely even in the middle of a cell.')


# ### utlities: DictArray 

# DataFrames are powerful but bewildering.
# DictArrays are simple python objects.  
# I sometimes find it convenient to convert a DataFrame into a DictArray.
# This is equivalent to __df.to_dict('records')__
# 

# In[ ]:


from collections import OrderedDict
"""
def DictArray(df):
    #DictArray converts a DataFrame into a Dictarray, just as DataFrame converts a DictArray into a Dataframe
    colheads = list(df.keys())
    values = df.values

    dictArray = []
    for rowIndex in range(len(df)):
        rowDict=OrderedDict()
        for colIndex,colhead in enumerate(colheads):
            rowDict[colhead] = values[rowIndex, colIndex]
        dictArray.append(rowDict)
    
    return dictArray
"""

def DictArray(df):
    return df.to_dict('records')

if EXPLAINING:
    showDF(df)
    da = DictArray(df)
    print(da)


# ### Convenience Functions for gridData: (da or df)
# DictArrays and DataFrames represent the same data. 
# Here are some convenient functions for dealing with such "gridData" regardless of their format.
# * show()*  displays DictArrays /or/ DataFrames 
# * rowValues()
# * colValues()

# In[ ]:


def show(gridData):
    "accept DF or DA"
    if type(gridData)==type([]):
        gridData=DataFrame(gridData)
    showDF(gridData)
    
DataFrameType = type(DataFrame())
    
def rowValues(gridData,index):
    """ Accepts DF or DA """
    if type(gridData) == DataFrameType:  # then convert to DictAraay
        gridData = DictArray(gridData)
    
    return list(gridData[index].values())
        
def colValues(gridData, key):
    """ Accepts DF or DA """
    if type(gridData) == DataFrameType:  
        gridData = DictArray(gridData)

    return list(DataFrame(gridData)[key])


if TESTING:
    dictArray = [dict(a=1,b=2),
                dict(a=3,b=4),
                dict(a=5,b=6)]
    
    df = DataFrame(dictArray)
    da = DictArray(df)
    print('We went from DictArray to DataFrame and back, and the data match:',  da==dictArray)
    print()
    print('print(df) and print(da)', 'provide different renderings:')
    print('print(df)\n', df)
    print()
    print('print(da)\n', da)
    print()
    print('But show(df) and show(da) provide the same pretty rendering:')
    show(df)
    show(da)
    print()
    print('Similarly, rowValues and colValues accept either DA or DF and return the same simple list of values:')
    print('\t rowValues with DF and DA:', rowValues(df,1),   rowValues(da,1) )  
    print('\t colValues with DF or DA:', colValues(df,'a'), colValues(da,'a') )


# ## Introducing Spreadsheets
# Modern spreadsheets are structured as workbooks comprised of multiple worksheets.
# 
# Each worksheet has its own name.
# Each worksheet is a dataGrid comprised of numbered rows and named columns.
# 
# We model this as a dictionary of DataFrames or "DataFrameBook".

# ### utilities: showDFB()
# for pretty printing of DFBs
# 

# In[ ]:


def showDFB(dfb,msg='showDFB: '):
    keys=list(dfb.keys())
    display(HTML('<i><b>' + msg + '</i></b>' + '  This DataFrameBook contains these sheets:  <b>' + str(keys) + '</b>...'))
    for key in keys:
        display(HTML('<hr/>'))
        display(HTML('<i>'+ key + '</i>'))
        show(dfb[key])
        
if TESTING:
    dfb=dict(Sheet1= DataFrame([   dict(a='Sheet1:row0a', b='Sheet1:row0b'),
                                   dict(a='Sheet1:row0a', b='Sheet1:row1b')]),

             Sheet2= DataFrame([ dict(a='Sheet2:row0a', b='Sheet2:row0b'),
                                 dict(a='Sheet2:row1a', b='Sheet2:row1b')] ))
    
    showDFB(dfb)
    
### NOTE: showDFB will actually work with DictArrays as well as DataFrames but that's not true for other functions below.


# ## enter GoogleSheets....
# 
# Unlike an XLS file, the natural habitat of a GoogleSheet is doc.google.com, not a local file system.
# 
# To access a googleSheet 
# * it must be readable by the public 
# * you need a credentials file.  
# * You need to have shared a googleSheet with the email address in the Credentials file.
# ([Here](https://gspread.readthedocs.io/en/latest/oauth2.html) is how to get a credentials file.)
# 
# Then you will be able will be able to use __get_DFB__ and __write_DFB__ to interact with GoogleSheets.

# Soon, you will use **get_dfb** and **write_dfb** to read and write DFBs to your local hard drive (as xls files) or to Google Drive (as Google Sheets)
# 
# But first we need to define the conversion Utilties that make this possible.  (You don't need to learn them).

# ## Internal utilities for converting between BDF and Excel, and for reading and writing to disk. 
#  
# <img src='https://www.evernote.com/l/AAEkhdm5GChC8ohEQuZgAPc1uvoABFMT-hIB/image.png' width=30% height=30%>

# #### internal function:  DBF <-> XL

# In[ ]:


#!pip3 install XlsxWriter

def DFB_to_XL(DFB, fileName='test.xlsx'):
    #https://xlsxwriter.readthedocs.io/example_pandas_multiple.html
    writer = pandas.ExcelWriter(fileName, engine='xlsxwriter')
    for name, df in DFB.items():
        df.to_excel(writer, sheet_name=name, index=False)
    writer.save()
    log('DFB_to_XL:  Wrote to', fileName)

def XL_to_DFB(fileName='df2gSheetTester.xlsx'):
    """take an xlsx file and return an ordered dict of DataFrames (DFB)"""
    log('xl_to_DFB:  Reading', fileName)
    return pandas.read_excel(fileName,sheet_name=None)

def test_DFB_XL_DFB():
    DFB_to_XL(dfb,'test.xls')
    new_dfb = XL_to_DFB('test.xls')
    
    showDFB(dfb, 'original dfb')
    print()
    showDFB(new_dfb, 'reconstructed dfb')
 
if TESTING:  
    test_DFB_XL_DFB()
    


# 

# #### Internal function: getGoogleSheets

# In[ ]:


#!pip3 install pygsheets
import pygsheets

CREDENTIALSFILE = 'enablebadger-b31383b767ef.json'
pySheetsTesterURL = 'https://docs.google.com/spreadsheets/d/1y19PMDTSqZkpW3UcwqG81tQMAH_g91iIqjrmBN_jTRw/edit#gid=0'

def getGoogleSheet(gsURL=pySheetsTesterURL,
                  credentialsFile = CREDENTIALSFILE):
    """
    1. Have credentials in credentials fileName.  [get google drive api credentials]
    2. Sheet must exist at URLofSheet
    3. Sheet must be shared with the email addess in the credentialsFile
    """
    googleClient = pygsheets.authorize(service_file=credentialsFile)
    log('getGoogleSheet:  ', googleClient)
    return googleClient.open_by_url(gsURL)

def test_getGoogleSheet():
    gs = getGoogleSheet()
    print(gs.title, gs.worksheets())

if TESTING: 
    test_getGoogleSheet()


# #### Internal functions:  DFB <-> GS

# In[ ]:


def GS_to_DFB(gsURL):
    """take the URL of a Gsheet and return a DFB"""
    global gs  #TODO remove this;  probably from testing 
    #global topRow
    gs=getGoogleSheet(gsURL)
    from collections import OrderedDict
    OD = OrderedDict()
    for worksheet in gs.worksheets():
        title = worksheet.title
        df = DataFrame(worksheet.get_all_records())
        
        ## note we are presuming that row 1 is column headings
        topRow=worksheet.get_row(1)
        while '' in topRow: # remove columns with no headings
            topRow.remove('')
        df = df.reindex(columns=topRow)
        
        OD[title] = df
    return OD 

def test_GS_to_DFB():
    URL = 'https://docs.google.com/spreadsheets/d/1y19PMDTSqZkpW3UcwqG81tQMAH_g91iIqjrmBN_jTRw/edit#gid=0'
    URL = 'https://docs.google.com/spreadsheets/d/1-EAzjpHbbaD6YWt0YGE2Vu3rXRXAvQIS4mq7uxKQmCs/edit#gid=0'
    dfb = GS_to_DFB(URL)
    showDFB(dfb, URL)

if TESTING: 
    test_GS_to_DFB()


# In[ ]:





# In[ ]:


def DFB_to_GS(DFB, gsURL):
    gs = getGoogleSheet(gsURL)
    while len(DFB)>len(gs.worksheets()): #add sheets if necessary
        gs.add_worksheet('temp',1,1)
    for i,title in enumerate(DFB.keys()):
        gs[i].clear()
        gs[i].title = title
        gs[i].set_dataframe(DFB[title],(1,1))# insert the data
    log('DFB_to_GS: updated', gs.title)

def test_DFB_to_GS(): #test: sheet in cloud changes
    DFB = XL_to_DFB('test.xls')
    gsURL='https://docs.google.com/spreadsheets/d/1y19PMDTSqZkpW3UcwqG81tQMAH_g91iIqjrmBN_jTRw/edit#gid=1856401945'
    DFB_to_GS(DFB,gsURL)

if TESTING: 
    test_DFB_to_GS()


# ## pySheets.get_DFB 

# In[ ]:


def get_DFB(fileName_or_gsURL):
    log('get_DFB:  working on', fileName_or_gsURL)
    if fileName_or_gsURL.startswith('http'):
        return GS_to_DFB(fileName_or_gsURL)
    else:
        return XL_to_DFB(fileName_or_gsURL)

def test_get_DFB():
    #dfb =  get_DFB('test2.xls')
    #showDFB(dfb, 'DFB from test2.xls')
    ProductionSheetURL = 'https://docs.google.com/spreadsheets/d/1-EAzjpHbbaD6YWt0YGE2Vu3rXRXAvQIS4mq7uxKQmCs/edit#gid=0'

    dfb = get_DFB( ProductionSheetURL )
    showDFB(dfb, 'DFB from google Sheet')

if TESTING: 
    test_get_DFB()


# ## pySheets.write_DFB

# In[ ]:


def write_DFB(dfb, fileName_or_gsURL):
    """Accepts filename or Google Sheets URL as destination"""
    
    log('write_DFB processing', dfb.keys(), fileName_or_gsURL)
    
    if fileName_or_gsURL.startswith('http'):
        DFB_to_GS(dfb,fileName_or_gsURL)
    else:
        DFB_to_XL(dfb,fileName_or_gsURL)


def test_write_DFB():
    write_DFB(dfb,'test.xls')
    
    gsURL='https://docs.google.com/spreadsheets/d/1y19PMDTSqZkpW3UcwqG81tQMAH_g91iIqjrmBN_jTRw/edit#gid=1856401945'
    write_DFB(dfb,gsURL)

if TESTING: 
    test_write_DFB()


# In[ ]:





# In[ ]:




