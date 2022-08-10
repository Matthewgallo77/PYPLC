from genericpath import getsize
from lib2to3.pgen2.token import NAME
from tkinter.tix import COLUMN
from tracemalloc import start
from turtle import st
from unicodedata import name
from uuid import NAMESPACE_DNS
from openpyxl import load_workbook
import snap7 # IMPORT SNAP7 LIBRARY
from snap7.util import *
import re
import struct # TO CONVERT BYTE ARRAY TO FLOATING POINT
from sys import getsizeof # IMPORT TO GET SIZE OF DATA

# DOCUMENTATION: https://python-snap7.readthedocs.io/en/1.0/client.html#snap7.client.Client.as_download   

excelWorkbook = load_workbook('ImpTags.xlsm')
uiSheet = excelWorkbook.active
lastRow = str(len(list(uiSheet.rows))) # GETS THE LAST NON EMPTY ROW

NAMES, DATATYPES, ADDRESSES = uiSheet['A2':'A'+lastRow], uiSheet['C2':'C'+lastRow], uiSheet['D2':'D'+lastRow] # ACCESS NECESSARY COLUMNS

nameList, datatypeList, addressList = [], [], [] # INITIALIZE LISTS

def getDATA(columnInput, infoList):
    infoList = []
    for CELL in columnInput:
        for INFO in CELL:
            infoList.append(INFO.value)
    return infoList

nameList, datatypeList, addressList  = getDATA(NAMES, nameList), getDATA(DATATYPES, datatypeList), getDATA(ADDRESSES, addressList)

address_dataType = {addressList[i]: datatypeList[i] for i in range(len(addressList))} # CONTAINS {ADDRESS: DATATYPE}

print(address_dataType)

def ReadMemory(PLC, byte, bit, datatype):
    result = PLC.read_area(snap7.types.Areas.MK, 0, byte, datatype)
    if datatype == 'Bool':
        return get_bool(result, 0, bit)
    elif datatype == 'Int':
        return get_int(result, 0)
    elif datatype =='Real':
        return get_real(result, 0)
    else:
        return 'ERROR: Data type has not been anticipated'

if __name__ == "__main__":
    PLC1 = snap7.client.Client()
    PLC2 = snap7.client.Client()
    try:
        PLC1.connect('192.168.10.2',0,1) # IP, RACK #, SLOT #
        PLC2.connect('192.168.10.2',0,1)
        print("CONNECTION STATUS: " + str(PLC1.get_connected())) # DISPLAYS IF CONNECTION TO PLC IS VALID
    except:
        print("CONNECTION STATUS: \n" "PLC1: " + str(PLC1.get_connected()) + "\n" + "PLC2: " + str(PLC2.get_connected()))





 
