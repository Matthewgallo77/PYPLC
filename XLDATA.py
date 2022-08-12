from genericpath import getsize
from lib2to3.pgen2.token import NAME
from tkinter.tix import COLUMN
from tracemalloc import start
from turtle import st
from unicodedata import name
from uuid import NAMESPACE_DNS
from numpy import record
from openpyxl import load_workbook
from snap7.util import *
import snap7 # IMPORT SNAP7 LIBRARY
import re
import struct # TO CONVERT BYTE ARRAY TO FLOATING POINT

# DOCUMENTATION: https://python-snap7.readthedocs.io/en/1.0/client.html#snap7.client.Client.as_download   

def dataCollect():
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

    for value in addressList:
        addressList = list(map(lambda x: x.replace(value, re.sub("[^\d\.]", "", value)), addressList))

    address_dataType = {addressList[i]: datatypeList[i] for i in range(len(addressList))} # CONTAINS {ADDRESS: DATATYPE}
   
    return address_dataType

def ReadMemory(PLC, byte, bit, datatype):
    if datatype == 'Bool':
        return get_bool(byteArray, byte, bit)
    elif datatype == 'Int':
        return get_int(result, byte)
    elif datatype =='Real':
        byteArray = PLC.read_area(snap7.types.Areas.MK, 0, byte, 4)
        value = round((struct.unpack('>f', byteArray))[0], 2)
        return value
    else:
        return 'ERROR: Data type has not been anticipated'

if __name__ == "__main__":
    
    PLC1 = snap7.client.Client()
    try:
        PLC1.connect('192.168.10.1',0,1) # IP, RACK #, SLOT #
        print("CONNECTION STATUS: \n PLC1: " + str(PLC1.get_connected())) # DISPLAYS IF CONNECTION TO PLC IS VALID
    except:
        print("CONNECTION STATUS: \n" "PLC1: " + str(PLC1.get_connected()))

    excelWorkbook = load_workbook('PLCTags.xlsx')
    uiSheet = excelWorkbook.active
    DATA = dataCollect()
    READ = str(ReadMemory(PLC1, 112, 0, DATA['112']))
    uiSheet.cell(row=13, column=9).value = READ
    excelWorkbook.save('PLCTags.xlsx')




 
