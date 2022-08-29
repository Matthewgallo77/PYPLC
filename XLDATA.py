
from genericpath import getsize
from lib2to3.pgen2.token import NAME
from numbers import Real
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
from snap7.common import check_error, ipv4, load_library
from snap7.types import S7WLBit, S7WLByte, S7WLInt, S7WLReal

# DOCUMENTATION: https://python-snap7.readthedocs.io/en/1.0/client.html#snap7.client.Client.as_download   

def excelConnect(fileName):
    excelWorkbook = load_workbook(fileName)
    uiSheet = excelWorkbook.active
    lastRow = str(len(list(uiSheet.rows))) # GETS THE LAST NON EMPTY ROW

    NAMES, DATATYPES, ADDRESSES = uiSheet['A2':'A'+lastRow], uiSheet['C2':'C'+lastRow], uiSheet['D2':'D'+lastRow] # ACCESS NECESSARY COLUMNS

    nameList, datatypeList, addressList = [], [], [] # INITIALIZE LISTS

    nameList, datatypeList, addressList  = getDATA(NAMES, nameList), getDATA(DATATYPES, datatypeList), getDATA(ADDRESSES, addressList)
    for value in addressList:
        addressList=list(map(lambda x: x.replace(value,re.sub("[^\d\.]", "", value)), addressList))
    address_dataType = {addressList[i]: datatypeList[i] for i in range(len(addressList))} # CONTAINS {ADDRESS: DATATYPE}

    return address_dataType


def getDATA(columnInput, infoList):
        infoList = []
        for CELL in columnInput:
            for INFO in CELL:
                infoList.append(INFO.value)
        return infoList

def ReadMerker(PLC, byte, bit, datatype):
    byteArray = PLC.read_area(snap7.types.Areas.MK, 0, byte, datatype)
    if datatype == snap7.types.S7WLBit:
        return get_bool(byteArray, byte, bit)
    elif datatype == snap7.types.S7WLInt:
        return get_int(byteArray, byte)
    elif datatype == snap7.types.S7WLReal:
        return get_real(byteArray, byte)
    else:
        return 'ERROR: Data type has not been anticipfated'

# def WriteMerker(PLC, byte, bit, datatype):
#     # GET VALUE FROM INPUT IN EXCEL VALUE CELL
#     # HAVE SCRIPT RUN WHEN VALUE CELL CHANGES... here is article:
#     # https://docs.microsoft.com/en-us/office/troubleshoot/excel/run-macro-cells-change
#     byteArray = PLC.read_area(snap7.types.Areas.MK, 0, byte, datatype)
    if datatype == snap7.types.S7WLBit:
        return get_bool(byteArray, byte, bit, value)
    elif datatype == snap7.types.S7WLInt:
        return get_int(byteArray, byte, value)
    elif datatype == snap7.types.S7WLWord:
        return get_real(byteArray, byte, value)
    else:
        return 'ERROR: Data type has not been anticipfated'
#     

if __name__ == "__main__":
    tagInfo = excelConnect('PLCTags.xlsx')
    PLC1 = snap7.client.Client()
    try:
        PLC1.connect('192.168.0.1',0,1) # IP, RACK #, SLOT #
        print("CONNECTION STATUS: \n" + "PLC1: " + str(PLC1.get_connected())) # DISPLAYS IF CONNECTION TO PLC IS VALID
    except:
        print("CONNECTION STATUS: \n" "PLC1: " + str(PLC1.get_connected()))
    
