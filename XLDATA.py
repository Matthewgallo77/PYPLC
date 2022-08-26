
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
from snap7.types import S7SZL, Areas, BlocksList, S7CpInfo, S7CpuInfo, S7DataItem
from snap7.types import S7OrderCode, S7Protection, S7SZLList, TS7BlockInfo, WordLen
from snap7.types import S7Object, buffer_size, buffer_type, cpu_statuses

# DOCUMENTATION: https://python-snap7.readthedocs.io/en/1.0/client.html#snap7.client.Client.as_download   

def excelConnect(fileName):
    excelWorkbook = load_workbook(fileName)
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
    for value in addressList:
        addressList=list(map(lambda x: x.replace(value,re.sub("[^\d\.]", "", value)), addressList))
    address_dataType = {addressList[i]: datatypeList[i] for i in range(len(addressList))} # CONTAINS {ADDRESS: DATATYPE}

def ReadMerker(PLC, byte, bit, datatype):
    byteArray = PLC.read_area(snap7.types.Areas.MK, 0, byte, datatype)
    if str(datatype) == 'Bool':
        return get_bool(byteArray byte, bit)
    elif str(datatype) == 'Int':
        return get_int(byteArray, byte)
    elif str(datatype) =='Real':
        return get_real(byteArray, byte)
    else:
        return 'ERROR: Data type has not been anticipfated'

def WriteMerker(PLC, byte, bit, datatype):
    # GET VALUE FROM INPUT IN EXCEL VALUE CELL
    # HAVE SCRIPT RUN WHEN VALUE CELL CHANGES... here is article:
    # https://docs.microsoft.com/en-us/office/troubleshoot/excel/run-macro-cells-change
    byteArray = PLC.read_area(snap7.types.Areas.MK, 0, byte, datatype)
    if str(datatype) == 'Bool':
        return set_bool(byteArray, 0, bit, value)
    elif str(datatype) == 'Int':
        return set_int(byteArray, 0, value)
    elif str(datatype) =='Real':
        return set_real(byteArray, 0, value)
    else:
        return 'ERROR: Data type has not been anticipfated'

if __name__ == "__main__":
    tagInfo = excelConnect('ImpTags.xlsm')
    PLC1 = snap7.client.Client()
    try:
        PLC1.connect('192.168.10.1',0,1) # IP, RACK #, SLOT #
        print("CONNECTION STATUS: \n" + "PLC1: " + str(PLC1.get_connected())) # DISPLAYS IF CONNECTION TO PLC IS VALID
    except:
        print("CONNECTION STATUS: \n" "PLC1: " + str(PLC1.get_connected()))


     

    
