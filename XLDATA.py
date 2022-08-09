
from genericpath import getsize
from lib2to3.pgen2.token import NAME
from tkinter.tix import COLUMN
from tracemalloc import start
from turtle import st
from unicodedata import name
from uuid import NAMESPACE_DNS
import warnings
from openpyxl import load_workbook
import snap7 # IMPORT SNAP7 LIBRARY
import struct # TO CONVERT BYTE ARRAY TO FLOATING POINT
from sys import getsizeof # IMPORT TO GET SIZE OF DATA

# DOCUMENTATION: https://python-snap7.readthedocs.io/en/1.0/client.html#snap7.client.Client.as_download

try:
    PLC_1 = snap7.client.Client()
    PLC_1.connect('192.168.10.1', 0, 1) # IP, RACK #, SLOT #
    print("PLC1 CONNECTION STATUS: " + str(PLC_1.get_connected())) # DISPLAYS IF CONNECTION TO PLC_1 IS VALID
except:
    print("PLC1 CONNECTION STATUS: "  + str(PLC_1.get_connected()))

# READ/WRITE
excelWorkbook = load_workbook('ImpTags.xlsm')
uiSheet = excelWorkbook.active


# RETREIVE DATA FROM EXCEL

NAMES, DATATYPES, ADDRESSES = uiSheet['A2':], uiSheet['C2':], uiSheet['D2':]

nameList = []
datatypeList = []
addressList= []

def getDATA(columnInput, infoList):
    infoList = []
    for CELL in columnInput:
        for INFO in CELL:
            infoList.append(INFO.value)
    return infoList

nameList = getDATA(NAMES, nameList)
datatypeList = getDATA(DATATYPES, datatypeList)
addressList = getDATA(ADDRESSES, addressList)

for dataType, address in zip(datatypeList, addressList):
    if dataType == 'Bool':
        STARTOFFSET = address.replace('%Q','').split('.')[0]
        BITOFFSET = address.replace('%Q','').split('.')[1]
    #     BYTEARRAY = PLC_1.
    # elif dataType == 'Real':

    # elif dataType == 'Int':





# def readMemory(ADDRESS, LENGTH): # REAL DATA READING!
#     PLC1_READ = PLC_1.read_area(snap7.types.Areas.MK, 0, ADDRESS, LENGTH) 
#     PLC1_FP = struct.unpack(">f", PLC1_READ)
#     print("Start Address: " + str(ADDRESS) + ' PLC1_FP: ' + str(PLC1_FP))

# def writeMemory(start_address, length, value):
#     PLC_1.mb_write(start_address, length, bytearray(struct.pack('>f', value)))
#     print('Start Address: ' + str(start_address) + ' Value: ' + str(value))

# readMemory(ADDRESS, LENGTH)
# writeMemory(start_address, length, value)





 