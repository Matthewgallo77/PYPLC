
import openpyxl
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import snap7 # IMPORT SNAP7 LIBRARY
import re
import xlwings as xw
from snap7.util import *
from sys import getsizeof # IMPORT TO GET SIZE OF DATA
from snap7.common import check_error, ipv4, load_library
from snap7.types import S7WLBit, S7WLByte, S7WLInt, S7WLReal
from win32com.client import Dispatch # pip install pywin32

# DOCUMENTATION: https://python-snap7.readthedocs.io/en/1.0/client.html#snap7.client.Client.as_download 


def getTagInfo():
    
    lastRow = str(len(list(uiSheet.rows))) # GETS THE LAST NON EMPTY ROW
    NAMES, DATATYPES, ADDRESSES = uiSheet['A2':'A'+lastRow], uiSheet['C2':'C'+lastRow], uiSheet['D2':'D'+lastRow] # ACCESS NECESSARY COLUMNS
    nameList, datatypeList, addressList = [], [], [] # INITIALIZE LISTS

    nameList, datatypeList, addressList  = dataList(NAMES, nameList), dataList(DATATYPES, datatypeList), dataList(ADDRESSES, addressList)

    datatypeList = [snap7.types.S7WLBit if dataType =='Bool' else snap7.types.S7WLWord if dataType == 'Word' else snap7.types.S7WLInt if dataType == 'Int' else snap7.types.S7WLReal for dataType in datatypeList] # 1 is Bool, 5 is Int, 8 is Real

    # for value in addressList:
        # addressList=list(map(lambda x: x.replace(value,re.sub("[^\d\.]", "", value)), addressList))
    address_dataType = {addressList[i]: datatypeList[i] for i in range(len(addressList))} # CONTAINS {ADDRESS: DATATYPE}
    # print(address_dataType)
    return address_dataType

def dataList(columnInput, infoList):
        infoList = []
        for CELL in columnInput:
            for INFO in CELL:
                infoList.append(INFO.value)
        return infoList

def ReadArea(PLC, byte, bit, datatype, area): # M, MB, MW, MD
    
    # print(byteArray)
    if datatype == snap7.types.S7WLBit:
        return get_bool(byteArray, byte, bit)
    elif datatype == snap7.types.S7WLInt or datatype == snap7.types.S7WLWord:
        return get_int(byteArray, byte)
    elif datatype == snap7.types.S7WLReal:
        return get_real(byteArray, byte)
    else:
        return 'ERROR: Data type has not been anticipated'


def ReadTags(PLC): # READS ALL TAGS AND RETRIEVES THERE VALUE
    valueList = []
    tagValue = None
    addressList = []
    for key, value in address_dataType.items():
        
        if 'I' in key:
            key = re.sub("[^\d\.]", "", key).split('.')
            byteArray = PLC.ab_read(area, byte, bit, datatype)
            if len(key) ==1:
                key.append(0)
            print(key[0],key[1])
            tagValue = ReadArea(PLC1, int(key[0]), int(key[1]), value, snap7.types.Areas.PE)
            print(tagValue)
            valueList.append(tagValue)
        elif 'Q' in key:
            byteArray = PLC.eb_read(area, byte, bit, datatype)
            key = re.sub("[^\d\.]", "", key).split('.')
            if len(key) ==1:
                key.append(0)
            tagValue = ReadArea(PLC1, int(key[0]), int(key[1]), value, snap7.types.Areas.PA)
            valueList.append(tagValue)
        elif 'M' in key:
            byteArray = PLC.mb_read(area, byte, bit, datatype)
            key = re.sub("[^\d\.]", "", key).split('.')
            if len(key) ==1:
                key.append(0)
            tagValue = ReadArea(PLC1, int(key[0]), int(key[1]), value, snap7.types.Areas.MK)
            valueList.append(tagValue)
                
    return valueList


# # def WriteMerker(PLC, byte, bit, datatype):
# #     # GET VALUE FROM INPUT IN EXCEL VALUE CELL
# #     # HAVE SCRIPT RUN WHEN VALUE CELL CHANGES... here is article:
# #     # https://docs.microsoft.com/en-us/office/troubleshoot/excel/run-macro-cells-change
# #     byteArray = PLC.read_area(snap7.types.Areas.MK, 0, byte, datatype)
#     if datatype == snap7.types.S7WLBit:
#         return get_bool(byteArray, byte, bit, value)
#     elif datatype == snap7.types.S7WLInt:
#         return get_int(byteArray, byte, value)
#     elif datatype == snap7.types.S7WLWord:
#         return get_real(byteArray, byte, value)
#     else:
#         return 'ERROR: Data type has not been anticipated'
# # 

def write_toExcel(PLC):
    
    wb = xw.Book('PLCTags.xlsx')
    sheet = xw.sheets[0]
    rowCount=2
    lastRow = len(list(uiSheet.rows))

    while PLC.get_connected():
        rowCount = 2 # RESET
        while rowCount<=lastRow: 
            valueList = ReadTags(PLC1)
            for value in valueList:
                sheet.range('K'+str(rowCount)).value = value
                rowCount+=1
    
if __name__ == "__main__":

    excelWorkbook = load_workbook('PilotTags.xlsx')
    uiSheet = excelWorkbook.active
    address_dataType = getTagInfo()
    
    PLC1 = snap7.client.Client()
    try:
        PLC1.connect('192.168.10.1',0,1) # IP, RACK #, SLOT #
        print("CONNECTION STATUS: \n" + "PLC1: " + str(PLC1.get_connected())) # DISPLAYS IF CONNECTION TO PLC IS VALID
    except:
        print("CONNECTION STATUS: \n" "PLC1: " + str(PLC1.get_connected()))
    values = ReadTags(PLC1)
