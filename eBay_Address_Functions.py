'''
Created on 8 Mar 2020

@author: Manjula Silva
'''
from PyQt5.QtWidgets import QWidget, QLabel, QTextEdit
from PyQt5.QtGui import QIcon, QPixmap
from pxData import create_Street_Abbreviations
import win32clipboard
import webbrowser, sys


# Functions .............................................................

def formatState(tState):
    strState = "STATE"
    if tState == "VIC" or tState == "NSW" or tState == "QLD" or tState == "SA" or tState == "TAS" or tState == "WA" or tState == "NT":
        strState = tState
    elif tState == "VICTORIA":
        strState = "VIC"
    elif tState == "NEW SOUTH WALES":
        strState = "NSW"
    elif tState == "QUEENSLAND":
        strState = "QLD"
    elif tState == "SOUTH AUSTRALIA":
        strState = "SA"
    elif tState == "TASMANIA":
        strState = "TAS"
    elif tState == "WESTERN AUSTRALIA":
        strState = "WA"
    elif tState == "NOTHERN TERRETORY":
        strState = "NT"
    elif tState == "ACT":
        strState = "ACT"
    else:
        strState = "UNKNOWN STATE"
    
    return strState

    
def pasteClipboard_To_InputAddressBox(tBox,vAddressOption):
    

    print(vAddressOption)
    
    if vAddressOption == 'Clipboard' :        
        win32clipboard.OpenClipboard()
        vDataFromClipboard = win32clipboard.GetClipboardData()
        win32clipboard.CloseClipboard()
        tBox.setText(vDataFromClipboard)    
    
    if vAddressOption == 'House Address' :
        tBox.setText("glenna hall")
        tBox.append("ebay:bp36lbs")
        tBox.append("10 banksia crescent")
        tBox.append("EAST DEVONPORT TAS 7310")
        
    if vAddressOption == 'Unit Address' :
        tBox.setText("aleksandra kirjanova")
        tBox.append("ebay:bp36lbs")
        tBox.append("unit 4, 122 Arthur st")
        tBox.append("surry hills NSW 2010")
        
    if vAddressOption == 'PO BOX Address' :
        tBox.setText("Jules Rayner")
        tBox.append("ebay:bp36lbs")
        tBox.append("PO BOX 5085")
        tBox.append("ALICE SPRINGS NT 0871")
            
    if vAddressOption == 'CnC Address' :
        tBox.setText("Jessica Reader")
        tBox.append("CnC Woolworths Toongabbie")
        tBox.append("eCP:PKCCDUGB 17 - 19 Aurelia St")
        tBox.append("Toongabbie NSW 2146")


def cmdCopyNewAddress(vOutPutBox):
    print(" .......... Copy New Address ..........")
    vNewAddress = vOutPutBox.toPlainText() #This is not ideal...
    print(vNewAddress)
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(vNewAddress)
    win32clipboard.CloseClipboard()

    
def doGoogleAddressSearch(vAddressToSearch):
    print('google :' + vAddressToSearch)
    url_part1 = "http://google.com/?q="
    #url_part1 = 'https://www.google.com/maps/place/'
    webbrowser.open(url_part1 + vAddressToSearch)


def getAddressType(pInputAddress):
    arrAddress = pInputAddress.splitlines()
    vAddressLine1 = arrAddress[2] # get 2nd address line this for Simple, Complex or PO Box types
    vAddressLineCnC1 = arrAddress[1] # get 2nd line just in case Click and collect type
    vAddressLineCnC2 = arrAddress[2] # get 2nd line just in case Click and collect type
    
    vAddressType = 'Unidentified'
    str1 = "p"
    str2 = "o"
    str3 = "box"
    
    vAddressLine1 = vAddressLine1.lower()
    iLocation1 = vAddressLine1.find(str1)
    iLocation2 = vAddressLine1.find(str2)
    iLocation3 = vAddressLine1.find(str3)
    
    str4 = 'shop'
    str5 = 'level'
    str6 = 'lvl'
    
    iLocation4 = vAddressLine1.find(str4)
    iLocation5 = vAddressLine1.find(str5)
    iLocation6 = vAddressLine1.find(str6)
    
    
    print(iLocation4,iLocation5,iLocation6)
    
    str7 = 'CnC'
    str8 = 'eCP'
    iLocation7 = vAddressLineCnC1.find(str7)
    iLocation8 = vAddressLineCnC2.find(str8)
    
    
    if iLocation3 != -1:
        if iLocation1 != -1 and iLocation2 != -1:
            if iLocation1 < iLocation2:
                vAddressType = 'PO_BOX'
    elif ( iLocation7 > -1 and iLocation8 > -1 ):
        vAddressType = 'CnC'
    elif ( iLocation4 > -1 or iLocation5 > -1 or iLocation6 > -1):
        vAddressType = 'Complex_Address_ShopOrHighRise'    
    else:
        vAddressType="Simple_House_Address"       
    
    print('address type identified :' + vAddressType)
    return vAddressType









def getBuyersNameAndAddress_in_Simple_House_Address_Format(pInputBoxText):    
    vError_Details = vBuyerName = vAddressLine1 = vAddressLine2 = ''
    
    #vBuyerName = 'John White'
    #vAddressLine1 = '33 High Street'
    #vAddressLine2 = 'Mount Waverley,  VIC  3155'
    
    
    
    arr_Street_Abbreviations = []
    arr_Street_Abbreviations = create_Street_Abbreviations(arr_Street_Abbreviations) # function from pxData module   
    
    arrAddress = pInputBoxText.splitlines()    
    # Format & Prepare Buyer's Name..................................................................
    
    vNameArray = arrAddress[0].split() #split the first line by spaces
    for vCount in vNameArray:
        vBuyerName = vBuyerName + " " + vCount.capitalize()
    
    vBuyerName = vBuyerName.strip()
    
    
    # Formating street address (line 2) ...........................................&&&&&&&
    
    vStreetArray = arrAddress[2].split() #split the second line by spaces
    vCapitalized_Street_Type =""
    vLast = len(vStreetArray)
    vLast = vLast -1
    vCurrent_Street_Type = vStreetArray[vLast]
    vCurrent_Street_Type = vCurrent_Street_Type.lower()
    vCapitalized_Street_Type = ""
    
    
    IsStreetFound=False
    # Searching street type in the array column 1
    iColumn1=0
    for i in arr_Street_Abbreviations:        
        if vCurrent_Street_Type == arr_Street_Abbreviations[iColumn1][0]:
            # Street TypeFound in column 1. Only Capitalization required
            #print(str(iColumn1) + " - " + arr_Street_Abbreviations[iColumn1][0])  >>>>>>>>>>>>>>>>>>>>2
            vCapitalized_Street_Type = vCurrent_Street_Type.capitalize()
            IsStreetFound = True
            break
        iColumn1=iColumn1+1
    
    
    # Searching street type in the array column 2
    iColumn1=0
    
    if not IsStreetFound:
        for i in arr_Street_Abbreviations:
            if vCurrent_Street_Type == arr_Street_Abbreviations[iColumn1][1]:
                # Street TypeFound in column 2. Re-wording + Capitalization required
                vCurrent_Street_Type = arr_Street_Abbreviations[iColumn1][0]
                vCapitalized_Street_Type = vCurrent_Street_Type.capitalize()
                IsStreetFound = True
                break
            iColumn1=iColumn1+1
        
    
  
    if not IsStreetFound:
        print(" Critical WARNING !!!!!!!   -   Missing Street Type" )
        vError_Details = "ERROR : ERROR IN STREET TYPE"

    
    vAddressLine1 =""
    y=0
    vGap1 = " " # First space
    vGap2 = " " # Second space
    while y < vLast:
        vAddressLine1 = vAddressLine1 + vGap1 + vGap2 + vStreetArray[y].capitalize()
        y=y+1
        if y == 2:
            vGap2 =''  # Delete 2nd space
    
    vAddressLine1 = vAddressLine1 + vGap1 + vCapitalized_Street_Type
    vAddressLine1 = vAddressLine1.strip()
    
    
    # Formating SUBURB, STATE & POSTCODE (line 3) .......................................
    vSuburbArray = arrAddress[3].split() #split the first line by spaces
    vSuburb_State_Postcode_CAPS =""
    vLast = len(vSuburbArray)
    vLast = vLast -1
    vPostCode = vSuburbArray[vLast]
    vLast = vLast -1
    vState = vSuburbArray[vLast]
    vState = formatState(vState)
    vSuburb =""
    
    
    y=0
    while y < vLast:
        vSuburb = vSuburb + vGap1 + vSuburbArray[y].upper()
        y=y+1
    
    vSuburb_State_Postcode_CAPS = vSuburb + ",     " + vState.upper() + "   " +vPostCode
    vSuburb_State_Postcode_CAPS = vSuburb_State_Postcode_CAPS.strip()
    vAddressLine2 = vSuburb_State_Postcode_CAPS
    
    
    return vError_Details, vBuyerName, vAddressLine1, vAddressLine2
    print('Simple_House_Address returned.........')
    
    
    
    
    
    
    
    
def getBuyersNameAndAddress_in_PO_BOX_Address_Format(pInputBoxText):
    
    vError_Details = vBuyerName = vAddressLine1 = vAddressLine2 = ''     
    arrAddress = pInputBoxText.splitlines()
    vGap1 = " " # First space
    
    # Format & Prepare Buyer's Name..................................................................
    
    vNameArray = arrAddress[0].split() #split the first line by spaces
    for vCount in vNameArray:
        vBuyerName = vBuyerName + " " + vCount.capitalize()
    
    vBuyerName = vBuyerName.strip()
    
    
    # Formating street address (line 2) ...........................................&&&&&&&
    
    vAddressLine1 = arrAddress[2]
    vAddressLine1 = vAddressLine1.lower()
    vPO_Box_Number = ''
    iBoxLocation = vAddressLine1.find('box')
    iStartingPoint = iBoxLocation + 3
    iEndingPoint = len(vAddressLine1)
    vPO_Box_Number = vAddressLine1[iStartingPoint:iEndingPoint]
    vPO_Box_Number = vPO_Box_Number.strip()
    vAddressLine1 = 'PO Box  ' + vPO_Box_Number
    
    
    # Formating SUBURB, STATE & POSTCODE (line 3) .......................................
    vSuburbArray = arrAddress[3].split() #split the first line by spaces
    vSuburb_State_Postcode_CAPS =""
    vLast = len(vSuburbArray)
    vLast = vLast -1
    vPostCode = vSuburbArray[vLast]
    vLast = vLast -1
    vState = vSuburbArray[vLast]
    vState = formatState(vState)
    vSuburb =""
    
    
    y=0
    while y < vLast:
        vSuburb = vSuburb + vGap1 + vSuburbArray[y].upper()
        y=y+1
    
    vSuburb_State_Postcode_CAPS = vSuburb + ",     " + vState.upper() + "   " +vPostCode
    vSuburb_State_Postcode_CAPS = vSuburb_State_Postcode_CAPS.strip()
    vAddressLine2 = vSuburb_State_Postcode_CAPS
    
    
    return vError_Details, vBuyerName, vAddressLine1, vAddressLine2
    print('PO_BOX_Address returned.........')















def getBuyersNameAndAddress_in_Complex_Address_Format(pInputBoxText):
    
    vError_Details = vBuyerName = vAddressLine1 = vAddressLine2 = ''     
    arrAddress = pInputBoxText.splitlines()
    vGap1 = " " # First space
    
    
    # Format & Prepare Buyer's Name..................................................................
    
    vNameArray = arrAddress[0].split() #split the first line by spaces
    for vCount in vNameArray:
        vBuyerName = vBuyerName + " " + vCount.capitalize()
    
    vBuyerName = vBuyerName.strip()
    
    
    # Formating street address (line 2) ...........................................&&&&&&&
    
    vLongAddressLine = ''
    vCurrentAddressLine = ''
    vNumOfLines_In_Address = len(arrAddress)
    vCurrent_Line = 1
    
    while vCurrent_Line < (vNumOfLines_In_Address - 1):
        vCurrentAddressLine = vCurrentAddressLine + vGap1 + arrAddress[vCurrent_Line]
        vCurrent_Line = vCurrent_Line + 1
        
    vCurrentAddressLine = vCurrentAddressLine.strip()
    vAddressArray = vCurrentAddressLine.split() #split the first line by spaces
    
    for vCount in vAddressArray:
        vLongAddressLine = vLongAddressLine + " " + vCount.capitalize()
    
    vLongAddressLine = vLongAddressLine.strip()
    print(vLongAddressLine)
    
    
    
    # Formating SUBURB, STATE & POSTCODE (line 3) .......................................
    vSuburbArray = arrAddress[vCurrent_Line].split() #split the first line by spaces
    vSuburb_State_Postcode_CAPS =""
    vLast = len(vSuburbArray)
    vLast = vLast -1
    vPostCode = vSuburbArray[vLast]
    vLast = vLast -1
    vState = vSuburbArray[vLast]
    vState = formatState(vState)
    vSuburb =""
    
    
    y=0
    while y < vLast:
        vSuburb = vSuburb + vGap1 + vSuburbArray[y].upper()
        y=y+1
    
    vSuburb_State_Postcode_CAPS = vSuburb + ",     " + vState.upper() + "   " +vPostCode
    vSuburb_State_Postcode_CAPS = vSuburb_State_Postcode_CAPS.strip()
    vAddressLine2 = vSuburb_State_Postcode_CAPS
    
    
    return vError_Details, vBuyerName, vLongAddressLine, vAddressLine2
    print('Complex_Address returned.........')
    





def getBuyersNameAndAddress_in_CnC_Format(pInputBoxText):
    
    vError_Details = vBuyerName = vAddressLine1 = vAddressLine2 = ''     
    arrAddress = pInputBoxText.splitlines()
    vGap1 = " " # First space
    
    
    # Format & Prepare Buyer's Name..................................................................
    
    vNameArray = arrAddress[0].split() #split the first line by spaces
    for vCount in vNameArray:
        vBuyerName = vBuyerName + " " + vCount.capitalize()
    
    vBuyerName = vBuyerName.strip()
    
    
    # Formating street address (line 2) ...........................................&&&&&&&
    
    vLongAddressLine = ''
    vCurrentAddressLine = ''
    vNumOfLines_In_Address = len(arrAddress)
    vCurrent_Line = 1
    
    while vCurrent_Line < (vNumOfLines_In_Address - 1):
        vCurrentAddressLine = vCurrentAddressLine + vGap1 + arrAddress[vCurrent_Line]
        vCurrent_Line = vCurrent_Line + 1
        
    vCurrentAddressLine = vCurrentAddressLine.strip()    
    vLongAddressLine = vCurrentAddressLine
    print(vLongAddressLine)

    
    
    
    # Formating SUBURB, STATE & POSTCODE (line 3) .......................................
    vSuburbArray = arrAddress[vCurrent_Line].split() #split the first line by spaces
    vSuburb_State_Postcode_CAPS =""
    vLast = len(vSuburbArray)
    vLast = vLast -1
    vPostCode = vSuburbArray[vLast]
    vLast = vLast -1
    vState = vSuburbArray[vLast]
    vState = formatState(vState)
    vSuburb =""
    
    
    y=0
    while y < vLast:
        vSuburb = vSuburb + vGap1 + vSuburbArray[y].upper()
        y=y+1
    
    vSuburb_State_Postcode_CAPS = vSuburb + ",     " + vState.upper() + "   " +vPostCode
    vSuburb_State_Postcode_CAPS = vSuburb_State_Postcode_CAPS.strip()
    vAddressLine2 = vSuburb_State_Postcode_CAPS
    
    
    return vError_Details, vBuyerName, vLongAddressLine, vAddressLine2
    print('CnC_Address returned.........') 