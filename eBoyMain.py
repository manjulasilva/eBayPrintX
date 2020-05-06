'''
Created on 8 Mar 2020

@author: Manjula Silva
'''

import sys
from PyQt5.QtWidgets import QApplication, QWidget, QGridLayout, QLabel, QRadioButton, QPushButton, QTextEdit, QComboBox,QFrame,QSplitter
from PyQt5.QtGui import QIcon, QPixmap, QFont
from eBay_Address_Functions import pasteClipboard_To_InputAddressBox, getAddressType, getBuyersNameAndAddress_in_Simple_House_Address_Format, getBuyersNameAndAddress_in_PO_BOX_Address_Format, getBuyersNameAndAddress_in_Complex_Address_Format, getBuyersNameAndAddress_in_CnC_Format, cmdCopyNewAddress,doGoogleAddressSearch
from eBay_RTF_Functions import open_RTF_Output_File_Preview, delete_RTF_Output_File,show_message_Complex_Address_Warning, createRTF_BuyersNameAddress_DL_House_Address, createRTF_BuyersNameAddress_C5_House_Address, createRTF_BuyersNameAddress_A4_DNB_House_Address, createRTF_BuyersNameAddress_A4_Express_House_Address, createRTF_BuyersNameAddress_A4_PostagePaid_DL_House_Address,createRTF_BuyersNameAddress_ParcelPost_SMALL_House_Address
from eBay_Automated_ByWin32ComClient import do_Auto_GoogleAddressSearch, print_OutputFile_GrayScale


# Declare global variables...
gEnvelope_Type = gAddressType = gBuyer_Name = gAddress_Line1 =  gAddress_Line2 = gAddress_Line3 = gAddress_Line4 = "unassigned yet"

def cmdChangeRadioImage(pEnvelopeType):
    label_Envelope_Size_Image.setPixmap(QPixmap('images/'+ pEnvelopeType + '.jpg'))
    global gEnvelope_Type # this says we will refer to global variable here.
    gEnvelope_Type = pEnvelopeType
    if not btn_Paste_Address.isEnabled():
        btn_Paste_Address.setEnabled(True)


def cmdPreview_Output_File():
    open_RTF_Output_File_Preview()

def cmdPasteAddress():
    pasteClipboard_To_InputAddressBox(txt_Address_Box_Input,cmb_Address_Option.currentText())
    setApplication_State_CLICK_btn_Paste_Address()

def cmdCurrentTest():
    delete_RTF_Output_File()
    print('No current test now.. 2')
    

def cmdGenerate_RTF_Output_File():
    print('Generate RTF output file ............')
    global gAddressType, gBuyer_Name, gAddress_Line1, gAddress_Line2
    
    if gEnvelope_Type == 'DL':
        iReturnCode = createRTF_BuyersNameAddress_DL_House_Address(gBuyer_Name, gAddress_Line1, gAddress_Line2)
    elif gEnvelope_Type == 'C5':
        iReturnCode = createRTF_BuyersNameAddress_C5_House_Address(gBuyer_Name, gAddress_Line1, gAddress_Line2)
    elif  gEnvelope_Type == 'A4_DONOTBEND':
        iReturnCode = createRTF_BuyersNameAddress_A4_DNB_House_Address(gBuyer_Name, gAddress_Line1, gAddress_Line2)
    elif  gEnvelope_Type == 'A4_Express':
        iReturnCode = createRTF_BuyersNameAddress_A4_Express_House_Address(gBuyer_Name, gAddress_Line1, gAddress_Line2)
    elif  gEnvelope_Type == 'A4_PostagePaid_DL':
        iReturnCode = createRTF_BuyersNameAddress_A4_PostagePaid_DL_House_Address(gBuyer_Name, gAddress_Line1, gAddress_Line2)
    elif gEnvelope_Type == 'ParcelPost_SMALL':
        iReturnCode = createRTF_BuyersNameAddress_ParcelPost_SMALL_House_Address(gBuyer_Name, gAddress_Line1, gAddress_Line2)
    
    if iReturnCode == 0:
        pixmap_GenerateRTF_Done = QPixmap('images/Green_Tick.png')
        label_GenerateRTF_Done_Image.setPixmap(pixmap_GenerateRTF_Done)


def cmdGet_FormattedAddress():
    
    # Get Address Type .........................................................
    strAddressType =''
    strAddressType = getAddressType(txt_Address_Box_Input.toPlainText())
    global gAddressType
    gAddressType = strAddressType
    
    strErrorReturned = ''
    global gBuyer_Name, gAddress_Line1, gAddress_Line2
        
    # If Address Type is "DL_House_Address" call "getBuyersNameAndAddress_DL_House_Address"
    if strAddressType == 'Simple_House_Address':        
        strErrorReturned, gBuyer_Name, gAddress_Line1, gAddress_Line2 = getBuyersNameAndAddress_in_Simple_House_Address_Format(txt_Address_Box_Input.toPlainText())        
        
    elif strAddressType == 'PO_BOX':
        strErrorReturned, gBuyer_Name, gAddress_Line1, gAddress_Line2 = getBuyersNameAndAddress_in_PO_BOX_Address_Format(txt_Address_Box_Input.toPlainText())
        
    elif strAddressType == 'CnC':
        print('CnC selected 888888')
        strErrorReturned, gBuyer_Name, gAddress_Line1, gAddress_Line2 = getBuyersNameAndAddress_in_CnC_Format(txt_Address_Box_Input.toPlainText())
        
    elif strAddressType == 'Complex_Address_ShopOrHighRise':
        strErrorReturned, gBuyer_Name, gAddress_Line1, gAddress_Line2 = getBuyersNameAndAddress_in_Complex_Address_Format(txt_Address_Box_Input.toPlainText())
        show_message_Complex_Address_Warning()
        
        
    txt_Address_Box_Output.setText(strErrorReturned)
    txt_Address_Box_Output.append(gBuyer_Name)
    txt_Address_Box_Output.append(gAddress_Line1)
    txt_Address_Box_Output.append(gAddress_Line2)
    
    setApplication_State_CLICK_btn_Format_Address()


def cmdPrintGrayscale():
    #pasteClipboard_To_InputAddressBox(txt_Address_Box_Input,cmb_Address_Option.currentText())
    #setApplication_State_CLICK_btn_Paste_Address()
    print("gray scale printing involked !!")
    print_OutputFile_GrayScale(gEnvelope_Type)

def setApplication_State_INIT():    
    btn_Paste_Address.setEnabled(False)
    btn_Copy_New_Address.setEnabled(False)
    
    btn_Format_Address.setEnabled(False)
    btn_Google_Address.setEnabled(False)
    btn_Create_Env_Label.setEnabled(False)
    btn_Preview_Env_Label.setEnabled(False)
    
    txt_Address_Box_Output.setText('')
    txt_Address_Box_Input.setText('')
    
    radio_01_DL.setAutoExclusive(False)
    radio_01_C5.setAutoExclusive(False)
    radio_01_A4.setAutoExclusive(False)    
    radio_01_DL.setChecked(False)
    radio_01_C5.setChecked(False)
    radio_01_A4.setChecked(False)
    radio_01_DL.setAutoExclusive(True)
    radio_01_C5.setAutoExclusive(True)
    radio_01_A4.setAutoExclusive(True)
    
    
    label_GenerateRTF_Done_Image.clear()
    pixmap_EnvelopeSize = QPixmap('images/Default.jpg')
    label_Envelope_Size_Image.setPixmap(pixmap_EnvelopeSize)
    
    
    
    print('........  Application Ready to Begin .......')
    

def setApplication_State_CLICK_btn_Paste_Address():    
    btn_Copy_New_Address.setEnabled(True)    
    btn_Format_Address.setEnabled(True)
    


def setApplication_State_CLICK_btn_Format_Address():   
    btn_Google_Address.setEnabled(True)
    btn_Create_Env_Label.setEnabled(True)
    btn_Preview_Env_Label.setEnabled(True)


def Create_Main_Window():
    win.setLayout(gridTop)
    win.show()
    setApplication_State_INIT() # this will enable/disable controls accordingly
    app.exec_()
    
  

# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
# .......................   GUI  ................................................................................
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


app = QApplication([]) #(sys.argv)
app.setStyle('Windows') #Fusion, Windows

# Window + Grid  ......................................................
win = QWidget() #is this window ??
win.setGeometry(0,0,1200,300)
win.setWindowTitle("Print X")
win.setWindowIcon(QIcon('images/PrintX_Icon.png'))

gridTop = QGridLayout() #This is the base layout


# Widgets .....................................................
label_TopLine_Image = QLabel()
label_Envelope_Size_Image = QLabel()
label_GenerateRTF_Done_Image = QLabel()

pixmap_TopLine = QPixmap('images/TopLine.jpg')
label_TopLine_Image.setPixmap(pixmap_TopLine)

pixmap_EnvelopeSize = QPixmap('images/Default.jpg')
label_Envelope_Size_Image.setPixmap(pixmap_EnvelopeSize)

pixmap_GenerateRTF_Done = QPixmap('images/Green_Tick.png')
label_GenerateRTF_Done_Image.setPixmap(pixmap_GenerateRTF_Done)


txt_Address_Box_Input = QTextEdit()

txt_Address_Box_Output= QTextEdit()
myFont = QFont()
myFont.setPointSize(16)
txt_Address_Box_Output.setFont(myFont)
#txt_Address_Box_Output.setPointSize(20)

#txt_Address_Box_Search= QTextEdit()


btn_Copy_New_Address = QPushButton(text="Copy New Address")
btn_Copy_New_Address.clicked.connect(lambda:cmdCopyNewAddress(txt_Address_Box_Output))

btn_Preview_Env_Label = QPushButton(text="Preview Label File")
btn_Preview_Env_Label.clicked.connect(lambda:cmdPreview_Output_File())

btn_Create_Env_Label = QPushButton(text="Create Label File")
btn_Create_Env_Label.clicked.connect(lambda:cmdGenerate_RTF_Output_File())

btn_Format_Address = QPushButton(text="Format Address")
btn_Format_Address.clicked.connect(lambda:cmdGet_FormattedAddress())

btn_Google_Address = QPushButton(text="Google Address")
#btn_Google_Address.clicked.connect(lambda:doGoogleAddressSearch(gAddress_Line1 + " " + gAddress_Line2))
btn_Google_Address.clicked.connect(lambda:do_Auto_GoogleAddressSearch(gAddress_Line1 + " " + gAddress_Line2))


btn_Paste_Address = QPushButton(text="Paste Address")
btn_Paste_Address.clicked.connect(lambda:cmdPasteAddress())


btn_Reset = QPushButton(text="Reset")
btn_Reset.clicked.connect(lambda:setApplication_State_INIT())

btn_CurrentTest = QPushButton(text="Current Test - Open RTF file")
btn_CurrentTest.clicked.connect(lambda:cmdCurrentTest())

btn_Print_Grayscale = QPushButton(text="Print Grayscale")
btn_Print_Grayscale.clicked.connect(lambda:cmdPrintGrayscale())


radio_01_DL = QRadioButton("DL Letter")
radio_01_C5 = QRadioButton("C5 Letter")
radio_01_A4 = QRadioButton("A4 with    \'DO NOT BEND PLEASE\'")
radio_01_A4_Express = QRadioButton("A4 Express")
radio_01_A4_PostagePaid_DL = QRadioButton("A4 Postage Paid DL")
radio_01_ParcelPost_SMALL = QRadioButton("Parcel Post (SMALL)")


radio_01_DL.toggled.connect(lambda:cmdChangeRadioImage('DL'))
radio_01_C5.toggled.connect(lambda:cmdChangeRadioImage('C5'))
radio_01_A4.toggled.connect(lambda:cmdChangeRadioImage('A4_DONOTBEND'))
radio_01_A4_Express.toggled.connect(lambda:cmdChangeRadioImage('A4_Express'))
radio_01_A4_PostagePaid_DL.toggled.connect(lambda:cmdChangeRadioImage('A4_PostagePaid_DL'))
radio_01_ParcelPost_SMALL.toggled.connect(lambda:cmdChangeRadioImage('ParcelPost_SMALL'))


cmb_Address_Option = QComboBox()
cmb_Address_Option.addItem("Clipboard")
cmb_Address_Option.addItem("House Address")
cmb_Address_Option.addItem("Unit Address")
cmb_Address_Option.addItem("PO BOX Address")
cmb_Address_Option.addItem("CnC Address")


# Left ............................................................................
if 1== 2:
    frame_left_column = QFrame()
    frame_left_column.setFrameShape(QFrame.StyledPanel)
    splitter_left_col = QSplitter()
    splitter_left_col.addWidget(frame_left_column)
    splitter_left_col.setStretchFactor(5, 5)
    layout_left_column = QGridLayout()
    layout_left_column.addWidget(label_Envelope_Size_Image,1,1)
    layout_left_column.addWidget(radio_01_DL,2,1)
    layout_left_column.addWidget(radio_01_C5,3,1)
    layout_left_column.addWidget(radio_01_A4,4,1)
    layout_left_column.addWidget(radio_01_A4_Express,5,1)
    frame_left_column.setLayout(layout_left_column)
    gridTop.addWidget(splitter_left_col,2,1,5,1)

# Add controls in GRID layout  ................................
# gridTop.addWidget( CONTROL NAME, ROW, COLUMN, Row span, Column span)

gridTop.addWidget(radio_01_DL,3,1)
gridTop.addWidget(radio_01_C5,4,1)
gridTop.addWidget(radio_01_A4,5,1)
gridTop.addWidget(radio_01_A4_Express,6,1)
gridTop.addWidget(radio_01_A4_PostagePaid_DL,7,1)
gridTop.addWidget(radio_01_ParcelPost_SMALL,8,1)

gridTop.addWidget(label_TopLine_Image,1,1,1,6)
gridTop.addWidget(label_Envelope_Size_Image,2,1)
gridTop.addWidget(label_GenerateRTF_Done_Image,5,6)

gridTop.addWidget(txt_Address_Box_Input,2,3,1,2)
gridTop.addWidget(txt_Address_Box_Output,2,5,1,2)
#gridTop.addWidget(txt_Address_Box_Search,6,4)

gridTop.addWidget(btn_Paste_Address,3,3)
#gridTop.addWidget(btn_CurrentTest,9,3)
#gridTop.addWidget(btn_Copy_New_Address,10,3)

gridTop.addWidget(btn_Format_Address,3,5)
gridTop.addWidget(btn_Google_Address,4,5)
gridTop.addWidget(btn_Create_Env_Label,5,5)
gridTop.addWidget(btn_Preview_Env_Label,6,5)
gridTop.addWidget(btn_Print_Grayscale,6,6)
gridTop.addWidget(btn_Reset,10,5)



gridTop.addWidget(cmb_Address_Option,3,4)


# Resize controls in GRID layout  ................................
btn_Google_Address.resize(175,30) # width , height
btn_Create_Env_Label.resize(175,30) # width , height
btn_Copy_New_Address.resize(175,30) # width , height
btn_Preview_Env_Label.resize(175,30) # width , height


  

if __name__ == '__main__':
    Create_Main_Window()
    

print(" - - End - - ")