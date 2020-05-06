'''
Created on 8 Mar 2020

@author: Manjula Silva
'''
import os
import sys
import fileinput
import subprocess as sp
import webbrowser
from PyQt5.QtWidgets import QMessageBox



#functions''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

def open_RTF_Output_File_Preview():
    #v_programName = "notepad.exe"
    #v_fileName = "test1.txt"
    #v_programName = "Winword.exe"
    v_fileName = "Output_file.rtf"
    #sp.Popen([v_programName, v_fileName])
    
    webbrowser.open(v_fileName)
    print ('Output file successfully opened')


def delete_RTF_Output_File():
    v_fileName = "Output_file.rtf"
    os.remove(v_fileName)
    print('output file deleted')


def show_message_OutputFileCreated():
    return
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Information)
    msg.setText("File created successfully")
    msg.setInformativeText("Please proceed with priview and print")
    msg.setWindowTitle("File created successfully")
    #msg.setDetailedText("The details are as follows:")
    msg.setStandardButtons(QMessageBox.Ok )
    #msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
    retval = msg.exec_()
    
    
def show_message_UnableToCreateOutputFile():
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Critical)
    msg.setText("Unable to generate envelope label file !")
    msg.setInformativeText("Please close all opened files and try again")
    msg.setWindowTitle("File create failure")
    #msg.setDetailedText("The details are as follows:")
    msg.setStandardButtons(QMessageBox.Ok )
    #msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
    retval = msg.exec_()

    
def show_message_Complex_Address_Warning():    
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Information)
    msg.setText("Given addresss is a complex address")
    msg.setInformativeText("Please review and make necessary changes before you print")
    msg.setWindowTitle("Complex Address")
    msg.setStandardButtons(QMessageBox.Ok )
    #msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
    retval = msg.exec_()


def createRTF_BuyersNameAddress_DL_House_Address(pBuyers_Name, pAddress_Line1, pAddress_Line2):
    print('Start generate RTF : DL House Address ..............................')
    
    fileToSearch = 'TemplateRTFs/DL_House_Address.rtf'
    txt_Name_To_Search = "@Buyer_Name"
    txt_Add1_To_Search = "@Address_Line1"
    txt_Add2_To_Search = "@Address_Line2"
    
    
    # Read in the file
    with open(fileToSearch, 'r') as file :
        filedata = file.read()
        
    # Replace the target string
    filedata = filedata.replace(txt_Name_To_Search, pBuyers_Name)
    filedata = filedata.replace(txt_Add1_To_Search, pAddress_Line1)
    filedata = filedata.replace(txt_Add2_To_Search, pAddress_Line2)
    
    # Write the file out again
    iStatus = -1
    try:
        with open('Output_file.rtf', 'w') as file:
            file.write(filedata)
        show_message_OutputFileCreated()
        iStatus = 0
        print('Output_file.rtf  successfully created for Buyer ' + pBuyers_Name)    
    except:
        show_message_UnableToCreateOutputFile()
        
    return iStatus    
   
  
def createRTF_BuyersNameAddress_C5_House_Address(pBuyers_Name, pAddress_Line1, pAddress_Line2):
    print('Start generate RTF : DL House Address ..............................')
    
    fileToSearch = 'TemplateRTFs/C5_House_Address.rtf'
    txt_Name_To_Search = "@Buyer_Name"
    txt_Add1_To_Search = "@Address_Line1"
    txt_Add2_To_Search = "@Address_Line2"
    
    # Read in the file
    with open(fileToSearch, 'r') as file :
        filedata = file.read()
        
    # Replace the target string
    filedata = filedata.replace(txt_Name_To_Search, pBuyers_Name)
    filedata = filedata.replace(txt_Add1_To_Search, pAddress_Line1)
    filedata = filedata.replace(txt_Add2_To_Search, pAddress_Line2)
    
    # Write the file out again
    iStatus = -1
    try:
        with open('Output_file.rtf', 'w') as file:
            file.write(filedata)
        show_message_OutputFileCreated()
        iStatus = 0
        print('Output_file.rtf  successfully created for Buyer ' + pBuyers_Name)    
    except:
        show_message_UnableToCreateOutputFile()
        
    return iStatus  
  
  
def createRTF_BuyersNameAddress_A4_DNB_House_Address(pBuyers_Name, pAddress_Line1, pAddress_Line2):
    print('Start generate RTF : DL House Address ..............................')
    
    fileToSearch = 'TemplateRTFs/A4_DONOTBEND.rtf'
    txt_Name_To_Search = "@Buyer_Name"
    txt_Add1_To_Search = "@Address_Line1"
    txt_Add2_To_Search = "@Address_Line2"
    
    # Read in the file
    with open(fileToSearch, 'r') as file :
        filedata = file.read()
        
    # Replace the target string
    filedata = filedata.replace(txt_Name_To_Search, pBuyers_Name)
    filedata = filedata.replace(txt_Add1_To_Search, pAddress_Line1)
    filedata = filedata.replace(txt_Add2_To_Search, pAddress_Line2)
    
    # Write the file out again
    iStatus = -1
    try:
        with open('Output_file.rtf', 'w') as file:
            file.write(filedata)
        show_message_OutputFileCreated()
        iStatus = 0
        print('Output_file.rtf  successfully created for Buyer ' + pBuyers_Name)    
    except:
        show_message_UnableToCreateOutputFile()
        
    return iStatus  
  

def createRTF_BuyersNameAddress_A4_Express_House_Address(pBuyers_Name, pAddress_Line1, pAddress_Line2):
    print('Start generate RTF : A4 Express House Address ..............................')
    
    fileToSearch = 'TemplateRTFs/A4_Express_House_Address.rtf'
    txt_Name_To_Search = "@Buyer_Name"
    txt_Add1_To_Search = "@Address_Line1"
    txt_Add2_To_Search = "@Address_Line2"
    
    # Read in the file
    with open(fileToSearch, 'r') as file :
        filedata = file.read()
        
    # Replace the target string
    filedata = filedata.replace(txt_Name_To_Search, pBuyers_Name)
    filedata = filedata.replace(txt_Add1_To_Search, pAddress_Line1)
    filedata = filedata.replace(txt_Add2_To_Search, pAddress_Line2)
    
    # Write the file out again
    iStatus = -1
    try:
        with open('Output_file.rtf', 'w') as file:
            file.write(filedata)
        show_message_OutputFileCreated()
        iStatus = 0
        print('Output_file.rtf  successfully created for Buyer ' + pBuyers_Name)    
    except:
        show_message_UnableToCreateOutputFile()
        
    return iStatus  


def createRTF_BuyersNameAddress_A4_PostagePaid_DL_House_Address(pBuyers_Name, pAddress_Line1, pAddress_Line2):
    print('Start generate RTF : DL House Address ..............................')
    
    fileToSearch = 'TemplateRTFs/A4_PostagePaid_DL_House_Address.rtf'
    txt_Name_To_Search = "@Buyer_Name"
    txt_Add1_To_Search = "@Address_Line1"
    txt_Add2_To_Search = "@Address_Line2"
    
    
    # Read in the file
    with open(fileToSearch, 'r') as file :
        filedata = file.read()
        
    # Replace the target string
    filedata = filedata.replace(txt_Name_To_Search, pBuyers_Name)
    filedata = filedata.replace(txt_Add1_To_Search, pAddress_Line1)
    filedata = filedata.replace(txt_Add2_To_Search, pAddress_Line2)
    
    # Write the file out again
    iStatus = -1
    try:
        with open('Output_file.rtf', 'w') as file:
            file.write(filedata)
        show_message_OutputFileCreated()
        iStatus = 0
        print('Output_file.rtf  successfully created for Buyer ' + pBuyers_Name + '333')    
    except:
        show_message_UnableToCreateOutputFile()
        
    return iStatus    



def createRTF_BuyersNameAddress_ParcelPost_SMALL_House_Address(pBuyers_Name, pAddress_Line1, pAddress_Line2):
    print('Start generate RTF : DL House Address ..............................')
    
    fileToSearch = 'TemplateRTFs/Parcel_Post_SMALL.rtf'
    txt_Name_To_Search = "@Buyer_Name"
    txt_Add1_To_Search = "@Address_Line1"
    txt_Add2_To_Search = "@Address_Line2"
    
    
    # Read in the file
    with open(fileToSearch, 'r') as file :
        filedata = file.read()
        
    # Replace the target string
    filedata = filedata.replace(txt_Name_To_Search, pBuyers_Name)
    filedata = filedata.replace(txt_Add1_To_Search, pAddress_Line1)
    filedata = filedata.replace(txt_Add2_To_Search, pAddress_Line2)
    
    # Write the file out again
    iStatus = -1
    try:
        with open('Output_file.rtf', 'w') as file:
            file.write(filedata)
        show_message_OutputFileCreated()
        iStatus = 0
        print('Output_file.rtf  successfully created for Buyer ' + pBuyers_Name + '444')    
    except:
        show_message_UnableToCreateOutputFile()
        
    return iStatus    




