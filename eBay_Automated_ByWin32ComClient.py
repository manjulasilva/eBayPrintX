'''
Created on 8 Mar 2020

@author: Manjula Silva
'''

import win32com.client
from win32com.client import Dispatch
import time
import clipboard
#from win32com.client import Dispatch


def do_Auto_GoogleAddressSearch(vAddressToSearch):
        print('Auto Google :' + vAddressToSearch)
        v_shell = win32com.client.Dispatch("WScript.Shell")
        #v_shell.Run("calc.exe")
        v_shell.Run("iexplore.exe")
        time.sleep(5)
        #v_shell.SendKeys(vAddressToSearch,0)
        clipboard.copy(vAddressToSearch)
        v_shell.SendKeys("^v",0)
        time.sleep(2)
        v_shell.SendKeys("{ENTER}",0)


def print_OutputFile_GrayScale(vEnvelope_Type):
    
    if vEnvelope_Type == 'DL':
        print("DL received")
        print_OutputFile_GrayScale_DL()
    elif vEnvelope_Type == 'C5':
        print("C5 received")
        print_OutputFile_GrayScale_C5()
    elif  vEnvelope_Type == 'A4_DONOTBEND':
        print("A4 donotbend received")
        print_OutputFile_GrayScale_C5() #A4_DONOTBEND also works like C5
    elif  vEnvelope_Type == 'A4_Express':
        print("A4 express received")
        print_OutputFile_GrayScale_C5() #A4_Express also works like C5
    elif vEnvelope_Type == 'A4_PostagePaid_DL':
        print('A4_PostagePaid_DL received')
        print_OutputFile_GrayScale_C5()  #'A4_PostagePaid_DL' also works like C5
    elif vEnvelope_Type == 'ParcelPost_SMALL':
        print('ParcelPost_SMALL received')
        print_OutputFile_GrayScale_C5()  #'A4_PostagePaid_DL' also works like C5
    
    print(" .......... Print GRAY SCALE DL/C5/ branched out..........")


def print_OutputFile_GrayScale_DL( ):    
    v_shell = win32com.client.Dispatch("WScript.Shell")
    #v_shell.Run("calc.exe")
    v_shell.Run("Output_file.rtf")
    time.sleep(5)
    v_shell.SendKeys("%f",0)  #  + for shift, ^ for CTRL, % for ALT,  ~ for ENTER
    time.sleep(2)
    v_shell.SendKeys("p",0)
    time.sleep(2)
    v_shell.SendKeys("r",0)
    time.sleep(1)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys(" ",0)   # Gray scale check box selected
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys(" ",0)  # press OK
    time.sleep(1)
    v_shell.SendKeys(" ",0)  # confirm change
    time.sleep(1)
    v_shell.SendKeys("%",0)
    v_shell.SendKeys("p",0)
    v_shell.SendKeys("p",0)
    print(" .......... Print GRAY SCALE (DL) completed!..........")


def print_OutputFile_GrayScale_C5( ):    
    v_shell = win32com.client.Dispatch("WScript.Shell")
    #v_shell.Run("calc.exe")
    v_shell.Run("Output_file.rtf")
    time.sleep(5)
    v_shell.SendKeys("%f",0)  #  + for shift, ^ for CTRL, % for ALT,  ~ for ENTER
    time.sleep(1)
    v_shell.SendKeys("p",0)
    time.sleep(2)
    v_shell.SendKeys("r",0)
    time.sleep(1)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys(" ",0)   # Gray scale check box selected
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys(" ",0)  # press OK
    time.sleep(1)
    #v_shell.SendKeys(" ",0)  # confirm change
    #time.sleep(1)
    v_shell.SendKeys("%",0)
    v_shell.SendKeys("p",0)
    v_shell.SendKeys("p",0)
    print(" .......... Print GRAY SCALE (DL) completed!..........")


def print_OutputFile_GrayScale_ORIGINAL( ):    
    v_shell = win32com.client.Dispatch("WScript.Shell")
    #v_shell.Run("calc.exe")
    v_shell.Run("Output_file.rtf")
    time.sleep(2)
    v_shell.SendKeys("%f",0)  #  + for shift, ^ for CTRL, % for ALT,  ~ for ENTER
    time.sleep(1)
    v_shell.SendKeys("p",0)
    time.sleep(2)
    v_shell.SendKeys("r",0)
    time.sleep(1)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys(" ",0)   # Gray scale check box selected
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys("{TAB}",0)
    v_shell.SendKeys(" ",0)  # press OK
    time.sleep(1)
    #v_shell.SendKeys(" ",0)  # confirm change
    #time.sleep(1)
    v_shell.SendKeys("%",0)
    v_shell.SendKeys("p",0)
    v_shell.SendKeys("p",0)
    
    
    
    print(" .......... Print GRAY SCALE..........")
    

    
