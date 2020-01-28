"""
***********************************************************************************************************************************
*                                                                                                                                 *   
*   Created By : -------------------  Date  ------------------ Defination -------------------- Comments ----------------- Version *
*   Devraj Bhattacharya             15/01/2020                 Main Module                                                 01     * 
*                                                                                                                                 *  
*   Modified By :                     Date                     Defination                      Comments                   Version *
*                                                                                                                                 * 
*                                                                                                                                 *         
***********************************************************************************************************************************
"""
import PySimpleGUI as sg
from plyer import notification
import getpass
import pandas as pd
import numpy as np
from datetime import datetime
from datetime import date
import win32com.client as win32
import warnings
import xlrd
import pdb
import csv
from openpyxl import load_workbook
from shutil import copyfile
import os
import shutil
import time
warnings.simplefilter(action='ignore', category=FutureWarning)
Activity= ''
TdLink = "\\ntbomfs001\L1.5 Management$"

def Send_Mail(subject,body,to,cc,path_attachment):

    import win32com.client as win32
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'devraj.bhattacharya@capgemini.com'
    mail.Subject = subject
    mail.Body = body
    mail.CC = cc
    #mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional

    # To attach a file to the email (optional):
    #attachment  = "Path to the attachment"
    #mail.Attachments.Add(attachment)
    mail.Send()

def PopUpWindowTimed(Message):
    sg.theme('Light Brown')
    layout10 = [      
            # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
            [sg.Text(Message, size=(60, 2),text_color= 'black', font=("Helvetica", 10))],      
            [sg.OK()]   
            
                  
        ]

    window10 = sg.Window('Level 1.5 Quality Control tool').Layout(layout10)
    button10, values10 = window10.Read(timeout=30000  )
    window10.Close()
    
cond = True
while cond is True :
    #sg.ChangeLookAndFeel('DarkBlue1')
    sg.theme('Material1')
    layout = [      
        # [sg.Text('EU General Data Protection Regulation (GDPR)', size=(40, 1), font=("Helvetica", 15))],      
        [sg.Text('Which group are you working for A/ B/ C?', size=(50, 2), font=("Helvetica", 10))],      
             
        [sg.InputText()],    
        [sg.Submit(tooltip='Click to submit this window'), sg.Cancel()]   
        
              
    ]

    window = sg.Window('Level 1.5 Quality Control tool', default_element_size=(40, 1)).Layout(layout)
    button, values = window.Read(timeout=30000  )
    window.Close()
    
    GROUP_Sel=values[0]   
    if button is None : 

        PopUpWindowTimed("Please enter a valid group.")
        #Send_Mail("EU General Data Protection Regulation (GDPR)","Resource has closed the window","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
        cond =True
    elif  button == 'Submit':
        group ='Group '+values[0].upper()
        print (group)
        df=pd.read_excel(os.getcwd()+'\Groups.xlsx',sheet_name = 'Sheet1')
        Rows = df.count
        for i in df['Number'] :
            K = df.loc[df['User ID']== getpass.getuser() , 'Group'].item()
                
        print(K)
        if K==group :
            PopUpWindowTimed("You have entered the correct Group. You can proceed with your task.")
            cond = False
        else :
            PopUpWindowTimed("Group entered is incorrect. Please enter the correct group.")
            cond = True                
    
    elif button== sg.TIMEOUT_KEY : 
        PopUpWindowTimed("Please enter a valid group")
        #Send_Mail("EU General Data Protection Regulation (GDPR)","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
        cond = True
        
    elif button== 'Cancel' :
        PopUpWindowTimed("You cannot Cancel. Enter a valid Group.")
        notification.notify(
        title='Hello ' + ' ' + getpass.getuser(),
        message='You have to enter a valid group.',
        )
        cond = True
        
      
    else :
        PopUpWindowTimed("Enter a valid Group.")
        #Send_Mail("EU General Data Protection Regulation (GDPR)","Resource has not made a proper selection","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
        cond = True
       


# Checking the last modified time of the Task distribution excel and replecing it with lastest file from shared folder to local folder.

if  os.path.exists('Task Distribution.xlsx'):
    mod_time=os.path.getmtime('Task Distribution.xlsx')
    modificationTime = time.strftime('%Y-%m-%d %H-%M-%S', time.localtime(mod_time))
    print("Last Modified Time : ", modificationTime)
    present_time =datetime.now().strftime('%Y-%m-%d %H-%M-%S')
    print(present_time)
    t1 =datetime.strptime(present_time,"%Y-%m-%d %H-%M-%S")
    t2=datetime.strptime(modificationTime,"%Y-%m-%d %H-%M-%S")
    time_diff = t1-t2
    print(time_diff)
    print(type(time_diff))
    
    """
    Below code snippet will check the last modified time of the Task Distribution excel in the local file and if it is more than 3 hours then it will
    fetch the latest one from shared folder. If the shared folder location is not accessible then resource has to manually copy the latest task distribution from
    shared folder and keep in the local folder.
    """
    if time_diff.seconds/3600 > 3 :
        #print("more than 9 hours")
        try : 
            src = r'\\ntbomfs001\L1.5 Management$\L1.5 GDPR\Task Distribution.xlsx'
            dst = os.getcwd()+'\Task Distribution.xlsx'
            #dst = r'C:\Users\devrbhat\Desktop\Python\GDPR\Task Distribution.xlsx'
            copyfile(src, dst)
            

        except Exception as e :
            PopUpWindowTimed("Unable to fetch the latest Task Distribution file from shared drive.Kindly copy the latest task distribution file and paste in the local file destination")
            time.sleep(5)
            #PopUpWindowTimed("Kindly copy the latest task distribution file and paste in the local file destination(C:\Users\devrbhat\Desktop\Python\Task Distribution.xlsx).")
        
else :
    PopUpWindowTimed("File does not exist. Retreiving latest Task Distribution File.") 
    try : 
        src = r'\\ntbomfs001\L1.5 Management$\L1.5 GDPR\Task Distribution.xlsx'
        dst = os.getcwd()+'\Task Distribution.xlsx'
        #dst = r'C:\Users\devrbhat\Desktop\Python\GDPR\Task Distribution.xlsx'
        copyfile(src, dst)
        

    except Exception as e :
        PopUpWindowTimed("Unable to fetch the latest Task Distribution file from shared drive.Kindly copy the latest task distribution file and paste in the local file destination")
        time.sleep(5)
        #PopUpWindowTimed("Kindly copy the latest task distribution file and paste in the local file destination(C:\Users\devrbhat\Desktop\Python\Task Distribution.xlsx)."    