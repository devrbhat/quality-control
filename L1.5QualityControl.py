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
import pdb
import csv
from openpyxl import load_workbook
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
"""
This method is validation task and roster combined. If the name of the resource there in the taskdistribution excel it validates the excel and then adds
all the task assigned on his/her name in a separate datatable. Now the task selected by the resource is checked the task list 
"""
def Task_Validation(Gname):
    rf=pd.read_excel('Resource List.xlsx',sheet_name= 'Sheet1')
    TD = pd.read_excel('Task Distribution.xlsx',sheet_name = 'Sheet1')
    New_Task= pd.DataFrame({"Tasks":[],"Names":[]})
    t=""  
    N=""                         
    #print(rf)
    for id in rf['User ID'] :
        R_Name= rf.loc[rf['User ID']== getpass.getuser().casefold() , 'Resource Name'].item()
        print(R_Name)
        for ind in TD.index:
            for nm in R_Name.split() :
                if nm in str(TD['Names'][ind]).split() :
                    print(ind)
                    t=TD['Tasks'][ind]
                    N=TD['Names'][ind]
                    print(t,N)
                    New_Task=New_Task.append([{"Tasks":t,"Names":N}])
                                    
        break       
    
    if len(New_Task)==0 :
        print("Roster Validation UnSuccessful")
        PopUpWindowTimed("Roster validation unsuccessful, kindly check your roster.")
        return ("Roster Validation UnSuccessful")
    else   :     
        print(R_Name)
        print(New_Task)
        eng_names = pd.DataFrame({"Group" : []})
        tf = pd.read_excel('Task List.xlsx',sheet_name = 'Sheet1')
        #print(tf)
        #print(tf.loc[7]) 
        for tk in New_Task['Tasks'] :
            for id in tf.index :
                if tf['Tasks'][id]==tk:
                    N=tf['Engagement'][id]
                    K=tf['Tasks'][id]
                    print(N,K)
                    eng_names = eng_names.append([{"Group" : N}])
    
    
    
        print(eng_names['Group'].str.upper()) 
        gname = Gname
        print(gname.upper())

        v= eng_names['Group'].str.contains(gname.upper()).sum()
        if v >0 :   
            return ("Validation Successful")
        else :
            return ("Validation UnSuccessful")                    
       
def PopUpWindowTimed(Message):
    sg.theme('Material1')
    layout10 = [      
            # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
            [sg.Text(Message, size=(60, 2), font=("Helvetica", 10))],      
            [sg.OK()]   
            
                  
        ]

    window10 = sg.Window('Level 1.5 Quality Control tool').Layout(layout10)
    button10, values10 = window10.Read(timeout=30000  )
    window10.Close()


def Activity_Check(Validity,condition) :
    valid = Validity
    cond = condition
    if valid == "Validation Successful" :
        PopUpWindowTimed("You have chosen the correct Engagement, kindly proceed.")
        layout5 = [
        [sg.Text('What activity are you doing ? ', size=(60, 1), font=("Helvetica", 10))],      
        [sg.InputOptionMenu(('Select','Application Monitoring','Performance checks & reporting','Idoc Reporting and Reprocessing','Mailbox Monitoring','Service breaks','Housekeeping','BOT monitoring & support','SOP based service requests (tickets)','Application failures/ incidents/ issues (tickets)','Ticket operations','Hotline support','Hypercare period support','Freeze support',))],
        [sg.Submit(tooltip='Click to submit this window'), sg.Cancel()]
            
        ]                   
        
        window5 = sg.Window('Level 1.5 Quality Control tool', default_element_size=(40, 1)).Layout(layout5)
        
        button5, values5 = window5.Read(timeout=30000 )
        window5.Close()
        #sg.Popup(button5,values5)
        print(values5[0])
        Activity = values5[0]
        if button5 is None or button5 == 'Cancel' or button5 == sg.TIMEOUT_KEY or values5[0] == 'Select' :
            PopUpWindowTimed('You have to make a proper selection. You cannot close or ignore the window')
            cond = True
               
        
        else :
            Activity = values5[0]
            Activity_Selection(Activity)
            cond =False
            
            
    elif valid == "Validation UnSuccessful":
        PopUpWindowTimed("Engagament selected is not as per your task assignment today. Please check your task list.")
        cond = True
    elif valid == "Roster Validation UnSuccessful": 
        cond = False
        Activity = "Roster Mismatch"
        
    return(cond)        
        
def Activity_Selection(activity):
    sg.theme('Light Brown')
    cond10= True
    while cond10 is True :
        layout11 = [      
            # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
            [sg.Text('Are you using your own credentials?', size=(60, 1),text_color='black', font=("Helvetica", 10))],      
            [sg.Button('Yes'),sg.Button('No')]   
            
                  
        ]

        window11 = sg.Window('Level 1.5 Quality Control tool').Layout(layout11)
        button11, values11 = window11.Read(timeout=30000  )
        window11.Close()
        if button11 is None :
            PopUpWindowTimed("You can not close the window.")
            cond10== True
        elif button11 == sg.TIMEOUT_KEY : 
            PopUpWindowTimed("Make a proper selection.")
            #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
            cond10= True
        elif button11 == 'Yes' :
            
            cond10=False
            
        elif button11== 'No':
            PopUpWindowTimed("Please do not proceed with this task. Sharing of credential is a violation of security policy. Please contact your team lead to secure your personal client id.")
            cond10 = False
    if activity == 'Application Monitoring' :
        sg.theme('Light Brown')
        cond13= True
        while cond13 is True :
            layout13 = [      
            # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
            [sg.Text(' Are there any process failures ?', size=(60, 1),text_color='black', font=("Helvetica", 10))],      
            [sg.Button('Yes'),sg.Button('No')]   
            
                
            ]
    
            window13 = sg.Window('Level 1.5 Quality Control tool').Layout(layout13)
            button13, values13 = window13.Read(timeout=30000  )
            window13.Close()
            if button13 is None :
                PopUpWindowTimed("You can not close the window")
                cond13== True
            elif button13 == sg.TIMEOUT_KEY : 
                PopUpWindowTimed("Make a proper selection")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                cond13= True
            elif button13 == 'Yes' :
                cond14= True
                while cond14 is True :
                    layout14 = [      
                    # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                    [sg.Text(' Is an Incident/Email sent  for failed process ?', size=(50, 1),text_color='black', font=("Helvetica", 10))],      
                    [sg.Button('Yes'),sg.Button('No'), sg.Button('NA')]   
                    
                        
                    ]
            
                    window14 = sg.Window('Level 1.5 Quality Control tool').Layout(layout14)
                    button14, values14 = window14.Read(timeout=30000  )
                    window14.Close()
                    if button14 is None :
                        PopUpWindowTimed("You can not close the window.")
                        cond14== True
                    elif button14 == sg.TIMEOUT_KEY : 
                        PopUpWindowTimed("Make a proper selection")
                        #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                        cond14= True
                    elif button14 == 'Yes' or button14 == 'NA' :
                        
                        cond14=False
                        
                    elif button14== 'No':
                        PopUpWindowTimed("Please check the standard operating procedure and inform the stakeholders about the failures.")
                        cond14 = False
                    cond13=False
                
            elif button13== 'No':
                cond13 = False    
        
        
                
    elif activity == 'Performance checks & reporting' : 
        sg.theme('Light Brown')
        cond16= True
        while cond16 is True :
            sg.theme('Light Brown')
            layout16 = [      
                # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                [sg.Text('How is the disk space utilisation?', size=(60, 1),text_color='black', font=("Helvetica", 10))],      
                [sg.Button('Within limits'),sg.Button('Exceeds the limit'),sg.Button('NA')]   
                
                      
            ]

            window16 = sg.Window('Level 1.5 Quality Control tool').Layout(layout16)
            button16, values16 = window16.Read(timeout=30000  )
            window16.Close()
            if button16 is None :
                PopUpWindowTimed("You can not close the window")
                cond16== True
            elif button16 == sg.TIMEOUT_KEY : 
                PopUpWindowTimed("Make a proper selection")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                cond16= True
            elif button16 == 'Within limits' or button16 == 'NA' :
                cond16=False
            elif button16== 'Exceeds the limit':
                PopUpWindowTimed("Please raise the incident/alert/email as per the SOP for this client. ")
                cond16 = False        
        
        cond17= True
        while cond17 is True :
            sg.theme('Light Brown')
            layout17 = [      
                # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                [sg.Text('Have you checked application server checks, lock entries, if lock entries found then have you deleted them?', size=(50, 2),text_color='black', font=("Helvetica", 10))],      
                [sg.Button('Yes'),sg.Button('No'),sg.Button('NA')]   
                
                      
            ]

            window17 = sg.Window('Level 1.5 Quality Control tool').Layout(layout17)
            button17, values17 = window17.Read(timeout=30000  )
            window17.Close()
            if button17 is None :
                PopUpWindowTimed("You can not close the window.")
                cond17== True
            elif button17 == sg.TIMEOUT_KEY : 
                PopUpWindowTimed("Make a proper selection")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                cond17= True
            elif button17 == 'Yes' or button17 == 'NA':
                cond17=False
            elif button17== 'No':
                PopUpWindowTimed("Please check the SOP and proceed.")
                cond17 = False

    elif activity == 'Idoc Reporting and Reprocessing' :
        sg.theme('Light Brown')
        cond18= True
        while cond18 is True :
            layout18 = [      
                # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                [sg.Text('Are there any IDOC failures ?', size=(60, 1),text_color='black', font=("Helvetica", 10))],      
                [sg.Button('Yes'),sg.Button('No')]   
                
                      
            ]

            window18 = sg.Window('Level 1.5 Quality Control tool').Layout(layout18)
            button18, values18 = window18.Read(timeout=30000  )
            window18.Close()
            if button18 is None :
                PopUpWindowTimed("You can not close the window.")
                cond18== True
            elif button18 == sg.TIMEOUT_KEY : 
                PopUpWindowTimed("Make a proper selection")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                cond18= True
            elif button18 == 'Yes' :
                cond19= True
                while cond19 is True :
                    layout19 = [      
                        # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                        [sg.Text('Have you reprocessed the failed IDOCs ?', size=(60, 1),text_color='black', font=("Helvetica", 10))],      
                        [sg.Button('Yes'),sg.Button('No'),sg.Button('NA')]   
                        
                              
                    ]

                    window19 = sg.Window('Level 1.5 Quality Control tool').Layout(layout19)
                    button19, values19 = window19.Read(timeout=30000  )
                    window19.Close()
                    if button19 is None :
                        PopUpWindowTimed("You can not close the window")
                        cond19== True
                    elif button19 == sg.TIMEOUT_KEY : 
                        PopUpWindowTimed("Make a proper selection.")
                        #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                        cond19= True
                    elif button19 == 'Yes': 
                        cond19=False
                    elif button19 == 'NA' :  
                        PopUpWindowTimed("Please inform the level 2 for next steps.")
                        cond19 = False
                    elif button19== 'No':
                        PopUpWindowTimed("Please check the SOP on how to reprocess the failed IDOCs.")
                        cond19 = False
                cond18=False
            elif button18== 'No':
                cond18 = False
                
    elif activity == 'Mailbox Monitoring' : 
        cond20= True
        while cond20 is True :
            sg.theme('Light Brown')
            layout20 = [      
                # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                [sg.Text('Are the sender and receiver for this mail are for the same and correct customer group?', size=(50, 2),text_color='black', font=("Helvetica", 10))],      
                [sg.Button('Yes'),sg.Button('No')]   
                
                      
            ]

            window20 = sg.Window('Level 1.5 Quality Control tool').Layout(layout20)
            button20, values20 = window20.Read(timeout=30000  )
            window20.Close()
            if button20 is None :
                PopUpWindowTimed("You can not close the window")
                cond20== True
            elif button20 == sg.TIMEOUT_KEY : 
                PopUpWindowTimed("Make a proper selection")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                cond20= True
            elif button20 == 'Yes' :
                cond20=False
            elif button20== 'No':
                PopUpWindowTimed("Please correct the sender and receiver as per SOP for this customer, before sending the email.")
                cond20 = False
        cond21= True
        while cond21 is True :
            sg.theme('Light Brown')
            layout21 = [      
                # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                [sg.Text('Have you received any email in past 1 hour?', size=(60, 1),text_color='black', font=("Helvetica", 10))],      
                [sg.Button('Yes'),sg.Button('No')]   
                
                      
            ]

            window21 = sg.Window('Level 1.5 Quality Control tool').Layout(layout21)
            button21, values21 = window21.Read(timeout=30000  )
            window21.Close()
            if button21 is None :
                PopUpWindowTimed("You can not close the window")
                cond21== True
            elif button21 == sg.TIMEOUT_KEY : 
                PopUpWindowTimed("Make a proper selection")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                cond21= True
            elif button21 == 'Yes':
                cond21=False
            elif button21== 'No':
                PopUpWindowTimed("Please check your network , VPN connection & Webmail link and ensure you are connected to the client environment.")
                cond21 = False  
        
    elif activity == 'Service breaks' :
        cond22= True
        while cond22 is True :
            sg.theme('Light Brown')
            layout22 = [      
                # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                [sg.Text('Are you aware of the timings when service break notification needs to be sent?', size=(60, 1),text_color='black', font=("Helvetica", 10))],      
                [sg.Button('Yes'),sg.Button('No')]   
                
                      
            ]

            window22 = sg.Window('Level 1.5 Quality Control tool').Layout(layout22)
            button22, values22 = window22.Read(timeout=30000  )
            window22.Close()
            if button22 is None :
                PopUpWindowTimed("You can not close the window")
                cond22== True
            elif button22 == sg.TIMEOUT_KEY : 
                PopUpWindowTimed("Make a proper selection")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                cond22= True
            elif button22 == 'Yes' :
                cond22=False
            elif button22== 'No':
                PopUpWindowTimed("Check the timing and set alert in Mailbox.")
                cond22 = False
        cond23= True
        while cond23 is True :
            sg.theme('Light Brown')
            layout23 = [      
                # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                [sg.Text('Have you checked the recipient list for service break notification?', size=(60, 1),text_color='black', font=("Helvetica", 10))],      
                [sg.Button('Yes'),sg.Button('No')]   
                
                      
            ]

            window23 = sg.Window('Level 1.5 Quality Control tool').Layout(layout23)
            button23, values23 = window23.Read(timeout=30000  )
            window23.Close()
            if button23 is None :
                PopUpWindowTimed("You can not close the window")
                cond23== True
            elif button23 == sg.TIMEOUT_KEY : 
                PopUpWindowTimed("Make A proper selection")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                cond23= True
            elif button23 == 'Yes' :
                cond23=False
            elif button23== 'No':
                PopUpWindowTimed("Please check with the SOP for the correct recipient list.")
                cond23 = False

    elif activity == 'Housekeeping' :
        cond24= True
        while cond24 is True :
            sg.theme('Light Brown')
            layout24 = [      
                # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                [sg.Text('Is your distribution  list updated?', size=(60, 1),text_color='black', font=("Helvetica", 10))],      
                [sg.Button('Yes'),sg.Button('No')]   
                
                      
            ]

            window24 = sg.Window('Level 1.5 Quality Control tool').Layout(layout24)
            button24, values24 = window24.Read(timeout=30000  )
            window24.Close()
            if button24 is None :
                PopUpWindowTimed("You can not close the window")
                cond24== True
            elif button24 == sg.TIMEOUT_KEY : 
                PopUpWindowTimed("Make a proper selection")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                cond24= True
            elif button24 == 'Yes' :
                cond24=False
            elif button24== 'No':
                PopUpWindowTimed("Kindly update the distribution list.")
                cond24 = False
        cond25= True
        while cond25 is True :
            sg.theme('Light Brown')
            layout25 = [      
                # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                [sg.Text('How many service break notifications have you sent today?', size=(60, 1),text_color='black', font=("Helvetica", 10))],      
                [sg.InputText()],
                [sg.Submit(tooltip='Click to submit this window')]   
                
                      
            ]

            window25 = sg.Window('Level 1.5 Quality Control tool').Layout(layout25)
            button25, values25 = window25.Read(timeout=30000  )
            window25.Close()
            if button25 is None :
                PopUpWindowTimed("You can not close the window")
                cond25== True
            elif button25 == sg.TIMEOUT_KEY : 
                PopUpWindowTimed("Make a proper selection")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                cond25= True
            elif button25 == 'Submit' :
                print(type(values25))
                if str(values25[0]).isnumeric():
                    print('Numeric')
                    cond25=False
                else :
                    PopUpWindowTimed("Enter a numeric value.")
                    print("not numeric")
                    cond25= True  
   
    elif activity == 'BOT monitoring & support' :
        sg.theme('Light Brown')
        cond27= True
        while cond27 is True :
            layout27 = [      
                # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                [sg.Text('Have BOTs failed today?', size=(60, 1), text_color='black',font=("Helvetica", 10))],      
                [sg.Button('Yes'),sg.Button('No')]   
                
                      
            ]

            window27 = sg.Window('Level 1.5 Quality Control tool').Layout(layout27)
            button27, values27 = window27.Read(timeout=30000  )
            window27.Close()
            if button27 is None :
                PopUpWindowTimed("You can not close the window")
                cond27== True
            elif button27 == sg.TIMEOUT_KEY : 
                PopUpWindowTimed("Make a proper selection")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                cond27= True
            elif button27 == 'Yes' :
                cond28 = True
                while cond28 is True :
                    layout28 = [      
                        # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                        [sg.Text('Have you executed the failed BOTS manually?', size=(60, 1),text_color='black', font=("Helvetica", 10))],      
                        [sg.Button('Yes'),sg.Button('No')]   
                        
                              
                    ]

                    window28 = sg.Window('Level 1.5 Quality Control tool').Layout(layout28)
                    button28, values28 = window28.Read(timeout=30000  )
                    window28.Close()
                    if button28 is None :
                        PopUpWindowTimed("You can not close the window.")
                        cond28== True
                    elif button28 == sg.TIMEOUT_KEY : 
                        PopUpWindowTimed("Make A proper selection")
                        #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                        cond28= True
                    elif button28 == 'Yes' :
                        
                        cond28=False
                    elif button28== 'No':
                        PopUpWindowTimed("Kindly proceed with the SOP for manual reports.")
                        cond28 = False
                    cond27=False
            elif button27== 'No':
                #PopUpWindowTimed("Keep checking the Dashboard.")
                cond27 = False
    elif activity == 'SOP based service requests (tickets)' :
        sg.theme('Light Brown')
        cond29= True
        while cond29 is True :
            layout29 = [      
                # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                [sg.Text('What type of Service request are you handling?', size=(60, 1),text_color='black', font=("Helvetica", 10))],      
                [sg.InputOptionMenu(('Select', 'Transport request', 'User authorization/ access control', 'Batch session deletion','Resend Fax','Others'))],
                [sg.Submit(tooltip='Click to submit this window')]    
                
                      
            ]

            window29 = sg.Window('Level 1.5 Quality Control tool').Layout(layout29)
            button29, values29 = window29.Read(timeout=30000  )
            window29.Close()
            if button29 is None :
                PopUpWindowTimed("You can not close the window")
                cond29== True
            elif button29 == sg.TIMEOUT_KEY : 
                PopUpWindowTimed("Make a proper selection")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                cond29= True
            elif button29 == 'Submit' :
                if values29[0] == 'Transport request':
                
                    cond40= True
                    while cond40 is True :
                        layout40 = [      
                            # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                            [sg.Text('Is the sequence of transport accurate as per the request?', size=(50, 1),text_color='black', font=("Helvetica", 10))],      
                           
                            [sg.Button('Yes'),sg.Button('No')]        
                            
                                  
                        ]

                        window40 = sg.Window('Level 1.5 Quality Control tool').Layout(layout40)
                        button40, values40 = window40.Read(timeout=30000  )
                        window40.Close()
                        #sg.Popup(button30, values30)
                        if button40 is None :
                            PopUpWindowTimed("You can not close the window")
                            cond40== True
                        elif button40 == sg.TIMEOUT_KEY : 
                            PopUpWindowTimed("Make a proper selection")
                            #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                            cond40= True
                        elif button40 == 'Yes' :
                            cond40=False
                        elif button40 == 'No' :  
                            PopUpWindowTimed('Please check the sequence before you import the transport. If already submitted & incorrect then please inform L2 support')
                            cond40=False
                    cond29=False
                elif values29[0] == 'User authorization/ access control':
                    cond30= True
                    while cond30 is True :
                        layout30 = [      
                            # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                            [sg.Text('Are you referring to the correct SOP for this type of request?', size=(50, 1),text_color='black', font=("Helvetica", 10))],      
                            
                            [sg.Button('Yes'),sg.Button('No')]        
                            
                                  
                        ]

                        window30 = sg.Window('Level 1.5 Quality Control tool').Layout(layout30)
                        button30, values30 = window30.Read(timeout=30000  )
                        window30.Close()
                        #sg.Popup(button30, values30)
                        if button30 is None :
                            PopUpWindowTimed("You can not close the window")
                            cond30== True
                        elif button30 == sg.TIMEOUT_KEY : 
                            PopUpWindowTimed("Make a proper selection")
                            #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                            cond30= True
                        elif button30 == 'Yes' :
                            cond30= False
                        elif button30 == 'No' :
                            PopUpWindowTimed('Please find the correct SOP for this request. If unclear contact L2 support')
                            cond30= False
                    cond29=False    
                elif values29[0] == 'Batch session deletion': 
                    cond41= True
                    while cond41 is True :
                        layout41 = [      
                            # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                            [sg.Text('Are you sure you are deleting the correct batches?', size=(50, 1),text_color='black', font=("Helvetica", 10))],      
                            
                            [sg.Button('Yes'),sg.Button('No')]        
                            
                                  
                        ]

                        window41 = sg.Window('Level 1.5 Quality Control tool').Layout(layout41)
                        button41, values41 = window41.Read(timeout=30000  )
                        window41.Close()
                        #sg.Popup(button41, values41)
                        if button41 is None :
                            PopUpWindowTimed("You can not close the window")
                            cond41== True
                        elif button41 == sg.TIMEOUT_KEY : 
                            PopUpWindowTimed("Make a proper selection")
                            #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                            cond41= True
                        elif button41 == 'Yes' :
                            cond41= False
                        elif button41 == 'No' :
                            PopUpWindowTimed('Please refer the SOP and the exact details as per request made.')
                            cond41= False
                    cond29=False    
                elif values29[0]== 'Resend Fax':
                    cond42= True
                    while cond42 is True :
                        layout42 = [      
                            # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                            [sg.Text('Is the date/time selected for resending the fax correct?', size=(60, 1),text_color='black', font=("Helvetica", 10))],      
                            #[sg.InputText()],
                            [sg.Button('Yes'),sg.Button('No')]        
                            
                                  
                        ]

                        window42 = sg.Window('Level 1.5 Quality Control tool').Layout(layout42)
                        button42, values42 = window42.Read(timeout=30000  )
                        window42.Close()
                        #sg.Popup(button42, values42)
                        if button42 is None :
                            PopUpWindowTimed("You can not close the window")
                            cond42== True
                        elif button42 == sg.TIMEOUT_KEY : 
                            PopUpWindowTimed("Make A proper selection")
                            #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                            cond42= True
                        elif button42 == 'Yes' :
                            cond42= False
                        elif button42 == 'No' :
                            PopUpWindowTimed('Please check the date/time stamp for the fax you are trying to resend.')
                            cond42= False
                    cond29=False        
                elif values29[0]== 'Others':
                    cond43= True
                    while cond43 is True :
                        sg.theme('Light Brown')
                        layout43 = [      
                            # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                            [sg.Text('Is there a documented SOP for this type of request?', size=(60, 1),text_color='black', font=("Helvetica", 10))],      
                            #[sg.InputText()],
                            [sg.Button('Yes'),sg.Button('No')]        
                            
                                  
                        ]

                        window43 = sg.Window('Level 1.5 Quality Control tool').Layout(layout43)
                        button43, values43 = window43.Read(timeout=30000  )
                        window43.Close()
                        #sg.Popup(button43, values43)
                        if button43 is None :
                            PopUpWindowTimed("You can not close the window")
                            cond43== True
                        elif button43 == sg.TIMEOUT_KEY : 
                            PopUpWindowTimed("Make a proper selection")
                            #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                            cond43= True
                        elif button43 == 'Yes' :
                            cond43= False
                        elif button43 == 'No' :
                            PopUpWindowTimed('Please find the correct SOP for this request. If unclear contact L2 support.')
                            cond43 = False
                    cond29 = False
    
    elif activity == 'Application failures/ incidents/ issues (tickets)' :
        cond31= True
        while cond31 is True :
            sg.theme('Light Brown')
            layout31 = [      
                # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                [sg.Text('Was there any Job failure?', size=(60, 1),text_color='black', font=("Helvetica", 10))],      
                [sg.Button('Yes'),sg.Button('No')]   
                
                      
            ]

            window31 = sg.Window('Level 1.5 Quality Control tool').Layout(layout31)
            button31, values31 = window31.Read(timeout=30000  )
            window31.Close()
            if button31 is None :
                PopUpWindowTimed("You can not close the window")
                cond31== True
            elif button31 == sg.TIMEOUT_KEY : 
                PopUpWindowTimed("Make a proper selection")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                cond31= True
            elif button31 == 'Yes' :
                PopUpWindowTimed("Please re-run the failed jobs as per SOP.")
                cond31=False
            elif button31== 'No':
                cond31 = False
        cond32= True
        while cond32 is True :
            sg.theme('Light Brown')
            layout32 = [      
                # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                [sg.Text('Have you re-run the Job succesfully ?', size=(60, 1),text_color='black', font=("Helvetica", 10))],      
                [sg.Button('Yes'),sg.Button('No'),sg.Button('NA')]   
                
                      
            ]

            window32 = sg.Window('Level 1.5 Quality Control tool').Layout(layout32)
            button32, values32 = window32.Read(timeout=30000  )
            window32.Close()
            if button32 is None :
                PopUpWindowTimed("You can not close the window")
                cond32== True
            elif button32 == sg.TIMEOUT_KEY : 
                PopUpWindowTimed("Make a proper selection")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                cond32= True
            elif button32 == 'Yes' :
                PopUpWindowTimed("Please update the job execution success status in scheduler.")
                cond32=False
            elif button32== 'No' or button32== 'NA':
                cond32 = False
    elif activity == 'Ticket operations' :
        cond33= True
        while cond33 is True :
            sg.theme('Light Brown')
            layout33 = [      
                # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                [sg.Text('Have you created any tickets today ?', size=(60, 1),text_color='black', font=("Helvetica", 10))],      
                [sg.Button('Yes'),sg.Button('No')]   
                
                      
            ]

            window33 = sg.Window('Level 1.5 Quality Control tool').Layout(layout33)
            button33, values33 = window33.Read(timeout=30000  )
            window33.Close()
            if button33 is None :
                PopUpWindowTimed("You can not close the window")
                cond33== True
            elif button33 == sg.TIMEOUT_KEY : 
                PopUpWindowTimed("Make a proper selection")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                cond33= True
            elif button33 == 'Yes' :
                PopUpWindowTimed('Please ensure correct customer ID is chosen for the ticket.')
                cond33=False
            elif button33== 'No':  
                cond33=False
        cond34= True
        while cond34 is True :
            sg.theme('Light Brown')
            layout34 = [      
                # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                [sg.Text('Have you assigned any tickets?', size=(60, 1),text_color='black', font=("Helvetica", 10))],      
                [sg.Button('Yes'),sg.Button('No')]   
                
                      
            ]

            window34 = sg.Window('Level 1.5 Quality Control tool').Layout(layout34)
            button34, values34 = window34.Read(timeout=30000  )
            window34.Close()
            if button34 is None :
                PopUpWindowTimed("You can not close the window")
                cond34== True
            elif button34 == sg.TIMEOUT_KEY : 
                PopUpWindowTimed("Make a proper selection")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                cond34= True
            elif button34 == 'Yes':
                PopUpWindowTimed('Please check if the assignment has been made to the right group/consultant for this client')
                cond34=False
            elif button34== 'No':
                cond34=False
    elif activity == 'Hotline support':
        cond35= True
        while cond35 is True :
            sg.theme('Light Brown')
            layout35 = [      
                # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                [sg.Text('Have you received any calls in the past two hours?', size=(60, 1),text_color='black', font=("Helvetica", 10))],      
                [sg.Button('Yes'),sg.Button('No')]   
                
                      
            ]

            window35 = sg.Window('Level 1.5 Quality Control tool').Layout(layout35)
            button35, values35 = window35.Read(timeout=30000  )
            window35.Close()
            if button35 is None :
                PopUpWindowTimed("You can not close the window")
                cond35== True
            elif button35 == sg.TIMEOUT_KEY : 
                PopUpWindowTimed("Make a proper selection")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                cond35= True
            elif button35 == 'Yes':
                PopUpWindowTimed('Please note the ticket number and inform the L2 support consultant on-call/in office.')
                cond35=False
            elif button35== 'No':
                cond35=False
    elif activity == 'Hypercare period support' :
        cond36= True
        while cond36 is True :
            sg.theme('Light Brown')
            layout36 = [      
                # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                [sg.Text('Are there any anomalies found during Hypercare support?', size=(60, 1),text_color='black', font=("Helvetica", 10))],      
                [sg.Button('Yes'),sg.Button('No')]   
                
                      
            ]

            window36 = sg.Window('Level 1.5 Quality Control tool').Layout(layout36)
            button36, values36 = window36.Read(timeout=30000  )
            window36.Close()
            if button36 is None :
                PopUpWindowTimed("You can not close the window")
                cond36== True
            elif button36 == sg.TIMEOUT_KEY : 
                PopUpWindowTimed("Make a proper selection")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                cond36= True
            elif button36 == 'Yes':
                PopUpWindowTimed('Please inform L2 support and raise ticket if needed as per SOP.')
                cond36=False
            elif button36== 'No':
                cond36=False
        cond37= True
        while cond37 is True :
            sg.theme('Light Brown')
            layout37 = [      
                # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                [sg.Text('Have you cleared all Hypercare emails?', size=(60, 1),text_color='black', font=("Helvetica", 10))],      
                [sg.Button('Yes'),sg.Button('No')]   
                
                      
            ]

            window37 = sg.Window('Level 1.5 Quality Control tool').Layout(layout37)
            button37, values37 = window37.Read(timeout=30000  )
            window37.Close()
            if button37 is None :
                PopUpWindowTimed("You can not close the window.")
                cond37== True
            elif button37 == sg.TIMEOUT_KEY : 
                PopUpWindowTimed("Make a proper selection")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                cond37= True
            elif button37 == 'Yes':
                cond37=False
            elif button37== 'No':
                PopUpWindowTimed('Kindly take action on priority.')
                cond37=False       
    elif activity == 'Freeze support' :
        cond38= True
        while cond38 is True :
            sg.theme('Light Brown')
            layout38 = [      
                # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                [sg.Text('Have you updated the freeze support checklist as per required frequency?', size=(60, 1),text_color='black', font=("Helvetica", 10))],      
                [sg.Button('Yes'),sg.Button('No')]   
                
                      
            ]

            window38 = sg.Window('Level 1.5 Quality Control tool').Layout(layout38)
            button38, values38 = window38.Read(timeout=30000  )
            window38.Close()
            if button38 is None :
                print("Here")
                PopUpWindowTimed("You can not close the window.")
                cond38== True
            elif button38 == sg.TIMEOUT_KEY : 
                PopUpWindowTimed("Make a proper selection")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                cond38= True
            elif button38 =='Yes':
                print("yes")
                cond38= False
            elif button38=='No':
                PopUpWindowTimed('Please update the freeze support checklist as per latest findings.')
                cond38= False
        
        cond39= True
        while cond39 is True :
            sg.theme('Light Brown')
            layout39 = [      
                # [sg.Text('Level 1.5 Quality Control tool', size=(40, 1), font=("Helvetica", 15))],      
                [sg.Text('Were there any anomalies found as per freeze support checklist?', size=(60, 1),text_color='black', font=("Helvetica", 10))],      
                [sg.Button('Yes'),sg.Button('No')]   
                
                      
            ]

            window39 = sg.Window('Level 1.5 Quality Control tool').Layout(layout39)
            button39, values39 = window39.Read(timeout=30000  )
            window39.Close()
            if button39 is None :
                PopUpWindowTimed("You can not close the window")
                cond39== True
            elif button39 == sg.TIMEOUT_KEY : 
                PopUpWindowTimed("Make a proper selection")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                cond39= True
            elif button39 == 'Yes':
                PopUpWindowTimed('Please inform the L2 support team about the issues found.')
                cond39=False
            elif button39== 'No':
                cond39=False 



#sg.ChangeLookAndFeel('TealMono')
try : 
    sg.theme('Material 1')
    df=pd.read_excel('Groups.xlsx',sheet_name = 'Sheet1')
    Rows = df.count
    for i in df['Number'] :

        K = df.loc[df['User ID']== getpass.getuser().casefold() , 'Group'].item()
            
    print(K)
    group = K
    if K== 'Group A' or K== 'group a':
        sg.theme('Material1')
        cond1= True 
        #print (group)
        while cond1 == True :
            layout1 = [
            [sg.Text('Select Engagement in which you are working.', size=(60, 1), font=("Helvetica", 10))],      
            [sg.InputOptionMenu(('Select', 'Kesko', 'Fiskars', 'Cargill','MHC','Campari','Infineon'))],
            [sg.Submit(tooltip='Click to submit this window'), sg.Cancel()]
              
            ]
            window2 = sg.Window('Level 1.5 Quality Control tool', default_element_size=(40, 1)).Layout(layout1)
        
            button2, values2 = window2.Read(timeout=30000 )
            window2.Close()  
            Eng = values2[0]
            #sg.Popup(button2, values2)
            if button2 is None :
                PopUpWindowTimed("You can not close the window")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection is made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                cond1 = True
            
            elif button2 == 'Submit':
                #sg.Popup(button2, values2)
                print(values2[0])
                if values2[0]=='Select' :
                    PopUpWindowTimed("make a proper selection",auto_close = 300)
                    notification.notify(
                    title='Hello ' + ' ' + getpass.getuser(),
                    message='Select a proper option',
                    )
                    cond1=True
                else : 
                
                    valid = Task_Validation(values2[0])
                    cond1 = Activity_Check(valid,cond1)
                       
            elif button2 == sg.TIMEOUT_KEY : 
                PopUpWindowTimed("Please make a proper selection.")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                cond1= True
            
            
            elif values2[0] == 'Cancel' :
                #PopUpWindowTimed("make a proper selection")
                notification.notify(
                title='Hello ' + ' ' + getpass.getuser(),
                message='You cannot Cancel',
                )
                cond1=True
            else : 
                notification.notify(
                title='Hello ' + ' ' + getpass.getuser(),
                message='Make sure you are using proper credntials for the task you are doing',
                )

    elif K== 'Group B' or K== 'group b' :
        sg.theme('Material1')
        cond3 = True
        while cond3 == True :
            layout3 = [
            [sg.Text('Select Engagement in which you are working.', size=(60, 1), font=("Helvetica", 10))],      
           #[sg.Button('Roukakesko'), sg.Button('Fiskars'), sg.Button('Cargill'), sg.Button('MHC'), sg.Button('Campari'), sg.Button('None')],      
            [sg.InputOptionMenu(('Select', 'SATO', 'Postnord',))],
            [sg.Submit(tooltip='Click to submit this window'), sg.Cancel()]
              
            ]
            window3 = sg.Window('Level 1.5 Quality Control tool', default_element_size=(40, 1)).Layout(layout3)
        
            button3, values3 = window3.Read(timeout=30000 )
            window3.Close()  
            Eng = values3[0]
            #sg.Popup(button3, values3)
            if button3 is None :
                PopUpWindowTimed("You cannot close the window.")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection is made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                cond3 = True
            
            elif button3 == 'Submit':
                #sg.Popup(button3, values3)
                print(values3[0])
                if values3[0]=='Select' :
                    PopUpWindowTimed("Make a proper selection")
                    notification.notify(
                    title='Hello ' + ' ' + getpass.getuser(),
                    message='Select a proper option',
                    )
                    cond3=True
                else : 
                
                    valid = Task_Validation(values3[0])
                    cond3 = Activity_Check(valid,cond3)
                    
            elif button3 == sg.TIMEOUT_KEY : 
                PopUpWindowTimed("Please Make a Proper Selection.")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                Activity == "Nothing Selected"
                cond3= True
            
            
            elif values3[0] == 'Cancel' :
                PopUpWindowTimed("Make a proper Selection")
                notification.notify(
                title='Hello ' + ' ' + getpass.getuser(),
                message='You cannot Cancel',
                )
                cond3=True
            else : 
                notification.notify(
                title='Hello ' + ' ' + getpass.getuser(),
                message='Make sure you are using proper credentials for the task you are doing..',
                )

    elif K=='Group C' or K== 'group c' :
        sg.theme('Material1')
        cond4=True
        while cond4 == True :
            layout4 = [
            [sg.Text('Select engagement in which you are working.', size=(60, 1), font=("Helvetica", 10))],      
            #[sg.Button('Roukakesko'), sg.Button('Fiskars'), sg.Button('Cargill'), sg.Button('MHC'), sg.Button('Campari'), sg.Button('None')],      
            [sg.InputOptionMenu(('Select', 'Volvo', 'Whirlpool','Valmet'))],
            [sg.Submit(tooltip='Click to submit this window'), sg.Cancel()]
              
            ]
            window4 = sg.Window('Level 1.5 Quality Control tool', default_element_size=(40, 1)).Layout(layout4)
        
            button4, values4 = window4.Read(timeout=30000 )
            window4.Close()  
            Eng = values4[0]
            #sg.Popup(button4, values4)
            if button4 is None:
                PopUpWindowTimed("You cannot close the window")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection is made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                cond4 = True
            
            elif button4 == 'Submit':
                #sg.Popup(button4, values4)
                print(values4[0])
                if values4[0]=='Select' :
                    PopUpWindowTimed("Make a Proper Selection.")
                    notification.notify(
                    title='Hello ' + ' ' + getpass.getuser(),
                    message='Select a proper option',
                    )
                    cond4=True
                else : 
                
                    valid = Task_Validation(values4[0])
                    cond4 = Activity_Check(valid,cond4)
                     
            elif button4 == sg.TIMEOUT_KEY : 
                PopUpWindowTimed("Please Make a Proper Selection.")
                #Send_Mail("Level 1.5 Quality Control tool","No Selection has been made","devraj.bhattacharya@capgemini.com","devraj.bhattacharya@capgemini.com;sanjukta.nath@capgemini.com",None)
                Activity == "Nothing Selected"
                cond4= True
            
            
            elif button4== 'Cancel' :
                #PopUpWindowTimed("make a proper selection")
                notification.notify(
                title='Hello ' + ' ' + getpass.getuser(),
                message='You cannot Cancel',
                )
                cond4=True
            else : 
                notification.notify(
                title='Hello ' + ' ' + getpass.getuser(),
                message='Make sure you are using proper credntials for the task you are doing',
                )
        
    else :
        PopUpWindowTimed(" Validation Failed")
        cond = True                

except Exception as e:
    print(e)
    PopUpWindowTimed("Unable to proceed. Kindly check your access and connections.")


