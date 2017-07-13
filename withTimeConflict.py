from Tkinter import *
import unittest
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pyautogui
import time
import sys
import win32gui, win32con

import tkMessageBox



class MainWindow():

    def __init__(self,master):
        self.master = master

        self.master.title('AMDOCS APP MANAGER')
        
        
        self.f1 = ""
        self.f2 = 0

        self.v = StringVar()
        self.v1 = StringVar()
        self.name = StringVar()
        self.startTime = StringVar()
        self.endTime = StringVar()
        self.startDate = StringVar()
     
        self.subject = StringVar()
       
        self.floor = StringVar()
        self.tower = StringVar()


        self.count = 0
    

        self.L1 = Label(self.master, text = 'BOT-->')
        self.L1.grid(row=1 , column=0)

        self.E1 = Entry(self.master, textvariable = self.v,font = "Helveticab 18", bd = 7,width=120,bg='green',fg='black')
       
        self.E1.grid(row=1 , column=1)
        self.E1.insert(0,"Hi, How can I help you?")


        self.L2 = Label(self.master, text = 'You--->')
        self.L2.grid(row=2 , column=0)

        self.E2 = Entry(self.master, textvariable = self.v1,font = "Helveticab 18", bd = 7,width=120,bg='yellow')
        
        self.E3 = Entry(self.master, textvariable = self.name,font = "Helveticab 18", bd = 7,width=120,bg='yellow')
       
        self.E4 = Entry(self.master, textvariable = self.startTime,font = "Helveticab 18", bd = 7,width=120,bg='yellow')
       
        self.E5 = Entry(self.master, textvariable = self.endTime, font = "Helveticab 18",bd = 7,width=120,bg='yellow')
       
        self.E6 = Entry(self.master, textvariable = self.startDate,font = "Helveticab 18", bd = 7,width=120,bg='yellow')
        self.E7 = Entry(self.master, textvariable = self.subject,font = "Helveticab 18", bd = 7,width=120,bg='yellow')
        self.E8 = Entry(self.master,font = "Helveticab 18", textvariable = self.tower, bd = 7,width=120,bg='yellow')
        self.E9 = Entry(self.master,font = "Helveticab 18", textvariable = self.floor, bd = 7,width=120,bg='yellow')

       
        self.E2.grid(row=2 , column=1)
  
  

        self.B1 = Button(self.master, text = 'Submit', command = self.userinput)
        self.B1.grid(row=3 , column=1)

        self.B2 = Button(self.master, text = 'Submit', command = self.takeName)
        self.B3 = Button(self.master, text = 'Submit', command = self.takeStartTime)
        self.B4 = Button(self.master, text = 'Submit', command = self.takeEndTime)
        self.B5 = Button(self.master, text = 'Submit', command = self.takeDate)
        self.B6 = Button(self.master, text = 'Submit', command = self.takeSub)
        self.B8 = Button(self.master, text = 'Submit', command = self.takeFloor)
        self.B7 = Button(self.master, text = 'Submit', command = self.takeTower)




    def callOutlook(self):

        global count
        global f1
        global f2

        pyautogui.click(1280,1200)
        pyautogui.keyDown('win')
        time.sleep(1)
        pyautogui.keyUp('win')
        pyautogui.typewrite('Outlook')
        time.sleep(1)
        pyautogui.press('enter')
        #pyautogui.click(142,504)
        time.sleep(4)
        window = win32gui.GetForegroundWindow()
        win32gui.ShowWindow(window, win32con.SW_MAXIMIZE)

        pyautogui.click(67,100)
        time.sleep(0.5)
        pyautogui.click(82,181)
        time.sleep(0.5)
        window = win32gui.GetForegroundWindow()
        win32gui.ShowWindow(window, win32con.SW_MAXIMIZE)

        y = name1.strip().split(",")
        pyautogui.click(375,207)

        for a in y:

            pyautogui.typewrite(a)
            pyautogui.keyDown('ctrl')
            pyautogui.keyDown('k')
            pyautogui.keyUp('ctrl')
            pyautogui.keyUp('k')
       
        pyautogui.click(392,238)
        time.sleep(1)
        

        pyautogui.typewrite(sub)

       

        pyautogui.click(236,292)
        for j in range(15):
            pyautogui.keyDown('backspace')
        pyautogui.keyUp('backspace')

        pyautogui.typewrite(date)

  
        time.sleep(1)




        #floor decision
        f1 = floor_[0]
        f2 = int(floor_[1])






        #new changes
        pyautogui.click(1155,395)
        pyautogui.click(1254,455)
        pyautogui.click(1254,455)
        pyautogui.click(1254,455)
        pyautogui.click(1254,455)
        pyautogui.click(1254,455)
        pyautogui.click(1254,455)
        pyautogui.click(1254,455)

        pyautogui.click(1188,414)

        if tower_=="T2":
            pyautogui.keyDown('down')
            pyautogui.keyUp('down')


            if f1 == 'C':
                pyautogui.keyDown('down')
                pyautogui.keyUp('down')


            if f1 == 'N':
                pyautogui.keyDown('down')
                pyautogui.keyUp('down')
                pyautogui.keyDown('down')
                pyautogui.keyUp('down')


            if f1 == 'S':
                pyautogui.keyDown('down')
                pyautogui.keyUp('down')
                pyautogui.keyDown('down')
                pyautogui.keyUp('down')
                pyautogui.keyDown('down')
                pyautogui.keyUp('down')

        
        if tower_=="T6":
            pyautogui.keyDown('down')
            pyautogui.keyUp('down')
            pyautogui.keyDown('down')
            pyautogui.keyUp('down')
            pyautogui.keyDown('down')
            pyautogui.keyUp('down')
            pyautogui.keyDown('down')
            pyautogui.keyUp('down')
            pyautogui.keyDown('down')
            pyautogui.keyUp('down')

        if tower_=="T7":
            pyautogui.keyDown('down')
            pyautogui.keyUp('down')
            pyautogui.keyDown('down')
            pyautogui.keyUp('down')
            pyautogui.keyDown('down')
            pyautogui.keyUp('down')
            pyautogui.keyDown('down')
            pyautogui.keyUp('down')
            pyautogui.keyDown('down')
            pyautogui.keyUp('down') 
            pyautogui.keyDown('down')
            pyautogui.keyUp('down')


        time.sleep(3)
        pyautogui.click(1163,635)
        pyautogui.click(1163,635)
        time.sleep(1)
        pyautogui.click(1254,617)
        pyautogui.click(1254,617)
        pyautogui.click(1254,617)
        pyautogui.click(1254,617)
        pyautogui.click(1254,617)
        pyautogui.click(1254,617)
        pyautogui.click(1254,617)
        
        time.sleep(3)
        if stime == "08:30 AM":
            count = 1

        if stime == "09:00 AM":
            count = 2
        if stime == "09:30 AM":
            count = 3
        if stime == "10:00 AM":
            count = 4
        if stime == "10:30 AM":
            count = 5
        if stime == "11:00 AM":
            count = 6
        if stime == "11:30 AM":
            count = 7
        if stime == "12:00 PM":
            count = 8
        if stime == "12:30 PM":
            count = 9
        if stime == "01:00 PM":
            count = 10
        if stime == "01:30 PM":
            count = 11
        if stime == "02:00 PM":
            count = 12
        if stime == "02:30 PM":
            count = 13
        if stime == "03:00 PM":
            count = 14
        if stime == "03:30 PM":
            count = 15
        if stime == "04:00 PM":
            count = 16
        if stime == "04:30 PM":
            count = 17
        


        print count


        for x in range(count):
            pyautogui.keyDown('down')
            pyautogui.keyUp('down')
        time.sleep(5)





        pyautogui.click(398,309)
        for j in range(15):
            pyautogui.keyDown('backspace')
        pyautogui.keyUp('backspace')
        pyautogui.typewrite(etime)


        pyautogui.click(1152,454)
        time.sleep(2)
        pyautogui.click(1280,1200)
        tkMessageBox.showinfo("Window", "Click on 'Send' to arrange a meeting")
       
        time.sleep(1)




    ''' pyautogui.click(1031,252)
        pyautogui.typewrite('^Ind_Magarpatta_T6_AU_ Colosseum_Big_Meeting_Room (VC)')
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.click(1124,625)
        time.sleep(3)
        im=ImageGrab.grab(bbox=(970,610,1087,645)) 
        im.save('no2.JPG')
        print "taken pic"
        pytesseract.pytesseract.tesseract_cmd = 'C:\Program Files (x86)\Tesseract-OCR\\tesseract.exe'
        
        im.show()
        text = pytesseract.image_to_string(Image.open('no2.JPG'),lang='eng')
        print "printing text"
        print(text)
        t = list(text)
        for a in range(len(t)):
            if t[a] == "N":
                if t[a+1] =="o":
                    pyautogui.click(11,233)

        
                    print "meeting can be arranged"
                
            
        print "out of loop"'''
        

    

    def takeName(self):
        global name1
        name1 = self.name.get()
        print name1
        self.E1.delete(0,END)
        self.E3.delete(0,END)
        self.E1.insert(0,"enter the start time: format - HH:MM AM/PM")
        self.E3.grid_forget()
        self.B2.grid_forget()
       
        self.E4.grid(row=2 , column=1)
        self.B3.grid(row=3 , column=1)
        

    def takeStartTime(self):
        global stime 
        stime = self.startTime.get()
        print stime ,"in stime"
        self.E1.delete(0,END)
        self.E4.delete(0,END)
        self.E1.insert(0,"enter the end time format - HH:MM AM/PM")
        self.E4.grid_forget()
        self.B3.grid_forget()
        
        self.E5.grid(row=2 , column=1)
        self.B4.grid(row=3 , column=1)
      
        

    def takeEndTime(self):
        global etime 
        etime = self.endTime.get()
        print etime
        self.E1.delete(0,END)
        self.E5.delete(0,END)
        self.E1.insert(0,"enter the Date format- MM/DD/YYYY")
        self.E5.grid_forget()
        self.B4.grid_forget()
        self.E6.grid(row=2 , column=1)
        self.B5.grid(row=3 , column=1)


    def takeDate(self):
        global date 
        date = self.startDate.get()
        print name1
        self.E1.delete(0,END)
        self.E6.delete(0,END)
        self.E1.insert(0,"enter the Subject")
        self.E6.grid_forget()
        self.B5.grid_forget()
        self.E7.grid(row=2,column=1)
        self.B6.grid(row=3 , column=1)

    def takeSub(self):
        global sub 
        sub = self.subject.get()
        print sub
        self.E1.delete(0,END)
        self.E7.delete(0,END)
        self.E1.insert(0,"Insert Tower Eg.- T2/T6/T7")
        self.B6.grid_forget()
        #self.callOutlook()
        self.E8.grid(row=2,column=1)
        self.B7.grid(row=3 , column=1)

    def takeTower(self):
        global tower_ 
        tower_ = self.tower.get()
        print tower_
        self.E1.delete(0,END)
        self.E8.delete(0,END)
        self.E1.insert(0,"enter the Floor Eg. N2/S3/C4 etc")
        self.E8.grid_forget()
        self.B7.grid_forget()
        self.E9.grid(row=2,column=1)
        self.B8.grid(row=3 , column=1)
        



    def takeFloor(self):
        global floor_ 
        floor_ = self.floor.get()
        print floor_
        self.E1.delete(0,END)
        self.E9.delete(0,END)
        self.E1.insert(0,"Arranging a meeting")
        self.E9.grid_forget()
        self.B5.grid_forget()
        self.callOutlook()



    def arrangeMeeting(self):
        print "in the meet"
        self.E1.delete(0,END)
        self.E2.delete(0,END)
        self.E1.insert(0,"Enter the name of Participants")
        self.E2.grid_forget()
        self.B1.grid_forget()
        #self.B2.grid(row=2 , column=2)
        self.E3.grid(row=2 , column=1)
        self.B2.grid(row=3 , column=1)
        
    def installChrome(self):

        driver= webdriver.Chrome("C:/chromedriver.exe")
            
        driver.get("https://www.google.com/chrome/browser/thankyou.html?statcb=1&installdataindex=defaultbrowser")

        window = win32gui.GetForegroundWindow()
        win32gui.ShowWindow(window, win32con.SW_MAXIMIZE)
                    

        time.sleep(1)
        
        pyautogui.click([327,958])
        
        driver.implicitly_wait(3)
        pyautogui.click([103,954])
        time.sleep(1)
        pyautogui.click([693,517])
        
   

    def userinput(self):
        global external_variable

        external_variable = self.v1.get()
        if external_variable in { "install a software" ,"Install Software" , "Install software" , "install Software","install software",\
        "I want to install a software","i wanna install a software","I want to install a software","want to install a software",\
        "wanna install a software","I want to install a software."}:
            self.E1.delete(0,END)
            self.E2.delete(0,END)
            self.E1.insert(0,"Which Software do you want me to install?")

        

        elif external_variable in {"having a trouble","I have a problem","I'm having an issue","have problem","can u solve my problem?"}:
            print "hell yeah"

        elif external_variable in {"arrange a meeting","arrange a meeting for me","can u arrange a meeting for me",\
        "i wanna arrange a meeting","I want you to arrange a meeting for me","arrange meeting","Arrange Meeting"}:
            self.arrangeMeeting()
            
            
        elif external_variable== "google chrome" or "chrome" or "Chrome" or "install google chrome":
            self.installChrome()
            

        else:
            self.E1.delete(0,END)
            self.E1.insert(0,"Sorry, I could not understand that!")
           

        
   

   
#----------------------------------------------------------------------
if __name__ == '__main__':
    master = Tk()

    MainWindow(master)
    master.wm_geometry("%dx%d+%d+%d" % (1000, 200, 40, 70))
    master.mainloop()

