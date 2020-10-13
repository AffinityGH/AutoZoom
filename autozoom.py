#############################################################
#\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\#
#------------------- Made by: AffinityGH -------------------#
#///////////////////////////////////////////////////////////#
#############################################################

from tkinter import *
from tkinter.ttk import *
import  datetime, time, subprocess, csv, os, webbrowser, ctypes, pyautogui
from threading import Thread
myappid = 'affinity.autozoom.tkinter.beta' # arbitrary string
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

class TKGUI:
    #--------------------------------------------------------------------------#
    #function related code
    def __init__(self):
        self.root= Tk() ##initilise the main window
        self.quit = False ##defaults the quit flag
        
        #GLOBAL VARIABLES#
        self.path = ''
        self.data=[]
        self.timeTable=[]

        #GLOBAL UI ELEMENTS#

        #style
        style = Style()
        style.configure("BW.TLabel", foreground="white", background="#2a2b2e")
        frame = Style()
        frame.configure('TFrame', background='#2a2b2e')
        
        self.separator1 = Frame(height=20,style="TFrame")
        self.separator2 = Frame(height=20,style="TFrame")
        self.separator3 = Frame(height=20,style="TFrame")
        self.separator4 = Frame(height=20,style="TFrame")
        

        self.intro1 = Label(self.root, text="Welcome to AutoZoom", style="BW.TLabel")
        self.intro2 = Label(self.root, text="Begin by choosing if you have an existing schedule, or if you need to make a new one.", style="BW.TLabel")


        self.separator1.pack(fill=X, padx=5, pady=5)
        self.intro1.pack()
        self.intro2.pack()
        self.separator2.pack(fill=X, padx=5, pady=5)

        #existing schedule
        self.exist = Button(self.root, text="Existing Schedule", command=self.onPressExisting)
        self.exist.pack()

        self.separator3.pack(fill=X, padx=5, pady=5)

        #existing
        self.search = Entry(self.root, textvariable=StringVar())
        self.submit = Button(self.root, text="Search", command=lambda: self.find(self.search.get()))
        
        #new schedule
        self.new = Button(self.root, text="New Schedule", command=self.onPressNew)
        self.new.pack()
        
        #new
        self.name = Entry(self.root, textvariable=StringVar())
        self.submitNew = Button(self.root, text="Create", command=lambda: self.save(self.name.get())) 

        self.separator4.pack()

        #new --> new file
        self.instruction1 = Label(self.root, text="Insert the link to your classes here.", style="BW.TLabel")
        self.entry1 = Entry(self.root, textvariable=StringVar())
        
        self.instruction2 = Label(self.root, text="Insert your password here. (Default is None)", style="BW.TLabel")
        self.entry2Var = StringVar()
        self.entry2Var.set('None')
        self.entry2 = Entry(self.root, textvariable=self.entry2Var)

        self.instruction3 = Label(self.root, text="Insert the time of your class here in the 24HR format: HH:MM", style="BW.TLabel")
        self.entry3 = Entry(self.root, textvariable=StringVar())
        
        self.submitToDoc = Button(self.root, text="Submit", command=self.editText)
        self.finishedNew = Button(self.root, text="Finished", command=self.newScheduleStart)

        #file found
        self.found = Label(self.root, text="File found!", style="BW.TLabel")
        self.notFound = Label(self.root, text="File not found...", style="BW.TLabel")

        #made by
        self.credits = Label(self.root, text="Made by AffinityGH", style="BW.TLabel")
        self.credits.pack()

        ###########PROTOCOL GUI############
        self.timerShow = False
        self.textVar = StringVar()
        self.timerLabel = Label(self.root, textvariable=self.textVar, style="BW.TLabel")

    def update_data(self):
        if self.data != []:
            self.waiting = True
            for d in self.data:
                Label(self.root, text=d, style="BW.TLabel").pack()
            self.data = []
            self.waiting = False

    def update_clock(self):
        if self.timerShow:
            self.timerLabel.pack()
            if self.timeTable != []:
                self.waitingClock = True
                for t in self.timeTable:
                    self.textVar.set(t)
                self.timeTable = []
                self.waitingClock = False
        else:
            self.timerLabel.pack_forget()

    #opens link in default browser
    def zoomLink(self, link):
        webbrowser.open(link)

    #gets the courses from the specified schedule name
    def get_courses(self, scheduleName):
        wrapper = []
        fp = open(scheduleName, 'r')
        try:
            courses = fp.readlines()
            for course in courses:
                list = course.split("-")
                zoom_link = list[0]
                course_time = list[2].strip().split(":")
                password = list[1]
                wrapper.append([zoom_link, int(course_time[0]), int(course_time[1]), False, password])
            fp.close()
        finally:
            fp.close()
        return wrapper

    #creates new text file
    def save(self, file_name):
        with open(file_name + '.txt', 'w') as file_object:
            file_object.close()
        self.path = file_name + '.txt'
        self.newEntryUser()

    def find(self, name): 
        file_name = name + ".txt"
        try:
            fp = open(file_name, 'r')
            self.fileFound()
            self.doneExisting()
            self.path = file_name
            self.protocolThread = Thread(target=self.protocol,args=(file_name,))
            self.protocolThread.daemon = True
            self.protocolThread.start()
        except FileNotFoundError as err:
            self.fileNotFound()
    
    def editText(self):
        piece1 = self.entry1.get()
        piece2 = self.entry2.get()
        piece3 = self.entry3.get()
        self.entry1.delete(0, 'end')
        self.entry2Var.set('None')
        self.entry3.delete(0, 'end')
        fullText = piece1 + "-" + piece2 + "-" + piece3
        doc = open(self.path, "a")
        if(os.path.getsize(self.path) > 0):
            doc.write("\n"+fullText)
        else:
            doc.write(fullText)
    
    def newScheduleStart(self):
        self.protocolThread = Thread(target=self.protocol,args=(self.path,))
        self.protocolThread.daemon = True
        self.protocolThread.start()
        self.cleanUpNew()

    def course_timer(self, hour, minute):
        now = datetime.datetime.now()
        if minute==0:
            if hour==0:
                hour=24
            minute=60
            hour = hour-1
        course_time = now.replace(hour=hour, minute=minute-2, second=0, microsecond=0)
        if now >= course_time:
            return True
        else:
            return False

    def time_till(self, hour, minute):
        now = datetime.datetime.now()
        if minute==0:
            if hour==0:
                hour=24
            minute=60
            hour = hour-1
        course_time = now.replace(hour=hour, minute=minute-2, second=0, microsecond=0)
        time0 = now.replace(hour=0, minute=0, second=0, microsecond=0)
        dtime = course_time - now
        dtime0 = time0-time0
        if dtime.total_seconds() <= 0:
            return dtime0
        else:
            return dtime
    
    def too_late(self, hour, minute):
        now = datetime.datetime.now()
        course_time = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
        time0 = now.replace(hour=0, minute=0, second=0, microsecond=0)
        dtime = course_time - now
        inversetime = now - course_time
        if inversetime.total_seconds()>300:
            return True
        else:
            return False

    def protocol(self, scheduleName):    
        self.data.append("Getting courses...")
        time.sleep(1)
        courses = self.get_courses(scheduleName)  # [[course1, 15:40], [course2, 08:40]]
        self.data.append("Done!")
        time.sleep(1)
        self.data.append("You're all set! You will automatically be logged into your class when the time comes.")
        time.sleep(1)
        self.data.append("Do NOT close this window! This will end the program.")
        time.sleep(1)
        while True:
            while self.waiting:
                time.sleep(0.5)
            all_done = True
            count = 1
            for course in courses:
                all_done = True
                if not course[3]:  # Not marked as done yet.
                    if not (self.too_late(course[1], course[2])):
                        all_done = False
                        self.data.append(f"Waiting for Zoom meeting #{count}.")
                        timeTill = self.time_till(course[1], course[2])
                        while not self.course_timer(course[1], course[2]):  # it's not this course's time yet.
                            self.timerShow = True
                            time.sleep(1)
                            timeTill = self.time_till(course[1], course[2])
                            timeTill = timeTill - datetime.timedelta(microseconds=timeTill.microseconds)
                            self.timeTable.append(f'{timeTill} seconds till Zoom meeting #{count}')
                            self.timerShow = False
                        self.zoomLink(course[0])
                        if not (course[4] == "None"):
                            self.data.append(f"Inputting password for class #{count}")
                            time.sleep(7)
                            pyautogui.typewrite(course[4])
                            pyautogui.press('enter')
                        self.data.append(f"Attended class #{count}")
                    else:
                        self.data.append(f"It's too late to attend class #{count}")
                    course[3] = True
                    count += 1
                else:   # This course has been attended.
                    continue
            if all_done:
                print("All done!")
                self.data.append("All classes have been finished. Thanks for using AutoZoom!")
                break



    #--------------------------------------------------------------------#
    #GUI RELATED CODE
    
    def onPressExisting(self):
        self.intro1.pack_forget()
        self.intro2.pack_forget()
        self.credits.pack_forget()
        self.exist.pack_forget()
        self.new.pack_forget()
        self.search.pack()
        self.submit.pack()

    def onPressNew(self):
        self.intro1.pack_forget()
        self.intro2.pack_forget()
        self.credits.pack_forget()
        self.new.pack_forget()
        self.exist.pack_forget()
        self.name.pack()
        self.submitNew.pack()
    
    def doneExisting(self):
        self.search.pack_forget()
        self.submit.pack_forget()
        self.separator1.pack_forget()
        self.separator2.pack_forget()

    def fileFound(self):
        self.found.pack()
        self.notFound.pack_forget()

    def fileNotFound(self):
        self.found.pack_forget()
        self.notFound.pack()
    
    def newEntryUser(self):
        self.name.pack_forget()
        self.submitNew.pack_forget()
        self.instruction1.pack()
        self.entry1.pack()
        self.instruction2.pack()
        self.entry2.pack()
        self.instruction3.pack()
        self.entry3.pack()
        self.submitToDoc.pack()
        self.finishedNew.pack()
    
    def cleanUpNew(self):
        self.entry1.pack_forget()
        self.entry2.pack_forget()
        self.entry3.pack_forget()
        self.instruction1.pack_forget()
        self.instruction2.pack_forget()
        self.instruction3.pack_forget()
        self.submitToDoc.pack_forget()
        self.finishedNew.pack_forget()
        self.separator1.pack_forget()
        self.separator2.pack_forget()

    ##########################################################################################
    ################################# MAIN APPLICATION METHOD ################################
    ##########################################################################################
    def run(self): ##main part of the application
        self.root.configure(bg="#2a2b2e") #sets the background to white rather than default gray.
        self.root.protocol("WM_DELETE_WINDOW", self.quitting) ##changes the X (close) Button to run a function instead.
        try:
            self.root.iconbitmap("autozoom.ico") ##sets the application Icon
        except:
            pass
        self.root.title("AutoZoom")
        self.root.geometry("800x600")

        #############################################################
        while not self.quit: ##flag to quit the application
            self.root.update_idletasks() #updates the root. same as root.mainloop() but safer and interruptable
            self.root.update() #same as above. This lest you stop the loop or add things to the loop.
            #add extra functions here if you need them to run with the loop#
            self.update_data()
            self.update_clock()
            time.sleep(0.05)

    def quitting(self): ##to set the quit flag
        self.quit = True
            
if __name__ == "__main__":
    app = TKGUI() ##creates instance of GUI class
    try:
        app.run()# starts the application
    except KeyboardInterrupt:
        app.quitting() ##safely quits the application when crtl+C is pressed
    except:
        raise #you can change this to be your own error handler if needed