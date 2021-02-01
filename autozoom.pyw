#############################################################
#\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\#
#------------------- Made by: AffinityGH -------------------#
#///////////////////////////////////////////////////////////#
#############################################################

from tkinter import *
from tkinter.ttk import *
import  datetime, time, subprocess, csv, os, webbrowser, ctypes, pyautogui, json, configparser
from threading import Thread
from configparser import ConfigParser
myappid = 'affinity.autozoom.tkinter.beta' # arbitrary string
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

#creates documents folder to store application data
autozoom_path = os.path.dirname(str(os.environ['USERPROFILE'])+"\\Documents\\") + "\\AutoZoom\\" #writes to Documents rather than System32 to combat need of admin priveleges
try:
    os.mkdir(autozoom_path)
except:
    print("Directory exists.")

#attempts creating default settings file if it already doesn't exist
if not os.path.exists(autozoom_path + 'config.ini'):
    with open(autozoom_path + 'config.ini', 'w') as file_object:
        configObj = ConfigParser()
        configObj["SETTINGS"] = {
            "early": 2,
            "autoclose": False,
            "toolate": 5
        }
        configObj.write(file_object)
        file_object.close()
else:
    print("File exists.")

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
        self.setForTmrw=False
        self.showButtons=False
        self.answered=False
        self.finished=False

        #GLOBAL UI ELEMENTS#

        #style
        style = Style()
        style.configure("BW.TLabel", foreground="white", background="#2a2b2e")
        frame = Style()
        frame.configure('TFrame', background='#2a2b2e')
        
        #inefficient and stupid spacing
        self.separator1 = Frame(height=20,style="TFrame")
        self.separator2 = Frame(height=20,style="TFrame")
        self.separator3 = Frame(height=20,style="TFrame")
        self.separator4 = Frame(height=20,style="TFrame")
        self.separator5 = Frame(height=20,style="TFrame")
        self.separator6 = Frame(height=20,style="TFrame")
        
        #introductory labels
        self.intro1 = Label(self.root, text="Welcome to AutoZoom", style="BW.TLabel")
        self.intro2 = Label(self.root, text="Begin by choosing if you have an existing schedule, or if you need to make a new one.", style="BW.TLabel")

        #literal garbage user interface code (tkinter kinda sucks lol)
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

        self.separator4.pack()

        #temporary schedule
        self.tempEvent = Button(self.root, text="Temporary Event", command=self.onPressTemp)
        self.tempEvent.pack()

        
        #new
        self.name = Entry(self.root, textvariable=StringVar())
        self.submitNew = Button(self.root, text="Create", command=lambda: self.save(self.name.get())) 

        self.separator5.pack()

        #settings button
        self.settingsButton = Button(self.root, text="Settings", command=self.onPressSettings)
        self.settingsButton.pack()

        self.separator6.pack()


        #new --> new file (to make new files)
        self.instruction1 = Label(self.root, text="Insert the link to your classes here.", style="BW.TLabel")
        self.entry1 = Entry(self.root, textvariable=StringVar())
        
        self.instruction2 = Label(self.root, text="Insert your password here. (Default is None)", style="BW.TLabel")
        self.entry2Var = StringVar()
        self.entry2Var.set('None')
        self.entry2 = Entry(self.root, textvariable=self.entry2Var)

        self.instruction3 = Label(self.root, text="Insert the time of your class here in the 24HR format: HH:MM", style="BW.TLabel")
        self.entry3 = Entry(self.root, textvariable=StringVar())
        
        #submission and finish buttons
        self.submitToDoc = Button(self.root, text="Submit", command=self.addEvent)
        self.finishedNew = Button(self.root, text="Finished", command=self.newScheduleStart)

        #file found
        self.found = Label(self.root, text="File found!", style="BW.TLabel")
        self.notFound = Label(self.root, text="File not found...", style="BW.TLabel")

        #made by yours truly (me affinity cool guy)
        self.credits = Label(self.root, text="Made by AffinityGH", style="BW.TLabel")
        self.credits.pack()

        ###########SETTINGS GUI###########
        self.settingLabel1 = Label(self.root, text="How early would you like to join class (in minutes)?", style="BW.TLabel")
        self.setting1 = StringVar()
        self.settingEntry1 = Entry(self.root, textvariable=self.setting1)
        
        self.settingLabel2 = Label(self.root, text="Would you like to close the AutoZoom application after done with the program?", style="BW.TLabel")
        self.setting2 = StringVar()
        self.settingEntry2 = Entry(self.root, textvariable=self.setting2)

        self.settingLabel3 = Label(self.root, text="How late is too late to join a class?", style="BW.TLabel")
        self.setting3 = StringVar()
        self.settingEntry3 = Entry(self.root, textvariable=self.setting3)

        self.settingsSaveButton = Button(self.root, text="Save", command=self.settingsSave)

        self.settingsNotice = Label(self.root, text="In order to make sure your changes are saved, please restart the program.", style="BW.TLabel")

        ###########TEMP EVENT GUI###########
        
        self.tempCreateButton = Button(self.root, text="Start Temp Event", command=self.tempEventCreate)

        ###########PROTOCOL GUI###########
        self.timerShow = False
        self.textVar = StringVar()
        self.timerLabel = Label(self.root, textvariable=self.textVar, style="BW.TLabel")

        #tomorrow or not
        self.waitForTmrw = Button(self.root, text="Yes", command=self.waitTmrw)
        self.contWithSchedule = Button(self.root, text="No", command=self.contSchedule)

    #temp event creation
    def tempEventCreate(self):
        piece1 = self.entry1.get()
        piece2 = self.entry2.get()
        piece3 = self.entry3.get()
        fullText = piece1 + "-" + piece2 + "-" + piece3
        with open(autozoom_path + 'temp.txt', 'w') as file_object:
            file_object.write(fullText)

        self.protocolThread = Thread(target=self.protocol,args=('temp.txt',))
        self.protocolThread.daemon = True
        self.protocolThread.start()
        self.cleanUpNew()

    #settings backend
    def settingsSave(self):
        #save what is inside the settings
        howEarly = self.settingEntry1.get()
        autoClose = self.settingEntry2.get()
        tooLate = self.settingEntry3.get()

        with open(autozoom_path + 'config.ini', 'w') as file_object:
            config_temp = ConfigParser()
            config_temp["SETTINGS"] = {
                "early": howEarly,
                "autoclose": autoClose,
                "toolate": tooLate
            }
            config_temp.write(file_object)
            file_object.close()
        
    #sets the wait for tomorrow property to true and says that the question is answered.
    def waitTmrw(self):
        self.setForTmrw = True
        self.answered = True
    
    #ignores the wait for tomorrow and continues with normal scheduling, also saying the question is answered.
    def contSchedule(self):
        self.setForTmrw = False
        self.answered = True

    #packages the buttons in the main GUI thread since it cannot be done from the protocol thread
    def buttonPackager(self):
        if self.showButtons:
            self.waitForTmrw.pack()
            self.contWithSchedule.pack()
        else: 
            self.waitForTmrw.pack_forget()
            self.contWithSchedule.pack_forget()

    #add labels to a queue to be rendered/packed
    def update_data(self):
        if self.data != []:
            self.waiting = True
            for d in self.data:
                Label(self.root, text=d, style="BW.TLabel").pack()
            self.data = []
            self.waiting = False

    #update the clock within the main GUI thread
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
        fp = open(scheduleName, 'r') #gets the text file with the schedule contents (don't add autozoom_path here since it's already included)
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
        with open(autozoom_path + file_name + '.txt', 'w') as file_object:
            file_object.close()
        self.path = autozoom_path + file_name + '.txt'
        self.newEntryUser()

    #finds the specified text file we are looking for
    def find(self, name): 
        file_name = name + ".txt"
        try:
            fp = open(autozoom_path + file_name, 'r')
            self.fileFound()
            self.doneExisting()
            self.path = autozoom_path + file_name
            self.protocolThread = Thread(target=self.protocol,args=(self.path,))
            self.protocolThread.daemon = True
            self.protocolThread.start()
        except FileNotFoundError as err:
            self.fileNotFound()
    
    #adds an event to the schedule being edited
    def addEvent(self):
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
    
    #starts the newly created schedule
    def newScheduleStart(self):
        self.protocolThread = Thread(target=self.protocol,args=(self.path,))
        self.protocolThread.daemon = True
        self.protocolThread.start()
        self.cleanUpNew()

    #checks whether it's time to attend the meeting
    def course_timer(self, hour, minute):
        now = datetime.datetime.now()
        
        config_temp = ConfigParser()
        config_temp.read(autozoom_path + 'config.ini')

        settings = config_temp["SETTINGS"]
        howEarly = int(settings["early"])
        
        if howEarly!=0:
            if minute==0:
                if hour==0:
                    hour=24
                minute=60
                hour = hour-1

        course_time = now.replace(hour=hour, minute=minute-howEarly, second=0, microsecond=0) #the -2 is to join 2 minutes early
        if now >= course_time:
            return True
        else:
            return False

    def time_till(self, hour, minute):
        now = datetime.datetime.now()

        config_temp = ConfigParser()
        config_temp.read(autozoom_path + 'config.ini')

        settings = config_temp["SETTINGS"]
        howEarly = int(settings["early"])

        if howEarly!=0:
            if minute==0:
                if hour==0:
                    hour=24
                minute=60
                hour = hour-1

        course_time = now.replace(hour=hour, minute=minute-howEarly, second=0, microsecond=0) #offset time to class by early period in config.ini
        time0 = now.replace(hour=0, minute=0, second=0, microsecond=0)
        dtime = course_time - now
        dtime0 = time0-time0
        if dtime.total_seconds() <= 0:
            return dtime0
        else:
            return dtime

    #checks time till midnight
    def time_till_midnight(self):
        now = datetime.datetime.now()
        midnight = now.replace(hour=23, minute=59, second=59, microsecond=0) #1 second before midnight
        time0 = now.replace(hour=0, minute=0, second=0, microsecond=0)
        dtime = midnight - now
        dtime0 = time0-time0
        if dtime.total_seconds() <= 0:
            return dtime0
        else:
            return dtime

    #checks whether it's too late to attend the meeting
    def too_late(self, hour, minute):
        now = datetime.datetime.now()
        course_time = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
        time0 = now.replace(hour=0, minute=0, second=0, microsecond=0)
        dtime = course_time - now
        inversetime = now - course_time
        
        config_temp = ConfigParser()
        config_temp.read(autozoom_path + 'config.ini')

        settings = config_temp["SETTINGS"]
        toolate = float(settings["toolate"]) * 60

        if inversetime.total_seconds()>toolate: #uses the toolate attribute in config.ini to see if it's too late to attend a class
            return True
        else:
            return False

    #main protocol to go through the entire schedule
    def protocol(self, scheduleName):    
        self.data.append("Getting courses...")
        time.sleep(0.5)
        courses = self.get_courses(scheduleName)  # [[course1, 15:40], [course2, 08:40]]
        self.data.append("Done!")
        time.sleep(0.5)
        self.data.append("You're all set! You will automatically be logged into your class when the time comes.")
        time.sleep(0.5)
        self.data.append("Do NOT close this window! This will end the program.")
        time.sleep(0.5)
        while True:
            while self.waiting:
                time.sleep(0.5)
            all_done = True
            count = 1
            late = 0
            for course in courses:
                all_done = True
                if not course[3]:  # Not marked as done yet.
                    if self.too_late(course[1], course[2]):
                        late += 1
                        if (late == 1):
                            self.data.append("You are too late to attend your first class.")
                            self.data.append("Would you like to set the scheduler for tomorrow instead?")
                            self.showButtons = True
                            while not self.answered:
                                time.sleep(0.1)
                            if self.setForTmrw:
                                self.data.append("Setting for tomorrow.")
                                self.showButtons = False
                                while not self.course_timer(23, 59):
                                    self.timerShow = True
                                    time.sleep(1)
                                    timeTill = self.time_till_midnight()
                                    timeTill = timeTill - datetime.timedelta(microseconds=timeTill.microseconds)
                                    self.timeTable.append(f'{timeTill} till midnight.')
                                    self.timerShow = False
                                time.sleep(2)
                                self.data.append("It is midnight, starting scheduling.")
                            else:
                                self.data.append("Continuing with schedule.")
                                self.showButtons = False
                    if not (self.too_late(course[1], course[2])):
                        all_done = False
                        self.data.append(f"Waiting for Zoom meeting #{count}.")
                        timeTill = self.time_till(course[1], course[2])
                        while not self.course_timer(course[1], course[2]):  # it's not this course's time yet.
                            self.timerShow = True
                            time.sleep(1)
                            timeTill = self.time_till(course[1], course[2])
                            timeTill = timeTill - datetime.timedelta(microseconds=timeTill.microseconds)
                            self.timeTable.append(f'{timeTill} till Zoom meeting #{count}')
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

                config_temp = ConfigParser()
                config_temp.read(autozoom_path + 'config.ini')

                settings = config_temp["SETTINGS"]
                autoClose = settings["autoclose"]

                if (autoClose):
                    self.data.append("This window will autoclose in 10 seconds.")
                    time.sleep(10)
                
                self.finished = True
                break

    #--------------------------------------------------------------------#
    #GUI RELATED CODE - this is absolutely awful I hate Tkinter so much
    
    def onPressExisting(self):
        self.intro1.pack_forget()
        self.intro2.pack_forget()
        self.credits.pack_forget()
        self.exist.pack_forget()
        self.new.pack_forget()
        self.search.pack()
        self.submit.pack()
        self.separator5.pack_forget()
        self.settingsButton.pack_forget()
        self.tempEvent.pack_forget()

    def onPressNew(self):
        self.intro1.pack_forget()
        self.intro2.pack_forget()
        self.credits.pack_forget()
        self.new.pack_forget()
        self.exist.pack_forget()
        self.name.pack()
        self.submitNew.pack()
        self.separator5.pack_forget()
        self.settingsButton.pack_forget()
        self.tempEvent.pack_forget()

    def onPressTemp(self):
        self.intro1.pack_forget()
        self.intro2.pack_forget()
        self.credits.pack_forget()
        self.new.pack_forget()
        self.exist.pack_forget()
        self.instruction1.pack()
        self.entry1.pack()
        self.instruction2.pack()
        self.entry2.pack()
        self.instruction3.pack()
        self.entry3.pack()
        self.tempEvent.pack_forget()
        self.separator5.pack_forget()
        self.settingsButton.pack_forget()
        self.tempCreateButton.pack()
    
    def onPressSettings(self):
        self.intro1.pack_forget()
        self.intro2.pack_forget()
        self.credits.pack_forget()
        self.new.pack_forget()
        self.exist.pack_forget()
        self.separator5.pack_forget()
        self.settingsButton.pack_forget()
        self.tempEvent.pack_forget()
        
        config_temp = ConfigParser()
        config_temp.read(autozoom_path + 'config.ini')

        settings = config_temp["SETTINGS"]
        howEarly = settings["early"]
        tooLate = settings["toolate"]
        autoClose = settings["autoclose"]

        self.settingLabel1.pack()
        self.setting1.set(howEarly)
        self.settingEntry1.pack()

        self.settingLabel2.pack()
        self.setting2.set(autoClose)
        self.settingEntry2.pack()

        self.settingLabel3.pack()
        self.setting3.set(tooLate)
        self.settingEntry3.pack()

        self.settingsSaveButton.pack()

        self.settingsNotice.pack()
    
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
        self.tempCreateButton.pack_forget()

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

        config_temp = ConfigParser()
        config_temp.read(autozoom_path + 'config.ini')

        settings = config_temp["SETTINGS"]
        autoclose = settings["autoclose"]

        #############################################################
        while not self.quit: ##flag to quit the application
            self.root.update_idletasks() #updates the root. same as root.mainloop() but safer and interruptable
            self.root.update() #same as above. This lest you stop the loop or add things to the loop.
            #add extra functions here if you need them to run with the loop#
            self.update_data()
            self.update_clock()
            self.buttonPackager()
            if self.finished and autoclose:
                self.quitting()
            time.sleep(0.07)

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