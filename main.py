# Time Clock Version 2.0.0
# Written By Dylan Smith
# Managed By FRC Team 1540 The Flaming Chickens
# Last Updated December 21, 2017

from graphics import *
from oauth2client.service_account import ServiceAccountCredentials
import gspread, os, json, time, threading, smtplib, httplib2, urllib2

# graphics is a graphics library
# oauth2client, gspread, and httplib2 are for uploading to the spreadsheet
# os is used for os.system, which runs shell script
# json is for reading and editing .json files
# time gets the time and date
# theading is for creating multiple threads
# urllib2 is for checking wifi
# smtplib is for sending emails

# user can change the width and height variable to whatever fits the screen best
width=800 #window width
height=600 #window height
makeGraphicsWindow(width,height)
setWindowTitle("Time Clock")

# button class
class Button:
    def __init__(self,x,y,w,h,color,text,where,action,rounded=False,cap=1000):
        # x-position, in percentage of screen-width
        self.x=x
        # y-position, in percentage of screen-height
        self.y=y
        # width, in percentage of screen-width
        self.w=w
        # height, in percentage of screen-height
        self.h=h
        # button color
        self.color=color
        # displayed text
        self.text=text
        # what page the button is viewed on
        self.where=where
        # what happens when button is pressed
        self.action=action
        # are the edges of the button rounded (or square)
        self.rounded=rounded
        # the maximum font size
        # for the cases where you want the text to not touch the edge of the button
        self.cap=cap

# this class represents one user loggin in or out
# each log is stored in a list, where it will be later removed while updating the spreadsheet
class Log:
    def __init__(self,args):
        # in/out: whether or not user is loggin in or out
        # when in, self.io="in"
        # when out, self.io="out"
        self.io=args[0]
        # ID of user logging in/out
        self.ID=args[1]
        # time user logged in/out
        # written as hour:minutes:seconds (hours are in 24-hour time)
        # e.g. 15:30:00 shows that a user logged in/out at exactly 3:30pm
        self.time=args[2]
        # name of user logging in/out
        self.name=args[4]
        # today's date
        self.date=args[3]
        # an integer representing the row on the main spreadsheet with this user
        self.idRow=int(args[5])
        # the user's email
        self.email=args[6]
        # number of hours the user obtained
        # calculated outside of __init__ function
        self.hours=None

################################################################
###################### Utility Functions #######################
################################################################

# text to speech
def say(string):
    if os.name=="posix": # macintosh
        os.system("say '"+string+"' -vSamantha")
    elif os.name=="nt": # windows
        # creates a file to read from, reads from it, then deletes the file
        open("tts.txt","w").write(string)
        os.system("cscript 'C:\Program Files\Jampal\ptts.vbs' < tts.txt -voice 'Microsoft Hazel Desktop'")
        os.remove("tts.txt")

# function that detects if point is inside of a box
def inbox(boxx,boxy,pointx,pointy,boxwidth=50,boxheight=50):
    if pointx>=boxx and pointx<=boxx+boxwidth:
        if pointy>=boxy and pointy<=boxy+boxheight:
            return True
        else:
            return False
    else:
        return False

# uses threading to make a *function* happen every number of *seconds*
def setInterval(function, seconds):
    def functionWrapper():
        setInterval(function, seconds)
        function()
    t = threading.Timer(seconds, functionWrapper)
    t.start()
    return t

# brings user back to home page
def reset():
    getWorld().page="home"
    getWorld().id=""
    getWorld().io=None

# writes a log file to the logs folder
def createFile(l):
    #finds a file name that doesn't exist
    count = 0
    while os.path.exists("logs/"+str(count)+".json"):
        count+=1
    open("logs/"+str(count)+".json","w").write(json.dumps(l))

################################################################
################### Spreadsheet Functions ######################
################################################################

# if a file was added while offline, obtains essential data
def obtainData(logList):
    try:
        w=getWorld()
        # logList is a list: ["incomplete",in/out,id,time,date]
        if logList[2] in w.ids:
            idRow = w.ids.index(logList[2])+w.labelRow+1 # row on hours/certs spreadsheet with the person
            name = w.sheet.cell(idRow,w.nameCol).value
            email = w.sheet.cell(idRow,w.emailCol).value
            for val in w.sheet2.col_values(w.nameCol2):
                # You Are Clocked In and Are Logging In
                if val==name and logList[1]=="in":
                    return False
                # You Are Clocked In and Are Logging Out
                elif val==name and logList[1]=="out":
                    return [logList[1],logList[2],logList[3],logList[4],name,str(idRow),email]
                # You Are Clocked Out and Are Logging In
                elif val=="" and logList[1]=="in":
                    return [logList[1],logList[2],logList[3],logList[4],name,str(idRow),email]
                # You Are Clocked Out and Are Logging Out
                elif val=="" and logList[1]=="out":
                    return False
        else:
            print "Attempt to log"+logList[1]+" with ID " + str(logList[2]) + " failed."
            return False
    except:
        checkConnection()
        return None

# adds data from the text files to the spreadsheet
def checkLogs():
    w=getWorld()
    checkConnection() # checks wifi
    if not w.running and w.connection:
        w.running=True
        # keys: number in each file name
        # values: a list with all of the data from the file
        logData = {}
        for f in os.listdir("logs"):
            # only accepts .json files
            if f[-5:]==".json":
                logList = json.load(open("logs/"+f,"r"))
                if logList[0]=="incomplete":
                    logList = obtainData(logList) # returns false if not a valid ID
                if logList==None:
                    break
                elif not logList==False:
                    logData[int(f[:-5])]=Log(logList)
                os.remove("logs/"+f)
        # the following line is to make the file sorted in numerical order (8,9,10,11) rather than string order (1,10,11,2)
        logs = sorted(logData.keys())
        while len(logs)>0:
            if checkConnection()==True:
                # log is a list of all the data
                log = logData.pop(logs[0])
                logs.pop(0)
                if log.io=="in":
                    login(log)
                else:
                    logout(log)
                # uploads data to the History Spreadsheet
                history(log)
            else:
                print "HI"
                val = 0
                for l in logs:
                    createFile(logData[l])
                    val+=1
                break
        w.running=False

# checks to make sure that nobody is clocked in past midnight
# if someone is clocked in past midnight, the system deletes them from the spreadsheet, then emails them telling them they forgot to clock out.
# this runs every two hours
def checkDates():
    w=getWorld()
    checkConnection() # checks wifi
    if w.connection:
        try:
            # if the first row on the spreadsheet contains a different date than todays date
            if not w.sheet2.cell(2,w.dateInCol).value==time.strftime("%x"):
                # the following code deletes all rows on spreadsheet and emails members
                while not w.sheet2.cell(2,w.dateInCol).value==time.strftime("%x"):
                    # first three lines are for logging in to 1540photo@gmail.com
                    server = smtplib.SMTP("smtp.gmail.com",587)
                    server.starttls()
                    server.login("1540photo@gmail.com","robotics1540")
                    # sends an email to the user
                    msg = "You forgot to sign out of the lab on "+w.sheet2.cell(2,w.dateInCol).value+". You cannot log these hours online. Next time, remember to log out."
                    server.sendmail("1540photo@gmail.com",w.sheet2.cell(2,w.emailCol2).value,msg)
                    server.quit()
                    # deletes the row
                    w.sheet2.delete_row(2)
                    return True
        except:
            checkConnection()

# the following uploads data to a history spreadsheet that shows the history of any member
def history(log):
    w=getWorld()
    checkConnection() # checks wifi
    if w.connection:
        try:
            count=1 # count will be the row-value with the member's name
            for n in w.sheet3.col_values(1):
                if n=="":
                    break
                elif n==log.name:
                    string = log.io + ";" + log.time + ";" + log.date # shows in/out, what time, and the date
                    count2=1 # count2 will be the col-value where the info will be pasted
                    add=False
                    for c in w.sheet3.row_values(count):
                        if c=="":
                            w.sheet3.update_cell(count,count2,string) # putting string on the spreadsheet
                            #if logging out, puts information about number of hours logged for that day
                            if not log.hours==None:
                                count2+=1
                                w.sheet3.update_cell(count,count2,log.hours+" hours")
                            add=True
                            break
                        count2+=1
                    # if there was no column availabe, add will be false
                    # from this point, it adds new columns along with putting data on the spreadsheet
                    if not add:
                        w.sheet3.add_cols(1)
                        w.sheet3.update_cell(count,count2,string)
                        if log.io=="out":
                            count2+=1
                            w.sheet3.add_cols(1)
                            w.sheet3.update_cell(count,count2,log.hours+" hours")
                    break
                count+=1
        except:
            checkConnection()

# puts data onto the spreadsheet when logging in
def login(log):
    w=getWorld()
    checkConnection() # checks wifi
    if w.connection:
        try:
            openRow = 1
            # this loop looks for the first open row on the lab attendance sheet
            # if the name is already on the spreadsheet, openRow=-1 and the name is not re-added to the spreadsheet
            for val in w.sheet2.col_values(1):
                if val=="":
                    break
                elif val==log.name:
                    openRow = -1
                    break
                openRow+=1
            if not openRow==-1:
                # uploading to the lab attendance sheet
                w.sheet2.insert_row([log.name,log.time,log.date,log.email],openRow)
        except:
            checkConnection()

# puts data onto the spreadsheet when logging out
def logout(log):
    w=getWorld()
    checkConnection() # checks wifi
    if w.connection:
        try:
            nameRow = 1
            # this loop looks for the row with the user's name on the lab attendance sheet
            # if the name is not on the spreadsheet, nameRow=-1 and the function ends
            for val in w.sheet2.col_values(w.nameCol2):
                if val=="":
                    nameRow = -1
                    break
                elif val==log.name:
                    break
                nameRow+=1
            if not nameRow==-1: # checks to see if nameRow is -1
                old = w.sheet2.cell(nameRow,w.timeInCol).value
                p1 = old.split(":") # the previous logged time in a list with format [hours, minutes, seconds]
                p2 = log.time.split(":") # the new logged time in a list with format [hours, minutes, seconds]
                val1 = float(p1[0]) + float(p1[1])/60.0 + float(p1[2])/360.0 # converts p1 into just hours
                val2 = float(p2[0]) + float(p2[1])/60.0 + float(p2[2])/360.0 # converts p2 into just hours
                total = val2 - val1 # finds the net time
                current = 0.0 # current is the time on the hours and certs spreadsheet
                if not w.sheet.cell(log.idRow,w.labHoursCol).value=="":
                    current = float(w.sheet.cell(log.idRow,w.labHoursCol).value)
                log.hours = str(round(current+total,2)) # rounds the total hours to two digits
                # updates the hours and certs spreadsheets
                w.sheet.update_cell(log.idRow,w.labHoursCol,log.hours)
                # removes member from the lab attendance spreadsheet
                w.sheet2.delete_row(nameRow)
        except:
            checkConnection()

################################################################
###################  Connection Functions  #####################
################################################################

# checks to see if you have wifi
def checkConnection():
    w=getWorld()
    try:
        # tries connecting to google.com
        urllib2.urlopen('https://www.google.com', timeout=1)
        if w.connection==False:
            w.connection=True
            connect()
        return True
    except urllib2.URLError:
        w.connection=False
        return False

# tries connecting to the google sheets server
def connect():
    w=getWorld()
    try:
        # sets up the client (only used in the next three lines)
        client = gspread.authorize(ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', ['https://spreadsheets.google.com/feeds']))
        # The Hours and Certifications Spreadsheet
        w.sheet = client.open("Hoursheet").sheet1
        # The Current Lab Attendance Spreadsheet
        w.sheet2 = client.open("Hoursheet").get_worksheet(1)
        # The Lab History Spreadsheet
        w.sheet3 = client.open("History").sheet1
        # finds where the labeled columns start
        count = 1
        for val in w.sheet.col_values(1):
            if val=="Name":
                w.labelRow = count
                break
            count+=1
        # finds the location of each row/col
        for s in [w.sheet.row_values(w.labelRow),w.sheet2.row_values(1)]:
            count = 1
            for i in s:
                if i=="":
                    break
                elif i=="Name":
                    if w.nameCol==None:
                        w.nameCol = count
                    else:
                        w.nameCol2 = count
                elif i=="Time In":
                    w.timeInCol = count
                elif i=="Lab Hours":
                    w.labHoursCol = count
                elif i=="Date In":
                    w.dateInCol = count
                elif i=="ID":
                    w.idCol = count
                elif i=="Email":
                    if w.emailCol==None:
                        w.emailCol = count
                    else:
                        w.emailCol2 = count
                count+=1

        # appends all IDs in order to w.ids
        seen=False
        for i in w.sheet.col_values(w.idCol):
            if i=="ID":
                seen=True
            elif not i=="":
                w.ids.append(i)
            elif seen==True:
                break
    except httplib2.ServerNotFoundError:
        w.connection = False

################################################################
########################## Mouse Press #########################
################################################################

# x is x-position of mouse
# y is y-position of mouse
# b is the button pressed on the mouse (left-click is 1)
def mousePress(w,x,y,b):
    for u in w.buttons:
        # checks to see if user clicked on the button
        if b==1 and w.page==u.where and inbox(u.x*width,u.y*height,x,y,u.w*width,u.h*height):
            return u.action(u)

onMousePress(mousePress)

################################################################
######################## Button Actions ########################
################################################################

# all of the following functions take b, the button, as an argument for consistency
# these are run when buttons are pressed

#sends user to login screen
def IN(b):
    getWorld().page="login/logout"
    getWorld().io="in"

#sends user to logout screen
def OUT(b):
    getWorld().page="login/logout"
    getWorld().io="out"

#adds the button's display to the id
def KEY(b):
    if len(getWorld().id)<4:
        getWorld().id+=b.text

#backspace key
def DELETE(b):
    if len(getWorld().id)>0:
        getWorld().id=getWorld().id[:-1]

#return to home page
def PASS(b):
    reset()

#checks to see if ID is an actual ID. if so, signs you in/out
def OK(b):
    w=getWorld()
    checkConnection() # checks wifi
    if w.connection:
        try:
            if w.id in w.ids:
                t = time.strftime("%H:%M:%S") # the time in hours:minutes:seconds
                d = time.strftime("%x") # the date
                idRow = w.ids.index(w.id)+w.labelRow+1 # row on hours/certs spreadsheet with the person
                name = w.sheet.cell(idRow,w.nameCol).value
                email = w.sheet.cell(idRow,w.emailCol).value
                # os.system("say 'string' -vSamantha") speaks 'string' using Siri's voice
                for val in w.sheet2.col_values(w.nameCol2):
                    # You Are Clocked In and Are Logging In
                    if val==name and w.io=="in":
                        say("You are already clocked in!")
                        break
                    # You Are Clocked In and Are Logging Out
                    elif val==name and w.io=="out":
                      # say("Goodbye "+name.split(" ")[0])
                        createFile([w.io,w.id,t,d,name,str(idRow),email])
                        reset()
                        break
                    # You Are Clocked Out and Are Logging In
                    elif val=="" and w.io=="in":
                       # say("Welcome "+name.split(" ")[0])
                        createFile([w.io,w.id,t,d,name,str(idRow),email])
                        reset()
                        break
                    # You Are Clocked Out and Are Logging Out
                    elif val=="" and w.io=="out":
                        say("You never signed in!")
                        break
        except:
            t = time.strftime("%H:%M:%S") # the time in hours:minutes:seconds
            d = time.strftime("%x") # the date
            createFile(["incomplete",w.io,w.id,t,d])
            reset()
        else:
            # if the ID is invalid, as it doesn't exist
            say(str(w.id)+"is not a valid ID!")
            print w.id+" is not a valid ID."
    else:
        t = time.strftime("%H:%M:%S") # the time in hours:minutes:seconds
        d = time.strftime("%x") # the date
        createFile(["incomplete",w.io,w.id,t,d])
        reset()

################################################################
################################################################
################################################################

# the start function runs when app is opened
def start(w):
    # what page you are on
    w.page = "home"
    # whether or not you are logging in or out
    # when in, w.io="in"
    # when out, w.io="out"
    w.io = None
    # the id presently being type
    w.id = ""
    # list of all IDs in same order as spreadsheet
    w.ids = []
    # list of all logs that still need to be uploaded to the spreadsheet
    w.logs = []
    # whether or not checkLogs() is running
    w.running = False

    w.connection = False
    w.sheet = None # The Hours and Certifications Spreadsheet
    w.sheet2 = None # The Current Lab Attendance Spreadsheet
    w.sheet3 = None # The Lab History Spreadsheet

    w.buttons = [
    Button(0.07,0.3,0.4,0.4,(3,155,229),"Log In","home",IN,cap=50),
    Button(0.53,0.3,0.4,0.4,(255,171,64),"Log Out","home",OUT,cap=50),
    Button(0.24,0.19,0.15,0.15,(208,211,216),"1","login/logout",KEY,cap=50),
    Button(0.42,0.19,0.15,0.15,(208,211,216),"2","login/logout",KEY,cap=50),
    Button(0.60,0.19,0.15,0.15,(208,211,216),"3","login/logout",KEY,cap=50),
    Button(0.24,0.39,0.15,0.15,(208,211,216),"4","login/logout",KEY,cap=50),
    Button(0.42,0.39,0.15,0.15,(208,211,216),"5","login/logout",KEY,cap=50),
    Button(0.60,0.39,0.15,0.15,(208,211,216),"6","login/logout",KEY,cap=50),
    Button(0.24,0.59,0.15,0.15,(208,211,216),"7","login/logout",KEY,cap=50),
    Button(0.42,0.59,0.15,0.15,(208,211,216),"8","login/logout",KEY,cap=50),
    Button(0.60,0.59,0.15,0.15,(208,211,216),"9","login/logout",KEY,cap=50),
    Button(0.24,0.79,0.15,0.15,(98,178,85),"OK","login/logout",OK,cap=50),
    Button(0.42,0.79,0.15,0.15,(208,211,216),"0","login/logout",KEY,cap=50),
    Button(0.60,0.79,0.15,0.15,(237,99,92),"Del","login/logout",DELETE,cap=50),
    Button(0.8,0.05,0.15,0.15,(3,155,229),"Cancel","login/logout",PASS,cap=30)
    ]

    # the following are just row/col ids for rows/cols.
    # e.g. if the column titled "Date In" is the third column, w.dateInCol will be 3.
    # w.nameCol must be the leftmost column in w.sheet
    w.labelRow = None
    w.nameCol = None
    w.nameCol2 = None
    w.timeInCol = None
    w.labHoursCol = None
    w.dateInCol = None
    w.idCol = None
    w.emailCol = None
    w.emailCol2 = None

    checkConnection()

    # checks to see if values currently on w.sheet2 are from today
    # if not, deletes them
    checkDates()
    # runs checkDates every .002 hours
    setInterval(checkDates,72)
    # runs checkLogs every 5 seconds
    setInterval(checkLogs,5)

def update(w):
    pass

# draw function for button
def button(x,y,w,h,color,text,rounded,cap):
    # draws a curved button by drawing to rectangles in a lowercase t shape
    # and then proceeds to draw arcs to fill up each corner
    if rounded:
        # center / t-shape
        fillRectangle(x*width+30,y*height,w*width-60,h*height,color=color)
        fillRectangle(x*width,y*height+30,w*width,h*height-60,color=color)
        # top-left
        fillCircle(30+x*width,30+y*height,30,color=color)
#        drawArcCircle(30+x*width,30+y*height,30,90,180)
        # top-right
        fillCircle((x+w)*width-30,30+y*height,30,color=color)
#        drawArcCircle((x+w)*width-30,30+y*height,30,0,90)
        # bottom-left
        fillCircle(30+x*width,30+(y+h)*height-60,30,color=color)
#        drawArcCircle(30+x*width,30+(y+h)*height-60,30,180,270)
        # bottom-right
        fillCircle((x+w)*width-30,30+(y+h)*height-60,30,color=color)
#        drawArcCircle((x+w)*width-30,30+(y+h)*height-60,30,270,360)
        # vertical-lines
#        drawLine(x*width,y*height+30,x*width,(y+h)*height-30)
#        drawLine((x+w)*width,y*height+30,(x+w)*width,(y+h)*height-30)
        # horizontal-lines
#        drawLine(x*width+30,y*height,(x+w)*width-30,y*height)
#        drawLine(x*width+30,(y+h)*height,(x+w)*width-30,(y+h)*height)
    else:
        # draws a rectangular button
        fillRectangle(x*width,y*height,w*width,h*height,color=color)
#        drawRectangle(x*width,y*height,w*width,h*height)
    if not text=="":
        # size represents the size that each character of the string is
        size = 1
        # currentStringSize is a tuple (width of string,height of string)
        currentStringSize = sizeString(text,font="Arial",size=size)
        # the following loop attempts to find a size for the text that will not exceed the size of the button
        while ((w*width)-currentStringSize[0]>20) and ((h*height)-currentStringSize[1]>20):
            currentStringSize = sizeString(text,font="Arial",size=size)
            # checks to see if font-size has reached the font-cap yet
            if size>=cap:
                break
            size+=1
        xpos = (x*width) + ((w*width)/2) - (currentStringSize[0]/2)
        ypos = (y*height) + ((h*height)/2) - (currentStringSize[1]/2)
        drawString(text,xpos,ypos,font="Arial",size=size)

# drawing all text/buttons
# this function clears the display before running
def draw(w):
    # draws each button
    for b in w.buttons:
        if b.where==w.page:
            text = b.text
            if b.text=="OK": # for the clarity of the login/logout screens
                text=w.io.capitalize()
            button(b.x,b.y,b.w,b.h,b.color,text,b.rounded,b.cap)
    # for the display on the login/logout page
    if w.page=="login/logout":
        s = sizeString(w.id,90,font="Arial")
        drawString(w.id,width/2-s[0]/2.0,height*.01,size=90,font="Arial")
    if w.connection==False:
        drawString("Please connect to Wifi.",10,10,size=20,font="Arial")

# runs the start, update, and draw function
runGraphics(start,update,draw)
