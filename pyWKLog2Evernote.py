# Python WKLog to Evernote
# By: TCC
# Date: 20140612
#
from twisted.internet import reactor, protocol, task
import socket
import sys
import time
import datetime
import hashlib
import binascii
import evernote.edam.userstore.constants as UserStoreConstants
import evernote.edam.type.ttypes as Types
import ConfigParser
import getopt
import os
import twitter


from evernote.api.client import EvernoteClient
import pyodbc


#database define
#DBfile = 'C:\pyWKLog2Evernote\msudbAS.mdb'
#DBfile = 'W:\MSUDBASE\WKLog3a\msudbAS.mdb'
#conn = pyodbc.connect('DRIVER={Microsoft Access Driver (*.mdb)};DBQ='+DBfile)
#conn = pyodbc.connect(DBfile)
#check database connection OK or not
#
DBfile = 'DSN=WKLOG'
#DBfile = 'DSN=WKLogDev'

try:
    conn = pyodbc.connect(DBfile)
    cursor = conn.cursor()
    #conn.close()
    print "Database: " + DBfile + " connection is OK."
except pyodbc.Error, err:
    print "ODBC connection error: %s", err
    quit()

    
now = datetime.datetime.now()
print "Year: " + str(now.year)

#SQL = "SELECT WKLogMain.[WKAutoNo] FROM WKLogMain ORDER BY WKLogMain.[WKAutoNo] DESC"
#SQL = "SELECT * FROM WKLogMain WHERE LogDate like '%2014%' ORDER BY WKAutoNo DESC;"
#SQL = "SELECT WKLogMain.[WKAutoNo] FROM WKLogMain ORDER BY WKLogMain.[WKAutoNo] DESC"
#SQL = "SELECT * FROM WKLogMain ORDER BY WKAutoNo DESC"
#SQL = "SELECT * FROM WKLogMain WHERE LogDate like '%2014%' ORDER BY WKAutoNo DESC;"    
SQL = "SELECT * FROM WKLogMain WHERE LogDate like " + "'%" + str(now.year) + "%'" + "ORDER BY WKAutoNo DESC;"
    
#Global variable
oldWKAutoNo = 0
lastWKAutoNo = 0
svrListenPort = 5000
elapseTime = 0
tryReconnect = 0
EverNoteErrorCount = 0
dbErrorCount = 0
consumer_key = ""
consumer_secret = ""
access_key = ""
access_secret = ""
auth_token = ""

#read lastAutoNo
try:
    file = open('lastAutoNo.txt','r')
    oldWKAutoNo  = int(file.read())
    print 'Old WKAutoNo: %s' % str(oldWKAutoNo)
    file.close()
except IOError:
    print "file read error"
    quit()

def readTokenFile():
    global consumer_key
    global consumer_secret
    global access_key
    global access_secret
    global auth_token
    
    try:
        file = open('token.txt','r')
        some_list = file.readlines()
        matching = [s for s in some_list if "consumer_key" in s]
        consumer_key = str(matching[0]).replace('\n','')
        consumer_key = consumer_key.replace('consumer_key ','')
        print 'Twitter consumer_key: %s' % consumer_key

        matching = [s for s in some_list if "consumer_secret" in s]
        consumer_secret = str(matching[0]).replace('\n','')
        consumer_secret = consumer_secret.replace('consumer_secret ','')
        print 'Twitter consumer_secret: %s' % consumer_secret

        matching = [s for s in some_list if "access_key" in s]
        access_key = str(matching[0]).replace('\n','')
        access_key = access_key.replace('access_key ','')
        print 'Twitter access_key: %s' % access_key

        matching = [s for s in some_list if "access_secret" in s]
        access_secret = str(matching[0]).replace('\n','')
        access_secret = access_secret.replace('access_secret ','')
        print 'Twitter access_secret: %s' % access_secret

        matching = [s for s in some_list if "auth_token" in s]
        auth_token = str(matching[0]).replace('\n','')
        auth_token = auth_token.replace('auth_token ','')
        print 'Evernote auth_token: %s' % auth_token

        file.close()
        
    except IOError:
        print '***Token file not exist***'
        quit()

                    
def writeLastAutoNo(autoNo):
    try:
        file = open('lastAutoNo.txt','w')
        file.write(str(autoNo))
        file.close()
    except IOError:
        print "file write error"


def postWKtwitter(message):
# post to @watchkeepermsu twitter accout
    encoding = None

    api = twitter.Api(consumer_key=consumer_key, consumer_secret=consumer_secret,
                    access_token_key=access_key, access_token_secret=access_secret,
                    input_encoding=encoding)

    try:
        status = api.PostUpdate(message)
    except UnicodeDecodeError:
        print "Your message could not be encoded.  Perhaps it contains non-ASCII characters? "
        print "Try explicitly specifying the encoding with the --encoding flag"
        sys.exit(2)
    print "%s just tweet %s" % (status.user.name, status.text) 
  

# Real applications authenticate with Evernote using OAuth, but for the
# purpose of exploring the API, you can get a developer token that allows
# you to access your own Evernote account. To get a developer token, visit
# https://sandbox.evernote.com/api/DeveloperToken.action
#
# add a new note into the Evernote Default Notebook
def addNewNote(row):
    global EverNoteErrorCount 

    if auth_token == "your developer token":
        print "Please fill in your developer token"
        print "To get a developer token, visit " \
            "https://sandbox.evernote.com/api/DeveloperToken.action"
        exit(1)

    # Initial development is performed on our sandbox server. To use the production
    # service, change sandbox=False and replace your
    # developer token above with a token from
    # https://www.evernote.com/api/DeveloperToken.action
    #client = EvernoteClient(token=auth_token, sandbox=True)
    client = EvernoteClient(token=auth_token, sandbox=False)

    try:
        user_store = client.get_user_store()
        print "Get user store OK"
    except user_store.Error:
        print "Get user store error"
        return -1
    

    try:
        version_ok = user_store.checkVersion(
            "Evernote EDAMTest (Python)",
            UserStoreConstants.EDAM_VERSION_MAJOR,
            UserStoreConstants.EDAM_VERSION_MINOR
        )
        print "Evernote API version up to date: ", str(version_ok)
        if not version_ok:
            print "My Evernote API version not up to date. ****"
            exit(1)
    except user_store.Error:
        print "Get Evernote API version error"
        return -1
    

    try:
        note_store = client.get_note_store()

        # List all of the notebooks in the user's account
        notebooks = note_store.listNotebooks()
        #print "Found ", len(notebooks), " notebooks:"
        notebookFound = False
        for notebook in notebooks:
            #print "  * ", notebook.name
            if (notebook.name == "S01.MSU.WKLog"):
                #print "found WKLog Notebook"
                note = Types.Note()
                note.notebookGuid = notebook.guid
                notebookFound = True

        if (notebookFound != True):
            print "**** Share notebook not found ****"
            quit()
            
        print "--- Creating a new note in the notebook: S01.MSU.WKLog --"

    except:
        print "Get note store error: %s", err
        EverNoteErrorCount = EverNoteErrorCount + 1
        return -1
   

    # To create a new note, simply create a new Note object and fill in
    # attributes such as the note's title.
    # note = Types.Note()
    note.title = row.WKRefNo

    print str(datetime.datetime.now())
    # check content of row is null
    if row.WKAutoNo is None:
        #error
        print "*** WKAutoNo Error ***"
        return

    if row.WKRefNo is None:
        #error
        print "*** WKRefNo Error ***"
    else:
        print 'WKRefNo: %s' % row.WKRefNo
        
    if row.LogDate is None:
        LogDate = 'NIL'
    else:
        LogDate = row.LogDate
        
    if row.LogTime is None:
        LogTime = 'NIL'
    else:
        LogTime = row.LogTime
        
    if row.RefDate is None:
        RefDate = 'NIL'
    else:
        RefDate = row.RefDate
        
    if row.RefTime is None:
        RefTime = 'NIL'
    else:
        RefTime = row.RefTime
        
    if row.AttendDate is None:
        AttendDate = 'NIL'
    else:
        AttendDate = row.AttendDate
        
    if row.AttendTime is None:
        AttendTime = 'NIL'
    else:
        AttendTime = row.AttendTime
        
    if row.ClrDate is None:
        ClrDate = 'NIL'
    else:
        ClrDate = row.ClrDate
        
    if row.ClrTime is None:
        ClrTime = 'NIL'
    else:
        ClrTime = row.ClrTime

    if row.RptBy is None:
        RptBy = 'NIL'
    else:
        RptBy = row.RptBy

    if row.AckBy is None:
        AckBy = 'NIL'
    else:
        AckBy = row.AckBy
        
    if row.Cat is None:
        Cat = 'NIL'
    else:
        Cat = row.Cat
        print "CAT: " + Cat
        
    if row.Site is None:
        Site = 'NIL'
    else:
        Site = row.Site
        if len(row.Site) < 3:
            Site = 'NIL'
            print "Site(<3): " + Site
        else:
            print "Site: " + Site
            
    if row.SubSys is None:
        SubSys = 'NIL'
    else:
        SubSys = row.SubSys
        if len(SubSys) < 3:
            SubSys = 'NIL'
            print "SubSys(<3): " + SubSys
        else:
            print "Subsys: " + SubSys

    # check special char & and replace    
    if row.Symptoms is None:
        Symptoms = 'NIL'
    else:
        Symptoms = row.Symptoms
        Symptoms = Symptoms.replace('&', 'and')
        tweetSymptoms = Symptoms[:120]
        print "Symptoms: " + Symptoms
    
        
    if row.Actions is None:
        Actions = 'NIL'
    else:
        Actions = row.Actions
        Actions = Actions.replace('&', 'and')
        print "Actions: " + Actions
        
    if row.Status is None:
        Status = 'NIL'
    else:
        Status = row.Status
        
    if row.RefTo is None:
        RefTo = 'NIL'
    else:
        RefTo = row.RefTo
           

    # The content of an Evernote note is represented using Evernote Markup Language
    # (ENML). The full ENML specification can be found in the Evernote API Overview
    # at http://dev.evernote.com/documentation/cloud/chapters/ENML.php
    #note.content = '<?xml version="1.0" encoding="UTF-8"?>'
    #note.content += '<!DOCTYPE en-note SYSTEM ' \
    #    '"http://xml.evernote.com/pub/enml2.dtd">'
    #note.content += '<en-note>Here is the Evernote logo:<br/>'
    #note.content += '<en-media type="image/png" hash="' + hash_hex + '"/>'
    #note.content += '</en-note>'
    note.content = '<?xml version="1.0" encoding="UTF-8"?>'
    note.content += '<!DOCTYPE en-note SYSTEM ' \
        '"http://xml.evernote.com/pub/enml2.dtd">'
    note.content += '<en-note>'
    note.content += '<span style="font-weight:bold;color:black;">' + 'Symptoms: ' + '</span>'  + Symptoms + '<br/>'
    note.content += '<span style="font-weight:bold;color:black;">' + 'Actions: ' + '</span>'  + Actions + '<br/>'
    note.content += '<span style="font-weight:bold;color:black;">' + 'WKAutoNo: ' + '</span>' + str(row.WKAutoNo) + '<br/>'
    note.content += '<span style="font-weight:bold;color:black;">' + 'WKRefNo: ' + '</span>' + row.WKRefNo + '<br/>'
    note.content += '<span style="font-weight:bold;color:black;">' + 'LogDateTime: ' + '</span>' + LogDate + ' ' + LogTime + '<br/>'
    note.content += '<span style="font-weight:bold;color:black;">' + 'RefDateTime: ' + '</span>' + RefDate + ' ' + RefTime + '<br/>'
    note.content += '<span style="font-weight:bold;color:black;">' + 'AttendDateTime: ' + '</span>' + AttendDate + ' ' + AttendTime + '<br/>'
    note.content += '<span style="font-weight:bold;color:black;">' + 'ClrDateTime: ' + '</span>'  + ClrDate + ' ' + ClrTime + '<br/>'
    note.content += '<span style="font-weight:bold;color:black;">' + 'RptBy: ' + '</span>'  + RptBy + '<br/>'
    note.content += '<span style="font-weight:bold;color:black;">' + 'AckBy: ' + '</span>'  + AckBy + '<br/>'
    note.content += '<span style="font-weight:bold;color:black;">' + 'Cat: ' + '</span>'  + Cat + '<br/>'
    note.content += '<span style="font-weight:bold;color:black;">' + 'Site: ' + '</span>'  + Site + '<br/>'
    note.content += '<span style="font-weight:bold;color:black;">' + 'SubSys: ' + '</span>'  + SubSys + '<br/>'
    note.content += '<span style="font-weight:bold;color:black;">' + 'Symptoms: ' + '</span>'  + Symptoms + '<br/>'
    #note.content += '<span style="font-weight:bold;color:black;">' + 'Actions: ' + '</span>'  + Actions + '<br/>'
    #note.content += '<span style="font-weight:bold;color:black;">' + 'Status: ' + '</span>'  + Status + '<br/>'
    note.content += '<span style="font-weight:bold;color:black;">' + 'RefTo: ' + '</span>'  + RefTo + '<br/>'
    note.content += '</en-note>'

    # add tag info
    note.tagNames = [SubSys, Site, Cat]
    
    # Finally, send the new note to Evernote using the createNote method
    # The new Note object that is returned will contain server-generated
    # attributes such as the new note's unique GUID.

    try:
        created_note = note_store.createNote(note)
        print "--- Successfully created a new note ---"
        print "###                                 ###"
        postWKtwitter(row.WKRefNo + ": " + tweetSymptoms)
        print "###                                 ###"
        print
        return 0
        
    except note_store.Error, err:
        print "Create note error: %s", err
        EverNoteErrorCount = EverNoteErrorCount + 1
        return -1
    

# update Evernote
def updateEvernote(row):
    print 'Create note: %s %s' % (str(row.WKAutoNo), row.WKRefNo)
    return addNewNote(row)
    

#init database while startup
def initDatabase():
    global oldWKAutoNo
    global dbErrorCount

    return 0
    try:
        cursor.execute(SQL)
        row = cursor.fetchone()
        if row:
            oldWKAutoNo = row.WKAutoNo
            print 'Found last WKAutoNo: %d' % oldWKAutoNo
            return 0
    except pyodbc.Error, err:
        print "ODBC connection error: %s", err
        dbErrorCount = dbErrorCount + 1
        quit()
        return -1



# reconnect database
def reconnect():
    global conn
    global cursor
    global tryReconnect
    global dbErrorCount

    try:
        print
        print "Try reconnect the Database"
        conn = pyodbc.connect(DBfile)
        cursor = conn.cursor()
        tryReconnect = 0
        dbErrorCount = dbErrorCount + 1
    except pyodbc.Error, err:
        print
        print "ODBC connection error: %s", err
        dbErrorCount = dbErrorCount + 1
        tryReconnect = 1
        
   
# check database
def checkDatabase():
    global oldWKAutoNo
    global lastWKAutoNo
    global elapseTime
    global tryReconnect

    try:
        cursor.execute(SQL)
        row = cursor.fetchone()
        if row:
            lastWKAutoNo = row.WKAutoNo
            # don't print the dbase info
            # print 'lastWKAutoNo: %d, dbErrorCount: %d, EverNoteErrorCount: %d' % (lastWKAutoNo, dbErrorCount, EverNoteErrorCount)
            NoRecordUpdate = lastWKAutoNo - oldWKAutoNo
            
        if (NoRecordUpdate == 0):
            return 0
        else:
            print
            print 'No of record to be updated: %d' % NoRecordUpdate
            a = []
            i = 1
            while i <= NoRecordUpdate:
                a.append(row)
                print "Record: " + str(i) + " " + str(a[i-1].WKAutoNo)
                row = cursor.fetchone()
                i+=1

            j = NoRecordUpdate
            while j >= 1:
                # print "Update record: " + str(j) + " " + str(a[j-1].WKAutoNo)
                state = updateEvernote(a[j-1])
                if (state <0):
                    return -1
                j -=1
        
            oldWKAutoNo = lastWKAutoNo

            #update to file
            writeLastAutoNo(lastWKAutoNo)
            
            elapseTime = 0
            return 0

    except pyodbc.Error, err:
        print
        print "ODBC connection error: %s", err
        conn.close()
        tryReconnect = 1
        return -1
                      
class MyProtocol(protocol.Protocol):
    
    def connectionMade(self):
        self.factory.clientConnectionMade(self)
        print "clients connection made: ", self.factory.clients
        print "no of clients connected: ", self.factory.numClients
        
    def connectionLost(self, reason):
        self.factory.clientConnectionLost(self)

    def dataReceived(self, data):
        print "raw data rx:", data
        
    def message(self, message):
        self.transport.write(message.encode('ascii','ignore') + '\n')



class MyFactory(protocol.Factory):
    protocol = MyProtocol
    
    def __init__(self):
        self.numClients = 0 
        self.clients = []
        self.lc = task.LoopingCall(self.announce)
        self.lc.start(10)

    def announce(self):
        global elapseTime
        # 5 sec task to check database
        if tryReconnect == 1:
            reconnect()
        else:
            state = ""
            state = checkDatabase()
            elapseTime = elapseTime + 10
            sys.stdout.write('ElapseTime: %ssec, dbErrorCount: %s, EverNoteErrorCount: %s\r'
                             % (str(elapseTime), str(dbErrorCount), str(EverNoteErrorCount)))
            sys.stdout.flush()
                             
    def clientConnectionMade(self, client):
        self.clients.append(client)
        self.numClients = self.numClients+1
        
    def clientConnectionLost(self, client):
        print 'client disconnect'
        self.clients.remove(client)
        self.numClients = self.numClients-1


# main program start here
print 'WKLog to Evernote started'
print 'My IP Address: ', socket.gethostbyname(socket.gethostname())
print 'I am listening on port: ', svrListenPort
readTokenFile()
initDatabase()
myfactory = MyFactory()
reactor.listenTCP(svrListenPort, myfactory)
reactor.run()
