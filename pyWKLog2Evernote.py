from twisted.internet import reactor, protocol, task
import socket
import sys
import time
import datetime
import hashlib
import binascii
import evernote.edam.userstore.constants as UserStoreConstants
import evernote.edam.type.ttypes as Types

from evernote.api.client import EvernoteClient
import pyodbc


#database define
#DBfile = 'C:\pyWKLog2Evernote\msudbAS.mdb'
#DBfile = 'W:\MSUDBASE\WKLog3a\msudbAS.mdb'
DBfile = 'DSN=WKLOG'

#conn = pyodbc.connect('DRIVER={Microsoft Access Driver (*.mdb)};DBQ='+DBfile)
conn = pyodbc.connect(DBfile)

cursor = conn.cursor()
print "Database: " + DBfile + " connected."

#Global variable
oldWKAutoNo = 0
lastWKAutoNo = 0
svrListenPort = 5000

# Real applications authenticate with Evernote using OAuth, but for the
# purpose of exploring the API, you can get a developer token that allows
# you to access your own Evernote account. To get a developer token, visit
# https://sandbox.evernote.com/api/DeveloperToken.action
#
# add a new note into the Evernote Default Notebook
def addNewNote(row):
    # watchkeeper token
    auth_token = "S=s76:U=84f6ef:E=14d7e735e85:C=14626c23289:P=1cd:A=en-devtoken:V=2:H=0616ed9537e7d30bd8bdfe27749c43aa"

    # tongcc8 token
    # auth_token = "S=s2:U=151bd:E=14d7f6cc0c6:C=14627bb94cd:P=1cd:A=en-devtoken:V=2:H=605e2dd576a2610d0888587b051389af"

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

    user_store = client.get_user_store()

    version_ok = user_store.checkVersion(
        "Evernote EDAMTest (Python)",
        UserStoreConstants.EDAM_VERSION_MAJOR,
        UserStoreConstants.EDAM_VERSION_MINOR
    )
    print "Is my Evernote API version up to date? ", str(version_ok)
    print ""
    if not version_ok:
        exit(1)

    note_store = client.get_note_store()

    # List all of the notebooks in the user's account
    notebooks = note_store.listNotebooks()
    print "Found ", len(notebooks), " notebooks:"
    for notebook in notebooks:
        print "  * ", notebook.name

    print
    print "Creating a new note in the default notebook"
    print

    # To create a new note, simply create a new Note object and fill in
    # attributes such as the note's title.
    note = Types.Note()
    #note.title = "Test note from pyWKLog2Evernote.py"
    note.title = row.WKRefNo

    # To include an attachment such as an image in a note, first create a Resource
    # for the attachment. At a minimum, the Resource contains the binary attachment
    # data, an MD5 hash of the binary data, and the attachment MIME type.
    # It can also include attributes such as filename and location.
    image = open('enlogo.png', 'rb').read()
    md5 = hashlib.md5()
    md5.update(image)
    hash = md5.digest()

    data = Types.Data()
    data.size = len(image)
    data.bodyHash = hash
    data.body = image

    resource = Types.Resource()
    resource.mime = 'image/png'
    resource.data = data

    # Now, add the new Resource to the note's list of resources
    note.resources = [resource]

    # To display the Resource as part of the note's content, include an <en-media>
    # tag in the note's ENML content. The en-media tag identifies the corresponding
    # Resource using the MD5 hash.
    hash_hex = binascii.hexlify(hash)

    # check content of row is null
    if row.WKAutoNo is None:
        #error
        print "*** WKAutoNo Error ***"
        return
    if row.WKRefNo is None:
        #error
        print "*** WKRefNo Error ***"
        return
    
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
        
    if row.Site is None:
        Site = 'NIL'
    else:
        Site = row.Site
        
    if row.SubSys is None:
        SubSys = 'NIL'
    else:
        SubSys = row.SubSys

    # check special char & and replace    
    if row.Symptoms is None:
        Symptoms = 'NIL'
    else:
        Symptoms = row.Symptoms
        Symptoms = Symptoms.replace('&', 'and')
        print "string & at Symptoms replaced: " + Symptoms
    
        
    if row.Actions is None:
        Actions = 'NIL'
    else:
        Actions = row.Actions
        Actions = Actions.replace('&', 'and')
        print "string & at Actions replaced: " + Actions
        
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
    note.content += '<span style="font-weight:bold;color:black;">' + 'Actions: ' + '</span>'  + Actions + '<br/>'
    note.content += '<span style="font-weight:bold;color:black;">' + 'Status: ' + '</span>'  + Status + '<br/>'
    note.content += '<span style="font-weight:bold;color:black;">' + 'RefTo: ' + '</span>'  + RefTo + '<br/>'
    note.content += '</en-note>'

    # add tag info
    note.tagNames = [SubSys, Site, Cat]
    
    # Finally, send the new note to Evernote using the createNote method
    # The new Note object that is returned will contain server-generated
    # attributes such as the new note's unique GUID.
    created_note = note_store.createNote(note)

    print "Successfully created a new note with GUID: ", created_note.guid

# update Evernote
def updateEvernote(row):
    print row.WKAutoNo, row.WKRefNo
    addNewNote(row)
    

#init database while startup
def initDatabase():
    global oldWKAutoNo
    
    SQL = "SELECT WKLogMain.[WKAutoNo] FROM WKLogMain ORDER BY WKLogMain.[WKAutoNo] DESC"
    cursor.execute(SQL)
    row = cursor.fetchone()
    if row:
        oldWKAutoNo = row.WKAutoNo
        print 'Found last WKAutoNo: %d' % oldWKAutoNo
        
        
   
# check database
def checkDatabase():
    global oldWKAutoNo
    global lastWKAutoNo
    global cursor

    #SQL = "SELECT WKLogMain.[WKAutoNo] FROM WKLogMain ORDER BY WKLogMain.[WKAutoNo] DESC"
    #SQL = "SELECT * FROM WKLogMain ORDER BY WKAutoNo DESC"
    SQL = "SELECT * FROM WKLogMain WHERE LogDate like '%2014%' ORDER BY WKAutoNo DESC;"
    
    cursor.execute(SQL)
    row = cursor.fetchone()
    
    if row:
        lastWKAutoNo = row.WKAutoNo
        #sys.stdout.write('.')
        #print 'Found last WKAutoNo: %d' % lastWKAutoNo
        
    if lastWKAutoNo == oldWKAutoNo:
        sys.stdout.write('.')
        #print "lastWKAutoNo equal oldWKAutoNo, no need to update"
        return 0
    else:
        print "lastWKAutoNo Not equal oldWKAutoNo, update Evernote"
        updateEvernote(row)
        oldWKAutoNo = lastWKAutoNo
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
        self.lc.start(3)

    def announce(self):
        # 3 sec task to check HMI action
        state = ""
        state = checkDatabase()
        #print state
            
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

initDatabase()
myfactory = MyFactory()
reactor.listenTCP(svrListenPort, myfactory)
reactor.run()
