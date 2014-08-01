'''
Created on Jul 26, 2014

@author: Oliver
'''

import OleFileIO_PL as converter
import sys
import Tkinter, tkFileDialog
from ExtractMsg import Message
import easygui as GUI



def checkFile(name):  
    try:  
        if not converter.isOleFile(name):
            sys.exit()       
    except:
        print "ERROR: %s is not a *.msg format file" % name
        sys.exit()
        
def getSubject(message):
    
    Date=''
    From=''
    To=''
    Subject=''
    
    if message.date:
        Date="Date: " + message.date
    if message.subject:
        Subject="Subject: " + message.subject
    if message.sender:
        From="From: " + message.sender
    if message.to:
        To="To: " + message.to
    
    msgString = "%s\n%s\n%s\n%s\n" % (Date,Subject,From,To)        
    
    return msgString

def doConversion(name):
    checkFile(name)
    ole = converter.OleFileIO(name)
    print ole.listdir(True, True)
    meta= ole.get_metadata()
    meta.dump()
    
if __name__ == '__main__':
    fileDialog = False
    fileLocation = ''
    
    if len(sys.argv) <= 1 :
        fileDialog = True
    elif len(sys.argv) == 2 :
        # For simplicity just take the first argument as file location
        fileLocation = sys.argv[1]
    if fileDialog:
        root = Tkinter.Tk()
        root.withdraw()
        fileLocation = tkFileDialog.askopenfilename()
    
    checkFile(fileLocation)
    
    # fileLocation is set by this point
    message = Message(fileLocation)
    
    subject = getSubject(message)
    
    GUI.textbox(subject, "", message.body, 0)