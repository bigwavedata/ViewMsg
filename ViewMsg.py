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
    assert converter.isOleFile("Example Message Subject.msg")

def doConversion(name):
    checkFile(name)
    ole = converter.OleFileIO(name)
    #data = ole.open('__nameid_version1.0')
    print ole.listdir(True, True)
    meta= ole.get_metadata()
    meta.dump()

# copy and import easyGUI in order to have the textbox 

    
    
if __name__ == '__main__':
    fileDialog = False
    fileLocation = ''
    #filename = "file1.msg" 
    #doConversion(filename)
    if len(sys.argv) <= 1 :
        fileDialog = True
    elif len(sys.argv) == 2 :
        # For simplicity just take the first argument as file location
        fileLocation = sys.argv[1]
    if fileDialog:
        root = Tkinter.Tk()
        root.withdraw()
        fileLocation = tkFileDialog.askopenfilename()
    
    # fileLocation is set by this point
    message = Message(fileLocation)
    
    print "the subject:"
    print "Value", message.subject
    print "Header",message.header
    print "the message:"
    print message.body
    GUI.textbox(message.subject, "", message.body, 0)