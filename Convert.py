'''
Created on Jul 11, 2014

@author: 
'''
import OleFileIO_PL as converter



def checkFile(name):
    print "hello"
    assert converter.isOleFile("file1.msg")

def doConversion(name):
    checkFile(name)
    ole = converter.OleFileIO(name)
    #data = ole.open('__nameid_version1.0')
    print ole.listdir(True, True)
    meta= ole.get_metadata()
    meta.dump()
    
    
if __name__ == '__main__':
    filename = "file1.msg"
    doConversion(filename)
    