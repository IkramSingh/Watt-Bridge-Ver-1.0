import WattBridgeGUI 
import Setup
import StartNewSequence
import wx
from openpyxl import Workbook
from openpyxl import load_workbook
import os.path
import os
import time
from rd31direct import *

class WattBridge(WattBridgeGUI.WattBridgeSoftware):
    '''Class that contains all of the functions to operate the WattBridge from the GUI.
    Creates the main GUI window as well as the Events Log window.'''
    def __init__( self, parent ):
        '''Constructor for WattBridge class.'''
        WattBridgeGUI.WattBridgeSoftware.__init__(self,parent)
        self.setupSpreadsheet()
        print os.path.abspath(os.curdir)
    def CheckConnectionsOnButtonClick( self, event ):
        '''Creates communication links between Watt Bridge software and all of the machines
        as well as checking to see if the communication links are successful'''
        self.WattBridgeEventsLog.AppendText("Setting Up Instruments and Checking Connections...\n")
        self.HP3458A_V = Setup.setup3458A() #Get HP3458A Visa object.
        HP3458A_V_ID = str(self.HP3458A_V.query('ID?')) #Check to see if HP3458A has been successfully connected.
        self.WattBridgeEventsLog.AppendText(HP3458A_V_ID)
        self.Ag53230A_V = Setup.setup53230A() #Get Ag53230A Visa object.
        Ag53230A_V_ID = str(self.Ag53230A_V.query('*IDN?')) #Check to see if Ag53230A has been successfully connected.
        self.WattBridgeEventsLog.AppendText(Ag53230A_V_ID)
        self.FLUKE_V = Setup.setup6105A() #Get 6105A Visa Object (Fluke V)
        FLUKE_V_ID = str(self.FLUKE_V.query('*IDN?')) #Check to see if 6105A has been successfully connected
        self.WattBridgeEventsLog.AppendText(FLUKE_V_ID)
        self.rd31 = RD31 () #Get RD31 Object as well as open its Serial port.
        rd31_ID = str(self.rd31.ask(0x02,0,"")) #Check to see if RD31 has been successfully connected
        self.rd31.port.close() #always close port after performing a command on RD31
        self.WattBridgeEventsLog.AppendText(rd31_ID)
        self.HP3478A_V = Setup.setup3478() #Get HP3478A Visa Object
        self.HP3478A_V.write('D23478A DONE.') #Display message on 3478A
        self.WattBridgeEventsLog.AppendText("\nGo and check 3478A for message '3478A DONE' being displayed. System will pause for 12 seconds...\n")
        time.sleep(12)
        self.HP3478A_V.write('D1') #Remove display to default
        self.RS232_6_WB = Setup.setupWB() #Get Watt Bridge Visa object
        self.RS232_6_WB.write("W0721\r")
        time.sleep(3)
        self.RS232_6_WB.write("V0127\r")
        self.WattBridgeEventsLog.AppendText("Check and see if Watt Bridge has been set to W0721 & V0127. System will pause for 12 seconds...\n")
        time.sleep(12)
        StartNewSequence.setInstruments(self.HP3458A_V,self.Ag53230A_V,self.FLUKE_V,self.rd31,self.HP3478A_V,self.RS232_6_WB) #Save Instrument objects in StartNewSequence class
        self.initialiseCounter() #Initialise the Ag53230A_V Frequency Counter
        self.WattBridgeEventsLog.AppendText("If 4 instruments ID are shown as well as 3478A and Watt Bridge being set correcly, the all connections are successful. \n")
    def WattBridgeSoftwareOnClose( self, event ):
        '''Closes all of the windows as well as Exists the Watt Bridge Software.'''
        self.Destroy()
    def AboutOnMenuSelection( self, event ):
        '''Displays information about the software to the user in a seperate window.'''
        WattBridgeGUI.About(None).Show(True) #Displays information about the software.
    def saveSpreadsheet(self):
        '''Saves the Excel file that was initially setup in the setupSpreadsheet function.'''
        self.WattBridgeEventsLog.AppendText("Saving Spreadsheet...\n") #Inform the user.
        self.wb.template=False #Make sure Excel file is saved as document not template
        self.wb.save(self.filename) #Save file with same name
        self.WattBridgeEventsLog.AppendText("Spreadsheet saved successfully\n") #Inform the user.
    def SaveDataOnButtonClick( self, event ):
        '''Accesses the saveSpreadsheet function when user presses "Save Data" button in the main GUI'''
        self.saveSpreadsheet()
    def StartNewSequenceOnButtonClick( self, event ):
        '''Start a new sequence when user presses the "Start New Sequence (from "Start Row")" button.'''
        StartNewSequence.startNewSequence(frame,self.ws,self.wsRS31Data) #Start of startNewSequence in StartNewSequence.py
    def setupSpreadsheet(self):
        '''Creates a link between the software and Excel sheet.'''
        self.WattBridgeEventsLog.AppendText("Setting Up Spreadsheet and creating user interface...\n") #Inform the user.
        print("Setting Up Spreadsheet and creating user interface...\n")
        self.filename="Book1.xlsx" #Excel filename
        self.wb=load_workbook(self.filename) #Open Excel file
        self.ws=self.wb.active #Make it active to work in
        self.wsRS31Data = self.wb.worksheets[1] #Second worksheet titled "RD31 Data"
        print("Spreadsheet setup successfully \n")
        self.WattBridgeEventsLog.AppendText("Spreadsheet setup successfully \n")#Inform the user.
    def initialiseCounter(self):
        '''Initialises the 53230A Frequency Counter.'''
        if self.SelectCounter.GetCurrentSelection()==0:
           self.Ag53230A_V.write('*rst;*cls;*sre 0;*ese 0;:stat:pres;:INP1:FILT:LPAS:STAT 1') 
           self.Ag53230A_V.write(':sens:freq:gate:sour time')
           self.Ag53230A_V.write(':inp1:rang 50')
           self.Ag53230A_V.write(':sens:rosc:sour int')
           if self.Channel1Filter.GetCurrentSelection()==1:
               self.Ag53230A_V.write(';:INP1:FILT:LPAS:STAT 1')
           else:
               self.Ag53230A_V.write(';:inp1:filt:lpas:stat 0')
           if self.Ch1TrigLevel.GetCurrentSelection()==0:
               self.Ag53230A_V.write(':inp1:coup dc;:inp1:slope pos;:inp1:lev:abs 3V;')
           elif self.Ch1TrigLevel.GetCurrentSelection()==1:
               self.Ag53230A_V.write(':inp1:coup dc;:inp1:slope pos;:inp1:lev:abs 6V;')
           if self.Ch2TrigLevel.GetCurrentSelection()==0:
               self.Ag53230A_V.write(':inp2:coup dc;:inp2:slope pos;:inp2:lev:abs 3V;')
           elif self.Ch2TrigLevel.GetCurrentSelection()==1:
               self.Ag53230A_V.write(':inp2:coup dc;:inp2:slope pos;:inp2:lev:abs 6V;')
           self.Ag53230A_V.write('DISP:TEXT "Counter Initialised"')
           time.sleep(3) #Delay for 3 seconds
           self.Ag53230A_V.write('DISP:TEXT:CLE')
#mandatory in wx, create an app, False stands for not deteriction stdin/stdout
#refer manual for details
app = wx.App(False)


#create an object of CalcFrame
frame = WattBridge(None)
#show the frame
frame.Show(True) 

#start the applications
app.MainLoop()