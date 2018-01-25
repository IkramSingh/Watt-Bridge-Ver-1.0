import WattBridgeGUI 
import Setup
import StartNewSequence
import wx
from openpyxl import Workbook
from openpyxl import load_workbook
import os.path
import os

class WattBridge(WattBridgeGUI.WattBridgeSoftware):
    HP3458A_V=0
    Ag53230A_V=0
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
        global HP3458A_V,Ag53230A_V
        self.WattBridgeEventsLog.AppendText("Setting Up Instruments and Checking Connections...\n")
        HP3458A_V = Setup.setup3458A() #Get HP3458A Visa object.
        HP3458A_V_ID = str(HP3458A_V.query('ID?')) #Check to see if HP3458A has been successfully connected.
        self.WattBridgeEventsLog.AppendText(HP3458A_V_ID)
        Ag53230A_V = Setup.setup53230A() #Get Ag53230A Visa object.
        Ag53230A_V_ID = str(Ag53230A_V.query('*IDN?')) #Check to see if Ag53230A has been successfully connected.
        self.WattBridgeEventsLog.AppendText(Ag53230A_V_ID)
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
    def initialiseCounter():
        if self.SelectCounter.GetCurrentSelection()==0:
            print("0")
#mandatory in wx, create an app, False stands for not deteriction stdin/stdout
#refer manual for details
app = wx.App(False)


#create an object of CalcFrame
frame = WattBridge(None)
#show the frame
frame.Show(True) 

#start the applications
app.MainLoop()