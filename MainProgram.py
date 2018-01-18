import WattBridgeGUI 
import Setup
import StartNewSequence
import wx
from openpyxl import Workbook
from openpyxl import load_workbook
import os.path
import os

class WattBridge(WattBridgeGUI.WattBridgeSoftware):
	'''Contains all of the functions to operate the WattBridge from the GUI.
	Creates the main GUI window as well as the Events Log window.'''
	def __init__( self, parent ):
		WattBridgeGUI.WattBridgeSoftware.__init__(self,parent)
		self.eventsLog = WattBridgeGUI.WattBridgeEventsLog(None)
		self.eventsLog.Show(True)
		self.setupSpreadsheet()
		print os.path.abspath(os.curdir)
	def WattBridgeSoftwareOnClose( self, event ):
		'''Closes all of the windows as well as Exists the Watt Bridge Software.'''
		self.Destroy()
		self.eventsLog.Destroy()

	def AboutOnMenuSelection( self, event ):
		'''Displays information about the software to the user in a seperate window.'''
		WattBridgeGUI.About(None).Show(True) #Displays information about the software.
	
	def StartNewSequenceOnButtonClick( self, event ):
		'''Start a new sequence when user presses the "Start New Sequence (from "Start Row")" button.'''
		StartNewSequence.startNewSequence(frame,self.eventsLog.EventsLog,self.ws) #Start of startNewSequence in StartNewSequence.py
		self.eventsLog.EventsLog.AppendText("Saving Spreadsheet...\n") #Inform the user.
		self.wb.template=False #Make sure Excel file is saved as document not template
		self.wb.save(self.filename) #Save file with same name
		self.eventsLog.EventsLog.AppendText("Spreadsheet saved successfully\n") #Inform the user.
		
	def setupSpreadsheet(self):
		'''Creates a link between the software and Excel sheet.'''
		self.eventsLog.EventsLog.AppendText("Setting Up Spreadsheet and creating user interface...\n") #Inform the user.
		#Location of Excel file
		self.filename="Book1.xlsx"
		self.wb=load_workbook(self.filename) #Open Excel file
		self.ws=self.wb.active #Make it active to work in
		self.eventsLog.EventsLog.AppendText("Spreadsheet setup successfully \n")#Inform the user.
	
#mandatory in wx, create an app, False stands for not deteriction stdin/stdout
#refer manual for details
app = wx.App(False)

#create an object of CalcFrame
frame = WattBridge(None)
#show the frame
frame.Show(True)
#start the applications
app.MainLoop()