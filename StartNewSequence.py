from openpyxl import Workbook
from openpyxl import load_workbook
import time

#-----Global Variables used by the Watt Bridge Software-----#
flukeError=0
FlukeErrorNumber=0
SetVoltsCell=0
SetAmpsCell=0
SetPhaseCell=0
SetVoltsCell=0
SetFrequencyCell=0
ReadingNumber=0
RowNumber=0
NumberOfReadings=0
DividerRange=0
Shunt=0
CTRatio=0
HEGFreq=0
SourceType=0
CHType=0
CalculationResults=0
RDPhase=0
SampleData=0
WCount=0
WSign=0
VCount=0
VSign=0
#---------------------------------------------------#
#Temporary variables. To be checked later
DCVRange=0
ActiveRow=0
	
def startNewSequence(wattBridgeGUI,eventsLog,ws):
	rowNumber = wattBridgeGUI.StartRow.GetValue() #Row number in excel sheet.
	continueSequence(wattBridgeGUI,eventsLog,rowNumber,ws)

def continueSequence(wattBridgeGUI,eventsLog,rowNumber,ws):
	global DCVRange,WCount,WSign,VCount,VSign,ReadingNumber,ActiveRow,RowNumber
	RowNumber=rowNumber
	eventsLog.AppendText("Initiating Radian \n") #Update event log.
	initialiseRadian() #Run Initialise Radian function. Must add later
	time.sleep(1) #Delay for 1 second
	Finished = 0 #End of process?
	while(Finished==0):
		wattBridgeGUI.CurrentRow.SetValue(RowNumber) #Show current Row in Excel file in main GUI to user.
		WCount=0 #Clear all W and V variables.
		WSign=0
		VCount=0
		VSign=0
		if wattBridgeGUI.SetDMMRangeRefVolts.GetCurrentSelection()==0: #Less than 0.7 V rms
			DCVRange=1
		elif wattBridgeGUI.SetDMMRangeRefVolts.GetCurrentSelection()==1: #Less than 7.0 V rms
			DCVRange=10
		eventsLog.AppendText("Reading: _ \n") #Update event log.
		ReadingNumber=1
		eventsLog.AppendText("Applying Power \n") #Update event log.
		ActiveRow=rowNumber
		Phase123Cell = ws.cell(row=ActiveRow,column=16).value #Obtain phase value from Excel sheet
		if Phase123Cell==123 or Phase123Cell==0:
			RDPhase=0
		else:
			RDPhase=Phase123Cell
		WattsOrVarsCell = str(ws.cell(row=ActiveRow,column=13).value)#Obtain watts/vars value from Excel sheet
		if WattsOrVarsCell[0]=="v": #If it is vars
			#Call Set Radian output pulse with Prog Radian ID,1,1,RD phase
			#Call RD Get Error Message with Prog Radian ID,RD 31 Error Message
			print("v")
		else:
			#Set Radian output pulse with Prog Radian ID,1,0,RD phase
			#Call RD Get Error Message with Prog Radian ID,RD 31 Error Message
			print("w")
		if wattBridgeGUI.ShuntVoltsTest.GetValue()==True: #If user has checked Shunt Volts Test
			#Output to HP3488A_V with "CRESET 4", term.=LF
			time.sleep(1)
			#Output to HP3488A_V with "CLOSE 404", term.=LF
			time.sleep(1)
			eventsLog.AppendText("Shunt Volts Test on \n")
		dateTime = str(time.asctime())
		ws.cell(row=ActiveRow,column=27,value=dateTime) #Set the time and date in Excel sheet.
		ws.cell(row=3,column=42,value=RowNumber) #Set the Row Number in Excel sheet.
		setPower(ws) #Execute Set Power function.
		print(DividerRange)
		print(Shunt)
		print(CTRatio)
		print(HEGFreq)
		Finished=1 #For testing purposes. Remove when not needed.
def initialiseRadian():
	pass
def setPower(ws):
	global DividerRange,Shunt,CTRatio,HEGFreq
	DividerRange = ws.cell(row=ActiveRow,column=6).value #Set the Divider Range cell value from Excel sheet
	Shunt = ws.cell(row=ActiveRow,column=7).value #Set the Shunt cell value from Excel sheet
	CTRatio = ws.cell(row=ActiveRow,column=8).value #Set the CT ratio cell from Excel sheet
	HEGFreq = ws.cell(row=ActiveRow,column=9).value #Set the Set frequency cell from Excel sheet