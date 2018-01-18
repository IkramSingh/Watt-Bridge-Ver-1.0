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
	'''startNewSequence function is executed when user presses the "Start New Sequence (from "Start Row")".
	Leads onto continueSequence function.'''
	rowNumber = wattBridgeGUI.StartRow.GetValue() #Row number in excel sheet.
	continueSequence(wattBridgeGUI,eventsLog,rowNumber,ws)

def continueSequence(wattBridgeGUI,eventsLog,rowNumber,ws):
	'''The core of the software. Contains all of the commands and function execution commands that performs
	all of the necessary measurements and calculations.'''
	global DCVRange,WCount,WSign,VCount,VSign,ReadingNumber,ActiveRow,RowNumber,SourceType
	global NumberOfReadings
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
			print("WattsOrVars: vars")
		else:
			#Set Radian output pulse with Prog Radian ID,1,0,RD phase
			#Call RD Get Error Message with Prog Radian ID,RD 31 Error Message
			print("WattsOrVars: watt")
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
		print("DividerRange: "+str(DividerRange))
		print("Shunt: "+str(Shunt))
		print("CTRatio: "+str(CTRatio))
		print("HEGFreq: "+str(HEGFreq))
		SourceType = ws.cell(row=ActiveRow,column=2).value #Get the Source Type from Excel sheet.
		print("SourceType: "+str(SourceType))
		if SourceType=="FLUKE" or SourceType=="FLUHIGH":
			print("FLUKE")
			#Execute Power Fluke
		elif SourceType=="CH":
			print("CH5500")
			#Execute CH5500
		elif SourceType=="HEG":
			#Execute PL10
			print("PL10")
		else:
			print("SourceType selected in Excel file doesnt exist.")
		ActiveRow=7 #Set ActiveRow to 7.
		eventsLog.AppendText("Finding Dial Settings \n") #Update event log.
		#Execute Find Dial Settings
		#Execute Refine Dial Settings
		ActiveRow=RowNumber
		#Execute Load Dial Settings
		NumberOfReadings = ws.cell(row=ActiveRow,column=11).value #Get the Number of Readings value from Excel sheet.
		print("NumberOfReadings: " + str(NumberOfReadings))
		if wattBridgeGUI.ShuntVoltsTest.GetValue()==True: #If user has checked Shunt Volts Test
			#Execute Test
			print("Test")
		#Output to HP3478A_V with "F4RAN5Z1", term.=LF
		time.sleep(3) #Delay for 3 seconds
		#Output to HP3478A_V with "T3", term.=LF
		#Enter from HP3478A_V up to 256 bytes, stop on EOS=LF
		#Set remote Excel link item Temperature Cell to HP3478A_V
		
		Finished=1 #For testing purposes. Remove when not needed.
def initialiseRadian():
	pass
def setPower(ws):
	'''Obtains the DividerRange,Shunt,CTRatio,HEGFreq variables from Excel file as well as outputting 
	various commands to the "RS232 6 WB".'''
	global DividerRange,Shunt,CTRatio,HEGFreq
	DividerRange = ws.cell(row=ActiveRow,column=6).value #Get the Divider Range cell value from Excel sheet
	Shunt = ws.cell(row=ActiveRow,column=7).value #Get the Shunt cell value from Excel sheet
	CTRatio = ws.cell(row=ActiveRow,column=8).value #Get the CT ratio cell from Excel sheet
	HEGFreq = ws.cell(row=ActiveRow,column=9).value #Get the Set frequency cell from Excel sheet
	# Open RS232 6 WB
	# Set mode of RS232 6 WB baud rate=9600, parity="N", bits=8, stop bits=1
	# Close RS232 6 WB
	# Output to RS232 6 WB with "DV", term.=CR, wait for completion?=1
	# Close RS232 6 WB
	# Output to RS232 6 WB with "DV", term.=CR, wait for completion?=1
	# Close RS232 6 WB
	# Output to RS232 6 WB with "W0000", term.=CR, wait for completion?=1
	# Close RS232 6 WB
	# Output to RS232 6 WB with "V0000", term.=CR, wait for completion?=1
	# Close RS232 6 WB
	# Output to RS232 6 WB with "A01", term.=CR, wait for completion?=1
	# Close RS232 6 WB
	# Output to RS232 6 WB with "B01", term.=CR, wait for completion?=1
	# Close RS232 6 WB
	if DividerRange==60:
		#Output to RS232 6 WB with "R" , "060", term.=CR, wait for completion?=1
		print("DividerRange is 60")
	else:
		#Output to RS232 6 WB with "R" , Divider Range, term.=CR, wait for completion?=1
		print("DividerRange is not 60")
	# Close RS232 6 WB
	# Output to RS232 6 WB with "WP-", term.=CR, wait for completion?=1
	# Close RS232 6 WB
	# Output to RS232 6 WB with "VP-", term.=CR, wait for completion?=1
	# Close RS232 6 WB
	time.sleep(0.5) #Delay for 0.5 seconds