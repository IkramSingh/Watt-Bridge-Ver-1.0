from openpyxl import Workbook
from openpyxl import load_workbook
import time
import threading
import SwerleinFreq

completedStartNewSequence=0
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
Finished=0
ACVoltsRms=0
UncalFreqy=0
FFTVolts=0
FFTPhase=0
#---------------------------------------------------#
#Temporary variables. To be checked later
DCVRange=0
ActiveRow=0
def continueSequence(wattBridgeGUI,rowNumber,ws,wsRS31Data):
    '''The core of the software. Contains all of the commands and function execution commands that performs
    all of the necessary measurements and calculations.'''
    global DCVRange,WCount,WSign,VCount,VSign,ReadingNumber,ActiveRow,RowNumber,SourceType
    global NumberOfReadings,Finished,ReadingNumber,ACVoltsRms,UncalFreqy,FFTVolts,FFTPhase
    RowNumber=rowNumber
    wattBridgeGUI.WattBridgeEventsLog.AppendText("Initiating Radian \n") #Update event log.
    initialiseRadian() #Run Initialise Radian function. Must add later
    time.sleep(1) #Delay for 1 second
    Finished = 0 #End of process?
    while(Finished==0):
        updateGUI(wattBridgeGUI) #Reupdate variables shown in main GUI.
        WCount=0 #Clear all W and V variables.
        WSign=0
        VCount=0
        VSign=0
        if wattBridgeGUI.SetDMMRangeRefVolts.GetCurrentSelection()==0: #Less than 0.7 V rms
            DCVRange=1
        elif wattBridgeGUI.SetDMMRangeRefVolts.GetCurrentSelection()==1: #Less than 7.0 V rms
            DCVRange=10
        wattBridgeGUI.WattBridgeEventsLog.AppendText("Reading: _ \n") #Update event log.
        ReadingNumber=1
        wattBridgeGUI.WattBridgeEventsLog.AppendText("Applying Power \n") #Update event log.
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
        wattBridgeGUI.WattBridgeEventsLog.AppendText("Finding Dial Settings \n") #Update event log.
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
        wsRS31Data.cell(row=ActiveRow,column=1,value=ActiveRow) #Set Row number cell in "RD31 Data" sheet to Active Row value
        for ReadingsLoop in range(NumberOfReadings):
            ReadingNumber=ReadingsLoop+1
            wattBridgeGUI.WattBridgeEventsLog.AppendText("Doing Reading "+str(ReadingNumber)+" \n") #Update event log.
            ActiveRow=RowNumber
            time.sleep(0.5) #Delay for 0.5 seconds
            #Output to RS232 6 WB with "DV", term.=CR, wait for completion?=1
            #Close RS232 6 WB
            #Output to RS232 6 WB with "A01", term.=CR, wait for completion?=1
            #Close RS232 6 WB
            #Output to RS232 6 WB with "B01", term.=CR, wait for completion?=1
            #Close RS232 6 WB
            time.sleep(0.5) #Delay for 0.5 seconds
            Ch1GateTimeCell = ws.cell(row=ActiveRow,column=10).value #Get Ch1 gate time cell value from Excel file.
            if Ch1GateTimeCell>0.1:
                if wattBridgeGUI.SelectCounter.GetCurrentSelection()==0:
                    #Output to Ag53230A_V with ":SENS:FREQ:GATE:TIME " , Excel link, term.=LF
                    #Output to Ag53230A_V with ":SENS:FUNC 'FREQ 1'", term.=LF
                    #Output to Ag53230A_V with ":INIT", term.=LF
                    print("Counter chosen is 53230A")
                elif wattBridgeGUI.SelectCounter.GetCurrentSelection()==1:
                    #Output to HP3131A_V with ":sens:freq:arm:stop:tim " , Excel link, term.=LF
                    #Output to HP3131A_V with ":func 'freq 1'", term.=LF
                    #Output to HP3131A_V with "init", term.=LF
                    print("Counter chosen is 3131A")
            wattBridgeGUI.WattBridgeEventsLog.AppendText("Collecting Swerlein Measurements \n") #Update event log.
            SwerleinFreq.OnStart() #Create a comunication link between Swerlein and 3458A.
            ACVoltsRms = SwerleinFreq.run() #Obtain Ac volts rms value using Swerleins Algorithm
            ws.cell(row=ActiveRow,column=35+7*(ReadingNumber-1),value=ACVoltsRms) #Set the Ac volts rms value in Excel sheet.
            UncalFreqy = SwerleinFreq.FNFreq() #Obtain the Frequency from 3458A
            ws.cell(row=ActiveRow,column=28,value=UncalFreqy) #Set the exact frequency value in Excel sheet.
            #Execute Set Up FFT Function
            wattBridgeGUI.WattBridgeEventsLog.AppendText("Reference Phase \n") #Update event log.
            #Execute FFT Volts & Phase function
            ws.cell(row=ActiveRow,column=36+7*(ReadingNumber-1),value=FFTVolts) #Set the FFT ref volts value in Excel sheet.
            ws.cell(row=ActiveRow,column=37+7*(ReadingNumber-1),value=FFTPhase) #Set the FFT ref phase value in Excel sheet.
            #Output to RS232 6 WB with "DD", term.=CR, wait for completion?=1
            #Close RS232 6 WB
            #Output to RS232 6 WB with "A33", term.=CR, wait for completion?=1
            #Close RS232 6 WB
            #Output to RS232 6 WB with "B33", term.=CR, wait for completion?=1
            #Close RS232 6 WB
            wattBridgeGUI.WattBridgeEventsLog.AppendText("Detector volts and phase \n") #Update event log.
            #Execute FFT Volts & Phase function
            ws.cell(row=ActiveRow,column=33+7*(ReadingNumber-1),value=FFTVolts) #Set the Det volts value in Excel sheet.
            ws.cell(row=ActiveRow,column=34+7*(ReadingNumber-1),value=FFTPhase) #Set the Det phase value in Excel sheet.
        Finished=1 #For testing purposes. Remove when not needed.
        wattBridgeGUI.WattBridgeEventsLog.AppendText("Completed collecting/measuring Data sequence \n") #Update event log.
        wattBridgeGUI.WattBridgeEventsLog.AppendText("Press 'Save Data' button to save back into original Excel file \n") #Update event log.
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
def updateGUI(wattBridgeGUI):
    '''Updates the values shown in the main GUI.'''
    wattBridgeGUI.CurrentRow.SetValue(RowNumber) #Show current Row in Excel file in main GUI to user.
    print("GUI updated")
def startNewSequence(wattBridgeGUI,ws,wsRS31Data):
    global RowNumber
    '''startNewSequence function is executed when user presses the "Start New Sequence (from "Start Row")".
    Leads onto continueSequence function. Contains 2 threads so that the "continueSequence" and "updateGUI"
    functions are executing simultaneously for the user.'''
    rowNumber = wattBridgeGUI.StartRow.GetValue() #Row number in excel sheet.
    RowNumber=rowNumber 
    t1=threading.Thread(target=updateGUI,args=(wattBridgeGUI,))
    t2=threading.Thread(target=continueSequence,args=(wattBridgeGUI,rowNumber,ws,wsRS31Data,))
    t1.start()
    t2.start()