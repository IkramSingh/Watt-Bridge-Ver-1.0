from openpyxl import Workbook
from openpyxl import load_workbook
import time
import threading
#import SwerleinFreq

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
def getExcelColumn(column):
    if column==35:
        return 'AI'
    if column==36:
        return 'AJ'
    if column==37:
        return 'AK'
    if column==33:
        return 'AG'
    if column==34:
        return 'AH'
    if column==42:
        return 'AP'
    if column==43:
        return 'AQ'
    if column==44:
        return 'AR'
    if column==40:
        return 'AN'
    if column==41:
        return 'AO'
    if column==38:
        return 'AL'
    if column==45:
        return 'AS'
    if column==39:
        return 'AM'
    if column==46:
        return 'AT'
    if column==47:
        return 'AU'
    if column==48:
        return 'AV'
    if column==49:
        return 'AW'
    if column==50:
        return 'AX'
    if column==51:
        return 'AY'
    if column==52:
        return 'AZ'
    if column==53:
        return 'BA'
    if column==54:
        return 'BB'
    if column==55:
        return 'BC'
    if column==56:
        return 'BD'
    if column==57:
        return 'BE'
    if column==58:
        return 'BF'
    if column==59:
        return 'BG'
    if column==60:
        return 'BH'
    if column==61:
        return 'BI'
    if column==62:
        return 'BJ'
    if column==63:
        return 'BK'
    if column==64:
        return 'BL'
    if column==65:
        return 'BM'
    if column==66:
        return 'BN'
    if column==67:
        return 'BO'
def continueSequence(wattBridgeGUI,rowNumber,ws,wsRS31Data):
    '''The core of the software. Contains all of the commands and function execution commands that performs
    all of the necessary measurements and calculations.'''
    global DCVRange,WCount,WSign,VCount,VSign,ReadingNumber,ActiveRow,RowNumber,SourceType
    global NumberOfReadings,Finished,ReadingNumber,ACVoltsRms,UncalFreqy,FFTVolts,FFTPhase
    RowNumber=int(rowNumber)
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
        ActiveRow=RowNumber
        Phase123Cell = ws['P'+str(ActiveRow)].value #Obtain phase value from Excel sheet
        if Phase123Cell==123 or Phase123Cell==0:
            RDPhase=0
        else:
            RDPhase=Phase123Cell
        WattsOrVarsCell = str(ws['M'+str(ActiveRow)].value)#Obtain watts/vars value from Excel sheet
        if WattsOrVarsCell[0]=="v": #If it is vars
            #Call Set Radian output pulse with Prog Radian ID,1,1,RD phase
            #Call RD Get Error Message with Prog Radian ID,RD 31 Error Message
            print("WattsOrVars: vars")
        elif WattsOrVarsCell[0]=="w": #If it is watts
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
        ws['AA'+str(ActiveRow)]=dateTime #Set the time and date in Excel sheet.
        ws['AP3']=RowNumber #Set the Row Number in Excel sheet.
        setPower(ws) #Execute Set Power function.
        print("DividerRange: "+str(DividerRange))
        print("Shunt: "+str(Shunt))
        print("CTRatio: "+str(CTRatio))
        print("HEGFreq: "+str(HEGFreq))
        SourceType = ws['B'+str(ActiveRow)].value #Get the Source Type from Excel sheet.
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
        NumberOfReadings = ws['K'+str(ActiveRow)].value #Get the Number of Readings value from Excel sheet.
        print("NumberOfReadings: " + str(NumberOfReadings))
        if wattBridgeGUI.ShuntVoltsTest.GetValue()==True: #If user has checked Shunt Volts Test
            #Execute Test
            print("Test")
        #Output to HP3478A_V with "F4RAN5Z1", term.=LF
        time.sleep(3) #Delay for 3 seconds
        #Output to HP3478A_V with "T3", term.=LF
        #Enter from HP3478A_V up to 256 bytes, stop on EOS=LF
        #Set remote Excel link item Temperature Cell to HP3478A_V
        wsRS31Data['A'+str(ActiveRow)]=ActiveRow #Set Row number cell in "RD31 Data" sheet to Active Row value
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
            Ch1GateTimeCell = ws['J'+str(ActiveRow)].value #Get Ch1 gate time cell value from Excel file.
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
            #SwerleinFreq.OnStart() #Create a comunication link between Swerlein and 3458A.
            ACVoltsRms = "Testing"#SwerleinFreq.run() #Obtain Ac volts rms value using Swerleins Algorithm
            ws[getExcelColumn(35+7*(ReadingNumber-1))+str(ActiveRow)]=ACVoltsRms #Set the Ac volts rms value in Excel sheet.
            UncalFreqy = "Testing"#SwerleinFreq.FNFreq() #Obtain the Frequency from 3458A
            ws['AB'+str(ActiveRow)]=UncalFreqy #Set the exact frequency value in Excel sheet.
            #Execute Set Up FFT Function
            wattBridgeGUI.WattBridgeEventsLog.AppendText("Reference Phase \n") #Update event log.
            #Execute FFT Volts & Phase function
            FFTVolts="Testing"
            FFTPhase="Testing"
            ws[getExcelColumn(36+7*(ReadingNumber-1))+str(ActiveRow)]=FFTVolts #Set the FFT ref volts value in Excel sheet.
            ws[getExcelColumn(37+7*(ReadingNumber-1))+str(ActiveRow)]=FFTPhase #Set the FFT ref phase value in Excel sheet.
            #Output to RS232 6 WB with "DD", term.=CR, wait for completion?=1
            #Close RS232 6 WB
            #Output to RS232 6 WB with "A33", term.=CR, wait for completion?=1
            #Close RS232 6 WB
            #Output to RS232 6 WB with "B33", term.=CR, wait for completion?=1
            #Close RS232 6 WB
            wattBridgeGUI.WattBridgeEventsLog.AppendText("Detector volts and phase \n") #Update event log.
            #Execute FFT Volts & Phase function
            ws[getExcelColumn(33+7*(ReadingNumber-1))+str(ActiveRow)]=FFTVolts #Set the Det volts value in Excel sheet.
            ws[getExcelColumn(34+7*(ReadingNumber-1))+str(ActiveRow)]=FFTPhase #Set the Det phase value in Excel sheet.
            wattBridgeGUI.WattBridgeEventsLog.AppendText("Read RD31 \n") #Update event log.
            #Execute Read Radian2 function.
            wattBridgeGUI.WattBridgeEventsLog.AppendText("Counter \n") #Update event log.
            GateTime = Ch1GateTimeCell
            if GateTime>0.1:
                if wattBridgeGUI.SelectCounter.GetCurrentSelection()==1:
                    #Output to HP3131A_V with "fetc?", term.=LF
                    time.sleep(0.25) #Delay for 0.25 seconds
                    #Enter from HP3131A_V up to 256 bytes, stop on EOS=LF
                    HP3131A_V="Testing"
                    ws[getExcelColumn(38+7*(ReadingNumber-1))+str(ActiveRow)]=HP3131A_V
                elif wattBridgeGUI.SelectCounter.GetCurrentSelection()==0:
                    #Output to Ag53230A_V with "fetc?", term.=LF
                    time.sleep(0.25) #Delay for 0.25 seconds
                    #Enter from Ag53230A_V up to 256 bytes, stop on EOS=LF
                    Ag53230A_V="Testing"
                    ws[getExcelColumn(38+7*(ReadingNumber-1))+str(ActiveRow)]=Ag53230A_V
                if wattBridgeGUI.CounterChannel.GetCurrentSelection()==0:
                    if wattBridgeGUI.SelectCounter.GetCurrentSelection()==1:
                        #Output to HP3131A_V with ":sens:freq:arm:stop:tim " , Excel link, term.=LF
                        #Output to HP3131A_V with ":func 'freq 2'", term.=LF
                        #Output to HP3131A_V with ":read?", term.=LF
                        time.sleep(2) #Delay for 2 seconds
                        #Enter from HP3131A_V up to 256 bytes, stop on EOS=LF
                        HP3131A_V="Testing"
                        ws[getExcelColumn(39+7*(ReadingNumber-1))+str(ActiveRow)]=HP3131A_V
                    elif wattBridgeGUI.SelectCounter.GetCurrentSelection()==0:
                        #Output to Ag53230A_V with ":SENS:FREQ:GATE:TIME " , Excel link, term.=LF
                        #Output to Ag53230A_V with ":SENS:FUNC 'FREQ 2'", term.=LF
                        #Output to Ag53230A_V with ":read?", term.=LF
                        time.sleep(2) #Delay for 2 seconds
                        #Enter from Ag53230A_V up to 256 bytes, stop on EOS=LF
                        Ag53230A_V="Testing"
                        ws[getExcelColumn(39+7*(ReadingNumber-1))+str(ActiveRow)]=Ag53230A_V
                else:
                    if WattsOrVarsCell[0]=="v": #If it is vars
                        #Call RD Get Instantaneous Data with Prog Radian ID,0,4,RD31 Total
                        #Call RD Get Error Message with Prog Radian ID,RD 31 Error Message
                        print("WattsOrVars: vars")
                    elif WattsOrVarsCell[0]=="w": #If it is watts
                        #RD Get Instantaneous Data with Prog Radian ID,0,2,RD31 Total
                        #RD Get Error Message with Prog Radian ID,RD 31 Error Message
                        print("WattsOrVars: watt")
                    RD31Total="Testing"
                    ws[getExcelColumn(39+7*(ReadingNumber-1))+str(ActiveRow)]=RD31Total
        #Execute Read Radian all Data
        if SourceType=="CH":
            #Output to CH5500_V with "S", term.=LF
            #Output to CH5050_V with "S", term.=LF
            print("SourceType = CH")
        elif SourceType=="HEG":
            #Output to PL10A_V with ":SOUR:OPER:STOP", term.=LF
            print("SourceType = HEG")
        elif SourceType=="FLUKE" or SourceType=="FLUHIGH":
           #Output to FLUKE_V with "OUTP:STAT OFF", term.=LF 
           print("SourceType = FLUKE or SourceType = FLUHIGH")
        #Output to RS232 6 WB with "DV", term.=CR, wait for completion?=1
        #Close RS232 6 WB
        #Output to RS232 6 WB with "A01", term.=CR, wait for completion?=1
        #Close RS232 6 WB
        #Output to RS232 6 WB with "B01", term.=CR, wait for completion?=1
        #Close RS232 6 WB
        time.sleep(1) #Delay for 1 second
        #Execute Paste Results function
        time.sleep(1) #Delay for 1 second
        RowNumber=RowNumber+1 #Increment to the next Row in excel sheet.
        ActiveRow = RowNumber
        if ws['A'+str(ActiveRow)].value==0: #Once reached end of Excel sheet.
            Finished=1
    wattBridgeGUI.WattBridgeEventsLog.AppendText("Completed collecting/measuring Data sequence \n") #Update event log.
    wattBridgeGUI.WattBridgeEventsLog.AppendText("Press 'Save Data' button to save back into original Excel file \n") #Update event log.
def initialiseRadian():
	pass
def setPower(ws):
    '''Obtains the DividerRange,Shunt,CTRatio,HEGFreq variables from Excel file as well as outputting 
    various commands to the "RS232 6 WB".'''
    global DividerRange,Shunt,CTRatio,HEGFreq
    DividerRange = ws['F'+str(ActiveRow)].value #Get the Divider Range cell value from Excel sheet
    Shunt = ws['G'+str(ActiveRow)].value #Get the Shunt cell value from Excel sheet
    CTRatio = ws['H'+str(ActiveRow)].value #Get the CT ratio cell from Excel sheet
    HEGFreq = ws['I'+str(ActiveRow)].value #Get the Set frequency cell from Excel sheet
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
    wattBridgeGUI.CurrentRow.SetValue(str(RowNumber)) #Show current Row in Excel file in main GUI to user.
    print("GUI updated")
def startNewSequence(wattBridgeGUI,ws,wsRS31Data):
    global RowNumber
    '''startNewSequence function is executed when user presses the "Start New Sequence (from "Start Row")".
    Leads onto continueSequence function. Contains 2 threads so that the "continueSequence" and "updateGUI"
    functions are executing simultaneously for the user.'''
    rowNumber = wattBridgeGUI.StartRow.GetValue() #Row number in excel sheet.
    RowNumber = rowNumber
    t1=threading.Thread(target=updateGUI,args=(wattBridgeGUI,))
    t2=threading.Thread(target=continueSequence,args=(wattBridgeGUI,rowNumber,ws,wsRS31Data,))
    t1.start()
    t2.start()
