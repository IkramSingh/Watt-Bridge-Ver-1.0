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
VRangeHigh=0
DCVoltageOffset=0
DCCurrentOffset=0
HighCurrentRange=0
IRangeLow=0
IRangeHigh=0
Chanel=0
SetVoltsPhase=0
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
def powerFluke(wattBridgeGUI,ws):
    global SetVoltsCell,SetPhaseCell,SetAmpsCell,SetFrequencyCell,VRangeHigh,DCVoltageOffset
    global HighCurrentRange,DCCurrentOffset,IRangeLow,IRangeHigh,Chanel,SetVoltsPhase,FlukeErrorNumber
    #Output to FLUKE_V with "*CLS", term.=LF
    #Output to FLUKE_V with "*RST", term.=LF
    #Output to FLUKE_V with "OUTP:SENS 0", term.=LF
    if wattBridgeGUI.FlukeRamp.GetValue()<2:
        wattBridgeGUI.WattBridgeEventsLog.AppendText("Error message: Ramp time less than 2 seconds \n") #Update event log.
    #Output to FLUKE_V with "OUTP:RAMP:TIME " , Fluke Ramp (s), term.=LF
    SetVoltsCell = ws['D'+str(ActiveRow)].value #Obtain set voltage value from Excel Sheet
    SetPhaseCell = ws['E'+str(ActiveRow)].value #Obtain set phase value from Excel Sheet
    #Execute Set Phases function
    SetAmpsCell = ws['C'+str(ActiveRow)].value #Obtain set amps value from Excel Sheet
    SetFrequencyCell = ws['I'+str(ActiveRow)].value #Obtain set frequency value from Excel Sheet
    if SetVoltsCell>1008:
        wattBridgeGUI.WattBridgeEventsLog.AppendText("Cause error: Voltage value out of range \n") #Update event log.
    if SetVoltsCell>250:
        wattBridgeGUI.WattBridgeEventsLog.AppendText("Cause error: Voltage selected is above 250V \n") #Update event log.
    if SetAmpsCell>120:
        wattBridgeGUI.WattBridgeEventsLog.AppendText("Cause error: Current value out of range \n") #Update event log.
    #Specific cases for SetVoltsCell
    if 0 <= SetVoltsCell <= 22.9:
        VRangeHigh=23
        DCVoltageOffset=-0.00081
    elif 22.9 <= SetVoltsCell <= 44.9:
        VRangeHigh=90
        DCVoltageOffset=0.00142
    elif 44.9 <= SetVoltsCell <= 89.9:
        VRangeHigh=45
        DCVoltageOffset=-0.00637
    elif 89.9 <= SetVoltsCell <= 179.9:
        VRangeHigh=180
        DCVoltageOffset=-0.0121
    elif 179.9 <= SetVoltsCell <= 359.9:
        VRangeHigh=360
        DCVoltageOffset=-0.01784
    elif 359.9 <= SetVoltsCell <= 649.9:
        VRangeHigh=650
        DCVoltageOffset=-0.02082
    elif 649.9 <= SetVoltsCell <= 1008:
        VRangeHigh=1008
        DCVoltageOffset=-0.01602
    if SourceType=="FLUHIGH":
        if 0 <= SetAmpsCell <= 1.999:
            HighCurrentRange=2
            DCCurrentOffset=-0.000173
        elif 1.999 <= SetAmpsCell <= 19.99:
            HighCurrentRange=20
            DCCurrentOffset=-0.00195
        elif 19.99 <= SetAmpsCell <= 120:
            HighCurrentRange=120
            DCCurrentOffset=0.006345
    else:
        if 0 <= SetAmpsCell <= 0.249:
            IRangeLow=0.05
            IRangeHigh=0.25
            DCCurrentOffset=1.785E-05
        elif 0.249 <= SetAmpsCell <= 0.499:
            IRangeLow=0.1
            IRangeHigh=0.5
            DCCurrentOffset=4.186E-05
        elif 0.499 <= SetAmpsCell <= 0.999:
            IRangeLow=0.05
            IRangeHigh=1
            DCCurrentOffset=8.285E-05
        elif 0.999 <= SetAmpsCell <= 1.999:
            IRangeLow=0.2
            IRangeHigh=2
            DCCurrentOffset=0.0001459
        elif 1.999 <= SetAmpsCell <= 4.99:
            IRangeLow=0.5
            IRangeHigh=5
            DCCurrentOffset=0.00042
        elif 4.99 <= SetAmpsCell <= 9.99:
            IRangeLow=1
            IRangeHigh=10
            DCCurrentOffset=0.0009386
        elif 9.99 <= SetAmpsCell <= 21:
            IRangeLow=2
            IRangeHigh=21
            DCCurrentOffset=0.002118
    if wattBridgeGUI.SourceCh1.GetValue()==True:
        Chanel = 1
        SetPhaseCell=SetPhaseCell
        SetVoltsPhase = 0
        #Execute Setup Chanel function
    if wattBridgeGUI.SourceCh2.GetValue()==True:
        Chanel = 2
        SetPhaseCell=SetPhaseCell
        SetVoltsPhase = -120
        #Execute Setup Chanel function 
    if wattBridgeGUI.SourceCh3.GetValue()==True:
        Chanel = 3
        SetPhaseCell=SetPhaseCell
        SetVoltsPhase = 120
        #Execute Setup Chanel function 
    FlukeErrorNumber=1
    while FlukeErrorNumber!=0:
        #Output to FLUKE_V with "SOUR:FREQ " , Set frequency cell, term.=LF
        #Output to FLUKE_V with "SYST:ERR?", term.=LF
        time.sleep(0.5) #Delay for 0.5 seconds
        #Enter from FLUKE_V(1) up to 512 bytes, stop on EOS=LF
        #Store in Fluke Error from FLUKE_V
        #Calculate Vector Index with v=Fluke Error i=0
        #Store in Fluke Error number from Vector Index
        #If/Then Fluke error with x=Fluke Error number
        #Calculate Vector Index with v=Fluke Error i=1
        #Cause error General Error code=20010, text=Vector Index
        #End If Fluke error
    #Output to FLUKE_V with "OUTP:STAT ON", term.=LF
    time.sleep(wattBridgeGUI.FlukeRamp.GetValue()) #Delay for Fluke Ramp (s) seconds
    #Output to FLUKE_V with "OUTP:STAT?", term.=LF
    time.sleep(0.5) #Delay for 0.5 seconds
    #Enter from FLUKE_V up to 256 bytes, stop on EOS=LF
    #Store in Is Fluke Off from FLUKE_V
    #If/Then Fluke not On with x=Is Fluke Off
    #Cause error Power Not On code=20006, text="Power output unsuccessful. Check connections."
    #End If Fluke not On
    time.sleep(30) #Delay for 30 seconds
def powerCH5500(ws):
    global CHType
    CHType = ws['B7'].value #Obtain CHType
    #Clear CH5500_V
    if CHType<99:
        #Clear CH5050_V
        #Output to CH5050_V with "S" , "I", term.=LF
    #Output to CH5500_V with "S", term.=LF
    SetVoltsCell = ws['D'+str(ActiveRow)].value #Obtain set voltage value from Excel Sheet
    SetVolts=0
    if 9 <CHType< 12:
        SetVolts = SetVoltsCell/2.497
        #Output to CH5500_V with "O" , 0, term.=LF
    else:
        SetVolts = SetVoltsCell/1
        #Output to CH5500_V with "O" , 180, term.=LF
    #Output to CH5500_V with "R" , Set volts, term.=LF
    SetAmpsCell = ws['C'+str(ActiveRow)].value #Obtain set amps value from Excel Sheet
    if SetAmpsCell<0:
        #Output to CH5500_V with "O" , 0, term.=LF
    Absolutevalue = abs(SetAmpsCell)
    #Output to CH5500_V with "V" , Absolute value, term.=LF
    SetPhaseCell = ws['E'+str(ActiveRow)].value #Obtain set phase value from Excel Sheet
    #Output to CH5500_V with "P" , SetPhaseCell, term.=LF
    SetFrequencyCell = ws['I'+str(ActiveRow)].value #Obtain set frequency value from Excel Sheet
    #Output to CH5500_V with "F" , SetFrequencyCell, term.=LF
    #Output to CH5500_V with "N", term.=LF
    if CHType<99:
        #Output to CH5050_V with "N", term.=LF
        #Output to CH5500_V with "N", term.=LF
        #Output to CH5050_V with "O", term.=LF
    time.sleep(60) #Delay for 60 seconds

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
            powerFluke(wattBridgeGUI,ws) #Execute powerFluke function.
        elif SourceType=="CH":
            print("CH5500")
            powerCH5500(ws) #Execute CH5500 function
        elif SourceType=="HEG":
            #Execute PL10 !!!Not needed functon!!!
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
            CHType = ws['B7'].value #Obtain CHType
            if CHType<99: #Perhaps not needed?
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
