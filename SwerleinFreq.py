##The standard swerlein algorithm, as he wrote it. has the basic corrections
##and verified against the testpoint code
##02/12/2016

import visa
import csv
import numpy as np
import time as time
import matplotlib.pyplot as plt

#-----Global Variable for Swerlein Measurements-----#
instrument=0
#---------------------------------------------------#
def FNFreq():
    global instrument
    Expect = float(instrument.query('FREQ')) #read frequency of signal
    instrument.write("TARM HOLD;LFILTER ON;LEVEL 0,DC;FSOURCE ACDCV")
    instrument.write("FREQ "+str(Expect*0.9))
    Cal = float(instrument.query("CAL? 245"))
    Freq=Expect/Cal
    return Freq

def FNVmeter_bw(Freq,Range):
    global instrument
    Lvfilter = 120000.0  #LOW VOLTAGE INPUT FILTER B.W.
    Hvattn=36000.0     #HIGH VOLTAGE ATTENUATOR B.W.(NUMERATOR)
    Gain100bw=82000.0   #AMP GAIN 100 B.W. PEAKING CORRECTION!
    if Range<=0.12:
        Bw_corr=(1+(Freq/Lvfilter)**2)/(1+(Freq/Gain100bw)**2)
        Bw_corr=np.sqrt(Bw_corr)
    elif Range<=12:
        Bw_corr=(1+(Freq/Lvfilter)**2)
        Bw_corr=np.sqrt(Bw_corr)
    elif Range>12:
        Bw_corr=(1+(Freq/Hvattn)**2)
        Bw_corr=np.sqrt(Bw_corr)
        
    return Bw_corr

## def OnStart():
##     global instrument
##     rm = visa.ResourceManager()
##     instrument = rm.open_resource('GPIB0::21::INSTR')
##     instrument.timeout = 10000

## def offsettimes(times):
##     array = np.array(times)
##     initial = times[0]
##     return array-initial

## def save(self,name,Acdc,Ac,Mean,Mem):
##     if name=="":
##         values = time.localtime(time.time())
##         file_name = str(values[0])+'.'+str(values[1])+'.'+str(values[2])+'.'+str(values[3])+'.'+str(values[4])+".csv"
##     else:
##         file_name = name
##     tosave = []
##     tosave.append(offsettimes(self.times))
##     if Acdc == True:
##         tosave.append(self.AcdcArray)
##     if Ac == True:
##         tosave.append(self.AcArray)
##     if Mean == True:
##         tosave.append(self.MeanArray)
##     tosave = np.array(tosave)
##     np.savetxt(file_name,tosave.transpose())

## def SaveMem(data):
##     values = time.localtime(time.time())
##     file_name = 'data.'+str(values[0])+'.'+str(values[1])+'.'+str(values[2])+'.'+str(values[3])+'.'+str(values[4])+".csv"
##     data = np.array(data)
##     np.savetxt(file_name,data)
##     #data.tofile(file_name,sep = ',')
##     
## def ReturnData(self):
##     return [offsettimes(self.times),self.readings]


## def SavePlot(self,name):
##     if name=="":
##         values = time.localtime(time.time())
##         file_name = str(values[0])+'.'+str(values[1])+'.'+str(values[2])+'.'+str(values[3])+'.'+str(values[4])
##     else:
##         file_name = name
##     plt.plot(offsettimes(self.times),self.readings)
##     plt.ylabel('Voltage [V]')
##     plt.xlabel('Time [ms]')
##     plt.savefig(str(file_name)+'.pdf')
    

def reset():
    global instrument
    instrument.timeout = 30000
    instrument.write('DISP OFF, RESET')
    instrument.write('RESET')
    instrument.write('end 2')
    instrument.write('DISP OFF, READY')
    
def run():
    reset()

    Force = False #input("force True/False: ")             #are variables forced or not
    Meas_time=15              # Target measure time
    Tsampforce=0.001           #  FORCED PARAMETER
    Aperforce=Tsampforce-(3e-5)#  FORCED PARAMETER
    Numforce=800.0              #  FORCED PARAMETER
    Aper_targ=0.001            #A/D APERTURE TARGET (SEC)
    Nharm_min=Nharm_set                   #MINIMUM # HARMONICS SAMPLED BEFORE ALIAS
    Nbursts=Nbursts_set                 #NUMBER OF BURSTS USED FOR EACH MEASUREMENT



    if (Force==False):
        #determine input signal RMS,range and frequency
        RMS = float(instrument.query('ACV')) #read value of AC voltage
        Range = 1.55*RMS
        instrument.write('DISP OFF, FREQUENCY')
        Expect_Freq = float(instrument.query('FREQ')) #read frequency of signal
        Freq = FNFreq()
        #SAMP PARAM
        Aper=Aper_targ
        Tsamp=(1e-7)*int((Aper+(3e-5))/(1e-7)+0.5) #rounds to 100ns
        #print("first Tsamp: "+str(Tsamp))
        Submeas_time=Meas_time/Nbursts #meas_time specified, this is target time per burst
        Burst_time=Submeas_time*Tsamp/(0.0015+Tsamp) #IT TAKES 1.5ms FOR EACH sample to compute Sdev
        Approxnum=int(Burst_time/Tsamp+0.5)
        #print(Approxnum)
        Ncycle=int(Burst_time*Freq+0.5) # NUMBER OF 1/Freq TO SAMPLE
        #print(" ")
        #print("Ncycle: "+str(Ncycle))
        if Ncycle==0:
            print("Ncycle was 0, set to 1")
            Ncycle=1
            Tsamp=(1e-7)*int(1.0/Freq/Approxnum/(1e-7)+0.5)
            Nharm=int(1.0/Tsamp/2.0/Freq)
            #print("Nharm: "+str(Nharm))
            if Nharm<Nharm_min:
                #print("Nharm too small, set to 6")
                Nharm=Nharm_min
                Tsamp=(1e-7)*int(1.0/2.0/Nharm/Freq/(1e-7)+0.5)
        else:
            Nharm=int(1/Tsamp/2/Freq)
            #print("Nharm: "+str(Nharm))
            if Nharm<Nharm_min:
                Nharm = Nharm_min
                #print("Nharm too small, set to 6")

            Tsamptemp=(1e-7)*int(1.0/2.0/Nharm/Freq/(1e-7)+0.5)
            Burst_time=Submeas_time*Tsamptemp/(0.0015+Tsamptemp)
            Ncycle=int(Burst_time*Freq+1) ##0.5 to 1
            Num=int(Ncycle/Freq/Tsamptemp+0.5)
            #print("Num: "+str(Num))
            
            if Ncycle>1:
                K=int(Num/20/Nharm+1) #0.5 to 1
                #print("K= "+str(K))
            else:
                K=0
                #print("K=0")
            Tsamp=(1e-7)*int(Ncycle/Freq/(Num-K)/(1e-7)+0.5)
            if Tsamp-Tsamptemp<(1e-7):
                #print("Tsamp increased from "+str(Tsamp)+"to "+str((Tsamp+1e-7)))
                Tsamp=Tsamp+1e-7

        Aper=Tsamp-(3e-5)
        Num=int(Ncycle/Freq/Tsamp+0.5)
        #print('NEW NUM '+str(Num))
        if Aper>1.0:
            Aper = 1.0
            print("Aperture too large, automatically set to 1")
        elif Aper<1e-6:
            print("A/D APERTURE IS TOO SMALL")
            print("LOWER Aper_targ, Nharm, OR INPUT Freq")
            print("Aperture set to 1e-6")
            Aper = 1e-6
    else:
        print("Using Forced Parameters")
        #Freq = input("Input signal frequency: ")
        Freq=50.0 #Forced frequency for testing.
        RMS = float(instrument.query('ACV')) #read value of AC voltage
        Range=1.55*RMS
        #Range = 1.55*float(input("RMS value of AC signal: "))
        Tsamp=Tsampforce
        Aper=Aperforce
        Num=Numforce
        Expect_Freq=Freq
        Submeas_time="Forced"
        Burst_time="Forced"
        Approxnum="Forced"
        Ncycle="Forced"
        Nharm="Forced"
        Tsamptemp="Forced"
        K="Forced"

    ######   setup machine
    instrument.write('DISP OFF,SETUP')
    instrument.write('TARM HOLD;AZERO OFF;DCV '+str(Range))
    instrument.write('APER '+str(Aper))
    instrument.write('TIMER '+str(Tsamp))
    instrument.write('NRDGS '+str(Num)+',TIMER')
    #print("TSAMP IS "+str(Tsamp))
    #self.instrument.write('SWEEP '+str(Tsamp)+','+str(Num))
    instrument.write('TRIG LEVEL;LEVEL 0,DC;DELAY 0;LFILTER ON')
    instrument.write('NRDGS '+str(Num)+',TIMER')
    memory_available = instrument.query('MSIZE?').split(',') #generates 2 element list, first element is memory available for measuremnts
    Storage = int( int(memory_available[0])/4 )
    #print(" ")
    #print("MACHINE SETUP COMPLETE")
    #print(" ")
    if Num>Storage:
        print("NOT ENOUGH VOLTMETER MEMORY FOR NEEDED SAMPLES")
        print("TRY A LARGER Aper_targ VALUE OR SMALLER Num")


    ######  PRELIMINARY COMPUTATIONS


    Bw_corr=FNVmeter_bw(Freq,Range)
    ##error estimate##
    if Range>120:
        Base = 15.0   #self heating and base error
    else:
        Base = 10.0
    #Vmeter_bw IS ERROR DUE TO UNCERTAINTY IN KNOWING THE HIGH FREQUENCY
    #RESPONSE OF THE DCV FUNCTION FOR VARIOUS RANGES AND FREQUENCIES
    #UNCERTAINTY IS 30% AND THIS ERROR IS RANDOM
    X1=FNVmeter_bw(Freq,Range)
    X2=FNVmeter_bw(Freq*1.3,Range)
    Vmeter_bw=int((1e+6)*abs(X2-X1))   #error due to meter band width, bw.
    #Aper_er IS THE DCV GAIN ERROR FOR VARIOUS A/D APERTURES
    #THIS ERROR IS SPECIFIED IN A GRAPH ON PAGE 11 OF THE DATA SHEET
    #THIS ERROR IS RANDOM
    Aper_er=int(1.0*.002/Aper)       #GAIN UNCERTAINTY - SMALL A/D APERTURE
    if Aper_er>30 and Aper>=1e-5:
        Aper_er=30
    if Aper<1e-5:
        Aper_er=10+int(0.0002/Aper)
    #Sincerr IS THE ERROR DUE TO THE APERTURE TIME NOT BEING PERFECTLY KNOWN
    #THIS VARIATION MEANS THAT THE Sinc CORRECTION TO THE SIGNAL FREQUENCY
    #IS NOT PERFECT.  ERROR COMPONENTS ARE CLOCK FREQ UNCERTAINTY(0.01%)
    #AND SWITCHING TIMING (50ns).  THIS ERROR IS RANDOM.
    X=np.pi*Aper*Freq       #USED TO CORRECT FOR A/D APERTURE ERROR
    Sinc = np.sin(X)/X
    Y=np.pi*Freq*(Aper*1.0001+5.0e-8)
    Sinc2=np.sin(Y)/Y
    Sincerr=int(1e+6*abs(Sinc2-Sinc))   #APERTURE UNCERTAINTY ERROR
    #Dist IS ERROR DUE TO UP TO 1% DISTORTION OF THE INPUT WAVEFORM
    #IF THE INPUT WAVEFORM HAS 1% DISTORTION, THE ASSUMPTION IS MADE
    #THAT THIS ENERGY IS IN THE THIRD HARMONIC.  THE APERTURE CORRECTION,
    #WHICH IS MADE ONLY FOR THE FUNDAMENTAL FREQUENCY WILL THEN BE
    #INCORRECT.  THIS ERROR IS RETURNED SEPERATELY.
    Sinc3=np.sin(3*X)/3/X      #SINC CORRECTION NEEDED FOR 3rd HARMONIC
    Harm_er=abs(Sinc3-Sinc)
    Dist=np.sqrt(1.0+(0.01*(1+Harm_er))**2)-np.sqrt(1.0+0.01**2)
    Dist=int(Dist*1e+6)
    ##Tim_er IS ERROR DUE TO MISTIMING.  IT CAN BE SHOWN THAT IF A
    ##BURST OF Num SAMPLES ARE USED TO COMPUTE THE RMS VALUE OF A SINEWAVE
    ##AND THE SIZE OF THIS BURST IS WITHIN 50ns*Num OF AN INTEGRAL NUMBER
    ##OF PERIODS OF THE SIGNAL BEING MEASURED, AN ERROR IS CREATED
    ##BOUNDED BY 100ns/4/Tsamp.  THIS ERROR IS DUE TO THE 100ns QUANTIZATION
    ##LIMITATION OF THE HP3458A TIME BASE.  IF THIS ERROR WERE ZERO, THEN
    ##Num*Tsamp= INTEGER/Freq, BUT WITH THIS ERROR UP TO 50ns OF TIMEBASE
    ##ERROR IS PRESENT PER SAMPLE, THEREFORE TOTAL TIME ERROR=50ns*Num
    ##THIS ERROR CAN ONLY ACCUMULATE UP TO 1/2 *Tsamp, AT WHICH POINT THE
    ##ERROR IS BOUNDED BY 1/4/Num
    ##THIS ERROR IS FURTHER REDUCED BY USING THE LEVEL TRIGGER
    ##TO SPACE Nbursts AT TIME INCREMENTS OF 1/Nbursts/Freq.  THIS
    ##REDUCTION IS SHOWN AS 20*Nbursts BUT IN FACT IS USUALLY MUCH BETTER
    ##THIS ERROR IS ADDED ABSOLUTELY TO THE Err CALCULATION
    #(Err,Dist_er,Freq,Range,Num,Aper,Nbursts) sent params
    #(Err,Dist,Freq,Range,Num,Aper,Nbursts) recieved renamed
    Tim_er=int((1e+6)*1e-7/4/(Aper+(3.0e-5))/20.0)#ERROR DUE TO HALF CYCLE ERROR
    Limit=int((1e+6)/4.0/Num/20.0)
    if Tim_er>Limit:
        Tim_er=Limit
    ##Noise IS THE MEASUREMENT TO MEASUREMENT VARIATIONS DUE TO THE
    ##INDIVIDUAL SAMPLES HAVING NOISE.  THIS NOISE IS UNCORRELATED AND
    ##IS THEREFORE REDUCED BY THE SQUARE ROOT OF THE NUMBER OF SAMPLES
    ##THERE ARE Nbursts*Num SAMPLES IN A MEASUREMENT.  THE SAMPLE NOISE IS
    ##SPECIFIED IN THE GRAPH ON PAGE 11 OF THE DATA SHEET.  THIS GRAPH
    ##SHOWS 1 SIGMA VALUES, 2 SIGMA VALUES ARE COMPUTED BELOW.
    ##THE ERROR ON PAGE 11 IS EXPRESSED AS A % OF RANGE AND IS MULTIPLIED
    ##BY 10 SO THAT IT CAN BE USED AS % RDG AT 1/10 SCALE.
    ##ERROR IS ADDED IN AN ABSOLUTE FASHION TO THE Err CALCULATION SINCE
    ##IT WILL APPEAR EVENTUALLY IF A MEASUREMENT IS REPEATED OVER AND OVER
    Noiseraw=0.9*np.sqrt(0.001/Aper)       #1 SIGMA NOISE AS PPM OF RANGE
    Noise=Noiseraw/np.sqrt(Nbursts*Num)  #REDUCTION DUE TO MANY SAMPLES
    Noise=10.0*Noise                   #NOISE AT 1/10 FULL SCALE
    Noise=2.0*Noise                    #2 SIGMA
    if Range<=0.12:               #NOISE IS GREATER ON 0.1 V RANGE
        Noise=7.0*Noise                  #DATA SHEET SAYS USE 20, BUT FOR SMALL
        Noiseraw=7.0*Noiseraw            #APERTURES, 7 IS A BETTER NUMBER
    Noise=int(Noise)+2.0                  #ERROR DUE TO SAMPLE NOISE
    ##Df_err IS THE ERROR DUE TO THE DISSIPATION FACTOR OF THE P.C. BOARD
    ##CAPACITANCE LOADING DOWN THE INPUT RESISTANCE.  THE INPUT RESISTANCE
    ##IS 10K OHM FOR THE LOW VOLTAGE RANGES AND 100K OHM FOR THE HIGH VOLTAGE
    ##RANGES (THE 10M OHM INPUT ATTENUATOR).  THIS CAPACITANCE HAS A VALUE
    ##OF ABOUT 15pF AND A D.F. OF ABOUT 1.0%.  IT IS SWAMPED BY 120pF
    ##OF LOW D.F. CAPACITANCE (POLYPROPALENE CAPACITORS) ON THE
    ##LOW VOLTAGE RANGES WHICH MAKES FOR AN EFFECTIVE D.F. OF ABOUT .11%.
    ##THIS CAPACITANCE IS SWAMPED BY 30pF LOW D.F. CAPACITANCE ON THE
    ##HIGH VOLTAGE RANGES WHICH MAKES FOR AN EFFECTIVE D.F. OF .33%.
    ##THIS ERROR IS ALWAYS IN THE NEGATIVE DIRECTION, SO IS ADDED ABSOLUTELY
    if Range<=12:
        Rsource=10000
        Cload=1.33e-10
        Df=1.1e-3          #0.11%
    else:
        Rsource=1.0e+5
        Cload=5.00e-11
        Df=3.3e-3
    Df_err=2*np.pi*Rsource*Cload*Df*Freq
    Df_err=int(1.0e+6*Df_err)#ERROR DUE TO TO PC BOARD DIELECTRIC ABSORBTION
    #Err IS TOTAL ERROR ESTIMATION.  RANDOM ERRORS ARE ADDED IN RSS FASHION
    Err=np.sqrt(Base**2+Vmeter_bw**2+Aper_er**2+Sincerr**2)
    Err=int(Err+Df_err+Tim_er+Noise)
##    print("SIGNAL FREQUENCY(Hz)= "+str(Freq))
##    print("Number of samples in each of "+str(Nbursts)+" bursts= "+str(Num))
##    print("Sample spacing(sec)= "+str(Tsamp))
##    print("A/D Aperture(sec)= "+str(Aper))
##    print("Measurement bandwidth(Hz)= "+str(int(5/Aper)/10))
##    print("ESTIMATED TOTAL SINEWAVE MEASUREMENT UNCERTAINTY(ppm)= "+str(Err))
##    print("ADDITIONAL ERROR FOR 1% DISTORTION(3rd HARMONIC)(ppm)= "+str(Dist))
##    print("NOTE: ERROR ESTIMATE ASSUMES (ACAL DCV) PERFORMED RECENTLY(24HRS)")

    ###measuremnts###
    Sum=0
    Sumsq=0
    
    
    
    for I in range(0,int(Nbursts)):
        instrument.write('DISP OFF,MEASURING')
        Delay=float(I)/Nbursts/Freq+(1e-6)
        instrument.write('DELAY '+str(Delay))
        instrument.write('TIMER '+str(Tsamp))
    
        #make measuremnts
        instrument.write('MEM FIFO;MFORMAT DINT') #first in first out for memeory 
        
        #clears memeory, sets to 4 bytes per reading
        instrument.write('TARM SGL')
        instrument.write('MMATH STAT')#Just makes the 3458A machine calculate the SDEV,MEAN,NSAMP,UPPER,LOWER. So no need to validate this.
        
        Sdev = float(instrument.query('RMATH SDEV'))
        Mean = float(instrument.query('RMATH MEAN'))
        
        single_burst = []

        Sdev=Sdev*np.sqrt((Num-1.0)/Num)     #CORRECT SDEV FORMULA
        Sumsq=Sumsq+Sdev*Sdev+Mean*Mean
        Sum=Sum+Mean
        Temp=Sdev*Bw_corr/Sinc
        Temp=Range/(1e+7)*int(Temp*(1e+7)/Range)#6 DIGIT TRUNCATION
        #print("Sdev: "+str(Sdev))
        #print("Mean: "+str(Mean))
        #print("RMS: "+str(np.sqrt(Sdev**2+Mean**2)))
        #print(Temp)
        I=I+1
    Dcrms=np.sqrt(Sumsq/Nbursts)
    #print("RMS value: "+str(Dcrms))
    #print(" ")
    Dc=Sum/Nbursts
    Acrms=np.sqrt(Dcrms**2-Dc**2)
    Acrms=Acrms*Bw_corr/Sinc  #CORRECT A/D Aper AND Vmeter B.W.
    Dcrms=np.sqrt(Acrms*Acrms+Dc*Dc)
    Acrms=Range/1e+7*int(Acrms*1e+7/Range+0.5)    #6 DIGIT TRUNCATION
    Dcrms=Range/1e+7*int(Dcrms*1e+7/Range+0.5)    #6 DIGIT TRUNCATION
    instrument.write('DISP ON')
    return Acrms

