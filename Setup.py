import visa

def setup3458A():
    '''Creates a communication link between Watt Bridge software and HP3458A DVM'''
    rm = visa.ResourceManager()
    instrument = rm.open_resource('GPIB0::22::INSTR')
    instrument.timeout = 30000
    instrument.write('DISP OFF, RESET')
    instrument.write('RESET')
    instrument.write('end 2')
    instrument.write('DISP OFF, READY')
    return instrument

def setup53230A():
    '''Creates a communication link between Watt Bridge software and Ag53230A Frequency Counter'''
    rm = visa.ResourceManager()
    instrument = rm.open_resource('GPIB0::3::INSTR')
    return instrument
