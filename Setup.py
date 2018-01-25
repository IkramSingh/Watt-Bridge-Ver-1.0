import visa

def setup3458A():
    rm = visa.ResourceManager()
    instrument = rm.open_resource('GPIB0::22::INSTR')
    instrument.timeout = 30000
    instrument.write('DISP OFF, RESET')
    instrument.write('RESET')
    instrument.write('end 2')
    instrument.write('DISP OFF, READY')
    return instrument

