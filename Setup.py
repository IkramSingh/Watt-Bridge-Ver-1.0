import visa

def setup3458A():
    rm = visa.ResourceManager()
    instrument = rm.open_resource('GPIB0::22::INSTR')
    instrument.timeout = 10000
    return instrument

