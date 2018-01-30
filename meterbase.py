#
# Copyright (C) 2006 Industrial Research Limited
#

#
# An abstract base class for meters.
#

from results import *

class MeterBase (object):
    def set_current_range (self, current):
        """Set the current range if necessary."""
        pass 
    def set_voltage_range (self, voltage):
        """Set the voltage range if necessary."""        
        pass
    def set_nominal_power (self, W):
        """Set an expected power to help the meter set itself up.
        This is necessary for the optimal setup of the pulse counter."""
        pass
    def set_counts_per_watthour (self, scale):
        """While this is in the base class it is really specific to only
        meters requiring the counter. It is here because the parameter
        is not restricted to pulse-output meters in the front-end."""
        pass
    def set_measurement_type (self, mtype):
        """Change the meter to accepts a new measurement type."""
        raise NotImplemented
    def get_voltage (self, phases = (1, 2, 3)):
        """Get the voltage for the specified phases."""
        raise NotImplemented
    def get_current (self, phases = (1, 2, 3)):
        """Get the current for the specified phases."""        
        raise NotImplemented
    def get_voltage_range (self, phases = (1, 2, 3)):
        """Get the voltage range for the specified phases."""
        raise NotImplemented
    def get_current_range (self, phases = (1, 2, 3)):
        """Get the current range for the specified phases."""        
        raise NotImplemented
    def get_phase (self, phases = (1, 2, 3)):
        """Get the I->V phase for the specified phases."""        
        raise NotImplemented
    def get_active_power (self, phases = (1, 2, 3)):
        """Get the Watts for the specified phases."""        
        raise NotImplemented
    def get_reactive_power (self, phases = (1, 2, 3)):
        """Get the Vars for the specified phases."""                
        raise NotImplemented
    def get_total_active_power (self):
        """Get the total Watts."""        
        raise NotImplemented
    def get_total_reactive_power (self):
        """Get the total Vars."""                
        raise NotImplemented    
    def get_apparent_power (self, phases = (1, 2, 3)):
        """Get the VAs for the specified phases."""                
        raise NotImplemented
    def get_frequency (self, phases = 1):
        """Get the fundamental frequency of the signal."""
        raise NotImplemented
    def get_results (self, mtype):
        """Return a results object with all the measurements."""
        self.set_measurement_type (mtype)
        # print "R " + str (mtype)
        r = Result (mtype)
        r["f"] = self.get_frequency ()
        r["V"] = self.get_voltage ()
        r["I"] = self.get_current ()
        r["phase"] = self.get_phase ()
        r["watts"] = self.get_active_power ()
        r["vars"] = self.get_reactive_power ()
        r["VA"] = self.get_apparent_power ()
        r["totalwatts"] = self.get_total_active_power ()
        r["totalvars"] = self.get_total_reactive_power ()
        r.fake_single_phase ()
        return r
    def get_all_metric (self, phases):
        """Get all metric measurement data that the meter can provide."""
        raise NotImplemented
    def autocal (self):
        """Run an autocal on the meter, if possible."""
        pass
    def initialise (self):
        """Initialise the meter."""
        pass
    def shutdown (self):
        """Perform any actions needed to shutdown the meter."""
        pass

