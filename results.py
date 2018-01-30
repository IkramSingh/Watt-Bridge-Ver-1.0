#
# Copyright (C) 2006 Industrial Research Limited
#
#
# Objects and methods to handle results.
#

from utilities import *
from measurement import *
import math

class Result (dict):
    """A basic set of results, frequency, Vrange, Voltage (x3), Irange, Current (x3),
    phase (x3), Watts (x3), Vars (x3) and VA (x3). A timestamp is
    also included."""
    
    def __init__ (self, mtype):
        """The default is initialised to all 0."""
        self["f"] = 0.0
        self["V"] = [ 0.0, 0.0, 0.0 ]
        self["I"] = [ 0.0, 0.0, 0.0 ]
        self["phase"] = [ 0.0, 0.0, 0.0 ]
        self["watts"] = [ 0.0, 0.0, 0.0 ]
        self["totalwatts"] = 0.0
        self["vars"] = [ 0.0, 0.0, 0.0 ]
        self["totalvars"] = 0.0
        self["VA"] = [ 0.0, 0.0, 0.0 ]
        self["timestamp"] = timestamp ()
        self["mtype"] = mtype

    def __str__ (self):
        """Throw out a string quitable for a CSV file."""
        return "%f, %f, %f, %f, %f, %f, %f, %f, %f, %f, %f, %f, %f, %f, %f, %f, %f, %f, %f, %f, %f, %s" % tuple (self["V"] + self["I"] + self["phase"] + [self["f"]] + self["watts"] + [self["totalwatts"]] + self["vars"] + [self["totalvars"]] + self["VA"] + [self["timestamp"]])

    def _test1 (self, a, b):
        """A helper function for _test. It applies the 10% test to only
        one pair."""
        try:
            ratio = a/b
            return (ratio > 0.9) and (ratio < 1.1)
        except ZeroDivisionError: 
            return False # It is very easy to have a target of 0 (e.g. single phase measurents) so ignore these.
        
    def _test (self, a, b):
        """Check whether a is within 10% of b."""
        result = []
        if hasattr (a, '__iter__'):
            if hasattr (b, '__iter__'):
                for x, y in zip (a, b):
                    result.append (self._test1 (x, y))
            else:
                for x in a:
                    result.append (self._test1 (x, b))
        else:
            if hasattr (b, '__iter__'):
                for y in b:
                    result.append (self._test1 (a, y))
            else:
                result.append (self._test1 (a, b))
        return result

    def fake_single_phase (self):
        """Explicitly set the fields that would be unused in a single-phase
        measurement to zero."""
        if self["mtype"] == MeasurementType.SinglePhase1:
            a, b = 1, 2
        elif self["mtype"] == MeasurementType.SinglePhase2:
            a, b = 0, 2
        elif self["mtype"] == MeasurementType.SinglePhase3:
            a, b = 0, 1
        else:
            return
        for res in ("V", "I", "phase", "watts", "vars", "VA"):
            self[res][a] = self[res][b] = 0.0        

    def adjust (self, calconstants):
        """*** DEPRECATED ***, Adjust the results given a set of calibration constants."""
        self["V"] = add_ppm (self["V"], calconstants[0:3])
        self["I"] = add_ppm (self["I"], calconstants[3:6])
        self["phase"] = add_ppm (self["phase"], calconstants[6:9]) # Should this be ppm or an absolute phase?
        self["f"] = add_ppm (self["f"], calconstants[9:10][0]) # Note the subscripting after the slice to make the result a scalar
        self["watts"] = add_ppm (self["watts"], calconstants[10:13])
        self["vars"] = add_ppm (self["vars"], calconstants[13:16])
        self["VA"] = add_ppm (self["VA"], calconstants[16:19])

    def check_sanity (self, measurement):
        """Test whether the results are reasonable given the attempted
        measurement. i.e. check if they are within 10% of nominal."""
        
        return True # Disable sanity checks. There are still issues.
        
        checks = []
        # The use of absolute value misses some obvious errors, but since pulse
        # counting is unsigned, this is necessary.        
        for i in xrange (0, 3):
            checks += self._test (abs (self["watts"][i]), abs (measurement.V[i]*measurement.I[i]*math.cos (measurement.theta)))
            checks += self._test (abs (self["vars"][i]), abs (measurement.V[i]*measurement.I[i]*math.sin (measurement.theta)))
        checks += self._test (abs (self["totalwatts"]), abs (3*measurement.V[i]*measurement.I[i]*math.cos (measurement.theta)))
        checks += self._test (abs (self["totalvars"]), abs (3*measurement.V[i]*measurement.I[i]*math.sin (measurement.theta)))        
        # We check only whether one answer was close. This is because our
        # lowest-common-denominator system is the pulse counter and it can
        # only measure one thing at a time.
        return reduce (lambda x, y: x or y, checks)

    def error (self, b):
        """Calculate the error between this result and another result."""
        e = Result(self["mtype"])
        # FIXME: Should the angle be off by ppm or an absolute phase?
        for a in self:
            if a == "timestamp" or a == "mtype":
                continue
            e[a] = ppm_error (self[a], b[a])
        e.fake_single_phase ()
        return tuple (e["V"] + e["I"] + e["phase"] + [e["f"]] + e["watts"] + [e["totalwatts"]] + e["vars"] + [e["totalvars"]] + e["VA"])

    def error_str (self, e):
        """Format the error in a form suitable for a CSV file."""
        return "%f, %f, %f, %f, %f, %f, %f, %f, %f, %f, %f, %f, %f, %f, %f, %f, %f, %f, %f, %f, %f" % e
