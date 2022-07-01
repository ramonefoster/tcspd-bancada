import sys
import math
import numpy as np
from datetime import datetime
import win32com.client      #needed to load COM objects
import ephem

util = win32com.client.Dispatch("ASCOM.Utilities.Util")

class TelescopeData():
    def ConvertCoord(TRa, Tdec, TLst, TUtc):
        #read telescopes data
        telRa = TRa
        telDec = Tdec
        telLST = TLst
        telUTC = TUtc

        #calculates HA
        telHA = telLST - telRa

        #convert format        
        raCoord = util.HoursToHMS(telRa, " ", " ", "", 2)
        decCoord = util.DegreesToDMS(telDec, " ", " ", "", 2)
        haCoord = util.HoursToHMS(telHA, " ", " ", "", 2)
        lstCoord = util.HoursToHMS(telLST, " ", " ", "", 2)
        utcCoord = telUTC

        return raCoord, decCoord, haCoord, lstCoord, utcCoord

    def pointConvert(ra_point, dec_point):
        Radouble = util.HMSToHours(ra_point)
        DecDouble = util.DMSToDegrees(dec_point)

        return Radouble, DecDouble
    
    def precess_coord(ra_target, dec_target):
        #teste pyephem
        OPD=ephem.Observer()
        OPD.lat='-22.5344'
        OPD.lon='-45.5825'
        OPD.date = datetime.utcnow()
        # %% these parameters are for super-precise estimates, not necessary.
        OPD.elevation = 1864 # meters
        star = ephem.FixedBody()
        star._ra  = ephem.hours(ra_target.replace(" ",":")) # in hours for RA
        star._dec = ephem.degrees(dec_target.replace(" ",":"))
        star.compute(OPD)
        ra_hms = str(star.ra).replace(":", " ")
        dec_dms = str(star.dec).replace(":", " ")

        return ra_hms, dec_dms
    
