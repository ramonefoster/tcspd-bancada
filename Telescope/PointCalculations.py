import sys
import math
import win32com.client      #needed to load COM objects

util = win32com.client.Dispatch("ASCOM.Utilities.Util")

#################################################
#Calculates Azimuth and Zenith
#################################################
def calcAzimuthAltura(pointRA, pointDEC, coordLat, sideral):
    DEG = 180 / math.pi
    RAD = math.pi / 180.0
    coordRA = util.HMSToHours(pointRA)
    coordDEC = util.DMSToDegrees(pointDEC)
    lst = util.HMSToHours(sideral)
    H = (lst - coordRA) * 15
    latitude = util.DMSToDegrees(coordLat)

    #altitude calc
    sinAltitude = (math.sin(coordDEC * RAD)) * (math.sin(latitude * RAD)) + (math.cos(coordDEC * RAD) * math.cos(latitude * RAD) * math.cos(H * RAD))
    altitude = math.asin(sinAltitude) * DEG #altura em graus

    #azimuth calc
    y = -1 * math.sin(H * RAD)
    x = (math.tan(coordDEC * RAD) * math.cos(latitude * RAD)) - (math.cos(H * RAD) * math.sin(latitude * RAD))

    #This AZposCalc is the initial AZ for dome positioning
    AZposCalc = math.atan2(y, x) * DEG

    #angle from Zenith
    zenith = round(90 - altitude, 2)

    #converting neg values to pos
    if (AZposCalc < 0) :
        AZposCalc = AZposCalc + 360

    if altitude < 0 :
        altIsOk = False 
        airmass = 0       
    else:
        altIsOk = True
        airmass = airMassCalc(90 - altitude)
    
    obsTime, isObservable, isPierWest = checkObsTime(altIsOk, coordRA, coordDEC, latitude, lst)

    return(zenith, altIsOk, AZposCalc, obsTime, isObservable, isPierWest, airmass)

#################################################
#calculates Airmass
#################################################
def airMassCalc(angle):    
        RAD = math.pi / 180.0
        airMass = 1 / (math.cos(angle * RAD) + (0.50572 * (96.07995 - angle) ** -1.6364))
        airmass = round(airMass, 2)
        return airmass

#################################################
#Check the time of observation of a given object 
# ################################################  
def checkObsTime(altIsOk, raPoint, decPoint, Latpoint, lst):
    #calculates if target is above the horizon, respectin the limits by the engineering team
    if (((raPoint - lst) > 6 and (raPoint - lst) <= (18)) or ((raPoint - lst) < -6 and (raPoint - lst) >= (-18)) or (decPoint - Latpoint) >= 90):
        isObservable = False
    elif altIsOk:
        isObservable = True
    else:
        isObservable = False

    #time of observation for an object'
    if raPoint > 0 and raPoint < 24 and isObservable and altIsOk:
        raX = raPoint
        
        if (raX - lst) >= 18:
            obsTime = util.HoursToHMS((raX - lst + 6 - 24), " ", " ", "", 2)
        elif (raX - lst) < -18:
            obsTime = util.HoursToHMS((raX - lst + 6 + 24), " ", " ", "", 2)
        else:
            obsTime = util.HoursToHMS((raX - lst + 6), " ", " ", "", 2)
    else:
        obsTime = util.HoursToHMS(0, " ", " ", "", 2)
    #side of pier
    if (raPoint - lst) > 0:
        isPierWest = True
    else:
        isPierWest = False
    

    return(obsTime, isObservable, isPierWest)