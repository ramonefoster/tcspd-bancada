from pickle import TRUE
import ephem
import re
import sys, os
from decimal import Decimal
import time
import math
import win32com.client      #needed to load COM objects
import win32gui
import win32con
from datetime import datetime, timedelta
import threading
import cv2
import numpy as np
import serial.tools.list_ports
from mpl_toolkits.mplot3d import Axes3D
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

from PyQt5 import QtCore, QtGui, QtWidgets, uic, QtWebEngineWidgets, QtTest
from PyQt5.QtCore import QTimer, QDateTime, QObject, QThread, pyqtSignal, QUrl, pyqtSlot, Qt, QSettings, QThreadPool
from PyQt5.QtGui import *
from PyQt5.QtWidgets import QMainWindow, QAction, qApp, QApplication, QStyle, QWidget, QLabel, QLineEdit, QTextEdit, QGridLayout, QMessageBox
import Telescope.telescopeCoord, Telescope.PointCalculations 
import controllers.MoveAxis as AxisDevice
import  shutil
import ftplib

util = win32com.client.Dispatch("ASCOM.Utilities.Util")

pyQTfileName = "TCSPD.ui"
Ui_MainWindow, QtBaseClass = uic.loadUiType(pyQTfileName)

class MyApp(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        self.thread_manager = QThreadPool()

        self.btnPoint.clicked.connect(self.point)
        self.btnPrecessP.clicked.connect(self.precess)
        self.btnAbort.clicked.connect(self.stop)
        self.btnReset.clicked.connect(self.reset_uc)

        global ss
        ss = self.txtPointRA.styleSheet() #original saved
        global statbuf
        statbuf = None
        self.device = None
        self.opd_device = None

        self.load_allsky()
        self.btnPrecess.clicked.connect(self.select_to_precess)
        self.btnBSC.clicked.connect(self.load_bsc_default)
        self.btnStart.clicked.connect(self.start_timer)
        self.listWidget.itemDoubleClicked.connect(self.select_to_precess)
        self.sliderTrack.valueChanged.connect(self.check_track)

        self.timer_update = QTimer()
        self.download_weather()
        self.update_weather()
                
    def start_timer(self):
        self.timer_update.timeout.connect(self.update_data)
        self.timer_update.stop()
        self.timer_update.start(1000)
        self.device = "AH"
        self.opd_device = AxisDevice.AxisControll(self.device, 'COM3', 9600)
        self.encDEC.setText('-22:32:04')
        self.statDEC.setStyleSheet("background-color: lightgreen")

    def check_track(self):
        if self.sliderTrack.value() == 1:
            try:
                self.opd_device.sideral_ligar()
            except Exception as e:
                print(e)
        elif self.sliderTrack.value() == 0:
            try:
                self.opd_device.sideral_desligar()
            except Exception as e:
                print(e)
                
    def load_allsky(self):
        """load allsky images"""
        urlAS = QUrl("http://200.131.64.237:8090/onlycam340c") 
        if urlAS.isValid():
            self.webAllsky.load(urlAS)         

    def select_to_precess(self):
        """select an object and send to precess area"""
        if self.listWidget.selectedItems():
            nameObj = ([item.text().split("\t")[0].strip() for item in self.listWidget.selectedItems()])[0]
            raObj = ([item.text().split("\t")[1].strip() for item in self.listWidget.selectedItems()])[0]
            decObj = ([item.text().split("\t")[2].strip() for item in self.listWidget.selectedItems()])[0]
            magObj = ([item.text().split("\t")[3].strip() for item in self.listWidget.selectedItems()])[0]
            ramObj = ([item.text().split("\t")[4].strip() for item in self.listWidget.selectedItems()])[0]
            decmObj = ([item.text().split("\t")[5].strip() for item in self.listWidget.selectedItems()])[0]
            self.set_precess(nameObj, raObj, decObj, magObj, ramObj, decmObj)
            self.objName.setText(nameObj)
            self.objRA.setText(raObj)
            self.objDEC.setText(decObj)
            self.tabWidget_2.setCurrentIndex(1)
    
    def get_sidereal(self):
        OPD=ephem.Observer()
        OPD.lat='-22.5344'
        OPD.lon='-45.5825'        
        utc = datetime.utcnow()
        OPD.date = utc
        # %% these parameters are for super-precise estimates, not necessary.
        OPD.elevation = 1864 # meters
        OPD.horizon = 0    
        sidereal_time = OPD.sidereal_time()
        return str(sidereal_time), str(utc)[12:19]

    def set_precess(self, nameObj, raObj, decObj, magObj, ramObj, decmObj):
        """Checks object if observable or not"""
        self.txtPointRA.setText(raObj)
        self.txtPointOBJ.setText(nameObj)
        self.txtPointDEC.setText(decObj)
        self.txtPointMag.setText(magObj)
        latitude = '-22:32:04'
        
        sideral, utc = self.get_sidereal()

        #self.working_area(raObj, decObj, latitude, sideral)
    
    def mover_rel(self):
        """fine movement by indexer"""
        if "AH" in self.device:
            dest_rel = self.txtIndexer.text()
            if len(dest_rel) > 2:
                self.opd_device.mover_rel(dest_rel)
        if "DEC" in self.device:
            dest_rel = self.txtIndexer_2.text()
            if len(dest_rel) > 2:
                self.opd_device.mover_rel(dest_rel)

    def stop(self):
        """stop any movement and abort slew"""
        try:
            self.opd_device.prog_parar()
            if "AH" in self.device:
                self.opd_device.sideral_desligar()
        except Exception as e:
            print(e)

    def tracking(self):
        """turns sidereal on and off"""
        if "AH" in self.device and statbuf:
            if statbuf[19] == "0":
                self.opd_device.sideral_ligar()
            elif statbuf[19] == "1":
                self.opd_device.sideral_desligar()

    def precess(self):
        """precess coordinates based on FINALS"""
        sideral, utc = self.get_sidereal()
        coord_format_ra = re.compile('.{2} .{2} .{4}')
        coord_format_dec = re.compile('.{3} .{2} .{2}')
        ra_p = self.txtPointRA.text()
        dec_p = self.txtPointDEC.text()
        if coord_format_ra.match(ra_p) is not None and coord_format_dec.match(dec_p) is not None:
            self.txtPointRA.setStyleSheet(ss) #back to original
            self.txtPointDEC.setStyleSheet(ss) #back to original
            self.txtTargetRA.setStyleSheet(ss) #back to original
            self.txtTargetDEC.setStyleSheet(ss) #back to original
            lat = '-22:32:04'
            self.working_area(ra_p, dec_p, lat, sideral)
            new_ra, new_dec = Telescope.telescopeCoord.TelescopeData.precess_coord(ra_p, dec_p)
            self.txtTargetRA.setText(new_ra.replace(",", "."))
            self.txtTargetDEC.setText(new_dec.replace(",", "."))
        else:
            self.txtPointRA.setStyleSheet("border: 1px solid red;") #changed
            self.txtPointDEC.setStyleSheet("border: 1px solid red;") #changed      
    
    def working_area(self, raObj, decObj, latitude, sideral):
        """check if object is in observable zone"""      
        zenith, is_altitude_ok, azimuth_calc, observation_time, is_observable, \
            is_pier_west, airmass = Telescope.PointCalculations.calcAzimuthAltura(raObj, decObj, latitude, sideral)
        if is_observable:
            self.txtPointWorkingArea.setText(" IN ")
            self.txtPointWorkingArea.setStyleSheet("background-color: lightgreen")
            self.txtZenitAngle.setText(str(zenith))
            self.txtPointObsTime.setText(str(observation_time))
            self.txtPointAirmass.setText(str(airmass))
        else:
            self.txtPointWorkingArea.setText(" OFF ")
            self.txtPointWorkingArea.setStyleSheet("background-color: indianred")
            self.txtZenitAngle.setText("")
            self.txtPointObsTime.setText("")
            self.txtPointAirmass.setText(str(""))
    
    def load_bsc_default(self):
        """loads BSC default file (by LNA)"""
        bsc_file = 'C:\\Users\\rguargalhone\\Documents\\BSC_08.txt'
        if bsc_file and os.path.exists(bsc_file):
            data = open(str(bsc_file), 'r')
            data_list = data.readlines()

            self.listWidget.clear()

            for eachLine in data_list:
                if len(eachLine.strip())!=0:
                    self.listWidget.addItem(eachLine.strip())

    def load_weather_file(self):
        """load weather txt file from weather station"""
        weather_file = 'C:\\Users\\rguargalhone\\Documents\\weatherData\\download.txt'
        if weather_file and os.path.exists(weather_file):
            try:
                data = open(str(weather_file), 'r')
                lines = data.read().splitlines()
                last_line = lines[-2]
                outside_temp = last_line.split()[2]
                wind_speed = last_line.split()[7]
                weather_bar = last_line.split()[15]
                Humidity = last_line.split()[5]  
                Dew = last_line.split()[6]  
                wind_dir = last_line.split()[8]              
                return (outside_temp, wind_speed, weather_bar, Humidity, Dew, wind_dir)
            except Exception as e:
                print("Error weather file: ", e)
                return ("0", "0", "0", "0")
        else:
            return ("0", "0", "0", "0")
    
    def update_weather(self):
        """loads file from weather station and updates"""
        temperature, windspeed, bar_w, humidity, dew, wind_dir = self.load_weather_file()
        self.txtTemp.setText(temperature)
        self.txtUmid.setText(humidity)
        self.txtWind.setText(windspeed)
        self.txtDew.setText(dew)
        self.txtWindDir.setText(wind_dir)
        """if humidity is higher than 90%, closes shutter"""
        if float(humidity) > 90:
            self.txtUmid.setStyleSheet("background-color: indianred")
            return(True)
        elif 80 < float(humidity) <= 90:
            self.txtUmid.setStyleSheet("background-color: gold")
            return(False)
        elif float(humidity) < 10:
            self.txtUmid.setStyleSheet("background-color: lightgrey")
        else:
            self.txtUmid.setStyleSheet("background-color: lightgrey")
            return(False)
    
    def update_data(self):
        """    Update coordinates every 1s    """
        # year = datetime.datetime.now().strftime("%Y")
        # month = datetime.datetime.now().strftime("%m")
        # day = datetime.datetime.now().strftime("%d")
        # hours = datetime.datetime.now().strftime("%H")
        # minute = datetime.datetime.now().strftime("%M")
        utc_time = str(datetime.utcnow().strftime('%H:%M:%S'))

        #ephem
        gatech = ephem.Observer()
        gatech.lon, gatech.lat = '-45.5825', '-22.534444'

        self.txtLST.setText(str(gatech.sidereal_time()))
        self.txtUTC.setText(utc_time)

        # #DATA
        self.get_status()
        if statbuf:            
            if "*" in statbuf:
                ra = statbuf[0:11]
                dec = self.txtPointDEC.text()
                self.encRA.setText(ra)
                lat = '-22:32:04'
                sideral, utc = self.get_sidereal()
                zenith, is_altitude_ok, azimuth_calc, observation_time, is_observable, \
                is_pier_west, airmass = Telescope.PointCalculations.calcAzimuthAltura(ra, dec, lat, sideral)
                self.bit_status()
                if "AH" in self.device:
                    self.ah_status()
    
    def ah_status(self):
        """shows ah statbuf and check sideral stat"""
        if statbuf[19] == "1":
            self.sliderTrack.setValue(1)
        else:
            self.sliderTrack.setValue(0)

    @pyqtSlot()
    def get_status(self):
        """calls threading stats"""
        self.thread_manager.start(self.get_prog_status)

    @pyqtSlot()
    def get_prog_status(self):
        """get statbuf from controller"""
        global statbuf
        statbuf = self.opd_device.progStatus()

    def bit_status(self):
        """sets the labels colors for each statbit"""
        if len(statbuf)>25:
            if statbuf[15] == "1":
                self.stat3.setStyleSheet("background-color: lightgreen")
            else:
                self.stat3.setStyleSheet("background-color: indianred")
            if statbuf[16] == "1":
                self.stat4.setStyleSheet("background-color: lightgreen")
            else:
                self.stat4.setStyleSheet("background-color: indianred")
            if statbuf[17] == "1":
                self.stat5.setStyleSheet("background-color: lightgreen")
                self.statSecurity.setStyleSheet("background-color: lightgreen")
            else:
                self.stat5.setStyleSheet("background-color: indianred")
                self.statSecurity.setStyleSheet("background-color: darkgreen")
            if statbuf[19] == "1":
                self.stat6.setStyleSheet("background-color: lightgreen")
            else:
                self.stat6.setStyleSheet("background-color: indianred")
            if statbuf[21] == "1":
                self.stat7.setStyleSheet("background-color: lightgreen")
                self.statGross.setStyleSheet("background-color: lightgreen")
            else:
                self.stat7.setStyleSheet("background-color: indianred")
                self.statGross.setStyleSheet("background-color: darkgreen")
            if statbuf[22] == "1":
                self.stat8.setStyleSheet("background-color: lightgreen")
                self.statGross.setStyleSheet("background-color: lightgreen")
            else:
                self.stat8.setStyleSheet("background-color: indianred")
                self.statGross.setStyleSheet("background-color: darkgreen")
            if statbuf[23] == "1":
                self.stat9.setStyleSheet("background-color: lightgreen")
                self.statFine.setStyleSheet("background-color: lightgreen")
            else:
                self.stat9.setStyleSheet("background-color: indianred")
                self.statFine.setStyleSheet("background-color: darkgreen")
            if statbuf[24] == "1":
                self.stat10.setStyleSheet("background-color: lightgreen")
            else:
                self.stat10.setStyleSheet("background-color: indianred")
            if statbuf[25] == "1":
                self.stat11.setStyleSheet("background-color: lightgreen")
            else:
                self.stat11.setStyleSheet("background-color: indianred")
            if statbuf[26] == "1":
                self.stat12.setStyleSheet("background-color: lightgreen")
            else:
                self.stat12.setStyleSheet("background-color: indianred")
            if statbuf[27] == "1":
                self.stat13.setStyleSheet("background-color: lightgreen")
            else:
                self.stat13.setStyleSheet("background-color: indianred")
            if statbuf[23] == "1" and statbuf[25] == "1":
                self.statRA.setStyleSheet("background-color: darkgreen")
            else:
                self.statRA.setStyleSheet("background-color: lightgreen")
    
    def reset_uc(self):
        try:
            self.opd_device.reset()
        except Exception as e:
            print(e)

    def download_weather(self):
        filename = 'downld02.txt'
        try:
            #t40
            ftp = ftplib.FTP("200.131.64.213") 
            ftp.login("tcspd", "opd@2020") 
            ftp.cwd('weather')
            local_filename = os.path.join('C:\\Users\\rguargalhone\\Documents\\', filename)
            ftp.retrbinary("RETR " + filename, open(local_filename, 'wb').write)
            ftp.quit()
            hour = datetime.now().hour
            minutes = datetime.now().minute
            if minutes < 10:
                minutes = '0' + str(minutes)
            print('['+str(hour)+':'+str(minutes)+']Weather T40 data downloaded')
            original = r'C:\Users\rguargalhone\Documents\downld02.txt'
            target = r'C:\Users\rguargalhone\Documents\weatherData\download.txt'
            shutil.copyfile(original, target)         
        except Exception as e:
            print("Error t40: ", e)

    def point(self):
        """Points the telescope to a given Target"""
        if "AH" in self.device:
            sid, utc = self.get_sidereal()            
            ra_txt = self.txtPointRA.text()
            lst = util.HMSToHours(sid)
            ra = util.HMSToHours(ra_txt)
            ra_txt = util.HoursToHMS(lst - ra, " ", " ", "", 2)
            if len(ra_txt) > 2:
                try:
                    if statbuf[25] == "0":
                        self.opd_device.sideral_ligar()
                        self.opd_device.mover_rap(ra_txt)
                    else:
                        print("erro")

                except Exception as exc_err:
                    print("error: ", exc_err)
            else:
                msg = "Ivalid RA"
                self.show_dialog(msg)
        elif "DEC" in self.device:
            dect_txt = self.txtPointDEC.text()
            if len(dect_txt) > 2:
                try:
                    if statbuf[25] == "0":
                        self.opd_device.mover_rap(dect_txt)
                    else:
                        print("erro")

                except Exception as exc_err:
                    print("error: ", exc_err)
            else:
                msg = "Ivalid DEC inputs"
                self.show_dialog(msg)

    def close_event(self, event):
        """shows a message to the user confirming closing application"""
        close = QMessageBox()
        close.setText("Are you sure?")
        close.setStandardButtons(QMessageBox.Yes | QMessageBox.Cancel)
        close = close.exec()

        if close == QMessageBox.Yes:
            self.disconnect_device()
            event.accept()
        else:
            event.ignore()

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MyApp()

    window.show()
    sys.exit(app.exec_())