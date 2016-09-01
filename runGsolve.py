"""
Created on Tue May 13 17:47:48 2014

@author: jackm and GOB
"""

from __future__ import with_statement
from distutils.core import setup

import sys
from PyQt4 import QtCore
from PyQt4 import QtGui
from Tkinter import Tk
from tkFileDialog import askopenfilename
import xlrd
import xlwt
import os.path
import numpy as np
from xlutils.copy import copy
import ctypes 
import GsolveGUIv3
import matplotlib
import matplotlib.pyplot
import gdal
gdal.UseExceptions()
import time
from matplotlib.backends.backend_qt4 import NavigationToolbar2QT as NavigationToolbar



class DesignerMainWindow(QtGui.QMainWindow, GsolveGUIv3.Ui_Gsolve):
    
    def __init__(self, parent=None):

        splash_pix = QtGui.QPixmap(r'images\\Gsolvelog.png')
        splash = QtGui.QSplashScreen(splash_pix, QtCore.Qt.WindowStaysOnTopHint)
        splash.setMask(splash_pix.mask())
        splash.show()
        app.processEvents()
        time.sleep(2)
        splash.close()
        
        super(DesignerMainWindow, self).__init__(parent)
         
        self.setupUi(self)
        self.showMaximized()
        #FullScreen()   
        self.connectActions()
        self.setWindowIcon(QtGui.QIcon("images\\gsolve.png"))
        # Initialise Calibration Table Data with data base
        AllwbCal = xlrd.open_workbook(os.path.join('Calibration Tables/G meter Calibration Table.xls'))
        
        CalibrationTableName = AllwbCal.sheet_names()
        NCaltabs = np.size(CalibrationTableName)

        for i in range(0, NCaltabs):
            self.CalibrationTabSelectview.addItem(str(CalibrationTableName[i]))
            self.CalTabSelect.addItem(str(CalibrationTableName[i]))

        CalTab = AllwbCal.sheet_by_index(0)
        N = CalTab.nrows
        self.CalibrationTabView.setRowCount(N)
        self.CalibrationTabView.setColumnCount(3)
        
        for i in range(0,N):
            self.CalibrationTabView.setItem(i, 0,
                        QtGui.QTableWidgetItem(str(CalTab.cell(i,0).value)))
            self.CalibrationTabView.setItem(i, 1,
                        QtGui.QTableWidgetItem(str(CalTab.cell(i,1).value)))
            self.CalibrationTabView.setItem(i, 2,
                        QtGui.QTableWidgetItem(str(CalTab.cell(i,2).value)))
        # Initialise Absolute Gravity Stations with data base
        ABSGw = xlrd.open_workbook(os.path.join('Absolute Gravity/Absolute Gravity.xls'))  
        AbsGDATA = ABSGw.sheet_by_name('Sheet1')
        N2 = AbsGDATA.nrows
        self.AbsGTabview.setRowCount(N2)

        for i in range(0, N2):
            self.AbsGTabview.setItem(i, 0,
                        QtGui.QTableWidgetItem(str(AbsGDATA.cell(i,0).value)))
            self.AbsGTabview.setItem(i, 1,
                        QtGui.QTableWidgetItem(str(AbsGDATA.cell(i,1).value)))
            self.AbsGTabview.setItem(i, 2,
                        QtGui.QTableWidgetItem(str(AbsGDATA.cell(i,2).value)))
            self.AbsGTabview.setItem(i, 3,
                        QtGui.QTableWidgetItem(str(AbsGDATA.cell(i,3).value)))
        

    def main(self):
        self.show()      

         
    def connectActions(self):
        self.ImportsureveyData.clicked.connect(self.Import_survey_file) 
        self.ImportNewCalTab.clicked.connect(self.Import_Cal_Tab) 
        self.CalibrationTabSelectview.currentIndexChanged.connect(self.View_Cal_Tab) 
        self.RemoveCalTab.clicked.connect(self.Remove_Cal_Tab) 
        self.AddNewAbsTie.clicked.connect(self.AddNewAbsTieForSolve) 
        self.CalculateBeta.stateChanged.connect(self.USE_BETA)
        self.ImportNewAbsG.clicked.connect(self.Import_AbsG_data) 
        self.RemoveAbsG.clicked.connect(self.Remove_AbsG_data)
        self.SolveAll.clicked.connect(self.SolveIT)
        self.UpdateEll.clicked.connect(self.SolveIT)
        self.AddNewSurveyMeasurement.clicked.connect(self.AddMeasurement)
        self.AddNewAbsG.clicked.connect(self.AddNewAbsGTie)
        self.SaveNewAbsG.clicked.connect(self.SaveAllAbsGTie)
        self.RemoveAbsTie.clicked.connect(self.RemoveAbsTieForSurvey)
        self.ExportResults.clicked.connect(self.ExportMainResults)
        self.ExportDetailedResults.clicked.connect(self.ExportDetailedResultstoEx)
        self.ExportGravityreductions.clicked.connect(self.ExportDetailedResultstoEx)
        self.RemoveSurveyMeasurement.clicked.connect(self.RemoveSurveyMeasurementsDO)      
        self.CDFplot.clicked.connect(self.CDFplots)
        self.driftplot_2.clicked.connect(self.Driftplot)
        self.actionExit.triggered.connect(self.Exitsurvey)
        self.actionNew.triggered.connect(self.Newsurvey)
        self.actionAbout.triggered.connect(self.AboutGsolve)
        self.SaveSurveyData.clicked.connect(self.SaveSurveyDatatoEx)
        self.ImportDEM.clicked.connect(self.ImportDEMdo)
        self.ImportdatapointsTC.clicked.connect(self.ImportdatapointsTCdo)
        self.CalcTC.clicked.connect(self.CalcTCdo)
        self.ExportTCresults.clicked.connect(self.ExportTCdo)

        
    def SaveSurveyDatatoEx(self):
        import tkFileDialog
        Tk().withdraw()
        filename = tkFileDialog.asksaveasfilename()
        
        w = xlwt.Workbook()
        Nrows = self.SurveyDataTab.rowCount()
        Outpage = w.add_sheet('Survey Data', cell_overwrite_ok=True)  
        Outpage2 = w.add_sheet('Locations', cell_overwrite_ok=True)   
        
        w.get_sheet(0).write(0,0,str('Station Name'))
        w.get_sheet(0).write(0,1,str('Day'))      
        w.get_sheet(0).write(0,2,str('Month'))       
        w.get_sheet(0).write(0,3,str('Year'))
        w.get_sheet(0).write(0,4,str('Hour'))
        w.get_sheet(0).write(0,5,str('Minute'))
        w.get_sheet(0).write(0,6,str('Dial Gravity'))
        w.get_sheet(0).write(0,7,str('Loop Number'))
        
        w.get_sheet(1).write(0,0,str('Station Name'))
        w.get_sheet(1).write(0,1,str('Latitude'))
        w.get_sheet(1).write(0,2,str('Longitude'))
        w.get_sheet(1).write(0,3,str('Elevation'))
        SurveyedNames = strs = ["" for x in range(Nrows)]
        
        for i in range(0,Nrows):
            Name = self.SurveyDataTab.item(i,0)
            Day = self.SurveyDataTab.item(i,5)
            Month = self.SurveyDataTab.item(i,6)
            Year = self.SurveyDataTab.item(i,7)
            Hour = self.SurveyDataTab.item(i,8)
            Minute = self.SurveyDataTab.item(i,9)
            Dial = self.SurveyDataTab.item(i,4)
            Loop = self.SurveyDataTab.item(i,10)
            w.get_sheet(0).write(int(i+1),0,str(QtGui.QTableWidgetItem.text(Name)))        
            w.get_sheet(0).write(int(i+1),1,float(str(QtGui.QTableWidgetItem.text(Day))))
            w.get_sheet(0).write(int(i+1),2,float(str(QtGui.QTableWidgetItem.text(Month))))        
            w.get_sheet(0).write(int(i+1),3,float(str(QtGui.QTableWidgetItem.text(Year))))
            w.get_sheet(0).write(int(i+1),4,float(str(QtGui.QTableWidgetItem.text(Hour))))
            w.get_sheet(0).write(int(i+1),5,float(str(QtGui.QTableWidgetItem.text(Minute))))
            w.get_sheet(0).write(int(i+1),6,float(str(QtGui.QTableWidgetItem.text(Dial))))
            w.get_sheet(0).write(int(i+1),7,float(str(QtGui.QTableWidgetItem.text(Loop))))  
            SurveyedNames[i] = str(QtGui.QTableWidgetItem.text(Name))
            
        SiteNames = np.unique(SurveyedNames)
        Nlocs = np.max(np.shape(SiteNames))
        print(Nlocs)
        
        for i in range(0,Nlocs):
            for j in range(0,Nrows):
                Namesj = self.SurveyDataTab.item(j,0)
                Namej = str(QtGui.QTableWidgetItem.text(Namesj))
                Latsj = self.SurveyDataTab.item(j,1)
                Latj = float(str(QtGui.QTableWidgetItem.text(Latsj)))
                Longsj = self.SurveyDataTab.item(j,2)
                Longj = float(str(QtGui.QTableWidgetItem.text(Longsj)))
                Elevsj = self.SurveyDataTab.item(j,3)
                Elevj = float(str(QtGui.QTableWidgetItem.text(Elevsj)))
                if Namej == SiteNames[i]:
                    w.get_sheet(1).write(int(i+1),0,Namej)
                    w.get_sheet(1).write(int(i+1),1,Latj)
                    w.get_sheet(1).write(int(i+1),2,Longj)
                    w.get_sheet(1).write(int(i+1),3,Elevj)
                    
        w.save(filename+'.xls')              

    
    def AboutGsolve(self):
        sw = ctypes.windll.user32.MessageBoxA(0,'Gsolve, python version 1.'+ \
        '\n'+'\n'+'Created by Jack McCubbine and Grant O\'Brien,'+'\n'+ \
        'May 2014-2016.',\
        'About', 0)
        print sw

              
    def Newsurvey(self):
        sw = ctypes.windll.user32.MessageBoxA(0, 'Are you sure you want to \
        start a new survey? This will clear all current survey data and \
        results.', 'New', 1)
        
        if sw == 1:
            #raise SystemExit 
            self.DetailedResultsTabview.setRowCount(0)
            self.MainResultsTabview.setRowCount(0)
            self.DriftTabview.setRowCount(0)
            self.SurveyDataTab.setRowCount(0)
            self.SurveyTies.setRowCount(0)
            self.GravityReductionsTab.setRowCount(0)
 
           
    def Exitsurvey(self):
        sw = ctypes.windll.user32.MessageBoxA(0, 'Are you sure you want to \
        exit?', 'Exit', 1)
        
        if sw == 1:
            #raise SystemExit
            print "Goodbye"
            os._exit(1)
            
            
###############################################################################            
    def Driftplot(self):
        self.BL = []
        self.TIM = []
        
        global BL #this imports the base lines and
        global TIM # and time in minutes
        
        Loopp = self.LoopSel_2.currentText()
        Nrows = self.DetailedResultsTabview.rowCount()
        if Loopp == 'All':
            BLp = BL
            CDialp = np.zeros((int(Nrows), 1))  
            TIDEp = np.zeros((int(Nrows), 1))  
            TIMp = np.zeros((int(Nrows), 1))
            LoopNamesp = strs = ["" for x in range(int(Nrows))]
            k2=-1
            for i in range(0, int(Nrows)):
                k2 = k2+1
                CDial = self.DetailedResultsTabview.item(i, 5)
                TIDE = self.DetailedResultsTabview.item(i, 11)
                LoopNames = self.DetailedResultsTabview.item(i, 0)
                    
                LoopNamesp[k2] = str(QtGui.QTableWidgetItem.text(LoopNames))
                CDialp[k2] = float(QtGui.QTableWidgetItem.text(CDial))
                TIDEp[k2] = float(QtGui.QTableWidgetItem.text(TIDE))
                TIMp[k2] = TIM[i]
                
            NrowsAll = self.MainResultsTabview.rowCount()
            GRAV = np.zeros((Nrows, 1))
            
            for i in range(0, Nrows):
                for j in range(0, int(NrowsAll)):
                    NameAllj = self.MainResultsTabview.item(j, 0)
                    Grav = self.MainResultsTabview.item(j, 1)
                    NameAlljreal = str(QtGui.QTableWidgetItem.text(NameAllj))
                    GravAlljreal = float(QtGui.QTableWidgetItem.text(Grav))
                    if NameAlljreal == LoopNamesp[i]:
                        GRAV[i] = GravAlljreal
                        
            Drift = self.DriftTabview.item(0, 1)  
            Driftp = -float(QtGui.QTableWidgetItem.text(Drift)) 
            
            Betat = self.BetaOut
            Beta = -float(QtGui.QLineEdit.text(Betat))
            Betatin = self.BetaIne
            Betain = -float(QtGui.QLineEdit.text(Betatin))
            x = self.CalculateBeta        
#            if Betain==Beta and QtGui.QCheckBox.isChecked(x)==False:
#                CDialp=CDialp-Beta*CDialp
#                Beta=0
            fig = matplotlib.pyplot.figure()
            ax = fig.add_subplot(111)
            x = TIMp/60
            y = GRAV-BLp-CDialp-Beta*CDialp+TIDEp+Driftp*TIMp[0]/60
            y2 = -Driftp*TIMp/60+Driftp*TIMp[0]/60
            
            ax.plot(x, y,'x')
            ax.plot(x, y2)
            title = 'Plot of the drift function and residuals for loop '+ \
            str(Loopp)
            
            matplotlib.pyplot.xlabel('Time in hours since start of loop')
            matplotlib.pyplot.ylabel('mGal')
            ax.set_title(title) 
            matplotlib.pyplot.show()  
            
        if Loopp != 'All':
            BLp = BL[int(Loopp)-1]
            k = 0
            for i in range(0, int(Nrows)):
                Loopqt = self.DetailedResultsTabview.item(i, 12)
                Loop = str(QtGui.QTableWidgetItem.text(Loopqt))  
                if Loop == Loopp:
                    k = k+1
            CDialp = np.zeros((k, 1))  
            TIDEp = np.zeros((k, 1))  
            TIMp = np.zeros((k, 1))   
            LoopNamesp = strs = ["" for x in range(k)]
        
            k2 = -1
            for i in range(0, int(Nrows)):
                Loopqt = self.DetailedResultsTabview.item(i, 12)
                Loop = str(QtGui.QTableWidgetItem.text(Loopqt))  
                if Loop == Loopp:
                    k2 = k2+1
                    CDial = self.DetailedResultsTabview.item(i, 5)
                    TIDE = self.DetailedResultsTabview.item(i, 11)
                    LoopNames = self.DetailedResultsTabview.item(i, 0)
                    
                    LoopNamesp[k2] = str(QtGui.QTableWidgetItem.text(LoopNames))
                    CDialp[k2] = float(QtGui.QTableWidgetItem.text(CDial))
                    TIDEp[k2] = float(QtGui.QTableWidgetItem.text(TIDE))
                    TIMp[k2] = TIM[i]
            
            NrowsAll = self.MainResultsTabview.rowCount()
            GRAV = np.zeros((k, 1))
            for i in range(0, len(LoopNamesp)):
                for j in range(0, int(NrowsAll)):
                    NameAllj = self.MainResultsTabview.item(j, 0)
                    Grav = self.MainResultsTabview.item(j, 1)
                    NameAlljreal = str(QtGui.QTableWidgetItem.text(NameAllj))
                    GravAlljreal = float(QtGui.QTableWidgetItem.text(Grav))
                    if NameAlljreal == LoopNamesp[i]:
                        GRAV[i] = GravAlljreal
                        
            Drift = self.DriftTabview.item(int(Loopp)-1, 1)  
            Driftp = -float(QtGui.QTableWidgetItem.text(Drift)) 
            
            Betat = self.BetaOut
            Beta = -float(QtGui.QLineEdit.text(Betat))
            Betatin = self.BetaIne
            Betain = -float(QtGui.QLineEdit.text(Betatin))
            x = self.CalculateBeta        
            #if Betain==Beta and QtGui.QCheckBox.isChecked(x)==False:
                 #CDialp=CDialp-Beta*CDialp
                #Beta=0
            CIt = self.ConfInt
            CI = float(QtGui.QLineEdit.text(CIt))
            
            fig = matplotlib.pyplot.figure()
            ax = fig.add_subplot(111)
            x = TIMp/60
            y = GRAV-BLp-CDialp-Beta*CDialp+TIDEp+Driftp*TIMp[0]/60
            y2 = -Driftp*TIMp/60+Driftp*TIMp[0]/60
            
            ax.plot(x, y,'x')
            ax.plot(x, y2)
            title = 'Plot of the drift function and residuals for loop ' \
            +str(Loopp)
            
            matplotlib.pyplot.xlabel('Time in hours since start of loop')
            matplotlib.pyplot.ylabel('mGal')
            ax.set_title(title) 
            matplotlib.pyplot.show()
            
###############################################################################            

    def CDFplots(self):
        Loopp = self.LoopCDFSel.currentText()
        Nrows = self.DetailedResultsTabview.rowCount()
        if Loopp == 'All':
            RESp = np.zeros((Nrows, 1))
            for i in range(0,int(Nrows)):
                Residual = self.DetailedResultsTabview.item(i, 13)
                RESp[i] = float(QtGui.QTableWidgetItem.text(Residual))
            fig = matplotlib.pyplot.figure()
            ax = fig.add_subplot(111)
            x = np.sort(RESp, axis=0)
            y=np.linspace(0, 1, np.size(x))
            ax.plot(x, y)
            ax.set_title('CDF plot for all loops, mean:' \
            +str(np.round(100*np.mean(RESp))/100)+ \
            ', standard deviation:'+str(np.round(100*np.std(RESp))/100))
            matplotlib.pyplot.show()        
            
        if Loopp != 'All':
            k = 0
            for i in range(0, int(Nrows)):
                Loopqt = self.DetailedResultsTabview.item(i, 12)
                Loop = str(QtGui.QTableWidgetItem.text(Loopqt))  
                if Loop == Loopp:
                    k = k+1
            RESp = np.zeros((k, 1))            
            k2 = -1
            for i in range(0, int(Nrows)):
                Loopqt = self.DetailedResultsTabview.item(i, 12)
                Loop = str(QtGui.QTableWidgetItem.text(Loopqt))  
                if Loop == Loopp:
                    k2 = k2+1
                    Residual = self.DetailedResultsTabview.item(i, 13)
                    RESp[k2] = float(QtGui.QTableWidgetItem.text(Residual))
            fig = matplotlib.pyplot.figure()
            ax = fig.add_subplot(111)
            x = np.sort(RESp, axis=0)
            y=np.linspace(0, 1, np.size(x))
            ax.plot(x, y)
            title='CDF plot for loop '+str(Loopp)
            ax.set_title(title+', mean: ' \
            +str(np.round(100*np.mean(RESp))/100)+ \
            ', standard deviation: '+str(np.round(100*np.std(RESp))/100))
            matplotlib.pyplot.show()

        
    def RemoveSurveyMeasurementsDO(self):
        Nrows = self.SurveyDataTab.rowCount()
        k = 0;
        SurveyedNames = strs = ["" for x in range(Nrows)]
        Lats = np.zeros((Nrows, 1))
        Longs = np.zeros((Nrows, 1))
        Elevs = np.zeros((Nrows, 1))
        Days = np.zeros((Nrows, 1))
        Days = np.zeros((Nrows, 1))
        Months = np.zeros((Nrows, 1))
        Years = np.zeros((Nrows, 1))
        Hours = np.zeros((Nrows, 1))
        Minutes = np.zeros((Nrows, 1))
        DialReadings = np.zeros((Nrows, 1))
        Loops = np.zeros((Nrows, 1))
        
        for i in range(0, Nrows):            
            SurveyedName = self.SurveyDataTab.item(i, 0)
            Lat = self.SurveyDataTab.item(i, 1)
            Long = self.SurveyDataTab.item(i, 2)
            Elev = self.SurveyDataTab.item(i, 3)
            DialReading = self.SurveyDataTab.item(i, 4)
            Day = self.SurveyDataTab.item(i, 5)
            Month = self.SurveyDataTab.item(i, 6)
            Year = self.SurveyDataTab.item(i, 7)
            Hour = self.SurveyDataTab.item(i, 8)
            Minute = self.SurveyDataTab.item(i, 9)
            Loop = self.SurveyDataTab.item(i, 10)
            SurveyedNames[i] = QtGui.QTableWidgetItem.text(SurveyedName)
            Lats[i] = float(QtGui.QTableWidgetItem.text(Lat))
            Longs[i] = float(QtGui.QTableWidgetItem.text(Long))
            Elevs[i] = float(QtGui.QTableWidgetItem.text(Elev))
            DialReadings[i] = float(QtGui.QTableWidgetItem.text(DialReading))
            Days[i] = float(QtGui.QTableWidgetItem.text(Day))
            Months[i] = float(QtGui.QTableWidgetItem.text(Month))
            Years[i] = float(QtGui.QTableWidgetItem.text(Year))
            Hours[i] = float(QtGui.QTableWidgetItem.text(Hour))
            Minutes[i] = float(QtGui.QTableWidgetItem.text(Minute))
            Loops[i] = float(QtGui.QTableWidgetItem.text(Loop))
            if self.SurveyDataTab.isItemSelected(SurveyedName) == False:
                k = k+1

        self.SurveyDataTab.setRowCount(k)
        k1 = -1
        
        for i in range(0, int(Nrows)):
            x = self.SurveyDataTab.item(i, 0)
            if self.SurveyDataTab.isItemSelected(x) == False:
                k1 = k1+1
                self.SurveyDataTab.setItem(int(k1),0,
                        QtGui.QTableWidgetItem(str(SurveyedNames[i])))
                self.SurveyDataTab.setItem(int(k1),1,
                        QtGui.QTableWidgetItem(str(float(Lats[i]))))
                self.SurveyDataTab.setItem(int(k1),2,
                        QtGui.QTableWidgetItem(str(float(Longs[i]))))
                self.SurveyDataTab.setItem(int(k1),3,
                        QtGui.QTableWidgetItem(str(float(Elevs[i]))))
                self.SurveyDataTab.setItem(int(k1),4,
                        QtGui.QTableWidgetItem(str(float(DialReadings[i]))))
                self.SurveyDataTab.setItem(int(k1),5,
                        QtGui.QTableWidgetItem(str(float(Days[i]))))
                self.SurveyDataTab.setItem(int(k1),6,
                        QtGui.QTableWidgetItem(str(float(Months[i]))))
                self.SurveyDataTab.setItem(int(k1),7,
                        QtGui.QTableWidgetItem(str(float(Years[i]))))
                self.SurveyDataTab.setItem(int(k1),8,
                        QtGui.QTableWidgetItem(str(float(Hours[i]))))
                self.SurveyDataTab.setItem(int(k1),9,
                        QtGui.QTableWidgetItem(str(float(Minutes[i]))))
                self.SurveyDataTab.setItem(int(k1),10,
                        QtGui.QTableWidgetItem(str(float(Loops[i]))))

 
    def ExportDetailedResultstoEx(self):
        import tkFileDialog
        Tk().withdraw()
        filename = tkFileDialog.asksaveasfilename()
        
        w = xlwt.Workbook()
        
        Outpage = w.add_sheet('Survey Measurements',
                              cell_overwrite_ok=True)
        Outpage = w.add_sheet('Gravity Results and reductions',
                              cell_overwrite_ok=True)
        Outpage = w.add_sheet('Gravimetre Drift',
                              cell_overwrite_ok=True)   
        Outpage = w.add_sheet('Beta Calibration Factor',
                              cell_overwrite_ok=True)   
        Outpage = w.add_sheet('Meta Data About Solution',
                              cell_overwrite_ok=True)   
        
        x = self.CalculateBeta        
        if QtGui.QCheckBox.isChecked(x) == True:
            CB = 1
        else:
            CB = 0 
        x = self.UseLoops                            
        if QtGui.QCheckBox.isChecked(x) == True:
            UL = 1
        else:
            UL = 0    
        w.get_sheet(4).write(0, 0, str('Method Used:'))
        Method = self.MethodSelect.currentText()
        w.get_sheet(4).write(int(1), 0, str(Method))
        w.get_sheet(4).write(2, 0, 
                       str('Beta Calibration Factor Calculated Or Provided?:'))
        if CB == 1:
            w.get_sheet(4).write(int(3), 0, str("Calculated"))
        if CB == 0:
            w.get_sheet(4).write(int(3), 0, str("Provided"))
        w.get_sheet(4).write(4, 0, str('Loops used?:'))
        if UL == 1:
            w.get_sheet(4).write(int(5), 0, str("Yes, "))
            Nloops = self.DriftTabview.rowCount()
            w.get_sheet(4).write(int(5), 1, str(str(Nloops)+ " Loops Used"))
        if UL == 0:
            w.get_sheet(4).write(int(5), 0, str("No"))

        w.get_sheet(4).write(6, 0, str('Ellipsoid Used for reductions:'))
        
        ELLname = self.Ellipsoidselect.currentText()
        w.get_sheet(4).write(7, 0, str(ELLname))
        w.get_sheet(3).write(0, 0, str('Beta Calibration factor'))

        Betat = self.BetaOut
        Beta = float(QtGui.QLineEdit.text(Betat))
        w.get_sheet(3).write(int(1), 0, str(Beta))        
            
        Nrows=self.DriftTabview.rowCount()
        w.get_sheet(2).write(0, 0, str('Loop ID'))
        w.get_sheet(2).write(0, 1, str('Drift (mGal per hour)'))

        for i in range(0, Nrows):
            Name = self.DriftTabview.item(i, 0)
            Drift = self.DriftTabview.item(i, 1)
         #   Weight=self.DetailedResultsTabview.item(i,13)
            w.get_sheet(2).write(int(i+1), 0,
                        str(QtGui.QTableWidgetItem.text(Name)))        
            w.get_sheet(2).write(int(i+1), 1,
                        float(str(QtGui.QTableWidgetItem.text(Drift))))        
        
        Nrows = self.DetailedResultsTabview.rowCount()
        
        w.get_sheet(0).write(0, 0, str('Name'))
        w.get_sheet(0).write(0, 1, str('Latitude'))      
        w.get_sheet(0).write(0, 2, str('Longitude'))       
        w.get_sheet(0).write(0, 3, str('Elevation'))     
        w.get_sheet(0).write(0, 4, str('Dial'))
        w.get_sheet(0).write(0, 5, str('Calibrated Dial'))
        w.get_sheet(0).write(0, 6, str('Day'))
        w.get_sheet(0).write(0, 7, str('Month'))
        w.get_sheet(0).write(0, 8, str('Year'))
        w.get_sheet(0).write(0, 9, str('Hour'))
        w.get_sheet(0).write(0, 10, str('Minute'))
        w.get_sheet(0).write(0, 11, str('Tidal Effect'))
        w.get_sheet(0).write(0, 12, str('Loop'))
        w.get_sheet(0).write(0, 13, str('Residual'))
       # w.get_sheet(0).write(0,13,str('Weight'))          

        for i in range(0, Nrows):
            Name = self.DetailedResultsTabview.item(i,0)
            Latitude = self.DetailedResultsTabview.item(i,1)
            Longitude = self.DetailedResultsTabview.item(i,2)
            Elev = self.DetailedResultsTabview.item(i,3)
            Dial = self.DetailedResultsTabview.item(i,4)
            CalibratedDial = self.DetailedResultsTabview.item(i,5)
            Day = self.DetailedResultsTabview.item(i,6)
            Month = self.DetailedResultsTabview.item(i,7)
            Year = self.DetailedResultsTabview.item(i,8)
            Hour = self.DetailedResultsTabview.item(i,9)
            Minute = self.DetailedResultsTabview.item(i,10)
            Tidal = self.DetailedResultsTabview.item(i,11)
            Loop = self.DetailedResultsTabview.item(i,12)
            Residual = self.DetailedResultsTabview.item(i,13)
         #   Weight=self.DetailedResultsTabview.item(i,13)
            w.get_sheet(0).write(int(i+1),0,
                        str(QtGui.QTableWidgetItem.text(Name)))        
            w.get_sheet(0).write(int(i+1),1,
                        float(str(QtGui.QTableWidgetItem.text(Latitude))))        
            w.get_sheet(0).write(int(i+1),2,
                        float(str(QtGui.QTableWidgetItem.text(Longitude))))        
            w.get_sheet(0).write(int(i+1),3,
                        float(str(QtGui.QTableWidgetItem.text(Elev))))            
            w.get_sheet(0).write(int(i+1),4,
                        float(str(QtGui.QTableWidgetItem.text(Dial))))            
            w.get_sheet(0).write(int(i+1),5,
                        float(str(QtGui.QTableWidgetItem.text(CalibratedDial))))            
            w.get_sheet(0).write(int(i+1),6,
                        float(str(QtGui.QTableWidgetItem.text(Day))))            
            w.get_sheet(0).write(int(i+1),7,
                        float(str(QtGui.QTableWidgetItem.text(Month))))            
            w.get_sheet(0).write(int(i+1),8,
                        float(str(QtGui.QTableWidgetItem.text(Year))))            
            w.get_sheet(0).write(int(i+1),9,
                        float(str(QtGui.QTableWidgetItem.text(Hour))))            
            w.get_sheet(0).write(int(i+1),10,
                        float(str(QtGui.QTableWidgetItem.text(Minute))))          
            w.get_sheet(0).write(int(i+1),11,
                        float(str(QtGui.QTableWidgetItem.text(Tidal))))          
            w.get_sheet(0).write(int(i+1),12,
                        float(str(QtGui.QTableWidgetItem.text(Loop))))          
            w.get_sheet(0).write(int(i+1),13,
                        float(str(QtGui.QTableWidgetItem.text(Residual))))      
          #  w.get_sheet(0).write(int(i+1),13,
          #              str(QtGui.QTableWidgetItem.text(Weight)))                           

        Nrows = self.MainResultsTabview.rowCount()
        w.get_sheet(1).write(0, 0, str('Name'))
        w.get_sheet(1).write(0, 1, str('Absolute Gravity'))      
        w.get_sheet(1).write(0, 2, str('Standard Error'))       
        w.get_sheet(1).write(0, 3, str('Number of Observations'))
        w.get_sheet(1).write(0, 4, str('Latitude'))
        w.get_sheet(1).write(0, 5, str('Longitude'))
        w.get_sheet(1).write(0, 6, str('Elevation'))
        w.get_sheet(1).write(0, 7, str('Ellipsoidal gravity'))
        w.get_sheet(1).write(0, 8, str('Free air effect'))
        w.get_sheet(1).write(0, 9, str('Free air anomaly'))
        
        for i in range(0, Nrows):
            Name = self.MainResultsTabview.item(i,0)
            Grav = self.MainResultsTabview.item(i,1)
            SE = self.MainResultsTabview.item(i,2)
            NO = self.MainResultsTabview.item(i,3)
            Lat = self.GravityReductionsTab.item(i,2)
            Long = self.GravityReductionsTab.item(i,3)
            Elev = self.GravityReductionsTab.item(i,4)
            Ellg = self.GravityReductionsTab.item(i,5)
            FAeff = self.GravityReductionsTab.item(i,6)
            FAanom = self.GravityReductionsTab.item(i,7)
            w.get_sheet(1).write(int(i+1),0,
                        str(QtGui.QTableWidgetItem.text(Name)))        
            w.get_sheet(1).write(int(i+1),1,
                        float(str(QtGui.QTableWidgetItem.text(Grav))))        
            w.get_sheet(1).write(int(i+1),2,
                        float(str(QtGui.QTableWidgetItem.text(SE))))        
            w.get_sheet(1).write(int(i+1),3,
                        float(str(QtGui.QTableWidgetItem.text(NO))))        
            w.get_sheet(1).write(int(i+1),4,
                        float(str(QtGui.QTableWidgetItem.text(Lat))))        
            w.get_sheet(1).write(int(i+1),5,
                        float(str(QtGui.QTableWidgetItem.text(Long))))        
            w.get_sheet(1).write(int(i+1),6,
                        float(str(QtGui.QTableWidgetItem.text(Elev))))        
            w.get_sheet(1).write(int(i+1),7,
                        float(str(QtGui.QTableWidgetItem.text(Ellg))))        
            w.get_sheet(1).write(int(i+1),8,
                        float(str(QtGui.QTableWidgetItem.text(FAeff))))        
            w.get_sheet(1).write(int(i+1),9,
                        float(str(QtGui.QTableWidgetItem.text(FAanom))))                           

        w.save(filename+'.xls')
              
        
    def ExportMainResults(self):
        import tkFileDialog
        Tk().withdraw()
        filename = tkFileDialog.asksaveasfilename()
        
        w = xlwt.Workbook()
        Nrows = self.MainResultsTabview.rowCount()
        Outpage = w.add_sheet('Absolute Gravity Results',
                              cell_overwrite_ok=True)   
        
        w.get_sheet(0).write(0,0,str('Name'))
        w.get_sheet(0).write(0,1,str('Absolute Gravity'))      
        w.get_sheet(0).write(0,2,str('Standard Error'))       
        w.get_sheet(0).write(0,3,str('Number of Observations'))
        
        for i in range(0, Nrows):
            Name = self.MainResultsTabview.item(i,0)
            Grav = self.MainResultsTabview.item(i,1)
            SE = self.MainResultsTabview.item(i,2)
            NO = self.MainResultsTabview.item(i,3)
            w.get_sheet(0).write(int(i+1),0,
                        str(QtGui.QTableWidgetItem.text(Name)))        
            w.get_sheet(0).write(int(i+1),1,
                        float(str(QtGui.QTableWidgetItem.text(Grav))))
            w.get_sheet(0).write(int(i+1),2,
                        float(str(QtGui.QTableWidgetItem.text(SE))))        
            w.get_sheet(0).write(int(i+1),3,
                        float(str(QtGui.QTableWidgetItem.text(NO))))
                           
        w.save(filename+'.xls')
     
           
    def RemoveAbsTieForSurvey(self):
        Nrows = self.SurveyTies.rowCount()
        k = 0;
        Gravs = np.zeros((Nrows,1))
        Names = strs = ["" for x in range(Nrows)]
        
        for i in range(0, Nrows):
            Name = self.SurveyTies.item(i,0)
            Grav = self.SurveyTies.item(i,1)
            Names[i] = QtGui.QTableWidgetItem.text(Name)
            Gravs[i] = float(QtGui.QTableWidgetItem.text(Grav))
            if self.SurveyTies.isItemSelected(Name) == False:
                k = k+1

        self.SurveyTies.setRowCount(k)
        k1 = -1       
        for i in range(0, int(Nrows)):
            x = self.SurveyTies.item(i,0)
            if self.SurveyTies.isItemSelected(x) == False:
                k1 = k1+1
                self.SurveyTies.setItem(int(k1),0, 
                                QtGui.QTableWidgetItem(str(Names[i])))
                self.SurveyTies.setItem(int(k1),1,
                                QtGui.QTableWidgetItem(str(float(Gravs[i]))))


    def SaveAllAbsGTie(self):
        Nrows = self.AbsGTabview.rowCount()
        w = xlwt.Workbook()
        CalTab2 = w.add_sheet('Sheet1', cell_overwrite_ok=True)          
        for i in range(0, Nrows):
            Name = self.AbsGTabview.item(i,0)
            Grav = self.AbsGTabview.item(i,1)
            Lat = self.AbsGTabview.item(i,2)
            Long = self.AbsGTabview.item(i,3)
            w.get_sheet(0).write(i,0,str(QtGui.QTableWidgetItem.text(Name)))        
            w.get_sheet(0).write(i,1,str(QtGui.QTableWidgetItem.text(Grav)))        
            w.get_sheet(0).write(i,2,str(QtGui.QTableWidgetItem.text(Lat)))        
            w.get_sheet(0).write(i,3,str(QtGui.QTableWidgetItem.text(Long)))
                           
        w.save('Absolute Gravity/Absolute Gravity.xls')
 
       
    def AddNewAbsGTie(self):        
        sw = ctypes.windll.user32.MessageBoxA(0,
        'A new space will be added to the Absolute Gravity database table. Remember to click save when you are finished',
        'New tie value', 1)
        if sw == 1:
            Nrows = self.AbsGTabview.rowCount()
            Nrowsnew = int(Nrows+1)
            self.AbsGTabview.setRowCount(Nrowsnew)
        
    
    def AddMeasurement(self):
        Nrows = self.SurveyDataTab.rowCount()
        Nrowsnew = int(Nrows+1)
        self.SurveyDataTab.setRowCount(Nrowsnew)


######################################################################################################
####################################  CORE CALCULATION ###############################################
######################################################################################################

    def SolveIT(self):
        QtGui.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
        # Get the survey data from the table
        Nmeasurements = self.SurveyDataTab.rowCount()
        Days = np.zeros(Nmeasurements) 
        Months = np.zeros(Nmeasurements)
        Years = np.zeros(Nmeasurements)
        Hours = np.zeros(Nmeasurements)
        Minutes = np.zeros(Nmeasurements)
        MeasurementLats = np.zeros(Nmeasurements)
        MeasurementLongs = np.zeros(Nmeasurements)
        MeasurementElevs = np.zeros(Nmeasurements)
        DialReadings = np.zeros(Nmeasurements)
        SurveyedNames = strs = ["" for x in range(Nmeasurements)]
        Loops = np.zeros(Nmeasurements)
        TidalEffects = np.zeros(Nmeasurements)
        TIM = np.zeros(Nmeasurements)
        
        for i in range (0, Nmeasurements):       
            SurveyedNamesi=self.SurveyDataTab.item(i,0)
            Latsi=self.SurveyDataTab.item(i,1)
            Longsi=self.SurveyDataTab.item(i,2)
            Elevsi=self.SurveyDataTab.item(i,3)
            DialReadingsi=self.SurveyDataTab.item(i,4)
            Daysi=self.SurveyDataTab.item(i,5)
            Monthsi=self.SurveyDataTab.item(i,6)
            Yearsi=self.SurveyDataTab.item(i,7)
            Hoursi=self.SurveyDataTab.item(i,8)
            Minutesi=self.SurveyDataTab.item(i,9)
            Loopsi=self.SurveyDataTab.item(i,10)

            SurveyedNames[i]=str(QtGui.QTableWidgetItem.text(SurveyedNamesi))
            MeasurementLats[i]=float(QtGui.QTableWidgetItem.text(Latsi))
            MeasurementLongs[i]=float(QtGui.QTableWidgetItem.text(Longsi))
            MeasurementElevs[i]=float(QtGui.QTableWidgetItem.text(Elevsi))
            DialReadings[i]=float(QtGui.QTableWidgetItem.text(DialReadingsi))
            Days[i]=float(QtGui.QTableWidgetItem.text(Daysi))
            Months[i]=float(QtGui.QTableWidgetItem.text(Monthsi))
            Years[i]=float(QtGui.QTableWidgetItem.text(Yearsi))
            Hours[i]=float(QtGui.QTableWidgetItem.text(Hoursi))
            Minutes[i]=float(QtGui.QTableWidgetItem.text(Minutesi))
            Loops[i]=float(QtGui.QTableWidgetItem.text(Loopsi))
        
        SiteNames=np.unique(SurveyedNames)
        # Get the tie data from the table
        Nties=self.SurveyTies.rowCount()
        TIENames=strs = ["" for x in range(Nties)]
        TIEVals=np.zeros(Nties)
        for i in range (0,Nties):       
            
            TIENamesi=self.SurveyTies.item(i,0)
            TIEValsi=self.SurveyTies.item(i,1)
            TIENames[i]=str(QtGui.QTableWidgetItem.text(TIENamesi))
            TIEVals[i]=float(QtGui.QTableWidgetItem.text(TIEValsi))
            
        # Calculate tidal effects TIDAL.TIDEFF(Latitude,Longitude,Day,Month,Year,Hour,Minute)
        # And calculate the time in minutes for the drift analysis
        import TIDAL
        for i in range (0,Nmeasurements):
            TidalEffects[i]=TIDAL.TIDEFF(MeasurementLats[i],MeasurementLongs[i],int(Days[i]),int(Months[i]),int(Years[i]),Hours[i],Minutes[i])    
            [JTi,TIMi]=TIDAL.TIMANDJT(int(Days[i]),int(Months[i]),int(Years[i]),Hours[i],Minutes[i])
            TIM[i]=TIMi
        # get calibration table
        wbCal =xlrd.open_workbook('Calibration Tables/G meter Calibration Table.xls')
        Sname=self.CalTabSelect.currentText()
        CalTab=wbCal.sheet_by_name(Sname)
        import CalibrateDial
        CalDials=CalibrateDial.Calibrate(DialReadings,CalTab)
        
        # Calculate tide free calibrated dial readings
        TideFreeCalDial=CalDials-TidalEffects
        # Calculate Absolute gravity at each site and the drift parameters and beta factor
        #self.UseLoops.
        x=self.UseLoops
        import GSOLVER
        if QtGui.QCheckBox.isChecked(x)==True:
            UL=1
        else:
            UL=0    
        Method=self.MethodSelect.currentText()
        x=self.CalculateBeta        
        if QtGui.QCheckBox.isChecked(x)==True:
            CB=1
        else:
            CB=0    
        
        if Method=='Method 1: Normal Least Squares' and CB==0:
            Betat=self.BetaIne
            Beta=float(QtGui.QLineEdit.text(Betat))
            TideFreeCalDial=TideFreeCalDial-np.dot(Beta,CalDials)
            [ABSG,RES,SiteVar,Drift,BL]=GSOLVER.GsolverM1(TideFreeCalDial,SurveyedNames,SiteNames,TIM,TIENames,TIEVals,Loops,UL)
            M=1
            self.BetaOut.setText(str(float(Beta)))
        if Method=='Method 1: Normal Least Squares' and CB==1:
            [ABSG,RES,SiteVar,Beta,Drift,BL]=GSOLVER.GsolverM1Beta(TideFreeCalDial,SurveyedNames,SiteNames,TIM,TIENames,TIEVals,Loops,UL,CalDials)
            M=1
            self.BetaOut.setText(str(float(Beta)))
        if Method=='Method 2: Decoupled Least Squares' and CB==0:
            Betat=self.BetaIne
            Beta=float(QtGui.QLineEdit.text(Betat))
            TideFreeCalDial=TideFreeCalDial-np.dot(Beta,CalDials)
            [ABSG,RES,SiteVar,Drift,BL]=GSOLVER.GsolverM2(TideFreeCalDial,SurveyedNames,SiteNames,TIM,TIENames,TIEVals,Loops,UL)
            M=2
            self.BetaOut.setText(str(float(Beta)))
        if Method=='Method 2: Decoupled Least Squares' and CB==1:
            [ABSG,RES,SiteVar,Beta,Drift,BL]=GSOLVER.GsolverM2Beta(TideFreeCalDial,SurveyedNames,SiteNames,TIM,TIENames,TIEVals,Loops,UL,CalDials)
            M=2
            self.BetaOut.setText(str(float(Beta)))
        if Method=='Method 3: Constrained Least Squares' and CB==0:
            Betat=self.BetaIne
            Beta=float(QtGui.QLineEdit.text(Betat))
            TideFreeCalDial=TideFreeCalDial-np.dot(Beta,CalDials)
            M=3
            [ABSG,RES,SiteVar,Drift,BL]=GSOLVER.GsolverM3(TideFreeCalDial,SurveyedNames,SiteNames,TIM,TIENames,TIEVals,Loops,UL)
            self.BetaOut.setText(str(float(Beta)))
        if Method=='Method 3: Constrained Least Squares' and CB==1:
            M=3
            [ABSG,RES,SiteVar,Beta,Drift,BL]=GSOLVER.GsolverM3Beta(TideFreeCalDial,SurveyedNames,SiteNames,TIM,TIENames,TIEVals,Loops,UL,CalDials)
            self.BetaOut.setText(str(float(Beta)))
##########################################################################################################
        # From the resdiuals removed those outside of the specified confidence interval and run again.
#########################################################################################################        
        RESobs=np.zeros(Nmeasurements)
        for k in range(0,Nmeasurements):
            RESobs[k]=RES[k]
        CIqt=self.ConfInt
        CI=float(QtGui.QLineEdit.text(CIqt))

        L=np.percentile(RESobs,(100-CI)/2)
        M=np.percentile(RESobs,100-(100-CI)/2)
        Gcount=0
        
        for i in range(0,Nmeasurements):
            if RESobs[i]>=L and RESobs[i]<=M:
                Gcount=Gcount+1
                
        TideFreeCalDial2=np.zeros(Gcount)
        SurveyedNames2=strs = ["" for x in range(Gcount)]
        Loops2=np.zeros(Gcount)
        CalDials2=np.zeros(Gcount)
        TIM2=np.zeros(Gcount)
        MeasurementLats2=np.zeros(Gcount)
        MeasurementLongs2=np.zeros(Gcount)
        MeasurementElevs2=np.zeros(Gcount)
        DialReadings2=np.zeros(Gcount)
        Days2=np.zeros(Gcount)
        Months2=np.zeros(Gcount)
        Years2=np.zeros(Gcount)
        Hours2=np.zeros(Gcount)
        Minutes2=np.zeros(Gcount)
        TidalEffects2=np.zeros(Gcount)
        
        Gcount2=-1
        for i in range(0,Nmeasurements):
            if RESobs[i]>=L and RESobs[i]<=M:
                Gcount2=Gcount2+1
                TideFreeCalDial2[Gcount2]=TideFreeCalDial[i]
                SurveyedNames2[Gcount2]=SurveyedNames[i]
                Loops2[Gcount2]=Loops[i]
                CalDials2[Gcount2]=CalDials[i]
                TIM2[Gcount2]=TIM[i]
                
                MeasurementLats2[Gcount2]=MeasurementLats[i]
                MeasurementLongs2[Gcount2]=MeasurementLongs[i]
                MeasurementElevs2[Gcount2]=MeasurementElevs[i]
                DialReadings2[Gcount2]=DialReadings[i]
                Days2[Gcount2]=Days[i]
                Months2[Gcount2]=Months[i]
                Years2[Gcount2]=Years[i]
                Hours2[Gcount2]=Hours[i]
                Minutes2[Gcount2]=Minutes[i]
                TidalEffects2[Gcount2]=TidalEffects[i]
                    
        Nmeasurements2=Gcount2
        SiteNames2=np.unique(SurveyedNames2)
        Nsites2=np.size(SiteNames2)
        Nsitesorg=np.size(SiteNames)
        Nties=np.size(TIENames)
        itsok=0
        Noties=0
        for j in range(0,Nsites2):
            for k in range(0,Nties):
                if SiteNames2[j]==TIENames[k]:
                    itsok=1
        for j in range(0,Nsitesorg):
            for k in range(0,Nties):
                if SiteNames[j]==TIENames[k]:
                    Noties=1
        if Nties==0 or Noties==0:
            sw=ctypes.windll.user32.MessageBoxA(0, 'Warning: Please add at least one absolute gravity tie station.', 'Warning', 0)
            return
        if itsok==0:
            sw=ctypes.windll.user32.MessageBoxA(0, 'Warning: The are no ties in the survey data after confidence interval filtering, please increase the confidence interval and try again.', 'Warning', 0)
            return        
        if Method=='Method 1: Normal Least Squares' and CB==0:
            Betat=self.BetaIne
            Beta=float(QtGui.QLineEdit.text(Betat))
            TideFreeCalDial=TideFreeCalDial2-np.dot(Beta,CalDials2)
            [ABSG2,RES2,SiteVar2,Drift2,BL]=GSOLVER.GsolverM1(TideFreeCalDial2,SurveyedNames2,SiteNames2,TIM2,TIENames,TIEVals,Loops2,UL)
            M=1
            self.BetaOut.setText(str(float(Beta)))
        if Method=='Method 1: Normal Least Squares' and CB==1:
            [ABSG2,RES2,SiteVar2,Beta2,Drift2,BL]=GSOLVER.GsolverM1Beta(TideFreeCalDial2,SurveyedNames2,SiteNames2,TIM2,TIENames,TIEVals,Loops2,UL,CalDials2)
            M=1
            self.BetaOut.setText(str(float(-Beta2)))
        if Method=='Method 2: Decoupled Least Squares' and CB==0:
            Betat=self.BetaIne
            Beta=float(QtGui.QLineEdit.text(Betat))
            TideFreeCalDial=TideFreeCalDial2-np.dot(Beta,CalDials2)
            [ABSG2,RES2,SiteVar2,Drift2,BL]=GSOLVER.GsolverM2(TideFreeCalDial2,SurveyedNames2,SiteNames2,TIM2,TIENames,TIEVals,Loops2,UL)
            M=2
            self.BetaOut.setText(str(float(Beta)))
        if Method=='Method 2: Decoupled Least Squares' and CB==1:
            [ABSG2,RES2,SiteVar2,Beta2,Drift2,BL]=GSOLVER.GsolverM2Beta(TideFreeCalDial2,SurveyedNames2,SiteNames2,TIM2,TIENames,TIEVals,Loops2,UL,CalDials2)
            M=2
            self.BetaOut.setText(str(float(-Beta2)))
        if Method=='Method 3: Constrained Least Squares' and CB==0:
            Betat=self.BetaIne
            Beta=float(QtGui.QLineEdit.text(Betat))
            TideFreeCalDial=TideFreeCalDial2-np.dot(Beta,CalDials2)
            M=3
            [ABSG2,RES2,SiteVar2,Drift2,BL]=GSOLVER.GsolverM3(TideFreeCalDial2,SurveyedNames2,SiteNames2,TIM2,TIENames,TIEVals,Loops2,UL)
            self.BetaOut.setText(str(float(Beta)))
        if Method=='Method 3: Constrained Least Squares' and CB==1:
            M=3
            [ABSG2,RES2,SiteVar2,Beta2,Drift2,BL]=GSOLVER.GsolverM3Beta(TideFreeCalDial2,SurveyedNames2,SiteNames2,TIM2,TIENames,TIEVals,Loops2,UL,CalDials2)
            self.BetaOut.setText(str(float(-Beta2)))
        TIM=TIM2            
        global TIM
#######################################################################################################
#######################################################################################################        
        
        ELLname=self.Ellipsoidselect.currentText()
        if ELLname=='GRS80':
            Gamma=978032.67715
            k=0.001931851353
            e2=0.00669438002290
        if ELLname=='GRS67':
            Gamma=978031.84558
            k=0.001931663383
            e2=0.00669460532856
        ABSG2=np.round(1000*ABSG2)/1000
        SE=np.round(1000*np.sqrt(SiteVar2))/1000
        Nsites=np.size(SiteNames2)
        self.MainResultsTabview.setRowCount(Nsites)
        self.GravityReductionsTab.setRowCount(Nsites)
        for i in range(0,Nsites):
            self.MainResultsTabview.setItem(i, 0, QtGui.QTableWidgetItem(str(SiteNames2[i])))
            self.MainResultsTabview.setItem(i, 1, QtGui.QTableWidgetItem(str(float(ABSG2[i]))))            
            self.GravityReductionsTab.setItem(i, 0, QtGui.QTableWidgetItem(str(SiteNames2[i])))
            self.GravityReductionsTab.setItem(i, 1, QtGui.QTableWidgetItem(str(float(ABSG2[i]))))
            TU=0  
            for j in range(0,Nties):
                if SiteNames[i]==TIENames[j] and M==3:
                    TU=1
            if TU==1:
                self.MainResultsTabview.setItem(i, 2, QtGui.QTableWidgetItem(str(np.sqrt(0))))
            else:
                self.MainResultsTabview.setItem(i, 2, QtGui.QTableWidgetItem(str(SE[i])))
            count=0
            for j in range(0,Nmeasurements2):
                if SiteNames2[i]==SurveyedNames2[j]:
                    count=count+1
                if SiteNames2[i]==SurveyedNames2[j]:
                    self.GravityReductionsTab.setItem(i, 2, QtGui.QTableWidgetItem(str(float(MeasurementLats2[j]))))
                    self.GravityReductionsTab.setItem(i, 3, QtGui.QTableWidgetItem(str(float(MeasurementLongs2[j]))))
                    self.GravityReductionsTab.setItem(i, 4, QtGui.QTableWidgetItem(str(float(MeasurementElevs2[j]))))
                    Latj=MeasurementLats2[j]
                    Ellj=MeasurementElevs2[j]
            Ellg=Gamma*((1+k*np.sin(Latj*np.pi/180)**2)/np.sqrt(1-e2*np.sin(Latj*np.pi/180)**2))
            FAEff=-0.3086*Ellj               
            self.GravityReductionsTab.setItem(i, 5, QtGui.QTableWidgetItem(str(float(Ellg))))
            self.GravityReductionsTab.setItem(i, 6, QtGui.QTableWidgetItem(str(float(FAEff))))
            self.GravityReductionsTab.setItem(i, 7, QtGui.QTableWidgetItem(str(float(ABSG2[i]-FAEff-Ellg))))
            self.MainResultsTabview.setItem(i, 3, QtGui.QTableWidgetItem(str(int(count))))
        
        self.DetailedResultsTabview.setRowCount(Nmeasurements2+1)    
        for i in range(0,Nmeasurements2+1):
            self.DetailedResultsTabview.setItem(i, 0, QtGui.QTableWidgetItem(str(SurveyedNames2[i])))
            self.DetailedResultsTabview.setItem(i, 1, QtGui.QTableWidgetItem(str(MeasurementLats2[i])))
            self.DetailedResultsTabview.setItem(i, 2, QtGui.QTableWidgetItem(str(MeasurementLongs2[i])))
            self.DetailedResultsTabview.setItem(i, 3, QtGui.QTableWidgetItem(str(MeasurementElevs2[i])))
            self.DetailedResultsTabview.setItem(i, 4, QtGui.QTableWidgetItem(str(DialReadings2[i])))
            self.DetailedResultsTabview.setItem(i, 5, QtGui.QTableWidgetItem(str(np.round(CalDials2[i]*1000)/1000)))
            self.DetailedResultsTabview.setItem(i, 6, QtGui.QTableWidgetItem(str(int(Days2[i]))))
            self.DetailedResultsTabview.setItem(i, 7, QtGui.QTableWidgetItem(str(int(Months2[i]))))
            self.DetailedResultsTabview.setItem(i, 8, QtGui.QTableWidgetItem(str(int(Years2[i]))))
            self.DetailedResultsTabview.setItem(i, 9, QtGui.QTableWidgetItem(str(int(Hours2[i]))))
            self.DetailedResultsTabview.setItem(i, 10, QtGui.QTableWidgetItem(str(int(Minutes2[i]))))
            self.DetailedResultsTabview.setItem(i, 11, QtGui.QTableWidgetItem(str(TidalEffects2[i])))
            self.DetailedResultsTabview.setItem(i, 12, QtGui.QTableWidgetItem(str(int(Loops2[i]))))
            self.DetailedResultsTabview.setItem(i, 13, QtGui.QTableWidgetItem(str(float(RES2[i]))))
            
            #self.MainResultsTabview.setItem(i, 3, QtGui.QTableWidgetItem(str()))
        Nloops=np.max(np.unique(Loops2))
        
        if UL==1:
            self.DriftTabview.setRowCount(Nloops) 
            for i in range(0,int(Nloops)):
                self.DriftTabview.setItem(i, 0, QtGui.QTableWidgetItem(str(np.round(i+1))))
                self.DriftTabview.setItem(i, 1, QtGui.QTableWidgetItem(str(float(np.dot(60,Drift2[i])))))
        if UL==0:
            self.DriftTabview.setRowCount(1) 
            self.DriftTabview.setItem(0, 0, QtGui.QTableWidgetItem(str(1)))
            self.DriftTabview.setItem(0, 1, QtGui.QTableWidgetItem(str(float(60*Drift2))))
            
        uloops=np.unique(Loops2)
        self.LoopCDFSel.clear()
        self.LoopCDFSel.addItem(str('All'))
        if UL==1:
            self.LoopSel_2.clear()
            for i in range(0,int(Nloops)):
                self.LoopCDFSel.addItem(str(int(uloops[i])))
                self.LoopSel_2.addItem(str(int(uloops[i])))
        if UL==0:
            self.LoopSel_2.clear()
            self.LoopSel_2.addItem(str('All'))
            
        global BL
        QtGui.QApplication.restoreOverrideCursor()
######################################################################################################
############################## End of CORE CALCULATION ###############################################
######################################################################################################
            
    def Remove_AbsG_data(self):
        
        Nrows=self.AbsGTabview.rowCount()
        k=0;
        for i in range(0,Nrows):
            x=self.AbsGTabview.item(i,0)
            if self.AbsGTabview.isItemSelected(x):
                k=k+1
                Names=QtGui.QTableWidgetItem.text(x)
        Names=strs = ["" for x in range(k)]
        k=-1
        for i in range(0,Nrows):
            x=self.AbsGTabview.item(i,0)
            if self.AbsGTabview.isItemSelected(x):
                k=k+1
                Names[k]=str(QtGui.QTableWidgetItem.text(x))
        
        sw=ctypes.windll.user32.MessageBoxA(0, 'Are you sure you want to remove these ties from the database', 'Warning', 4)
        if sw==6:
            
            # Update Database of Calibration Tables
            w1 =xlrd.open_workbook('Absolute Gravity/Absolute Gravity.xls')
            w=xlwt.Workbook()
            CalTab2=w.add_sheet('Sheet1',cell_overwrite_ok=True)
            k=-1
            for i in range(0,Nrows):
                R=0                   
                x=self.AbsGTabview.item(i,0)
                for j in range(0,np.size(Names)):
                          if Names[j]==str(QtGui.QTableWidgetItem.text(x)):
                              R=1
                if R==0:
                    Name=self.AbsGTabview.item(i,0)
                    Grav=self.AbsGTabview.item(i,1)
                    Lat=self.AbsGTabview.item(i,2)
                    Long=self.AbsGTabview.item(i,3)
                    k=k+1
                    w.get_sheet(0).write(k,0,str(QtGui.QTableWidgetItem.text(Name)))        
                    w.get_sheet(0).write(k,1,str(QtGui.QTableWidgetItem.text(Grav)))        
                    w.get_sheet(0).write(k,2,str(QtGui.QTableWidgetItem.text(Lat)))        
                    w.get_sheet(0).write(k,3,str(QtGui.QTableWidgetItem.text(Long)))        
                    
            w.save('Absolute Gravity/Absolute Gravity.xls')
        # Initialise Absolute Gravity Stations with data base
        ABSGw = xlrd.open_workbook(os.path.join('Absolute Gravity/Absolute Gravity.xls'))  
        AbsGDATA=ABSGw.sheet_by_name('Sheet1')
        N2=AbsGDATA.nrows
        self.AbsGTabview.setRowCount(N2)

        for i in range(0,N2):
            self.AbsGTabview.setItem(i, 0, QtGui.QTableWidgetItem(str(AbsGDATA.cell(i,0).value)))
            self.AbsGTabview.setItem(i, 1, QtGui.QTableWidgetItem(str(AbsGDATA.cell(i,1).value)))
            self.AbsGTabview.setItem(i, 2, QtGui.QTableWidgetItem(str(AbsGDATA.cell(i,2).value)))
            self.AbsGTabview.setItem(i, 3, QtGui.QTableWidgetItem(str(AbsGDATA.cell(i,3).value)))

            
    def Import_AbsG_data(self):
        Tk().withdraw()
        filenameCal = askopenfilename(title='Select Absolute Gravity data file')
        wbAbsG = xlrd.open_workbook(os.path.join(filenameCal))
        
        # Update Database of Calibration Tables
        w1 =xlrd.open_workbook('Absolute Gravity/Absolute Gravity.xls')
        w=copy(w1)
        
        AllImportDataTab=wbAbsG.sheet_by_index(0) 
        Wsheet=w1.sheet_by_name('Sheet1')        
        N=AllImportDataTab.nrows
        NAC=Wsheet.nrows
        NACi=NAC-1
        for i in range(NAC,NAC+N):
            sw=0
            rj=-1
            for j in range(0,NAC):
                if AllImportDataTab.cell(i-NAC,0).value==Wsheet.cell(j,0).value:
                    sw=ctypes.windll.user32.MessageBoxA(0,'Station in import file already exist in the data base, over-write?', 'Warning', 1)
                    rj=j
            if sw==1:
                w.get_sheet(0).write(rj,0,AllImportDataTab.cell(i-NAC,0).value)        
                w.get_sheet(0).write(rj,1,AllImportDataTab.cell(i-NAC,1).value)        
                w.get_sheet(0).write(rj,2,AllImportDataTab.cell(i-NAC,2).value)        
                w.get_sheet(0).write(rj,3,AllImportDataTab.cell(i-NAC,3).value)        
            if sw==0:
                NACi=int(NACi+1)
                w.get_sheet(0).write(NACi,0,AllImportDataTab.cell(i-NAC,0).value)        
                w.get_sheet(0).write(NACi,1,AllImportDataTab.cell(i-NAC,1).value)        
                w.get_sheet(0).write(NACi,2,AllImportDataTab.cell(i-NAC,2).value)        
                w.get_sheet(0).write(NACi,3,AllImportDataTab.cell(i-NAC,3).value)        
        w.save('Absolute Gravity/Absolute Gravity.xls')
        # Initialise Absolute Gravity Stations with data base
        ABSGw = xlrd.open_workbook(os.path.join('Absolute Gravity/Absolute Gravity.xls'))  
        AbsGDATA=ABSGw.sheet_by_name('Sheet1')
        N2=AbsGDATA.nrows
        self.AbsGTabview.setRowCount(N2)

        for i in range(0,N2):
            self.AbsGTabview.setItem(i, 0, QtGui.QTableWidgetItem(str(AbsGDATA.cell(i,0).value)))
            self.AbsGTabview.setItem(i, 1, QtGui.QTableWidgetItem(str(AbsGDATA.cell(i,1).value)))
            self.AbsGTabview.setItem(i, 2, QtGui.QTableWidgetItem(str(AbsGDATA.cell(i,2).value)))
            self.AbsGTabview.setItem(i, 3, QtGui.QTableWidgetItem(str(AbsGDATA.cell(i,3).value)))


    def USE_BETA(self):
        x=self.CalculateBeta.isChecked()
        if x==True:
            N=self.SurveyTies.rowCount()
            if N<2:
                ctypes.windll.user32.MessageBoxA(0, 'Must be using at least 2 tie stations to calculate the beta calibration factor.', 'Warning', 0)
                self.CalculateBeta.setCheckState(0)
 
               
    def Import_survey_file(self):
        global CalTab 
        # Select Survey file and import data
        Tk().withdraw()
        filename = askopenfilename(title='Select Survey Data')
        wb = xlrd.open_workbook(os.path.join(filename))
        SURVEYDATA=wb.sheet_by_name('Survey Data')
        Locations=wb.sheet_by_name('Locations')
        Nlocations=Locations.nrows-1
        Lats=np.zeros(Nlocations)
        Longs=np.zeros(Nlocations)
        Elev=np.zeros(Nlocations)
        SiteNames=strs = ["" for x in range(Nlocations)]
        for i in range (0,Nlocations):
            SiteNames[i]=Locations.cell(i+1,0).value
            Lats[i]=Locations.cell(i+1,1).value
            Longs[i]=Locations.cell(i+1,2).value
            Elev[i]=Locations.cell(i+1,3).value
        # Survey Data
        Nmeasurements=SURVEYDATA.nrows-1
        Days=np.zeros(Nmeasurements)
        Months=np.zeros(Nmeasurements)
        Years=np.zeros(Nmeasurements)
        Hours=np.zeros(Nmeasurements)
        Minutes=np.zeros(Nmeasurements)
        MeasurementLats=np.zeros(Nmeasurements)
        MeasurementLongs=np.zeros(Nmeasurements)
        MeasurementElev=np.zeros(Nmeasurements)
        DialReadings=np.zeros(Nmeasurements)
        Loop=np.zeros(Nmeasurements)
        SurveyedNames=strs = ["" for x in range(Nmeasurements)]
        TidalEffects=np.zeros(Nmeasurements)
        TIM=np.zeros(Nmeasurements)
        for i in range (0,Nmeasurements):
            SurveyedNames[i]=SURVEYDATA.cell(i+1,0).value
            Days[i]=SURVEYDATA.cell(i+1,1).value
            Months[i]=SURVEYDATA.cell(i+1,2).value
            Years[i]=SURVEYDATA.cell(i+1,3).value
            Hours[i]=SURVEYDATA.cell(i+1,4).value
            Minutes[i]=SURVEYDATA.cell(i+1,5).value
            DialReadings[i]=SURVEYDATA.cell(i+1,6).value
            Loop[i]=SURVEYDATA.cell(i+1,7).value
        for i in range (0,Nmeasurements):
            for j in range (0,Nlocations):
                if SurveyedNames[i]==SiteNames[j]:
                    MeasurementLats[i]=Lats[j]
                    MeasurementLongs[i]=Longs[j]
                    MeasurementElev[i]=Elev[j]
        self.SurveyDataTab.setRowCount(Nmeasurements)
        for i in range(0,Nmeasurements):
            self.SurveyDataTab.setItem(i, 0, QtGui.QTableWidgetItem(str(SurveyedNames[i])))
            self.SurveyDataTab.setItem(i, 1, QtGui.QTableWidgetItem(str(MeasurementLats[i])))
            self.SurveyDataTab.setItem(i, 2, QtGui.QTableWidgetItem(str(MeasurementLongs[i])))
            self.SurveyDataTab.setItem(i, 3, QtGui.QTableWidgetItem(str(MeasurementElev[i])))
            self.SurveyDataTab.setItem(i, 4, QtGui.QTableWidgetItem(str(DialReadings[i])))
            self.SurveyDataTab.setItem(i, 5, QtGui.QTableWidgetItem(str(int(Days[i]))))
            self.SurveyDataTab.setItem(i, 6, QtGui.QTableWidgetItem(str(int(Months[i]))))
            self.SurveyDataTab.setItem(i, 7, QtGui.QTableWidgetItem(str(int(Years[i]))))
            self.SurveyDataTab.setItem(i, 8, QtGui.QTableWidgetItem(str(int(Hours[i]))))
            self.SurveyDataTab.setItem(i, 9, QtGui.QTableWidgetItem(str(int(Minutes[i]))))
            self.SurveyDataTab.setItem(i, 10, QtGui.QTableWidgetItem(str(int(Loop[i]))))
 
       
    def Import_Cal_Tab(self):
        Tk().withdraw()
        filenameCal = askopenfilename(title='Select Calibration Table')
        wbCal = xlrd.open_workbook(os.path.join(filenameCal))
        
        CalibrationTableName=wbCal.sheet_names()
        # Update Database of Calibration Tables
        w1 =xlrd.open_workbook('Calibration Tables/G meter Calibration Table.xls')
        w=copy(w1)
        Allsheet=w1.sheet_names()
        
        Nsheets=np.size(Allsheet)
        
        sw=0
        for k in range(0,Nsheets):
            if str(Allsheet[k])==str(CalibrationTableName[0]):
                sw=ctypes.windll.user32.MessageBoxA(0, 'You already have a calibration table by that name in the data base', 'Warning', 0)
        if sw==0:
            CalTab=wbCal.sheet_by_index(0)
            N=CalTab.nrows
            NLB=self.CalibrationTabSelectview.count()
            self.CalibrationTabSelectview.addItem(str(CalibrationTableName[0]))
            self.CalTabSelect.addItem(str(CalibrationTableName[0]))
            self.CalibrationTabSelectview.setCurrentIndex(int(NLB))
            self.CalibrationTabView.setRowCount(N)
            self.CalibrationTabView.setColumnCount(3)
            self.CalibrationTabView.setHorizontalHeaderItem(0, QtGui.QTableWidgetItem("Lower"))
            self.CalibrationTabView.setHorizontalHeaderItem(1, QtGui.QTableWidgetItem("Upper"))
            self.CalibrationTabView.setHorizontalHeaderItem(2, QtGui.QTableWidgetItem("Factor"))
            for i in range(0,N):
                self.CalibrationTabView.setItem(i, 0, QtGui.QTableWidgetItem(str(CalTab.cell(i,0).value)))
                self.CalibrationTabView.setItem(i, 1, QtGui.QTableWidgetItem(str(CalTab.cell(i,1).value)))
                self.CalibrationTabView.setItem(i, 2, QtGui.QTableWidgetItem(str(CalTab.cell(i,2).value)))
        
            CalTab2=w.add_sheet(str(CalibrationTableName[0]),cell_overwrite_ok=True)
            for i in range(0,N):
                w.get_sheet(Nsheets).write(i,0,CalTab.cell(i,0).value)
                w.get_sheet(Nsheets).write(i,1,CalTab.cell(i,1).value)
                w.get_sheet(Nsheets).write(i,2,CalTab.cell(i,2).value)
        
            w.save('Calibration Tables/G meter Calibration Table.xls')
 
   
    def View_Cal_Tab(self):
        # Update Database of Calibration Tables
        wbCal =xlrd.open_workbook('Calibration Tables/G meter Calibration Table.xls')
        Sname=self.CalibrationTabSelectview.currentText()
        CalTab=wbCal.sheet_by_name(Sname)
        N=CalTab.nrows
        self.CalibrationTabView.setRowCount(N)
        self.CalibrationTabView.setColumnCount(3)
        self.CalibrationTabView.setHorizontalHeaderItem(0, QtGui.QTableWidgetItem("Lower"))
        self.CalibrationTabView.setHorizontalHeaderItem(1, QtGui.QTableWidgetItem("Upper"))
        self.CalibrationTabView.setHorizontalHeaderItem(2, QtGui.QTableWidgetItem("Factor"))
        for i in range(0,N):
            self.CalibrationTabView.setItem(i, 0, QtGui.QTableWidgetItem(str(CalTab.cell(i,0).value)))
            self.CalibrationTabView.setItem(i, 1, QtGui.QTableWidgetItem(str(CalTab.cell(i,1).value)))
            self.CalibrationTabView.setItem(i, 2, QtGui.QTableWidgetItem(str(CalTab.cell(i,2).value)))

        
    def Remove_Cal_Tab(self):        
        sw=ctypes.windll.user32.MessageBoxA(0, 'Are you sure you want to delete this calibration table?', 'Warning', 4)
        if sw==6:
            w1 =xlrd.open_workbook('Calibration Tables/G meter Calibration Table.xls')
            Sname=self.CalibrationTabSelectview.currentText()
            Allsheet=w1.sheet_names()
            N=np.size(Allsheet)
            for i in range(0,N):
                if str(Sname)==str(Allsheet[i]):
                    Ni=i
            w=xlwt.Workbook()
            k=-1
            for i in range(0,N):
                if i!=Ni:
                    k=k+1
                    CalTab2=w.add_sheet(str(Allsheet[i]),cell_overwrite_ok=True)
                    N2=w1.sheet_by_name(str(Allsheet[i])).nrows
                    WorkingSheet=w1.sheet_by_name(str(Allsheet[i]))
                    for j in range(0,N2):
                        w.get_sheet(k).write(j,0,WorkingSheet.cell(j,0).value)
                        w.get_sheet(k).write(j,1,WorkingSheet.cell(j,1).value)
                        w.get_sheet(k).write(j,2,WorkingSheet.cell(j,2).value)

            w.save('Calibration Tables/G meter Calibration Table.xls')
            # Initialise Calibration Table Data with data base
            self.CalibrationTabSelectview.clear()
            self.CalTabSelect.clear()

            AllwbCal = xlrd.open_workbook(os.path.join('Calibration Tables/G meter Calibration Table.xls'))
        
            CalibrationTableName=AllwbCal.sheet_names()
            NCaltabs=np.size(CalibrationTableName)
            for i in range(0,NCaltabs):
                self.CalibrationTabSelectview.addItem(str(CalibrationTableName[i]))
                self.CalTabSelect.addItem(str(CalibrationTableName[i]))
            CalTab=AllwbCal.sheet_by_index(0)
            N=CalTab.nrows
            self.CalibrationTabView.setRowCount(N)
            self.CalibrationTabView.setColumnCount(3)
            for i in range(0,N):
                self.CalibrationTabView.setItem(i, 0, QtGui.QTableWidgetItem(str(CalTab.cell(i,0).value)))
                self.CalibrationTabView.setItem(i, 1, QtGui.QTableWidgetItem(str(CalTab.cell(i,1).value)))
                self.CalibrationTabView.setItem(i, 2, QtGui.QTableWidgetItem(str(CalTab.cell(i,2).value)))

     
    def AddNewAbsTieForSolve(self):
        sw=ctypes.windll.user32.MessageBoxA(0, 'Get ties stations from the absolute gravity data base', 'Import', 4)
        if sw==6:
            Nmeasurements=self.SurveyDataTab.rowCount()              
            SurveyedNamesAll=strs = ["" for x in range(Nmeasurements)]
            
            for i in range(0,Nmeasurements):
                x=self.SurveyDataTab.item(i,0)
                SurveyedNamesAll[i]=str(QtGui.QTableWidgetItem.text(x))
            SurveyedNames=np.unique(SurveyedNamesAll)
            NSurveyedNames=np.size(SurveyedNames)
            N=self.AbsGTabview.rowCount()              
            k=0
            for i in range(0,NSurveyedNames):
                for j in range(0,N):
                    x=self.AbsGTabview.item(j,0)
                    AbsNames=str(QtGui.QTableWidgetItem.text(x))
                    if SurveyedNames[i]==AbsNames:
                        k=k+1
            self.SurveyTies.setRowCount(k)
            if k==0:
                sw=ctypes.windll.user32.MessageBoxA(0,'No surveyed tie stations in the absolute gravity data base.'
                +'\n'+'Please add the tie data manually or add new stations to the absolute gravity data base.', 'Import Error', 0)
                
            nk=-1
            for i in range(0,NSurveyedNames):
                for j in range(0,N):
                    absGname=self.AbsGTabview.item(j,0)
                    absGval=self.AbsGTabview.item(j,1)
                    AbsNames=str(QtGui.QTableWidgetItem.text(absGname))
                    AbsG=float(QtGui.QTableWidgetItem.text(absGval))
                    if SurveyedNames[i]==AbsNames:
                        nk=nk+1
                        self.SurveyTies.setItem(nk,0, QtGui.QTableWidgetItem(AbsNames))
                        self.SurveyTies.setItem(nk,1, QtGui.QTableWidgetItem(str(AbsG)))
        if sw==7:
            sw2=ctypes.windll.user32.MessageBoxA(0, 'Add new row and insert manually?', 'Import', 4)
            if sw2==6:
                Nrows=self.SurveyTies.rowCount()
                Nrowsnew=int(Nrows+1)
                self.SurveyTies.setRowCount(Nrowsnew)


#    def import_coastline(self):
#        """
#        display a coastline sourced from http://www.gadm.org
#        *** Note http://www.gadm.org commercial licence for USA ***
#        """
#        from matplotlib.patches import Polygon
#        from matplotlib.collections import PatchCollection
#
#        myCountry = []
#        m = Basemap()
#        patch = []
#        shapefile = r'country_shapefiles/world/countries.shp'
#        countries = m.readshapefile(shapefile, 'countries', drawbounds=False)
#        for info, shape in zip(m.countries_info, m.countries):
#            if info['NAME'] == myCountry:
#                patch.append(Polygon(np.array(shape)))


    #def draw_coastline(self)
        
        
    def ImportDEMdo(self):
        """
        Use gdal to read grid files \n
        Many available raster formats!!!
        """
        QtGui.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)        
        #QtGui.QApplication.restoreOverrideCursor()
        
        Tk().withdraw()
        DEMfile = askopenfilename(title='Select DEM file')
        #print DEMfile
        if DEMfile.endswith('.xyz'):
            Easting, Northing, Height = np.loadtxt(DEMfile,
                                                   unpack=True)
            N = Easting.shape[0]
            print("N", N)

            self.Easting_vals, self.Easting_idx = np.unique(Easting, 
                                                        return_inverse=True)        
            self.Northing_vals, self.Northing_idx = np.unique(Northing, 
                                                        return_inverse=True)       
            self.Height_vals, self.Height_idx = np.unique(Height, 
                                                        return_inverse=True)
            self.Height_array = np.empty(self.Northing_vals.shape + \
                                         self.Easting_vals.shape)
            self.Height_array.fill(np.nan)
            self.Height_array[self.Northing_idx, self.Easting_idx] = Height
            #self.Easting, self.Northing = np.meshgrid(self.Easting_vals,
            #                                          self.Northing_vals)
            self.Height_array[self.Height_array == 0] = np.nan
            self.Height_array = np.flipud(self.Height_array)
            #Height = self.Height_array
        else:
            DEM = gdal.Open(DEMfile)
            self.raster_srs = DEM.GetProjectionRef()
            trans = DEM.GetGeoTransform()
            extent = (trans[0], trans[0] + DEM.RasterXSize*trans[1],
                      trans[3] + DEM.RasterYSize*trans[5], trans[3])
            Easting = np.arange(extent[0], extent[1], trans[1])
            Northing = np.arange(extent[2], extent[3], trans[1])
            Northing = np.flipud(Northing)
            Height_array = DEM.ReadAsArray()
            ### need to get and set no data values to nan's
            #Height_array[Height_array < -9999.] = np.nan
            nodatavalue = DEM.GetRasterBand(1).GetNoDataValue()
            if nodatavalue is not None:
                self.Height_array = np.ma.masked_equal(Height_array, nodatavalue)
            elif nodatavalue is None:
                self.Height_array = np.ma.masked_equal(Height_array, 0.)
        
            self.Easting_vals, self.Easting_idx = np.unique(Easting,
                                                 return_inverse=True)
            self.Northing_vals, self.Northing_idx = np.unique(Northing,
                                                 return_inverse=True)
            self.Height_vals, self.Height_idx = np.unique(Height_array,
                                                 return_inverse=True)        
            Height = np.ravel(Height_array)
            Easting, Northing = np.meshgrid(Easting, Northing)
            Easting = np.ravel(Easting)
            Northing = np.ravel(Northing)

        self.figDiv = 1000.
        #self.mplwidget.figure.clf()
        self.mplwidget.axes.cla()
        im = self.mplwidget.axes.imshow(self.Height_array,
                                   extent=[min(self.Easting_vals/self.figDiv),
                                   max(self.Easting_vals/self.figDiv),
                                   min(self.Northing_vals/self.figDiv), 
                                   max(self.Northing_vals/self.figDiv)],)
        
        #self.mplwidget.axes.figure.colorbar(im, shrink=0.5).remove()
        #if cb is not None:
        #    cb.remove()
        cb = self.mplwidget.axes.figure.colorbar(im, shrink=0.5)
        
        ### add matplotlib navigation toolbar ###
        self.navi_toolbar = NavigationToolbar(self.mplwidget, self)
        self.addToolBar(self.navi_toolbar)    
        
        #self.mplwidget.axes.set_aspect('equal')        
        self.mplwidget.axes.set_title('DEM height (m)') 
        self.mplwidget.axes.set_xlabel('Easting (km)', fontsize=9)
        self.mplwidget.axes.set_ylabel('Northing (km)', fontsize=9)
        self.mplwidget.draw()
        self.Resolution.setText(str(self.Easting_vals[2]
                                    -self.Easting_vals[1]))
        self.innerradTC.setText(str((self.Easting_vals[2]
                                     -self.Easting_vals[1])/2))
        
        global Easting
        global Northing 
        global Height 
        global Height_array 
        global Easting_vals 
        global Northing_vals
        
        QtGui.QApplication.restoreOverrideCursor()
        

    def transform_coordinates(self, lon, lat):
        """
        terrian corrections only work in meters, \n
        NOT is decimal degrees
        """
        # get SRS from input raster
        targetSR = osr.SpatialReference()
        targetSR.ImportFromWkt(self.raster_srs)
        self.raster_srs = targetSR.ExportToProj4()  
        # wgs84
        p1 = pyproj.Proj("+init=EPSG:4326")
        # whatever raster is in
        p2 = pyproj.Proj(self.raster_srs)
        x, y = pyproj.transform(p1, p2, x, y)
        return x, y 

        
    def ImportdatapointsTCdo(self):  
        global Easting
        global Northing 
        global Height
        global Height_array 
        global Easting_vals 
        global Northing_vals

        Tk().withdraw()
        filename = askopenfilename(title='Select data points')
        wb = xlrd.open_workbook(os.path.join(filename))
        Locations = wb.sheet_by_name('Locations')
        Nlocations = Locations.nrows-1
        Lat = np.zeros(Nlocations)
        Long = np.zeros(Nlocations)
        elev = np.zeros(Nlocations)
        DEMelev = np.zeros(Nlocations)
        east = np.zeros(Nlocations)
        north = np.zeros(Nlocations)
        Names = strs = ["" for x in range(Nlocations)]
               
        self.EastingsNorthingsTC.setRowCount(Nlocations)
        self.EastingsNorthingsTC_results.setRowCount(Nlocations)
        sw = ctypes.windll.user32.MessageBoxA(0, 'Estimate DEM Height?',
                                              'Observation Height', 4)
        if sw == 7: 
            sw2 = ctypes.windll.user32.MessageBoxA(0, 'Get DEM height from file?',
                                                   'Observation Height', 4)
            if sw2 == 7: 
                sw3 = ctypes.windll.user32.MessageBoxA(0, 'Using given height',
                                                       'Observation Height', 0)                        
        QtGui.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
        for i in range (0,Nlocations):
            elev[i] = Locations.cell(i+1,3).value
            Lat[i] = Locations.cell(i+1,1).value
            Long[i] = Locations.cell(i+1,2).value
            Names[i] = Locations.cell(i+1,0).value
            
            east[i] = Locations.cell(i+1,4).value
            north[i] = Locations.cell(i+1,5).value
            
            if sw == 7:
                if sw2 == 6:
                    DEMelev[i] = Locations.cell(i+1,6).value
                if sw2 == 7:
                    DEMelev[i] = Locations.cell(i+1,3).value         
            if sw == 6:
                #print Northing.shape, Easting.shape
                Rad = np.sqrt(np.add(np.square(Easting-east[i]),
                                     np.square(Northing-north[i])))
                                     
                DEMelev[i] = Height[np.where(Rad == Rad.min())]
            self.DEMelev = DEMelev
        QtGui.QApplication.restoreOverrideCursor()    
            
        for i in range(0,Nlocations):
            self.EastingsNorthingsTC.setItem(i, 0, QtGui.QTableWidgetItem(str(Names[i])))
            self.EastingsNorthingsTC.setItem(i, 1, QtGui.QTableWidgetItem(str(Lat[i])))
            self.EastingsNorthingsTC.setItem(i, 2, QtGui.QTableWidgetItem(str(Long[i])))
            self.EastingsNorthingsTC.setItem(i, 3, QtGui.QTableWidgetItem(str(east[i])))
            self.EastingsNorthingsTC.setItem(i, 4, QtGui.QTableWidgetItem(str(north[i])))
            self.EastingsNorthingsTC.setItem(i, 5, QtGui.QTableWidgetItem(str(elev[i])))
            self.EastingsNorthingsTC.setItem(i, 6, QtGui.QTableWidgetItem(str(DEMelev[i])))
            
            self.EastingsNorthingsTC_results.setItem(i, 0, QtGui.QTableWidgetItem(str(Names[i])))
            self.EastingsNorthingsTC_results.setItem(i, 1, QtGui.QTableWidgetItem(str(Lat[i])))
            self.EastingsNorthingsTC_results.setItem(i, 2, QtGui.QTableWidgetItem(str(Long[i])))
            self.EastingsNorthingsTC_results.setItem(i, 3, QtGui.QTableWidgetItem(str(east[i])))
            self.EastingsNorthingsTC_results.setItem(i, 4, QtGui.QTableWidgetItem(str(north[i])))
            self.EastingsNorthingsTC_results.setItem(i, 5, QtGui.QTableWidgetItem(str(elev[i])))
            self.EastingsNorthingsTC_results.setItem(i, 6, QtGui.QTableWidgetItem(str(DEMelev[i])))
        
        global east
        global north
        global DEMelev
        
        self.TCplot2.axes.hold(True)
        im2 = self.TCplot2.axes.imshow(self.Height_array, 
                                extent=[min(self.Easting_vals)/self.figDiv, 
                                        max(self.Easting_vals)/self.figDiv,
                                        min(self.Northing_vals)/self.figDiv,
                                        max(self.Northing_vals)/self.figDiv])       
        cb2 = self.TCplot2.axes.figure.colorbar(im2, shrink=0.5)
        ### add matplotlib navigation toolbar ###
        #self.navi_toolbar = NavigationToolbar(self.mplwidget, self)
        #self.addToolBar(self.navi_toolbar)
        
        self.TCplot2.axes.plot(east/self.figDiv, north/self.figDiv, marker='x',
                               color='w', markersize=10, linestyle='')
        self.TCplot2.axes.plot(east/self.figDiv, north/self.figDiv, marker='x',
                               color='k', markersize=5, linestyle='')
        #self.mplwidget.axes.set_aspect('equal')
        outrad = float(self.outterradTC.text())/self.figDiv
        self.TCplot2.axes.axis([min(east)/self.figDiv-outrad,
                                max(east)/self.figDiv+outrad,
                                min(north)/self.figDiv-outrad,
                                max(north)/self.figDiv+outrad])#"tight")
        self.TCplot2.axes.set_title('DEM height (m)') 
        self.TCplot2.axes.set_xlabel('Easting (km)', fontsize=9)
        self.TCplot2.axes.set_ylabel('Northing (km)', fontsize=9)
        self.TCplot2.draw()
        self.TCplot2.axes.hold(False)
        


    def CalcTCdo(self):
        QtGui.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)        
        
        #import np
        import TerrCorr
        
        global east
        global north
        global DEMelev
        global Easting
        global Northing 
        global Height
        global Height_array 
        global Easting_vals 
        global Northing_vals
        
        outrad = float(self.outterradTC.text())
        innerrad = float(self.innerradTC.text())
        reso = float(self.Resolution.text())
        rho = float(self.density.text())
        self.TCplot2.axes.hold(True)
        self.TCplot2.axes.cla()
        
        im3 = self.TCplot2.axes.imshow(self.Height_array,
                                       extent=[min(self.Easting_vals)/self.figDiv,
                                               max(self.Easting_vals)/self.figDiv,
                                               min(self.Northing_vals)/self.figDiv,
                                               max(self.Northing_vals)/self.figDiv])
        self.TCplot2.axes.plot(east/self.figDiv, north/self.figDiv,'kx')
        TC = np.zeros(np.size(east))
           
        for i in range (0, east.size):
            Radi = np.sqrt(np.add(np.square(Easting-east[i]),
                                  np.square(Northing-north[i])))                
            easti = east[i]
            northi = north[i]
            hi = DEMelev[i]
            Eastingv = Easting[np.where(Radi < outrad)]-easti
            Northingv = Northing[np.where(Radi < outrad)]-northi
            Heightv = Height[np.where(Radi < outrad)]-hi
            Radi2 = np.sqrt(np.add(np.square(Eastingv), np.square(Northingv)))                
            Eastingv2 = Eastingv[np.where(Radi2 > innerrad)]
            Northingv2 = Northingv[np.where(Radi2 > innerrad)]
            Heightv2 = np.abs(Heightv[np.where(Radi2 > innerrad)])
            Heightv2[np.isnan(Heightv2)] = 0
            TC[i] = TerrCorr.TerrCorrg(np.abs(Eastingv2),
                                       np.abs(Northingv2),
                                       np.abs(Heightv2),
                                       rho,reso)
            self.EastingsNorthingsTC_results.setItem(i, 7,
                                        QtGui.QTableWidgetItem(str(TC[i])))
            self.TCplot2.axes.plot(easti/self.figDiv, northi/self.figDiv, 'wx') 
            Vplot = "%.3f" % TC[i]
            self.TCplot2.axes.text(easti/self.figDiv, northi/self.figDiv, Vplot)
            self.TCplot2.axes.axis([min(east)/self.figDiv-outrad/self.figDiv,
                                    max(east)/self.figDiv+outrad/self.figDiv,
                                    min(north)/self.figDiv-outrad/self.figDiv,
                                    max(north)/self.figDiv+outrad/self.figDiv])#"tight")
            self.TCplot2.axes.set_title('DEM height (m)') 
            self.TCplot2.axes.set_xlabel('Easting (km)', fontsize=9)
            self.TCplot2.axes.set_ylabel('Northing (km)', fontsize=9)
            self.TCplot2.draw()
        self.TCplot2.axes.hold(False)
        print(TC)
        
        #self.mplwidget.axes.set_aspect('equal')
        self.TCparameter_results.setItem(0, 0, QtGui.QTableWidgetItem(str(rho)))
        self.TCparameter_results.setItem(0, 1, QtGui.QTableWidgetItem(str(reso)))
        self.TCparameter_results.setItem(0, 2, QtGui.QTableWidgetItem(str(innerrad)))
        self.TCparameter_results.setItem(0, 3, QtGui.QTableWidgetItem(str(outrad)))
        
        QtGui.QApplication.restoreOverrideCursor()
           
    def ExportTCdo(self):        
        import tkFileDialog
        Tk().withdraw()
        filename = tkFileDialog.asksaveasfilename()
        
        w = xlwt.Workbook()
        Nrows = self.EastingsNorthingsTC_results.rowCount()
        Outpage = w.add_sheet('Locations', cell_overwrite_ok=True)   
        Outpage = w.add_sheet('Parameters', cell_overwrite_ok=True)   
        
        w.get_sheet(0).write(0,0,str('Name'))
        w.get_sheet(0).write(0,1,str('Latitude'))      
        w.get_sheet(0).write(0,2,str('Longitude'))       
        w.get_sheet(0).write(0,3,str('Height'))       
        w.get_sheet(0).write(0,4,str('Easting'))       
        w.get_sheet(0).write(0,5,str('Northing'))       
        w.get_sheet(0).write(0,6,str('DEM Height'))       
        w.get_sheet(0).write(0,7,str('Terrain Correction (mGal)'))       
        
        w.get_sheet(1).write(0,0,str('Density (kg/m^3)'))
        w.get_sheet(1).write(0,1,str('Resolution (m)'))      
        w.get_sheet(1).write(0,2,str('Inner Radius (m)'))       
        w.get_sheet(1).write(0,3,str('Outter Radius (m)'))
        
        rho = self.TCparameter_results.item(0,0)
        res = self.TCparameter_results.item(0,1)
        inn = self.TCparameter_results.item(0,2)
        out = self.TCparameter_results.item(0,3)
        
        w.get_sheet(1).write(1,0,float(str(QtGui.QTableWidgetItem.text(rho))))
        w.get_sheet(1).write(1,1,float(str(QtGui.QTableWidgetItem.text(res))))      
        w.get_sheet(1).write(1,2,float(str(QtGui.QTableWidgetItem.text(inn))))       
        w.get_sheet(1).write(1,3,float(str(QtGui.QTableWidgetItem.text(out))))
        
        for i in range(0,Nrows):
            Name = self.EastingsNorthingsTC_results.item(i,0)
            Lat = self.EastingsNorthingsTC_results.item(i,1)
            Long = self.EastingsNorthingsTC_results.item(i,2)
            H1 = self.EastingsNorthingsTC_results.item(i,3)
            x = self.EastingsNorthingsTC_results.item(i,4)
            y = self.EastingsNorthingsTC_results.item(i,5)
            H2 = self.EastingsNorthingsTC_results.item(i,6)
            TC1 = self.EastingsNorthingsTC_results.item(i,7)
            w.get_sheet(0).write(int(i+1),0,str(QtGui.QTableWidgetItem.text(Name)))       
            w.get_sheet(0).write(int(i+1),1,float(str(QtGui.QTableWidgetItem.text(Lat))))        
            w.get_sheet(0).write(int(i+1),2,float(str(QtGui.QTableWidgetItem.text(Long))))                           
            w.get_sheet(0).write(int(i+1),3,float(str(QtGui.QTableWidgetItem.text(H1))))                           
            w.get_sheet(0).write(int(i+1),4,float(str(QtGui.QTableWidgetItem.text(x))))                           
            w.get_sheet(0).write(int(i+1),5,float(str(QtGui.QTableWidgetItem.text(y))))                           
            w.get_sheet(0).write(int(i+1),6,float(str(QtGui.QTableWidgetItem.text(H2))))                           
            w.get_sheet(0).write(int(i+1),7,float(str(QtGui.QTableWidgetItem.text(TC1))))                           
        w.save(filename+'.xls')


if __name__=='__main__':
    app = QtGui.QApplication(sys.argv)
    
    # start program #    
    dmw = DesignerMainWindow()
    dmw.show()
    sys.exit(app.exec_())
    
    