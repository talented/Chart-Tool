# -*- coding: utf-8 -*-
"""
Created on Thu May 03 16:47:32 2018

@author: Özgür Yarikkas
"""

import sys, os
import sip
sip.setapi('QString', 2)
from PyQt4 import QtGui
from PyQt4.QtCore import QTimer
import updstart, about
from v3_ui import draw_pie, draw_reason, draw_xy, get_column_dict, get_plot_style

#-------------------------------------------------------------------------------------#

class About(QtGui.QMainWindow, about.Ui_MainWindow):
    def __init__(self, parent=None):
        super(About, self).__init__(parent)
        self.setupUi(self)

class ExampleApp(QtGui.QMainWindow, updstart.Ui_MainWindow):
    def __init__(self, parent=None):
        super(ExampleApp, self).__init__(parent)
        self.setupUi(self)
        self.setWindowTitle("Chart Tool")
        self.pushButton_browse.clicked.connect(self.browse_folder)
        self.pushButton_exit.clicked.connect(self.close)
        self.pushButton_Bprocess.clicked.connect(self.Bprocessed)
        self.pushButton_Pprocess.clicked.connect(self.Pprocessed)
        self.statusBar().showMessage('Ready')

        self.lineEdit_Btitle.setPlaceholderText("Default: Column 1 text")
        self.lineEdit_Bcolumn1.setPlaceholderText("Enter Column Name")
        self.lineEdit_Bcolumn2.setPlaceholderText("Enter Column Name")
        self.lineEdit_Bfield.setPlaceholderText("Enter Field Description")
        text1 = 'Mandatory Field'
        text2 = 'Mandatory Field if\n2 Column or Batch Process\noption selected'
        text3 = 'Mandatory field if 2 Column option selected.\nWrite a field name from given column 1.\nLeave blank for Batch Process'
        text4 = 'If blank, Column 1 text will be used as Title'
        self.label_2.setToolTip(text1)
        self.lineEdit_Bcolumn1.setToolTip(text1)
        self.label_3.setToolTip(text2)
        self.lineEdit_Bcolumn2.setToolTip(text2)
        self.label_6.setToolTip(text3)
        self.lineEdit_Bfield.setToolTip(text3)
        self.lineEdit_Btitle.setToolTip(text4)
        self.lineEdit_Ptitle.setToolTip(text4)
        self.checkBox_Psorted.setToolTip('If not checked, Bars will be sorted\naccording to the order in column 1')
        self.checkBox_Bpercentage.setToolTip('If not checked, Bars which is representing\nthe Percentages, will not be shown')

        self.lineEdit_Ptitle.setPlaceholderText("Default: Column 1 text")
        self.lineEdit_Pcolumn1.setPlaceholderText("Enter Column Name")
        self.lineEdit_Pcolumn2.setPlaceholderText("Enter Column Name")
        self.lineEdit_Pfield.setPlaceholderText("Enter Field Description")
        self.label_5.setToolTip(text1)
        self.lineEdit_Pcolumn1.setToolTip(text1)
        self.label_4.setToolTip(text2)
        self.lineEdit_Pcolumn2.setToolTip(text2)
        self.lineEdit_Pfield.setToolTip(text3)
        self.lineEdit_Pfield.setToolTip(text3)
        self.checkBox_Pexploded.setToolTip('Highlights the part with highest percentage')
        self.label_11.setToolTip('Items which their Percentages\nare less than given value, will be\ncollapsed and shown as "Others"')

        get_plot_style()
        self.file_path = ''
        self.data = ''
        self.selected = ''
        self.path = ''
        self.base = ''
        self.batchCompleted = 0
        self.batchQuantity = 0
        self.batchCount = 0

#        Timer
        self.timer = QTimer()

        #Menubar
        self.actionExit.setShortcut('Alt+Q')
        self.actionExit.setStatusTip('Exit application')
        self.actionExit.triggered.connect(self.close)

        self.actionOpen.setShortcut('Alt+B')
        self.actionOpen.setStatusTip('Load Excel File')
        self.actionOpen.triggered.connect(self.browse_folder)

        self.actionHelp.setShortcut('F1')
        self.actionHelp.setStatusTip('How to use the Chart Tool')
        self.actionHelp.triggered.connect(self.helpfile)

#        self.actionAbout.setShortcut('F1')
#        self.actionHelp.setStatusTip('How to use the Chart Tool')
        self.actionAbout.triggered.connect(self.Aabout)
        self.dialog = About(self)

#-------------------------------------------------------------------------------------#

    def Aabout(self):
        self.dialog.show()

#-------------------------------------------------------------------------------------#

    def helpfile(self):
        self.statusBar().showMessage('Still in preparation..')
        pass
#-------------------------------------------------------------------------------------#

    def count(self):
        self.statusBar().showMessage('Ready')
        self.progressBar.setValue(0)
        print 'reset'

#-------------------------------------------------------------------------------------#

    def message_box(self):
        self.setWindowTitle('Exit')

#-------------------------------------------------------------------------------------#

    def closeEvent(self, event):
        reply = QtGui.QMessageBox.question(self, 'Message',
            "Are you sure to quit?", QtGui.QMessageBox.Yes |
            QtGui.QMessageBox.No, QtGui.QMessageBox.No)

        if reply == QtGui.QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

#-------------------------------------------------------------------------------------#

    def browse_folder(self):
#        self.listWidget.clear() # In case there are any existing elements in the list

        filename = QtGui.QFileDialog.getOpenFileName(self,'Open File', None, 'Excel File (*.xlsx)')

        if filename:
            try:
                self.lineEdit_selected.setText(filename)
                self.path = os.path.dirname(self.lineEdit_selected.text())
                self.base = os.path.basename(self.lineEdit_selected.text())

            except:
                self.statusBar().showMessage('Failed to read file!')
#                pass
#                self.QmessageBox("Open Source File", "Failed to read file \n'%s'"%filename, self.QMessageBox.Warning, self.QMessageBox.Ok, 0, 0)
#                return
#        self.lineEdit
        # execute getExistingDirectory dialog and set the directory variable to be equal
        # to the user selected directory

#        if directory: # if user didn't pick a directory don't continue
#            for file_name in os.listdir(directory): # for all files, if any, in the directory
##                self.listWidget.addItem(file_name)  # add file to the listWidget
#                print file_name

#-------------------------------------------------------------------------------------#

    def success(self, tip):
        self.statusBar().showMessage('Success! ' + tip + ' chart created under ' + self.path +'/charts')

#-------------------------------------------------------------------------------------#

    def Bprocessed(self):
        self.progressBar.setValue(0)
        if self.path != '':
            if self.comboBox_bar.currentText() == 'Simple Column':
                self.Bar_simple_column()
            elif self.comboBox_bar.currentText() == '2 Column':
                self.Bar_2_column()
            elif self.comboBox_bar.currentText() == 'Batch Process':
                self.Bar_batch_process()
        else:
            self.statusBar().showMessage('No excel file selected!')

#-------------------------------------------------------------------------------------#

    def Bar_settings(self):
#       Bar settings
        self.hidden = self.lineEdit_Bxlabel.text()
        self.barTitle = True
        self.BgivenTitle = None
        self.Bangle = self.spinBox_Bangle.value()
        self.Bpercentage = True
        self.Blegend = True
        self.Blegendloc = self.comboBox_Blegend.currentText()
        self.xlegend = None
        self.Bsorted = True
        self.Bbcolor = self.comboBox_Bbcolor.currentText()
        self.Bpcolor = self.comboBox_Bpcolor.currentText()

        #self.Bcolor = self.comboBox_Pcolor.currentText()
        self.Bcolumn1 = self.lineEdit_Bcolumn1.text()
        self.Bcolumn2 = self.lineEdit_Bcolumn2.text()
        self.Bfield = self.lineEdit_Bfield.text()
        self.Bxlabel = False
        self.BgivenXlabel = None
        self.Bylabel = False
        self.BgivenYlabel = None

        if self.Bcolumn1 != '':
            if not self.checkBox_Btitle.isChecked():
                self.barTitle = False
            if self.lineEdit_Btitle.text():
                self.BgivenTitle = self.lineEdit_Btitle.text()

            if self.checkBox_Bxlabel.isChecked():
                self.Bxlabel = True
            if self.lineEdit_Bxlabel.text():
                self.BgivenXlabel = self.lineEdit_Bxlabel.text()

            if self.checkBox_Bylabel.isChecked():
                self.Bylabel = True
            if self.lineEdit_Bylabel.text():
                self.BgivenYlabel = self.lineEdit_Bylabel.text()

            if self.Blegend and self.lineEdit_Blegend.text():
                self.xlegend = self.lineEdit_Blegend.text()
            elif self.Blegend and not self.lineEdit_Blegend.text():
                self.xlegend = self.Bcolumn1

            if not self.checkBox_Blegend.isChecked():
                self.Blegend = False

            if not self.checkBox_Psorted.isChecked():
                self.Bsorted = False

            if not self.checkBox_Bpercentage.isChecked():
                self.Bpercentage = False

#-------------------------------------------------------------------------------------#

    def Bar_simple_column(self):
        self.Bar_settings()
 
#        1
        if self.Bcolumn1 != '':
            try:
                draw_xy(self.base, self.Bcolumn1, titleS=self.barTitle, titleT=self.BgivenTitle, currdir=self.path, angle=self.Bangle,
                        percentage=self.Bpercentage, legendloc=self.Blegendloc, Bsorted=self.Bsorted, xlabel=self.Bxlabel,
                        xlabelT=self.BgivenXlabel, ylabel=self.Bylabel, ylabelT=self.BgivenYlabel, xlegend=self.xlegend,
                        Bbcolor=self.Bbcolor, Bpcolor=self.Bpcolor, hidden=self.hidden)

                #self.increaseValue('BAR')
                #self.batchCount += 1
                # check every second
#                self.timer.start(1000*1)
                self.success('BAR')
                self.timer.singleShot(6000, self.count)

            except:
                self.statusBar().showMessage('Please double check the entered column name!')
                pass
        else:
            self.statusBar().showMessage('Column1 can not be blank!')

#-------------------------------------------------------------------------------------#

    def Bar_2_column(self):
        self.Bar_settings()

#         2
        if self.Bcolumn1 != '' and self.Bcolumn2 !='' and self.Bfield != '':
            try:
                draw_reason(self.base,self.Bcolumn1, self.Bcolumn2, self.Bfield, titleS=self.barTitle, titleT=self.BgivenTitle,
                            currdir=self.path, percentage=self.Bpercentage, angle=self.Bangle, Bsorted=self.Bsorted, xlabel=self.Bxlabel,
                            xlabelT=self.BgivenXlabel, ylabel=self.Bylabel, ylabelT=self.BgivenYlabel, xlegend=self.xlegend,
                            legendloc=self.Blegendloc, Bbcolor=self.Bbcolor, Bpcolor=self.Bpcolor, hidden=self.hidden)
                self.success('BAR')
                self.timer.singleShot(6000, self.count)
            except:
                self.statusBar().showMessage('Please double check the entered column name!')
                pass
        else:
            self.statusBar().showMessage('Column 1, Column 2 and Field sections can not be blank!')

#-------------------------------------------------------------------------------------#

    def Bar_batch_process(self):
        self.Bar_settings()
        os.chdir(self.path)
        call = get_column_dict(self.base, self.Bcolumn1).keys()
        self.batchQuantity = len(call)

        if self.Bcolumn1 != '' and self.Bcolumn2 !='':
    #        3. Batch Process
            try:
                for param in call:
                    draw_reason(self.base,self.Bcolumn1, self.Bcolumn2, param, titleS=self.barTitle, titleT=self.BgivenTitle,
                            currdir=self.path, percentage=self.Bpercentage, angle=self.Bangle, Bsorted=self.Bsorted, xlabel=self.Bxlabel,
                            xlabelT=self.BgivenXlabel, ylabel=self.Bylabel, ylabelT=self.BgivenYlabel, xlegend=self.xlegend,
                            legendloc=self.Blegendloc, Bbcolor=self.Bbcolor, Bpcolor=self.Bpcolor, hidden=self.hidden)
                    self.increaseValueBatch()
                    self.batchCount += 1
                self.success('BAR')
                self.timer.singleShot(6000, self.count)

            except:
                self.statusBar().showMessage('Please double check the entered column names and field name')

                self.progressBar.setValue(0)
                pass
        else:
            self.statusBar().showMessage('Column1, Column2 can not be blank!')

#-------------------------------------------------------------------------------------#

    def Pprocessed(self):
        self.progressBar.setValue(0)
        if self.path != '':
            if self.comboBox_pie.currentText() == 'Simple Column':
                self.Pie_simple_column()
            elif self.comboBox_pie.currentText() == '2 Column':
                self.Pie_2_column()
            elif self.comboBox_pie.currentText() == 'Batch Process':
                self.Pie_batch_process()
        else:
            self.statusBar().showMessage('No excel file selected!')

#-------------------------------------------------------------------------------------#

    def Pie_settings(self):
#       Pie settings
        self.pieTitle = True
        self.PgivenTitle = None
        self.Plegendloc = self.comboBox_Plegend.currentText()
        self.Pexploded = False
        self.Pprocent = self.spinBox_Pcollapsed.value()
        self.Pcolor = self.comboBox_Pcolor.currentText()

        self.Pcolumn1 = self.lineEdit_Pcolumn1.text()
        self.Pcolumn2 = self.lineEdit_Pcolumn2.text()
        self.Pfield = self.lineEdit_Pfield.text()

        if self.Pcolumn1 != '':
            if not self.checkBox_Ptitle.isChecked():
                self.pieTitle = False

            if self.lineEdit_Ptitle.text():
                self.PgivenTitle = self.lineEdit_Ptitle.text()

            if self.checkBox_Pexploded.isChecked():
                self.Pexploded = True

#-------------------------------------------------------------------------------------#

    def Pie_simple_column(self):
        self.Pie_settings()
        if self.Pcolumn1 != '':
            try:
                draw_pie(self.base, self.Pcolumn1, currdir=self.path, pieTitle=self.pieTitle, givenTitle=self.PgivenTitle,
                         randomColors=False, collapsed=True, exploded=self.Pexploded, legendloc=self.Plegendloc,
                         procent=self.Pprocent, pctdistance=0.6, Pcolor=self.Pcolor)
                self.success('Pie')
                self.timer.singleShot(6000, self.count)
            except:
                self.statusBar().showMessage('Please double check the entered column name!')
                pass
        else:
            self.statusBar().showMessage('Column 1 section can not be blank!')

#-------------------------------------------------------------------------------------#

    def Pie_2_column(self):
        self.Pie_settings()
        self.PgivenTitle = self.Pcolumn2
        if self.Pcolumn1 != '' and self.Pcolumn2 !='' and self.Pfield != '':
            try:

        #        2
                draw_pie(self.base, self.Pcolumn1, divergences=self.Pcolumn2, field=self.Pfield, currdir=self.path,
                         pieTitle=self.pieTitle, givenTitle=self.PgivenTitle, randomColors=False,
                         collapsed=True, exploded=self.Pexploded, procent=self.Pprocent, legendloc=self.Plegendloc,
                         Pcolor=self.Pcolor)
                self.success('PIE')
                self.timer.singleShot(6000, self.count)
            except:
                self.statusBar().showMessage('Please double check the entered column names and field name')
                pass
        else:
            self.statusBar().showMessage('Column 1, Column 2 and Field sections can not be blank!')

#-------------------------------------------------------------------------------------#

    def Pie_batch_process(self):
        self.Pie_settings()
        self.PgivenTitle = self.Pcolumn2
        os.chdir(self.path)
        call = get_column_dict(self.base, self.Pcolumn1).keys()
        self.batchQuantity = len(call)

        if self.Pcolumn1 != '' and self.Pcolumn2 !='':
    #        3. Batch Process
            try:
                for param in call:
        #            
                    draw_pie(self.base, self.Pcolumn1, divergences=self.Pcolumn2, field=param, currdir=self.path,
                             pieTitle=self.pieTitle, givenTitle=self.PgivenTitle, randomColors=False,
                             collapsed=True, exploded=self.Pexploded, procent=self.Pprocent, legendloc=self.Plegendloc,
                             Pcolor=self.Pcolor)
                    self.increaseValueBatch()
                    self.batchCount += 1
                self.success('PIE')
                self.timer.singleShot(6000, self.count)

            except:
                self.statusBar().showMessage('Please double check the entered column names and field name')

                self.progressBar.setValue(0)
                pass
        else:
            self.statusBar().showMessage('Column1, Column2 can not be blank!')

#-------------------------------------------------------------------------------------#

#    PROGRESS BAR
    def increaseValueBatch(self):
        self.batchCompleted += int(100 / self.batchQuantity)

        self.progressBar.setValue(self.batchCompleted)
        self.statusBar().showMessage('Creating Charts...')
        if self.batchCount == self.batchQuantity - 1:
            self.progressBar.setValue(100)

 #-------------------------------------------------------------------------------------#

def main():

    app = QtGui.QApplication.instance()
    if app is None:
        app = QtGui.QApplication(sys.argv)

    form = ExampleApp()
    form.show()
    app.exec_()

if __name__ == '__main__':
    main()

#    # Create and display the splash screen
#    splash_pix = QtGui.QPixmap('preloader.gif')
#
#    splash = QtGui.QSplashScreen(splash_pix, Qt.WindowStaysOnTopHint)
##    splash = QtGui.QSplashScreen(splash_pix, Qt.WindowStaysOnTopHint)
#    splash.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)
#    splash.setEnabled(False)
#    # splash = QSplashScreen(splash_pix)
#    # adding progress bar
#    pgBar = QtGui.QProgressBar(splash)
#    pgBar.setMaximum(10)
#    pgBar.setGeometry(0, splash_pix.height() - 50, splash_pix.width(), 20)
#
##     splash.setMask(splash_pix.mask())
#
#    splash.show()
##    splash.showMessage("<h1><font color='green'>Please Wait...</font></h1>", Qt.AlignTop | Qt.AlignCenter, Qt.black)
#
#    for i in range(1, 11):
#        pgBar.setValue(i)
#        t = time.time()
#        while time.time() < t + 0.7:
#           app.processEvents()
#
#    # Simulate something that takes time
#    time.sleep(1)
#
#    form = ExampleApp()
#    form.show()
#    splash.finish(form)
#    sys.exit(app.exec_())
