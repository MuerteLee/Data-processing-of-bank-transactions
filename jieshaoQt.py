import sys,os,platform
import math
from PyQt5.QtCore import QDir, Qt
from PyQt5.QtGui import QFont, QPalette
from PyQt5.QtWidgets import (QApplication, QCheckBox, QDialog,
         QFileDialog, QFrame, QGridLayout,QDialogButtonBox,
        QInputDialog, QLabel,  QPushButton, QMessageBox)
import os, sys
import xlrd
import xlwt
from xlutils.copy import copy
from xlrd import *
from fractions import Fraction

# -*- coding: utf-8 -*-
class createExcelModule():
    def __init__(self,excelFile):
        if os.path.isfile(excelFile):
            #print("Error: Please specical the other file, because this file had existed!\n")
            try:
                data = xlrd.open_workbook(excelFile);
            except Exception:
                print ('Error: Cannot open workbook \n');
            table = data.sheets()[0];
            #colnames =  table.row_values(0);
            #print(table.cell(0,1).value)
            colnames = ['商户编号', '消费总笔数', '有效笔数', '交易总金额','结算金额']
            for i in range(0,len(colnames)):
                #print(str(table.cell(0,i).value), colnames[i])
                if str(table.cell(0,i).value) != colnames[i]:
                    w = copy(open_workbook(excelFile))
                    w.get_sheet(0).write(0,i,colnames[i])
                    w.save(excelFile)
                    print("Error: The excel file was destoried! I had format!!!\n")
                    print(colnames)
        else:
            style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',num_format_str='#,##0.00');
            file = xlwt.Workbook()
            sheet = file.add_sheet('交易记录')
            sheet.write(0,0,'商户编号',style0)
            sheet.write(0,1,'消费总笔数')
            sheet.write(0,2,'有效笔数')
            sheet.write(0,3,'交易总金额')
            sheet.write(0,4,'结算金额')
            file.save(excelFile)

class initExcel(createExcelModule):
    def __init__(self, execelFile):
        createExcelModule.__init__(self, execelFile);

class initFileData():
    def __init__(self,fileName):
        parse = [ ]
        self.guestId = ''
        self.bargainRecode = {} 
        bargainNum = []
        barTmp1 = []
        bargain = 0
        try:
            with open(fileName, 'r') as file:
                for line in file:
                    line = line.strip('\n').strip();
                    if line != '':
                        parse.append(line)
                self.guestId = (parse[1].split('         ')[0].split(':'))[0].strip().split('      ')[1];
        finally:
            file.close();
        for i in range(3, len(parse)):
            if int(parse[i].find('_____________________')) == 0:
                if bargain < 2:
                    bargainNum.append(i);
                    bargain = bargain + 1;
                else:
                    break;
        #print(parse)
        for i in range(bargainNum[0]+1, bargainNum[1]-1):
            location = parse[i].find("*")
            num = parse[i].count("*")
            #print(location,num)
            splitValue = '';
            #print(parse[i])
            for j in range(location-6,location+num+4):
                #print(parse)
                splitValue = splitValue + parse[i][j]
            #print(splitValue)
            #barTmp = parse[i].split(splitValue)
            #print(barTmp)
            #barTmp = parse[i].split('        ')
            barTmp = parse[i].split(' ')
            #print(barTmp)
            for ii in range(0, len(barTmp)):
                if str(barTmp[ii]) != '':
                    barTmp1.append(barTmp[ii].strip())
            #print(barTmp1)
            #sprint(i-bargainNum[0] -1)
            self.bargainRecode[i-bargainNum[0]-1] = barTmp1;
            barTmp1 = []
        #print(self.bargainRecode)

class parseFileData(initFileData):
    def __init__(self, fileName, validValue):
        initFileData.__init__(self, fileName);
        #print(self.guestId)
        self.effectTraceNum = 0
        traceSum = 0.00
        cardTraceM = 0.00
        accountMTmp = 0.00
        self.accountM = 0.00
        self.TotalTraceM = 0.00
        self.TotalTraceNum = 0.0
        for index in range(0, len(self.bargainRecode)):
            value = float(self.bargainRecode[index][4]);
            #bankValue = self.bargainRecode[index][3]
            #accountMTMP = float(self.bargainRecode[index][4].split('  ')[0]);
            
            if value >= validValue:
                self.effectTraceNum = self.effectTraceNum + 1

            if value > 0 and value < 3334:
                self.accountM = self.accountM + value*0.9922
                self.TotalTraceNum = self.TotalTraceNum + 1
            elif value >= 3334:
                self.TotalTraceNum = self.TotalTraceNum + 1
                self.accountM = self.accountM + value - 26
            elif value < 0:
                if math.fabs(value) >= validValue and self.TotalTraceNum > 0:
                    self.effectTraceNum = self.effectTraceNum - 1
                self.TotalTraceNum = self.TotalTraceNum - 1
                if  math.fabs(value) < 3334:
                    self.accountM = self.accountM + math.fabs(value)*0.0078 + value 
                elif math.fabs(value) >= 3334:
                    self.accountM = self.accountM + value + 26
                
            #print(self.accountM)
            #cardTraceM = cardTraceM + float(bankValue)
            traceSum = traceSum + value
                #print(self.bargainRecode[index][2].strip())
        self.TotalTraceM = traceSum + cardTraceM
#        print(traceSum,cardTraceM)
#        print(self.TotalTraceM, self.effectTraceNum)
        
class writeExcelData(initExcel, parseFileData):
    def __init__(self, excelFile, fileName, validValue):
        initExcel.__init__(self, excelFile);
        parseFileData.__init__(self, fileName, validValue)
        self.excelFile = excelFile
    def wirteExcelData(self):
        try:
            data = xlrd.open_workbook(self.excelFile);
            table = data.sheets()[0];
            i = table.nrows;
            w = copy(data)
            w.get_sheet(0).write(i,0,self.guestId);
            w.get_sheet(0).write(i,1,self.TotalTraceNum);
            #w.get_sheet(0).write(i,1,len(self.bargainRecode));
            w.get_sheet(0).write(i,2,self.effectTraceNum);
            w.get_sheet(0).write(i,3,float(self.TotalTraceM));
            w.get_sheet(0).write(i,4,float(self.accountM));
        finally:
            w.save(self.excelFile)
        
class Dialog(QDialog):
    def __init__(self, parent=None):
        super(Dialog, self).__init__(parent)
        self.directoryValue =''
        self.effectValue = 0.00
        self.fileValue = ''
        self.openFilesPath = ''
        self.saveFileName = ''

        self.setFixedSize(800,600)
        frameStyle = QFrame.Sunken | QFrame.Panel

        self.directoryLabel = QLabel()
        self.directoryLabel.setFrameStyle(frameStyle)
        self.directoryButton = QPushButton("指定文件夹路径：")

        #self.directoryLabelReset = QLabel()
        #self.directoryLabelReset.setFrameStyle(frameStyle)
        self.directoryButtonReset = QPushButton("Reset")

        self.openFileNamesLabel = QLabel()
        self.openFileNamesLabel.setFrameStyle(frameStyle)
        self.openFileNamesButton = QPushButton("指定文件路径：")
        self.openFileNamesButtonReset = QPushButton("Reset")

        self.doubleLabel = QLabel()
        self.doubleLabel.setFrameStyle(frameStyle)
        self.doubleButton = QPushButton("有效交易值：")
        self.doubleLabel.setText("$3333")

        self.saveFileNameLabel = QLabel()
        self.saveFileNameLabel.setFrameStyle(frameStyle)
        self.saveFileNameButton = QPushButton("生成excel路径")
        self.saveFileNameButtonReset = QPushButton("Reset")

        buttonBox = QDialogButtonBox(QDialogButtonBox.Cancel)

        #self.okdirectoryLabel = QLabel()
        #self.okdirectoryLabel.setFrameStyle(frameStyle)
        self.okdirectoryButton = QPushButton("OK：")

        #buttonBox.accepted.connect(self.accept)
        buttonBox.rejected.connect(self.reject)

        self.doubleButton.clicked.connect(self.setDouble)
        self.openFileNamesButton.clicked.connect(self.setOpenFileNames)
        self.openFileNamesButtonReset.clicked.connect(self.setOpenFileNamesReset)
        self.saveFileNameButton.clicked.connect(self.setSaveFileName)
        self.saveFileNameButtonReset.clicked.connect(self.setSaveFileNameReset)
        self.directoryButton.clicked.connect(self.setExistingDirectory)
        self.directoryButtonReset.clicked.connect(self.setExistingDirectoryReset)
        self.okdirectoryButton.clicked.connect(self.okButton)

        self.native = QCheckBox()
        self.native.setText("Use native file dialog.")
        self.native.setChecked(True)
        if sys.platform not in ("win32", "darwin"):
            self.native.hide()

        layout = QGridLayout()
        layout.setColumnStretch(1, 1)
        layout.setColumnMinimumWidth(1, 250)

        layout.addWidget(self.directoryButton, 0, 0)
        layout.addWidget(self.directoryButtonReset, 0, 2)
        layout.addWidget(self.directoryLabel, 0, 1)
        #layout.addWidget(self.directoryLabelReset, 0, 2)

        layout.addWidget(self.openFileNamesButton, 1, 0)
        layout.addWidget(self.openFileNamesButtonReset, 1, 2)
        layout.addWidget(self.openFileNamesLabel, 1, 1)

        layout.addWidget(self.doubleButton, 2, 0)
        layout.addWidget(self.doubleLabel, 2, 1)

        layout.addWidget(self.saveFileNameButton, 3, 0)
        layout.addWidget(self.saveFileNameButtonReset, 3, 2)
        layout.addWidget(self.saveFileNameLabel, 3, 1)

        layout.addWidget(self.okdirectoryButton, 4, 0)
        #layout.addWidget(self.okdirectoryLabel, 4, 1)

        layout.addWidget(self.native, 15, 0)
        layout.addWidget(buttonBox)
        self.setLayout(layout)

    def checkResourceFile(self, fileName):
            parse = [ ]
            self.checkFileType = 0
            try:
                with open(fileName, 'r') as file:
                    for line in file:
                        line = line.strip('\n').strip();
                        if line != '':
                            parse.append(line)
                    #print((parse[1].split('         ')[0].split(':'))[0].strip().split('      ')[1]);
                    self.checkFileType = ((parse[1].split('         ')[0].split(':'))[0].strip().split('      ')[1]);
                    #print(self.checkFileType)
            finally:
                file.close();
            return self.checkFileType
        
    def okButton(self):      
        fileList=[]
        
        if not self.effectValue:
            self.effectValue = 3333;
            
        if len(self.fileValue):
            for i in self.fileValue:
                fileList.append(i)

        if len(self.directoryValue):
            for file in os.listdir(self.directoryValue):                
                file = self.directoryValue+"\\"+file;
                if file.find('.') == -1:
                    #print(self.checkResourceFile(file))
                    if  os.path.isfile(file) and int(self.checkResourceFile(file)) > 0:
                        print(file)
                        fileList.append(file)
        if not self.saveFileName:
                self.saveFileName = 'C:\\自动生成对账单.xls'

        if len(fileList):
            for file in fileList:
                parsefile = writeExcelData(self.saveFileName, file, self.effectValue)
                parsefile.wirteExcelData()

        
        #print('self.fileValue =%s\n' %self.fileValue)
        #print('self.effectValue=%s\n' %self.effectValue)
        MESSAGE = "<p>Excel 已经生成！</p>"
        if os.path.isfile(self.saveFileName):
            reply = QMessageBox.information(self,
                "生成Excel状态提示", MESSAGE)

    def setSaveFileNameReset(self):    
        self.saveFileNameLabel.setText('')

    def setSaveFileName(self):    
        options = QFileDialog.Options()
        if not self.native.isChecked():
            options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getSaveFileName(self,
                "QFileDialog.getSaveFileName()",
                self.saveFileNameLabel.text(),
                "All Files (*);;Text Files (*.txt)", options=options)
        if fileName:
            if platform.system() == 'Windows':
                fileName = fileName.replace('/','\\');
            if fileName.find('.xls') <= 0:
               fileName = fileName.split('.')[0]+'.xls'
            self.saveFileNameLabel.setText(fileName)
            self.saveFileName = fileName
            
    def setDouble(self):    
        d, ok = QInputDialog.getDouble(self, "QInputDialog.getDouble()",
                "Amount:", 3333.00, -10000, 1000000, 1)
        if ok:
            self.doubleLabel.setText("$%g" % d)
            self.effectValue = d

    def setOpenFileNamesReset(self):    
        self.openFileNamesLabel.setText('');

    def setOpenFileNames(self):    
        options = QFileDialog.Options()
        if not self.native.isChecked():
            options |= QFileDialog.DontUseNativeDialog
        files, _ = QFileDialog.getOpenFileNames(self,
                "QFileDialog.getOpenFileNames()", self.openFilesPath,
                "All Files (*);;Text Files (*.txt)", options=options)
        if files:             
            self.openFilesPath = files[0]
            if platform.system() == 'Windows':
                self.openFileNamesLabel.setText(str(files).replace('/','\\'))
            else:
                self.openFileNamesLabel.setText("[%s]" % ', '.join(files))
            self.fileValue = files 

    def setExistingDirectoryReset(self):    
        self.directoryLabel.setText('')

    def setExistingDirectory(self):    
        options = QFileDialog.DontResolveSymlinks | QFileDialog.ShowDirsOnly
        directory = QFileDialog.getExistingDirectory(self,
                "QFileDialog.getExistingDirectory()",
                self.directoryLabel.text(), options=options)
        if directory:
            if platform.system() == 'Windows':
                directory = directory.replace('/','\\')
            self.directoryValue = directory
            self.directoryLabel.setText(directory)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    dialog = Dialog()
    dialog.show()
    sys.exit(app.exec_())
