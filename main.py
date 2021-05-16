import os
import re
from time import strftime
import datetime
import xlsxwriter
import sys
from tkinter import filedialog
from tkinter import *
from megastudy import GoMegastudy
from megastudy import SetMegastudy
from megastudy import CalcMegastudy
from daesung import GoDaesung
from daesung import SetDaesung
from daesung import CalcDaesung
from ebs import Ebs
from skyedu import GoSkyedu
from skyedu import SetSkyedu
from skyedu import CalcSkyedu
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5 import QtCore
#from win32process import CREATE_NO_WINDOW

from selenium import webdriver

form_class = uic.loadUiType("GUIMegastudy.ui")[0]

site = {'메가스터디':'MEGA', 'EBSi':'EBS', '대성 마이맥':'DS', '스카이 에듀':'SKY'}
department = {
				# 'MEGA':
				# 	{'국어' : 'divMenuTecList1', '수학' : 'divMenuTecList2', '영어' : 'divMenuTecList3', '한국사' : 'divMenuTecList4', '사회' : 'divMenuTecList5', '과학' : 'divMenuTecList6', '대학별고사' : 'divMenuTecList7', '제2외국어한문' : 'divMenuTecList8'}
				# ,
				'MEGA':
					{'국어' : '0', '수학' : '1', '영어' : '2', '한국사' : '3', '사회' : '4', '과학' : '5', '대학별고사' : '6', '제2외국어한문' : '7'}
				,'DS':
					{'국어' : 'li01', '수학' : 'li02', '영어' : 'li03', '한국사' : 'li10', '사회' : 'li04', '과학' : 'li05', '대학별고사' : 'li07', '제2외국어한문' : 'li06'}
				,'SKY':
					{'국어' : '국어', '수학' : '수학', '영어' : '영어', '한국사' : '한국사', '사회' : '사회', '과학' : '과학', '대학별고사' : '대학별고사', '제2외국어' : '제2외국어', '월간대치동':'월간대치동', '내신전문':'내신전문'}
			 }

SelectedSite = []
MegasubjectObject = []  # 사용자가 선택한 과목들만 이 리스트에 담아줘야함
DSsubjectObject = []
SKYsubjectObject = []
EBSsubjectObject = []

korean = None
math = None
english = None
korhistory = None
society = None
science = None
univ = None
foreign = None

# 대성마이맥 객체
dskorean = None
dsmath = None
dsenglish = None
dskorhistory = None
dssociety = None
dsscience = None
dsuniv = None
dsforeign = None
# 대성마이맥 객체

# EBS 객체
ebs = Ebs('http://www.ebsi.co.kr/index.jsp')
# EBS 객체

# 스카이에듀 객체
skykorean = None
skymath = None
skyenglish = None
skykorhistory = None
skysociety = None
skyscience = None
skyuniv = None
skyforeign = None
skydaechi = None
skynaesin = None
# 스카이에듀 객체

EBS_ID = ""
EBS_PW = ""

OPT1 = ''
delayTime = 0
startDate = 0
endDate = 0
reserveDate = 0
reserveTime = 0
selectedParseList = []
selectedParseList2 = []
selectedParseList3 = []
selectedParseList4 = []
selectedParseListforRemoveEBS = []
parsingMode = 0  # 파싱모드 , 0 : 전체파싱  1: 선생별 개별파싱
filepath = ""
mythread = None
labelstatus = None
labelstatus2 = None

labelstatus15 = None
labelstatus16 = None
labelstatus17 = None
labelstatus18 = None
labelstatus19 = None
labelstatus20 = None
labelstatus21 = None
labelstatus22 = None

listWidget = None
listWidget2 = None
listWidget3 = None
listWidget4 = None

startButton = None
pauseButton = None
resetButton = None
syncTList = None
addButton = None
delButton = None

tabWidget = None
tabWidgetIndex = None
reserveOPT = None

is_pause= 0 #processing pause : hk.kim-18.01.28
check_stop_class = None #processing pause : hk.kim-18.01.28
threadSelector = ""

driver_global = None

def setWebDriver(Option):
	labelstatus2.setText('네트워크 접속중..')
	global driver_global
	if Option == 'OFF':
		options = webdriver.ChromeOptions()                        	# 헤드리스 옵션!
		options.add_argument("headless")							# 헤드리스 옵션!
		options.add_argument("--disable-gpu")
		driver = webdriver.Chrome(chrome_options=options)		# 헤드리스 옵션!
		# print('Chrome Headless Webdriver load --- ok')
		# global driver_global
		driver_global =  driver
		return driver
	elif Option == 'ON' :
		driver = webdriver.Chrome()
		# print('Chrome Webdriver load --- ok')
		# global driver_global
		driver_global =  driver
		return driver
	labelstatus2.setText('네트워크 접속 완료')


class MyWindow(QMainWindow, form_class):

	def __init__(self):
		super().__init__()
		self.setupUi(self)
		self.setGeometry(1000, 50, 337, 860)
		self.setFixedSize(907, 696)
		app.aboutToQuit.connect(self.closeEvent)
		self.label_Status.setText('System Ready')
		self.label_Status_2.setText('')
		self.label_22.setText('')
		# self.autoAddTeacher()

		####
		#elf.thread = MyThread()

		#self.thread.threadEvent.connect(self.threadEventHandler)
		self.pushButton.clicked.connect(self.runAnalyze)
		self.pushButton_2.clicked.connect(self.threadStop)
		self.pushButton_2.setDisabled(True)
		self.pushButton_6.clicked.connect(self.exit)
		self.pushButton_6.setDisabled(True)
		####

		self.listWidget.setSelectionMode(QAbstractItemView.MultiSelection)
		self.listWidget_2.setSelectionMode(QAbstractItemView.MultiSelection)
		self.listWidget_3.setSelectionMode(QAbstractItemView.MultiSelection)
		self.listWidget_4.setSelectionMode(QAbstractItemView.MultiSelection)
		self.listWidget_5.setSelectionMode(QAbstractItemView.MultiSelection)
		self.listWidget_6.setSelectionMode(QAbstractItemView.MultiSelection)
		self.listWidget_7.setSelectionMode(QAbstractItemView.MultiSelection)
		self.listWidget_8.setSelectionMode(QAbstractItemView.MultiSelection)
		#self.listWidget.show()
		#self.listWidget_2.show()

		#self.pushButton_2.clicked.connect(self.exit)
		self.checkBox_MEGA.stateChanged.connect(self.checkBoxMEGA)
		self.checkBox_EBS.stateChanged.connect(self.checkBoxEBS)
		self.checkBox_DS.stateChanged.connect(self.checkBoxDS)
		self.checkBox_SKY.stateChanged.connect(self.checkBoxSKY)
		self.checkBox1.stateChanged.connect(self.checkBoxState1)
		self.checkBox2.stateChanged.connect(self.checkBoxState2)
		self.checkBox3.stateChanged.connect(self.checkBoxState3)
		self.checkBox4.stateChanged.connect(self.checkBoxState4)
		self.checkBox5.stateChanged.connect(self.checkBoxState5)
		self.checkBox6.stateChanged.connect(self.checkBoxState6)
		self.checkBox7.stateChanged.connect(self.checkBoxState7)
		self.checkBox8.stateChanged.connect(self.checkBoxState8)
		self.checkBox9.stateChanged.connect(self.checkBoxState9)
		self.checkBox10.stateChanged.connect(self.checkBoxState10)
		self.checkBoxAll.stateChanged.connect(self.checkBoxStateAll)
		self.dateEdit.setDate(QtCore.QDate.currentDate().addDays(-7))
		self.dateEdit_2.setDate(QtCore.QDate.currentDate().addDays(-1))
		self.dateEdit.dateChanged.connect(self.dateChange)
		self.dateEdit_2.dateChanged.connect(self.dateChange)
		self.dateTimeEdit.setDate(QtCore.QDate.currentDate())
		self.dateTimeEdit.setTime(QtCore.QTime.currentTime())
		self.dateTimeEdit.dateTimeChanged.connect(self.reserveDateChange)
		self.checkBox_OPT1.stateChanged.connect(self.checkBoxOPT)
		self.checkBox_OPT2.stateChanged.connect(self.checkBoxOPT)
		self.checkBox_OPT3.stateChanged.connect(self.checkBoxOPT)
		self.spinBox.valueChanged.connect(self.delayTimeChange)

		self.pushButton_3.clicked.connect(self.adding)
		self.pushButton_4.clicked.connect(self.removing)
		self.pushButton_5.clicked.connect(self.pathChange)

		self.pushButton_7.clicked.connect(self.addTeacherList)
		self.pushButton_8.clicked.connect(self.addSiteTeacherList)

		mythread.finished.connect(self.unlock_All)

		global labelstatus
		labelstatus = self.label_Status
		global labelstatus2
		labelstatus2 = self.label_Status_2
		global labelstatus15
		labelstatus15 = self.label_Status_15
		global labelstatus16
		labelstatus16 = self.label_Status_16
		global labelstatus17
		labelstatus17 = self.label_Status_17
		global labelstatus18
		labelstatus18 = self.label_Status_18
		global labelstatus19
		labelstatus19 = self.label_Status_19
		global labelstatus20
		labelstatus20 = self.label_Status_20
		global labelstatus21
		labelstatus21 = self.label_Status_21
		global labelstatus22
		labelstatus22 = self.label_Status_22

		global listWidget
		listWidget = self.listWidget
		global listWidget2
		listWidget2 = self.listWidget_2
		global listWidget3
		listWidget3 = self.listWidget_3
		global listWidget4
		listWidget4 = self.listWidget_4

		global pauseButton
		pauseButton = self.pushButton_2
		global startButton
		startButton = self.pushButton
		global resetButton
		resetButton = self.pushButton_6
		global addButton
		addButton = self.pushButton_3
		global delButton
		delButton = self.pushButton_4
		global tabWidget
		tabWidget = self.tabWidget

		global reserveOPT
		reserveOPT = self.checkBox_OPT3

		global filepath
		f = open('path.txt', 'r')
		filepath = f.read()
		self.label_11.setText(filepath)

		global EBS_ID
		global EBS_PW
		f = open('EBS_login.txt', 'r')
		EBS_ID = f.readline().strip().split(':')[1]
		EBS_PW = f.readline().split(':')[1]
		self.label_23.setText(EBS_ID)

		self.dateChange()
		self.checkBoxOPT()

		self.webdriver = webdriver

	def closeEvent(self, event):
		print('Close button pressed')
		print(driver_global)
		driver_global.quit()
		import sys
		sys.exit(0)


	def addTeacherList(self):
		global threadSelector
		threadSelector = "TeacherList"
		self.threadStart()

	def addSiteTeacherList(self):
		global threadSelector
		threadSelector = "SiteTeacherList"
		global tabWidgetIndex
		tabWidgetIndex = self.tabWidget.currentIndex()
		self.threadStart()

	def runAnalyze(self):
		global threadSelector
		threadSelector = "runAnalyze"
		self.threadStart()

	def adding(self):
		if self.tabWidget.currentIndex() == 0:  # 메가스터디 탭
			self.listWidget_5.clear()
			global selectedParseList
			selectedParseList = []
			selected = [item.text() for item in self.listWidget.selectedItems()]
			#print(selected)
			for i in range(0, len(selected)):
				self.listWidget_5.addItem(selected[i])
				selectedParseList.append(selected[i])
			self.listWidget.clearSelection()
			#print(selectedParseList)

		elif self.tabWidget.currentIndex() == 1:  # EBS 탭
			self.listWidget_6.clear()
			global selectedParseList2
			selectedParseList2 = []
			selected = [item.text() for item in self.listWidget_2.selectedItems()]
			for i in range(0, len(selected)):
				self.listWidget_6.addItem(selected[i])
				selectedParseList2.append(selected[i])
			self.listWidget_2.clearSelection()
			#print(selectedParseList2)

		elif self.tabWidget.currentIndex() == 2:  # 대성마이맥 탭
			self.listWidget_7.clear()
			global selectedParseList3
			selectedParseList3 = []
			selected = [item.text() for item in self.listWidget_3.selectedItems()]
			for i in range(0, len(selected)):
				self.listWidget_7.addItem(selected[i])
				selectedParseList3.append(selected[i])
			self.listWidget_3.clearSelection()
			#print(selectedParseList3)

		elif self.tabWidget.currentIndex() == 3:  # 스카이에듀 탭
			self.listWidget_8.clear()
			global selectedParseList4
			selectedParseList4 = []
			selected = [item.text() for item in self.listWidget_4.selectedItems()]
			for i in range(0, len(selected)):
				self.listWidget_8.addItem(selected[i])
				selectedParseList4.append(selected[i])
			self.listWidget_4.clearSelection()
			#print(selectedParseList4)

		if len(selectedParseList) > 0 or len(selectedParseList2) > 0 or len(selectedParseList3) > 0 or len(selectedParseList4) > 0:
			self.lock_CheckBox()
			global parsingMode
			parsingMode = 1

	def removing(self):
		if self.tabWidget.currentIndex() == 0:  # 메가스터디 탭
			selected = [item.text() for item in self.listWidget_5.selectedItems()]
			self.listWidget_5.clear()
			for i in range(0, len(selected)):
				selectedParseList.remove(selected[i])
			for j in range(0, len(selectedParseList)):
				self.listWidget_5.addItem(selectedParseList[j])
			#print(selectedParseList)

		elif self.tabWidget.currentIndex() == 1:  # EBS 탭
			selected = [item.text() for item in self.listWidget_6.selectedItems()]
			self.listWidget_6.clear()
			for i in range(0, len(selected)):
				selectedParseList2.remove(selected[i])
			for j in range(0, len(selectedParseList2)):
				self.listWidget_6.addItem(selectedParseList2[j])
			#print(selectedParseList2)

		elif self.tabWidget.currentIndex() == 2:  # 대성마이맥 탭
			selected = [item.text() for item in self.listWidget_7.selectedItems()]
			self.listWidget_7.clear()
			for i in range(0, len(selected)):
				selectedParseList3.remove(selected[i])
			for j in range(0, len(selectedParseList3)):
				self.listWidget_7.addItem(selectedParseList3[j])
			#print(selectedParseList3)

		elif self.tabWidget.currentIndex() == 3:  # 스카이에듀 탭
			selected = [item.text() for item in self.listWidget_8.selectedItems()]
			self.listWidget_8.clear()
			for i in range(0, len(selected)):
				selectedParseList4.remove(selected[i])
			for j in range(0, len(selectedParseList4)):
				self.listWidget_8.addItem(selectedParseList4[j])
			#print(selectedParseList4)

		if len(selectedParseList) == 0 and len(selectedParseList2) == 0 and len(selectedParseList3) == 0 and len(selectedParseList4) == 0:
			self.unlock_CheckBox()
			global parsingMode
			parsingMode = 0

	def unlock_All(self):
		self.unlock_CheckBox()
		self.unlock_Date_and_Option()

	def lock_CheckBox(self):
		self.checkBox_MEGA.setEnabled(False)
		self.checkBox_EBS.setEnabled(False)
		self.checkBox_DS.setEnabled(False)
		self.checkBox_SKY.setEnabled(False)
		self.checkBox1.setEnabled(False)
		self.checkBox2.setEnabled(False)
		self.checkBox3.setEnabled(False)
		self.checkBox4.setEnabled(False)
		self.checkBox5.setEnabled(False)
		self.checkBox6.setEnabled(False)
		self.checkBox7.setEnabled(False)
		self.checkBox8.setEnabled(False)
		self.checkBox9.setEnabled(False)
		self.checkBox10.setEnabled(False)
		self.checkBoxAll.setEnabled(False)

	def unlock_CheckBox(self):
		self.checkBox_MEGA.setEnabled(True)
		self.checkBox_EBS.setEnabled(True)
		self.checkBox_DS.setEnabled(True)
		self.checkBox_SKY.setEnabled(True)
		self.checkBox1.setEnabled(True)
		self.checkBox2.setEnabled(True)
		self.checkBox3.setEnabled(True)
		self.checkBox4.setEnabled(True)
		self.checkBox5.setEnabled(True)
		self.checkBox6.setEnabled(True)
		self.checkBox7.setEnabled(True)
		self.checkBox8.setEnabled(True)
		self.checkBox9.setEnabled(True)
		self.checkBox10.setEnabled(True)
		self.checkBoxAll.setEnabled(True)

	def lock_Date_and_Option(self):
		self.dateEdit.setDisabled(True)
		self.dateEdit_2.setDisabled(True)
		self.checkBox_OPT1.setDisabled(True)
		self.checkBox_OPT2.setDisabled(True)
		self.checkBox_OPT3.setDisabled(True)
		self.dateTimeEdit.setDisabled(True)
		addButton.setDisabled(True)
		delButton.setDisabled(True)

	def unlock_Date_and_Option(self):
		self.dateEdit.setEnabled(True)
		self.dateEdit_2.setEnabled(True)
		self.checkBox_OPT1.setEnabled(True)
		self.checkBox_OPT2.setEnabled(True)
		self.checkBox_OPT3.setEnabled(True)
		self.dateTimeEdit.setEnabled(True)
		addButton.setEnabled(True)
		delButton.setEnabled(True)

	def exit(self):
		global mythread
		mythread.terminate()
		global is_pause
		is_pause = 0
		self.pushButton.setEnabled(True)
		self.pushButton_2.setDisabled(True)
		self.pushButton_2.setText('일시 중지')
		self.label_Status_2.setText('')
		self.unlock_CheckBox()
		self.unlock_Date_and_Option()

	def dateChange(self):
		global startDate
		global endDate
		startDate = int(self.dateEdit.text().replace("-", ""))
		endDate = int(self.dateEdit_2.text().replace("-", ""))
		#startDate = 20180128
		#endDate = 20180201
		#print(startDate, endDate)

	def reserveDateChange(self):
		global reserveTime
		global reserveDate
		userInput = self.dateTimeEdit.text()
		reserveDate = userInput.split(' ')[0]
		reserveTime = userInput.split(' ')[1]
		#print(QDate.currentDate().toString("yyyy-MM-dd") , reserveDate)
		#print(QTime.currentTime().toString("hh:mm") , reserveTime)

	def pathChange(self):
		global filepath
		self.root = Tk()
		self.root.dirName = filedialog.askdirectory()
		self.label_11.setText(self.root.dirName)
		filepath = self.root.dirName + '/'
		f = open("path.txt",'w')
		f.write(filepath)
		self.root.destroy()

	def delayTimeChange(self):
		global delayTime
		delayTime = int(self.spinBox.value())

	def checkBoxMEGA(self):
		if self.checkBox_MEGA.isChecked() == True:
			SelectedSite.append('MEGA')
		elif self.checkBox_MEGA.isChecked() == False:
			SelectedSite.remove('MEGA')
	def checkBoxEBS(self):
		if self.checkBox_EBS.isChecked() == True:
			SelectedSite.append('EBS')
		elif self.checkBox_EBS.isChecked() == False:
			SelectedSite.remove('EBS')
	def checkBoxDS(self):
		if self.checkBox_DS.isChecked() == True:
			SelectedSite.append('DS')
		elif self.checkBox_DS.isChecked() == False:
			SelectedSite.remove('DS')
	def checkBoxSKY(self):
		if self.checkBox_SKY.isChecked() == True:
			SelectedSite.append('SKY')
		elif self.checkBox_SKY.isChecked() == False:
			SelectedSite.remove('SKY')

	def checkBoxState1(self):
		if self.checkBox1.isChecked() == True:
			MegasubjectObject.append(korean)
			DSsubjectObject.append(dskorean)
			SKYsubjectObject.append(skykorean)
			EBSsubjectObject.append('국어')
		elif self.checkBox1.isChecked() == False:
			MegasubjectObject.remove(korean)
			DSsubjectObject.remove(dskorean)
			SKYsubjectObject.remove(skykorean)
			EBSsubjectObject.remove('국어')
	def checkBoxState2(self):
		if self.checkBox2.isChecked() == True:
			MegasubjectObject.append(math)
			DSsubjectObject.append(dsmath)
			SKYsubjectObject.append(skymath)
			EBSsubjectObject.append('수학')
		elif self.checkBox2.isChecked() == False:
			MegasubjectObject.remove(math)
			DSsubjectObject.remove(dsmath)
			SKYsubjectObject.remove(skymath)
			EBSsubjectObject.remove('수학')
	def checkBoxState3(self):
		if self.checkBox3.isChecked() == True:
			MegasubjectObject.append(english)
			DSsubjectObject.append(dsenglish)
			SKYsubjectObject.append(skyenglish)
			EBSsubjectObject.append('영어')
		elif self.checkBox3.isChecked() == False:
			MegasubjectObject.remove(english)
			DSsubjectObject.remove(dsenglish)
			SKYsubjectObject.remove(skyenglish)
			EBSsubjectObject.remove('영어')
	def checkBoxState4(self):
		if self.checkBox4.isChecked() == True:
			MegasubjectObject.append(korhistory)
			DSsubjectObject.append(dskorhistory)
			SKYsubjectObject.append(skykorhistory)
			EBSsubjectObject.append('한국사')
		elif self.checkBox4.isChecked() == False:
			MegasubjectObject.remove(korhistory)
			DSsubjectObject.remove(dskorhistory)
			SKYsubjectObject.remove(skykorhistory)
			EBSsubjectObject.remove('한국사')
	def checkBoxState5(self):
		if self.checkBox5.isChecked() == True:
			MegasubjectObject.append(society)
			DSsubjectObject.append(dssociety)
			SKYsubjectObject.append(skysociety)
			EBSsubjectObject.append('사회')
		elif self.checkBox5.isChecked() == False:
			MegasubjectObject.remove(society)
			DSsubjectObject.remove(dssociety)
			SKYsubjectObject.remove(skysociety)
			EBSsubjectObject.remove('사회')
	def checkBoxState6(self):
		if self.checkBox6.isChecked() == True:
			MegasubjectObject.append(science)
			DSsubjectObject.append(dsscience)
			SKYsubjectObject.append(skyscience)
			EBSsubjectObject.append('과학')
		elif self.checkBox6.isChecked() == False:
			MegasubjectObject.remove(science)
			DSsubjectObject.remove(dsscience)
			SKYsubjectObject.remove(skyscience)
			EBSsubjectObject.remove('과학')
	def checkBoxState7(self):
		if self.checkBox7.isChecked() == True:
			MegasubjectObject.append(univ)
			DSsubjectObject.append(dsuniv)
			SKYsubjectObject.append(skyuniv)
			# SKYsubjectObject.append(skydaechi)
			# SKYsubjectObject.append(skynaesin)
        	# 수정 - hk.kim 20.05.17
			EBSsubjectObject.append('대학별고사')
		elif self.checkBox7.isChecked() == False:
			MegasubjectObject.remove(univ)
			DSsubjectObject.remove(dsuniv)
			SKYsubjectObject.remove(skyuniv)
			# SKYsubjectObject.remove(skydaechi)
			# SKYsubjectObject.remove(skynaesin)
        	# 수정 - hk.kim 20.05.17
			EBSsubjectObject.remove('대학별고사')
	def checkBoxState8(self):
		if self.checkBox8.isChecked() == True:
			MegasubjectObject.append(foreign)
			DSsubjectObject.append(dsforeign)
			SKYsubjectObject.append(skyforeign)
			EBSsubjectObject.append('제2외국어')
		elif self.checkBox8.isChecked() == False:
			MegasubjectObject.remove(foreign)
			DSsubjectObject.remove(dsforeign)
			SKYsubjectObject.remove(skyforeign)
			EBSsubjectObject.remove('제2외국어')
	def checkBoxState9(self):
		if self.checkBox9.isChecked() == True:
			EBSsubjectObject.append('직업')
		elif self.checkBox9.isChecked() == False:
			EBSsubjectObject.remove('직업')
	def checkBoxState10(self):
		if self.checkBox10.isChecked() == True:
			EBSsubjectObject.append('일반/진로/교양')
		elif self.checkBox10.isChecked() == False:
			EBSsubjectObject.remove('일반/진로/교양')
	def checkBoxStateAll(self):
		if self.checkBoxAll.isChecked() == True:
			self.checkBox1.setChecked(True)
			self.checkBox2.setChecked(True)
			self.checkBox3.setChecked(True)
			self.checkBox4.setChecked(True)
			self.checkBox5.setChecked(True)
			self.checkBox6.setChecked(True)
			self.checkBox7.setChecked(True)
			self.checkBox8.setChecked(True)
			self.checkBox9.setChecked(True)
			self.checkBox10.setChecked(True)

		elif self.checkBoxAll.isChecked() == False:
			print('test1')
			self.checkBox1.setChecked(False)
			print('test2')
			self.checkBox2.setChecked(False)
			self.checkBox3.setChecked(False)
			print('test3')
			self.checkBox4.setChecked(False)
			self.checkBox5.setChecked(False)
			print('test4')
			self.checkBox6.setChecked(False)
			self.checkBox7.setChecked(False)
			print('test5')
			self.checkBox8.setChecked(False)
			print('test6')
			self.checkBox9.setChecked(False)
			print('test7')
			self.checkBox10.setChecked(False)

	def checkBoxOPT(self):
		global delayTime
		global OPT1
		if self.checkBox_OPT1.isChecked() == True:
			OPT1 = 'ON'
		elif self.checkBox_OPT1.isChecked() == False:
			OPT1 = 'OFF'
		
		if self.checkBox_OPT2.isChecked() == True:
			self.spinBox.setEnabled(True)
		elif self.checkBox_OPT2.isChecked() == False:
			delayTime = 0
			self.spinBox.setDisabled(True)

		if self.checkBox_OPT3.isChecked() == True:
			self.dateTimeEdit.setEnabled(True)
		elif self.checkBox_OPT3.isChecked() == False:
			self.dateTimeEdit.setDisabled(True)

	def threadStop(self):
        #processing pause : hk.kim-18.01.28
		global is_pause
		if is_pause == 0:
			is_pause = 1
			self.pushButton_2.setText('재개')
			self.pushButton_6.setEnabled(True)
		else:
			is_pause = 0
			self.pushButton_2.setText('일시 중지')
			self.pushButton_6.setDisabled(True)
		self.label_Status.setText('집계가 종료되었습니다.')

	def threadStart(self):
         #processing pause : hk.kim-18.01.28
         self.lock_CheckBox()
         self.lock_Date_and_Option()
         self.pushButton_2.setEnabled(True)
         global mythread
         mythread.start()


class DataAnalyze(QThread):
	def __init__(self, parent=None):
		super().__init__()
		  ########################### Test
	# global OPT1
	# global startDate
	# global endDate
	# global parsingMode

	def run(self):
		global threadSelector
		global labelstatus
		if threadSelector == "runAnalyze":
			if reserveOPT.isChecked():
				timeMatched = 0
				while timeMatched == 0:
					labelstatus.setText('예약 실행 대기중')
					labelstatus2.setText('예약시간: ' + reserveDate + ' ' + reserveTime)
					if QDate.currentDate().toString("yyyy-MM-dd") == reserveDate and QTime.currentTime().toString("hh:mm") == reserveTime :
						timeMatched = 1
				labelstatus.setText('집계를 시작합니다.')
				startButton.setDisabled(True)
				self.analyzeStart()
			else:
				labelstatus.setText('집계를 시작합니다.')
				startButton.setDisabled(True)
				self.analyzeStart()
		elif threadSelector == "TeacherList":
			self.updateTList()
			startButton.setEnabled(True)
			pauseButton.setDisabled(True)
			resetButton.setDisabled(True)
		elif threadSelector == "SiteTeacherList":
			self.driver = setWebDriver('OFF')
			if tabWidgetIndex == 0:
				# 와드
				try:
					labelstatus2.setText('메가스터디 목록 불러오는 중..')
					self.settingMega()
					labelstatus2.setText('메가스터디 목록 불러오기 완료')
					self.driver.quit()
				# 와드
				except:
					labelstatus2.setText('메가스터디 목록 불러오기 실패')
					self.driver.quit()
			elif tabWidgetIndex == 1:
				try:
					labelstatus2.setText('EBS 목록 불러오는 중..')
					self.settingEBS()
					labelstatus2.setText('EBS 목록 불러오기 완료')
					self.driver.quit()
				except:
					labelstatus2.setText('EBS 목록 불러오기 실패')
					self.driver.quit()
			elif tabWidgetIndex == 2:
				try:
					labelstatus2.setText('대성마이맥 목록 불러오는 중..')
					self.settingDS()
					labelstatus2.setText('대성마이맥 목록 불러오기 완료')
					self.driver.quit()
				except:
					labelstatus2.setText('대성마이맥 목록 불러오기 실패')
					self.driver.quit()
			elif tabWidgetIndex == 3:
				try:
					labelstatus2.setText('스카이에듀 목록 불러오는 중..')
					self.settingSKY()
					labelstatus2.setText('스카이에듀 목록 불러오기 완료')
					self.driver.quit()
				except:
					labelstatus2.setText('스카이에듀 목록 불러오기 실패')
					self.driver.quit()
			

	def updateTList(self):
		self.driver = setWebDriver('OFF')
		errorMessage = []
		labelstatus2.setText('메가스터디 목록 불러오는 중')
		try:
			self.settingMega()
		except:
			errorMessage.append('메가')
		labelstatus2.setText('EBS 목록 불러오는 중')
		try:
			self.settingEBS()
		except:
			errorMessage.append('EBS')
		labelstatus2.setText('대성마이맥 목록 불러오는 중')
		try:
			self.settingDS()
		except:
			errorMessage.append('대성')
		labelstatus2.setText('스카이에듀 목록 불러오는 중')
		try:
			self.settingSKY()
		except:
			errorMessage.append('스카이')
		if len(errorMessage) > 0:
			labelstatus2.setText(','.join(errorMessage) +' 로딩 실패')
		else:
			labelstatus2.setText('선생님 목록 불러오기 완료')
		self.driver.quit()

	def settingMega(self):
		self.driver = setWebDriver("OFF")
		labelstatus2.setText('메가스터디 목록 불러오는 중')
		Mega = GoMegastudy(self.driver)
		tpage = Mega.tpage()
		global korean
		global math
		global english
		global korhistory
		global society
		global science
		global univ
		global foreign
		korean = SetMegastudy(department['MEGA']['국어'], tpage)
		math = SetMegastudy(department['MEGA']['수학'], tpage)
		english = SetMegastudy(department['MEGA']['영어'], tpage)
		korhistory = SetMegastudy(department['MEGA']['한국사'], tpage)
		society = SetMegastudy(department['MEGA']['사회'], tpage)
		science = SetMegastudy(department['MEGA']['과학'], tpage)
		univ = SetMegastudy(department['MEGA']['대학별고사'], tpage)
		foreign = SetMegastudy(department['MEGA']['제2외국어한문'], tpage)

		MegasubjectObject = [korean, math, english, korhistory, society, science, univ, foreign]
		lenMega = len(MegasubjectObject)
		MegaTeacherList = []
		for i in range(0, lenMega):
			subjectTList = MegasubjectObject[i].getFullList()
			for j in range(0, len(subjectTList)):
				MegaTeacherList.append(subjectTList[j])
		listWidget.clear()
		for x in range(0, len(MegaTeacherList)):
			listWidget.addItem(MegaTeacherList[x])
		MegasubjectObject = []
		#self.driver.quit()

	def settingEBS(self):
		self.driver = setWebDriver("OFF")
		labelstatus2.setText('EBS 목록 불러오는 중')
		ebs.set_driver(self.driver)
		ebs.go_to_url_page('http://www.ebsi.co.kr/ebs/pot/potn/retrieveTchrSubMain.ebs?Clickz=G001', 0)
		lecture_array = ebs.get_lecture_list()
		teacher_array = ebs.get_teacher_list(lecture_array)
		listWidget2.clear()
		for h in range(0, len(teacher_array)):
			listWidget2.addItem(str(teacher_array[h].get_full_info()))

	def settingDS(self):
		self.driver = setWebDriver("OFF")
		labelstatus2.setText('대성마이맥 목록 불러오는 중')
		Mega = GoDaesung(self.driver)
		tpage = Mega.tpage()
		global dskorean
		global dsmath
		global dsenglish
		global dskorhistory
		global dssociety
		global dsscience
		global dsuniv
		global dsforeign
		dskorean = SetDaesung(department['DS']['국어'], tpage)
		dsmath = SetDaesung(department['DS']['수학'], tpage)
		dsenglish = SetDaesung(department['DS']['영어'], tpage)
		dskorhistory = SetDaesung(department['DS']['한국사'], tpage)
		dssociety = SetDaesung(department['DS']['사회'], tpage)
		dsscience = SetDaesung(department['DS']['과학'], tpage)
		dsuniv = SetDaesung(department['DS']['대학별고사'], tpage)
		dsforeign = SetDaesung(department['DS']['제2외국어한문'], tpage)

		DSsubjectObject = [dskorean, dsmath, dsenglish, dskorhistory, dssociety, dsscience, dsuniv, dsforeign]
		lenDS = len(DSsubjectObject)
		DSTeacherList = []
		for q in range(0, lenDS):
			DSsubjectTList = DSsubjectObject[q].getFullList()
			for w in range(0, len(DSsubjectTList)):
				DSTeacherList.append(DSsubjectTList[w])
		listWidget3.clear()
		for e in range(0, len(DSTeacherList)):
			listWidget3.addItem(DSTeacherList[e])
		DSsubjectObject = []

	def settingSKY(self):
		self.driver = setWebDriver('OFF')
		labelstatus2.setText('스카이에듀 목록 불러오는 중')
		Sky = GoSkyedu(self.driver)  # 스카이에듀는 드라이버 없이 beautifulsoup으로만 가져올 수 있음
		tpage = Sky.tpage()
		global skykorean
		global skymath
		global skyenglish
		global skykorhistory
		global skysociety
		global skyscience
		global skyuniv
		global skyforeign
		global skydaechi
		global skynaesin
		skykorean = SetSkyedu(department['SKY']['국어'], tpage)
		skymath = SetSkyedu(department['SKY']['수학'], tpage)
		skyenglish = SetSkyedu(department['SKY']['영어'], tpage)
		skykorhistory = SetSkyedu(department['SKY']['한국사'], tpage)
		skysociety = SetSkyedu(department['SKY']['사회'], tpage)
		skyscience = SetSkyedu(department['SKY']['과학'], tpage)
		skyuniv = SetSkyedu(department['SKY']['대학별고사'], tpage)
		skyforeign = SetSkyedu(department['SKY']['제2외국어'], tpage)
		# skydaechi = SetSkyedu(department['SKY']['월간대치동'], tpage)
		# skynaesin = SetSkyedu(department['SKY']['내신전문'], tpage)

		# SKYsubjectObject = [skykorean, skymath, skyenglish, skykorhistory, skysociety, skyscience, skyuniv, skyforeign, skydaechi, skynaesin]
		# SKYsubjectObject = [skykorean, skymath, skyenglish, skykorhistory, skysociety, skyscience, skyuniv, skyforeign, skynaesin]
		SKYsubjectObject = [skykorean, skymath, skyenglish, skykorhistory, skysociety, skyscience, skyuniv, skyforeign]
		lenSKY = len(SKYsubjectObject)
		SKYTeacherList = []
		for a in range(0, lenSKY):
			SKYsubjectTList = SKYsubjectObject[a].getFullList()
			for s in range(0, len(SKYsubjectTList)):
				SKYTeacherList.append(SKYsubjectTList[s])
		listWidget4.clear()
		for d in range(0, len(SKYTeacherList)):
			listWidget4.addItem(SKYTeacherList[d])
		SKYsubjectObject = []


	def analyzeStart(self):
		driver = setWebDriver(OPT1)  # 페이지 넘어가면서 파싱을 위해 웹드라이버 셋팅
		# driver.set_page_load_timeout(15)
		if parsingMode == 0:
			# 와드
			# try:
				for site in SelectedSite:
					if site == 'MEGA':
						labelstatus15.setText(str(strftime("%Y-%m-%d %H:%M")))
						self.analyzeMega(driver)
						labelstatus16.setText(str(strftime("%Y-%m-%d %H:%M")))
					elif site == 'EBS':
						labelstatus17.setText(str(strftime("%Y-%m-%d %H:%M")))
						self.analyzeEBS(driver)
						labelstatus18.setText(str(strftime("%Y-%m-%d %H:%M")))
					elif site == 'DS':
						labelstatus19.setText(str(strftime("%Y-%m-%d %H:%M")))
						self.analyzeDS(driver)
						labelstatus20.setText(str(strftime("%Y-%m-%d %H:%M")))
					elif site == 'SKY':
						labelstatus21.setText(str(strftime("%Y-%m-%d %H:%M")))
						self.analyzeSKY(driver)
						labelstatus22.setText(str(strftime("%Y-%m-%d %H:%M")))
				labelstatus.setText('집계 완료. 엑셀 파일을 확인해주세요')
				startButton.setEnabled(True)
				pauseButton.setDisabled(True)
				resetButton.setDisabled(True)
				driver.quit()
			# 와드
			# except:
			# 	labelstatus.setText('집계 오류 - main-908')
			# 	startButton.setEnabled(True)
			# 	pauseButton.setDisabled(True)
			# 	resetButton.setDisabled(True)
			# 	driver.quit()
		elif parsingMode == 1:
			# try:
			if len(selectedParseList) > 0:
				labelstatus15.setText(str(strftime("%Y-%m-%d %H:%M")))
				self.analyzeMega(driver)
				labelstatus16.setText(str(strftime("%Y-%m-%d %H:%M")))
			if len(selectedParseList2) > 0:
				labelstatus17.setText(str(strftime("%Y-%m-%d %H:%M")))
				self.analyzeEBS(driver)
				labelstatus18.setText(str(strftime("%Y-%m-%d %H:%M")))
			if len(selectedParseList3) > 0:
				labelstatus19.setText(str(strftime("%Y-%m-%d %H:%M")))
				self.analyzeDS(driver)
				labelstatus20.setText(str(strftime("%Y-%m-%d %H:%M")))
			if len(selectedParseList4) > 0:
				labelstatus21.setText(str(strftime("%Y-%m-%d %H:%M")))
				self.analyzeSKY(driver)
				labelstatus22.setText(str(strftime("%Y-%m-%d %H:%M")))
			startButton.setEnabled(True)
			pauseButton.setDisabled(True)
			resetButton.setDisabled(True)
			labelstatus.setText('개별 집계 완료. 엑셀 파일을 확인해주세요')
			driver.quit()
			# except:
			# 	startButton.setEnabled(True)
			# 	pauseButton.setDisabled(True)
			# 	resetButton.setDisabled(True)
			# 	labelstatus.setText('집계 오류 - main-940')
			# 	driver.quit()
	def analyzeMega(self, driver):
		global labelstatus
		global check_stop_class
		def workBook(filename):
			workbook = xlsxwriter.Workbook(filename)  # 'math.xlsx'
			return workbook
		if parsingMode == 0 :
			labelstatus.setText('메가스터디 집계를 시작합니다.')
			
			subjectresultForExcel = []

			selectedPersonNum = 0
			progress = 0
			for u in range(0, len(MegasubjectObject)):
				selectedPersonNum = len(MegasubjectObject[u].getIDList()) + selectedPersonNum
			
			for x in range(0, len(MegasubjectObject)):
				IDList = MegasubjectObject[x].getIDList()
				NameList = MegasubjectObject[x].getNameList()
				SubjectList = MegasubjectObject[x].getSubjectList()
				
				for i in range(0, len(IDList)):
					labelstatus.setText('메가스터디 ' + NameList[i] + ' 선생님 집계중  ' + str(progress + i + 1) + '/' + str(selectedPersonNum))
					calcBoards = CalcMegastudy(IDList[i], endDate, startDate, delayTime, driver)  # url, startdate, enddate, waitTime, chromedriver
					semiresult = calcBoards.calcBoard(check_stop_class, labelstatus2)  # processing pause : hk.kim-18.01.28
					result = calcBoards.dataResult(semiresult, NameList[i], SubjectList[0])  # calcBoardResult, teacherName, subjectName
					for j in range(0, len(result)):
						subjectresultForExcel.append(result[j])
					excelFile = workBook(filepath + '메가스터디_' + str(startDate) + '-' + str(endDate) + '.xlsx')
					calcBoards.xlsxWrite(excelFile, subjectresultForExcel)
					excelFile.close()
					labelstatus2.setText('')
					labelstatus.setText('메가스터디 ' + NameList[i] + ' 선생님 집계 완료')
				progress = progress + len(IDList)
			labelstatus.setText('메가스터디 집계 종료 : 엑셀파일을 확인해주세요')

		elif parsingMode == 1:
			labelstatus.setText('메가스터디 개별 집계를 시작합니다.')
			MEGAdic = self.IdNameDicMEGA()
			IDList = []
			NameList = []
			SubjectList = []
			for t in range(0, len(selectedParseList)):
				split = selectedParseList[t].split(':')  #국어: 김선겸
				IDList.append(MEGAdic[split[1].lstrip()])
				NameList.append(split[1].lstrip())
				SubjectList.append(split[0])
			# print(IDList)
			subjectresultForExcel = []
			for i in range(0, len(IDList)):
				# print(NameList[i], IDList[i])
				labelstatus.setText('메가스터디 ' + NameList[i] + ' 선생님 집계중')
				calcBoards = CalcMegastudy(IDList[i], endDate, startDate, delayTime, driver)  # url, startdate, enddate, waitTime, chromedriver
				semiresult = calcBoards.calcBoard(check_stop_class, labelstatus2) #processing pause : hk.kim-18.01.28
				result = calcBoards.dataResult(semiresult, NameList[i], SubjectList[i])  # calcBoardResult, teacherName, subjectName
				for j in range(0, len(result)):
					subjectresultForExcel.append(result[j])
				excelFile = workBook(filepath + '메가스터디_개별_' + str(startDate) + '-' + str(endDate) + '.xlsx')
				calcBoards.xlsxWrite(excelFile, subjectresultForExcel)
				excelFile.close()
				labelstatus.setText('메가스터디 ' + NameList[i] + ' 선생님 집계 완료')
				labelstatus2.setText('')
			labelstatus.setText('메가스터디 개별 집계 종료 : 엑셀파일을 확인해주세요')

	def analyzeEBS(self, driver):
		def workBook(filename):
			workbook = xlsxwriter.Workbook(filename)  # 'math.xlsx'
			return workbook
		if parsingMode == 0:
			labelstatus.setText('EBS 집계를 시작합니다.')
			ebs.set_driver(driver)
			ebs.go_to_url_page('http://www.ebsi.co.kr/ebs/pot/potl/login.ebs?destination=/ebs/pot/potn/retrieveTchrSubMain.ebs%3FClickz%3DG001&alertYn=N', 0)
			driver.implicitly_wait(delayTime)
			ebs.login(EBS_ID, EBS_PW, labelstatus2)
			ebs.go_to_url_page('http://www.ebsi.co.kr/ebs/pot/potn/retrieveTchrSubMain.ebs?Clickz=G001', 0)
			# 과목 가져오기
			lecture_array = ebs.get_lecture_list()
			# 선생님 리스트 가져오기
			teacher_array = ebs.get_teacher_list(lecture_array)
			add_subject_name_list = EBSsubjectObject[:]
			teacher_array = ebs.add_teacher_by_subject(teacher_array, add_subject_name_list)
			subjectresultForExcel = []
			for teachers in teacher_array:
				labelstatus.setText('EBS ' + str(teachers.get_name()) + ' 선생님 집계중')
				# print(str(teachers.get_full_info()))
				ebs.go_to_url_page('http://www.ebsi.co.kr/ebs/pot/potg/retrieveStdQnaList.ebs?teacherId='
								   + teachers.get_code() + '&bbsType=H600&targetCode=null#;', 1)
				startdate = str(startDate)[2:4] + '.' + str(startDate)[4:6] + '.' + str(startDate)[6:]
				enddate = str(endDate)[2:4] + '.' + str(endDate)[4:6] + '.' + str(endDate)[6:]
				#date count : hk.kim-18.01.29

				bbs_counts = ebs.get_bbs_count(delayTime, startdate, enddate, check_stop_class, labelstatus2)  # processing pause : hk.kim-18.01.28
				#date count : hk.kim-18.01.29

				#dategroup = []
				#date_countgroup = {}
				#for counter in bbs_counts:
				#	dateMac = str(counter.date).replace(".", "")
				#	date = '20' + dateMac[0:2] + dateMac[2:4] + dateMac[4:6] # 20180130
				#	dategroup.append(date)
				#	date_countgroup[date] = str(counter.num)
					# subjectresultForExcel.append(str(teachers.get_subject()) + ':' + str(teachers.get_name()) + ':' + date + ':' + str(counter.num))
					# 수학:심주석:20180130:1
					#preResultForExcel.append(str(counter.date))
				# EBS에서는 startdate가 과거날짜, enddate가 오늘날짜(최근날짜)
				startdate_ = datetime.date(int(str(startDate)[0:4]), int(str(startDate)[4:6].lstrip('0')), int(str(startDate)[6:].lstrip('0')))
				enddate_ = datetime.date(int(str(endDate)[0:4]), int(str(endDate)[4:6].lstrip('0')), int(str(endDate)[6:].lstrip('0')))
				# duration = int(str(enddate_ - startdate_).split(" ")[0]) + 1

				date_diff = enddate_ - startdate_
				if startDate == endDate:
					date_diff = 0
				else:
					date_diff = str(enddate_ - startdate_).split(" ")[0]

				duration = int(date_diff) + 1

				for z in range(0, duration):
					dates = startdate_ + datetime.timedelta(z)
					datesFormat = int(str(dates).split('-')[0] + str(dates).split('-')[1] + str(dates).split('-')[2])
					#count = dategroup.count(str(datesFormat))
					# date count : hk.kim-18.01.29
					count = 0
					for count_info in bbs_counts:
						if count_info.date == str(datesFormat):
							count = count_info.num

					subjectresultForExcel.append(str(teachers.get_subject()) + ':' + str(teachers.get_name()) + ':' + str(datesFormat) + ':' + str(count))

					#if count == 0:
					#	subjectresultForExcel.append(str(teachers.get_subject()) + ':' + str(teachers.get_name()) + ':' + str(datesFormat) + ':' + str(0))
					#else:
					#	count_ = date_countgroup[str(datesFormat)]
					#	subjectresultForExcel.append(str(teachers.get_subject()) + ':' + str(teachers.get_name()) + ':' + str(datesFormat) + ':' + count_)

				driver.implicitly_wait(delayTime)
				labelstatus.setText('EBS ' + str(teachers.get_name()) + ' 선생님 집계 완료')
				excelFile = workBook(filepath + 'EBS_' + str(startDate) + '-' + str(endDate) + '.xlsx')
				Ebs.xlsxWrite(excelFile, subjectresultForExcel)
				excelFile.close()
			labelstatus.setText('EBS 집계 종료 : 엑셀파일을 확인해주세요')
			#엑셀 출력 코드 추가 필요

		elif parsingMode == 1:
			labelstatus.setText('EBS 개별 집계를 시작합니다.')
			ebs.set_driver(driver)
			ebs.go_to_url_page('http://www.ebsi.co.kr/ebs/pot/potl/login.ebs?destination=/ebs/pot/potn/retrieveTchrSubMain.ebs%3FClickz%3DG001&alertYn=N', 0)
			driver.implicitly_wait(delayTime)
			ebs.login(EBS_ID, EBS_PW, labelstatus2)
			ebs.go_to_url_page('http://www.ebsi.co.kr/ebs/pot/potn/retrieveTchrSubMain.ebs?Clickz=G001', 0)
			lecture_array = ebs.get_lecture_list()
			teacher_array = ebs.get_teacher_list(lecture_array)
			add_teacher_name_list = []
			for i in range(0, len(selectedParseList2)):
				split = selectedParseList2[i].split(':')
				name = split[1].lstrip()
				add_teacher_name_list.append(name)
			teacher_array = ebs.add_teacher_by_name(teacher_array, add_teacher_name_list)
			subjectresultForExcel = []
			for teachers in teacher_array:
				labelstatus.setText('EBS ' + str(teachers.get_name()) + ' 선생님 집계중')
				ebs.go_to_url_page('http://www.ebsi.co.kr/ebs/pot/potg/retrieveStdQnaList.ebs?teacherId='
								   + teachers.get_code() + '&bbsType=H600&targetCode=null#;', 1)
				startdate = str(startDate)[2:4] + '.' + str(startDate)[4:6] + '.' + str(startDate)[6:]
				enddate = str(endDate)[2:4] + '.' + str(endDate)[4:6] + '.' + str(endDate)[6:]
				#date count : hk.kim-18.01.29
				
				bbs_counts = ebs.get_bbs_count(delayTime, startdate, enddate, check_stop_class, labelstatus2) #processing pause : hk.kim-18.01.28

				#date count : hk.kim-18.01.29
				#dategroup = []
				#date_countgroup = {}
				#for counter in bbs_counts:
				#	dateMac = str(counter.date).replace(".", "")
				#	date = '20' + dateMac[0:2] + dateMac[2:4] + dateMac[4:6] # 20180130
				#	dategroup.append(date)
				#	date_countgroup[date] = str(counter.num)
					# subjectresultForExcel.append(str(teachers.get_subject()) + ':' + str(teachers.get_name()) + ':' + date + ':' + str(counter.num))
					# 수학:심주석:20180130:1
					#preResultForExcel.append(str(counter.date))
				# EBS에서는 startdate가 과거날짜, enddate가 오늘날짜(최근날짜)
				startdate_ = datetime.date(int(str(startDate)[0:4]), int(str(startDate)[4:6].lstrip('0')), int(str(startDate)[6:8].lstrip('0')))
				enddate_ = datetime.date(int(str(endDate)[0:4]), int(str(endDate)[4:6].lstrip('0')), int(str(endDate)[6:8].lstrip('0')))
				# duration = int(str(enddate_ - startdate_).split(" ")[0]) + 1

				date_diff = enddate_ - startdate_
				if startDate == endDate:
					date_diff = 0
				else:
					date_diff = str(enddate_ - startdate_).split(" ")[0]

				duration = int(date_diff) + 1

				for z in range(0, duration):
					dates = startdate_ + datetime.timedelta(z)
					datesFormat = int(str(dates).split('-')[0] + str(dates).split('-')[1] + str(dates).split('-')[2])
					# date count : hk.kim-18.01.29
					count = 0
					for count_info in bbs_counts:
						if count_info.date == str(datesFormat):
							count = count_info.num
					subjectresultForExcel.append(str(teachers.get_subject()) + ':' + str(teachers.get_name()) + ':' + str(datesFormat) + ':' + str(count))

					#if count == 0:
					#	subjectresultForExcel.append(str(teachers.get_subject()) + ':' + str(teachers.get_name()) + ':' + str(datesFormat) + ':' + str(0))
					#else:
					#	count_ = date_countgroup[str(datesFormat)]
					#	subjectresultForExcel.append(str(teachers.get_subject()) + ':' + str(teachers.get_name()) + ':' + str(datesFormat) + ':' + count_)

				driver.implicitly_wait(delayTime)
				labelstatus.setText('EBS ' + str(teachers.get_name()) + ' 선생님 집계 완료')
				excelFile = workBook(filepath + 'EBS_개별_' + str(startDate) + '-' + str(endDate) + '.xlsx')
				Ebs.xlsxWrite(excelFile, subjectresultForExcel)
				excelFile.close()
			labelstatus.setText('EBS 개별 집계 종료 : 엑셀파일을 확인해주세요')
			#엑셀 출력 코드 추가 필요

	def analyzeDS(self, driver):
		def workBook(filename):
			workbook = xlsxwriter.Workbook(filename)  # 'math.xlsx'
			return workbook
		if parsingMode == 0:
			labelstatus.setText('대성마이맥 집계를 시작합니다.')
			subjectresultForExcel = []
			selectedPersonNum = 0
			progress = 0
			for u in range(0, len(DSsubjectObject)):
				# 중복 제거
				list_dic = {}
				for t in range(0, len(DSsubjectObject[u].getIDList())):
					list_dic[DSsubjectObject[u].getIDList()[t]] = ""

				selectedPersonNum = len(list_dic) + selectedPersonNum
				# selectedPersonNum = len(DSsubjectObject[u].getIDList()) + selectedPersonNum
			for x in range(0, len(DSsubjectObject)):
				IDList = DSsubjectObject[x].getIDList()
				NameList = DSsubjectObject[x].getNameList()
				SubjectList = DSsubjectObject[x].getSubjectList()

				i_am_already_processed = []
				
				for i in range(0, len(IDList)):
					# 이미 집계한 적이 있는 ID면 continue, 없으면 pass해서 계속 처리 
					try:
						index = i_am_already_processed.index(IDList[i])
						continue
					except ValueError:
						pass
					labelstatus.setText('대성마이맥 ' + NameList[i] + ' 선생님 집계중  ' + str(progress + i + 1) + '/' + str(selectedPersonNum))
					BoardAddr = DSsubjectObject[x].getIndivBoardAddress(IDList[i])
					calcBoards = CalcDaesung(BoardAddr, endDate, startDate, delayTime, driver)  # url, startdate, enddate, waitTime, chromedriver
					semiresult = calcBoards.calcBoard(check_stop_class, labelstatus2) #processing pause : hk.kim-18.01.28
					result = calcBoards.dataResult(semiresult, NameList[i], SubjectList[0])  # calcBoardResult, teacherName, subjectName
					for j in range(0, len(result)):
						subjectresultForExcel.append(result[j])
					labelstatus.setText('메가스터디 ' + NameList[i] + ' 선생님 집계 완료')
					labelstatus2.setText('')
					excelFile = workBook(filepath + '대성마이맥_' + str(startDate) + '-' + str(endDate) + '.xlsx')
					calcBoards.xlsxWrite(excelFile, subjectresultForExcel)
					excelFile.close()
					i_am_already_processed.append(IDList[i])
				progress = progress + len(IDList)
			labelstatus.setText('대성마이맥 집계 종료 : 엑셀파일을 확인해주세요')

		elif parsingMode == 1:
			labelstatus.setText('대성마이맥 개별 집계를 시작합니다.')
			DSdic = self.IdNameDicDS()
			IDList = []
			NameList = []
			SubjectList = []
			print(DSdic)
			for i in range(0, len(selectedParseList3)):
				split = selectedParseList3[i].split(':')
				print(selectedParseList3[i])
				IDList.append(DSdic[selectedParseList3[i]])
				NameList.append(selectedParseList3[i].split(':')[1].lstrip())
				SubjectList.append(selectedParseList3[i].split(':')[0])
			subjectresultForExcel = []
			for i in range(0, len(IDList)):
				labelstatus.setText('대성마이맥 ' + NameList[i] + ' 선생님 집계중')
				BoardAddr = dskorean.getIndivBoardAddress(IDList[i])
				calcBoards = CalcDaesung(BoardAddr, endDate, startDate, delayTime, driver)  # url, startdate, enddate, waitTime, chromedriver
				semiresult = calcBoards.calcBoard(check_stop_class, labelstatus2) #processing pause : hk.kim-18.01.28
				result = calcBoards.dataResult(semiresult, NameList[i], SubjectList[i])  # calcBoardResult, teacherName, subjectName
				for j in range(0, len(result)):
					subjectresultForExcel.append(result[j])
				labelstatus.setText('메가스터디 ' + NameList[i] + ' 선생님 집계 완료')
				labelstatus2.setText('')
				excelFile = workBook(filepath + '대성마이맥_개별_' + str(startDate) + '-' + str(endDate) + '.xlsx')
				calcBoards.xlsxWrite(excelFile, subjectresultForExcel)
				excelFile.close()
			labelstatus.setText('대성마이맥 개별 집계 종료 : 엑셀파일을 확인해주세요')

	def analyzeSKY(self, driver):
		def workBook(filename):
			workbook = xlsxwriter.Workbook(filename)  # 'math.xlsx'
			return workbook
		if parsingMode == 0:
			labelstatus.setText('스카이에듀 집계를 시작합니다.')
			subjectresultForExcel = []
			selectedPersonNum = 0
			progress = 0
			print("SKYsubjectObject",SKYsubjectObject)
			for u in range(0, len(SKYsubjectObject)):
				selectedPersonNum = len(SKYsubjectObject[u].getIDList()) + selectedPersonNum
			for subject in SKYsubjectObject:  # 과목 루프
				IDList = subject.getIDList()
				NameList = subject.getNameList()
				SubjectList = subject.getSubjectList()
				QAAddressList = subject.getBoardAddressList(IDList)
				if len(IDList) != len(NameList) and len(IDList) != len(QAAddressList):
					labelstatus.setText('선생님 정보 개편됨 : 개발사에 문의해주세요')
					break
				for i, address in enumerate(QAAddressList):  # 선생님 루프
					if address != None:
						labelstatus.setText('스카이에듀 ' + NameList[i] + ' 선생님 집계중  ' + str(progress + i + 1) + '/' + str(selectedPersonNum))
						calcBoards = CalcSkyedu(address, endDate, startDate, delayTime, driver)  # url, startdate, enddate, waitTime, chromedriver
						semiresult = calcBoards.calcBoard(check_stop_class, labelstatus2) #processing pause : hk.kim-18.01.28
						result = calcBoards.dataResult(semiresult, NameList[i], SubjectList[0])  # calcBoardResult, teacherName, subjectName
						for parseddata in result:
							subjectresultForExcel.append(parseddata)
						labelstatus.setText('스카이에듀 ' + NameList[i] + ' 선생님 집계 완료')
						labelstatus2.setText('')
						excelFile = workBook(filepath + '스카이에듀_' + str(startDate) + '-' + str(endDate) + '.xlsx')
						calcBoards.xlsxWrite(excelFile, subjectresultForExcel)
						excelFile.close()
					elif address == None:
						pass
				progress = progress + len(IDList)
			labelstatus.setText('스카이에듀 집계 종료 : 엑셀파일을 확인해주세요')

		elif parsingMode == 1:
			labelstatus.setText('스카이에듀 개별 집계를 시작합니다.')
			SKYdic = self.IdNameDicSKY()
			print("SKYdic", SKYdic)
			IDList = []
			NameList = []
			SubjectList = []
			for i in range(0, len(selectedParseList4)):
				# split = selectedParseList4[i].split(':')  # 국어: 김선겸
				# IDList.append(SKYdic[split[1].lstrip()])
				IDList.append(SKYdic[selectedParseList4[i].lstrip()])
				NameList.append(selectedParseList4[i].split(':')[1].lstrip())
				SubjectList.append(selectedParseList4[i].split(':')[0])
			subjectresultForExcel = []
			for i in range(0, len(IDList)):
				labelstatus.setText('스카이에듀 ' + NameList[i] + ' 선생님 집계중')
				if IDList[i]!=None:
					BoardAddr = skykorean.getIndivBoardAddress(IDList[i])
					calcBoards = CalcSkyedu(BoardAddr, endDate, startDate, delayTime, driver)  # url, startdate, enddate, waitTime, chromedriver
					semiresult = calcBoards.calcBoard(check_stop_class, labelstatus2) #processing pause : hk.kim-18.01.28
					result = calcBoards.dataResult(semiresult, NameList[i], SubjectList[i])  # calcBoardResult, teacherName, subjectName
					for j in range(0, len(result)):
						subjectresultForExcel.append(result[j])
					labelstatus.setText('스카이에듀 ' + NameList[i] + ' 선생님 집계 완료')
					labelstatus2.setText('')
					excelFile = workBook(filepath + '스카이에듀_개별_' + str(startDate) + '-' + str(endDate) + '.xlsx')
					calcBoards.xlsxWrite(excelFile, subjectresultForExcel)
					excelFile.close()
				else:
					pass
			labelstatus.setText('스카이에듀 개별 집계 종료 : 엑셀파일을 확인해주세요')

	def IdNameDicMEGA(self):  # 이름 : ID 딕셔너리
		dicIdName = {}
		MegasubjectObject = [korean, math, english, korhistory, society, science, univ, foreign]
		for i in range(0, len(MegasubjectObject)):
			tIdList = MegasubjectObject[i].getIDList()
			tNameList = MegasubjectObject[i].getNameList()
			#print(len(tIdList), len(tNameList))
			for j in range(0, len(tIdList)):
				if len(tNameList[j]) > 0:
					dicIdName[tNameList[j]] = tIdList[j]
		#print(dicIdName)
		return dicIdName

	def IdNameDicDS(self):
		dicIdName = {}
		DSsubjectObject = [dskorean, dsmath, dsenglish, dskorhistory, dssociety, dsscience, dsuniv, dsforeign]
		for i in range(0, len(DSsubjectObject)):
			tIdList = DSsubjectObject[i].getIDList()
			tNameList = DSsubjectObject[i].getNameList()
			tSubjectList = DSsubjectObject[i].getSubjectList()
			for j in range(0, len(tIdList)):
				dicIdName[tSubjectList[j] + ': ' + tNameList[j]] = tIdList[j]
		#print(dicIdName)
		return dicIdName

	def IdNameDicSKY(self):
		dicIdName = {}
		# SKYsubjectObject = [skykorean, skymath, skyenglish, skykorhistory, skysociety, skyscience, skyuniv, skyforeign, skydaechi, skynaesin]
		# SKYsubjectObject = [skykorean, skymath, skyenglish, skykorhistory, skysociety, skyscience, skyuniv, skyforeign, skynaesin]
		SKYsubjectObject = [skykorean, skymath, skyenglish, skykorhistory, skysociety, skyscience, skyuniv, skyforeign]
		for i in range(0, len(SKYsubjectObject)):
			# tIdList = SKYsubjectObject[i].getIDList()
			# tNameList = SKYsubjectObject[i].getNameList()
			# print(len(tIdList), tIdList)
			# print(len(tNameList), tNameList)
			# for j in range(0, len(tIdList)):
			# 	dicIdName[tNameList[j]] = tIdList[j]
			tIdList = SKYsubjectObject[i].getIDList()
			tNameList = SKYsubjectObject[i].getFullList()
			print(len(tIdList), tIdList)
			print(len(tNameList), tNameList)
			for j in range(0, len(tIdList)):
				dicIdName[tNameList[j]] = tIdList[j]
		print("dicIdName",dicIdName)
		return dicIdName

#processing pause : hk.kim-18.01.28
class CheckPauseClass():
	def get_is_pause(self):
		global is_pause
		return is_pause

if __name__ == "__main__":
	# global mythread
	# global check_stop_class
	app = QApplication(sys.argv)
	mythread = DataAnalyze()
	check_stop_class = CheckPauseClass() #processing pause : hk.kim-18.01.28
	myWindow = MyWindow()
	myWindow.show()
	app.exec_()
