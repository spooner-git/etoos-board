#-*-coding: utf-8
import os
import re
import time
import xlsxwriter
import datetime
from collections import OrderedDict
from urllib.request import urlopen
from bs4 import BeautifulSoup
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

print('시스템 준비중...')

class GoDaesung:
	def __init__(self, driver):
		url = "http://www.mimacstudy.com/tcher/home/tcherHomeMain.ds?requestMenuId=MNMN_M004"

		process = False
		while process == False:
			try:
				driver.get(url)
				WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, 'ulM000000186')))
				process = True
			except:
				print("pass!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
				# pass
		source = driver.page_source
		self.soup = BeautifulSoup(source, "html.parser")
		driver.close()

	def tpage(self):
		return self.soup


class SetDaesung:  # 초반 선생정보 가져오는 Setting 클래스
	def __init__(self, department, tpage):
		self.department = department
		self.subject_teacherList = tpage.find('li', id=self.department)  # 태그 ID를 찾아간다.
		self.subject_teachers = self.subject_teacherList.find_all('a') #
		if len(self.subject_teachers) == 0:
			return False

	def getFullList(self):  # '과목선생님 : 이름  개인주소 : http://~~~~~' 요소 형식으로 리스트 리턴
		result = []
		for i in range(1, len(self.subject_teachers)):
			subjectName = self.subject_teachers[0].getText()
			tName = self.subject_teachers[i].getText()
			tAddress = self.subject_teachers[i].get('href')
			# result.append(subjectName + '^' + tName + '^' + tAddress)
			'''
			if subjectName == '사회탐구':
				subjectName = '사회'
			elif subjectName == '과학탐구':
				subjectName = '과학'
			'''
			if subjectName + ': ' + tName not in result:
				result.append(subjectName + ': ' + tName)
		return result

	def getSubjectList(self):
		result = []
		for i in range(1, len(self.subject_teachers)):
			subjectName = self.subject_teachers[0].getText()
			'''
			if subjectName == '사회탐구':
				subjectName = '사회'
			elif subjectName == '과학탐구':
				subjectName = '과학'
			'''
			result.append(subjectName)
		return result

	def getIDList(self):
		result = []
		for i in range(1, len(self.subject_teachers)):
			splitAddress = self.subject_teachers[i].get('href').split('&')
			splitAddress2 = splitAddress[0].split('?')
			teacherID = splitAddress2[1]
			result.append(teacherID)
		return result

	def getNameList(self):
		result = []
		for i in range(1, len(self.subject_teachers)):
			tName = self.subject_teachers[i].getText()
			result.append(tName)
		return result

	def getAddressList(self):
		result = []
		for i in range(1, len(self.subject_teachers)):
			tAddress = self.subject_teachers[i].get('href')
			result.append(tAddress)
		return result

	def getBoardAddressList(self, idArray):
		result = []
		for i in range(0, len(idArray)):
			boardaddressbeforeID = 'http://www.mimacstudy.com/tcher/studyQna/getStudyQnaList.ds?'
			boardaddress = boardaddressbeforeID + idArray[i]
			result.append(boardaddress)
		return result

	def getIndivBoardAddress(self, id):
			boardaddressbeforeID = 'http://www.mimacstudy.com/tcher/studyQna/getStudyQnaList.ds?'
			boardaddress = boardaddressbeforeID + id
			return boardaddress


class CalcDaesung:  # 선생,과목, 게시판주소등을 토대로 본격적으로 긁어오고 결과를 엑셀로 출력하는 것까지 관련된 클래스
	def __init__(self, url, startdate, enddate, waitTime, chromedriver):
		self.url = url
		self.startdate = startdate
		self.enddate = enddate
		self.waitTime = waitTime
		self.driver = chromedriver

	def calcBoard(self, check_stop_class, labelstatus): #processing pause : hk.kim-18.01.28
		print('Parsing Start************************')
		self.questionCount = []  # 게시판의 모든 날짜가 여기 담은 후, 중복 갯수를 세어서 날짜별 게시글 수를 센다.
		A = 0
		# #############################게시판 Searching 옵션
		startpage = 1
		endpage = 999
		url = self.url  # getIndivBoardAddress에 선생님List의 Index를 넣을것
		# waitTime = [0.2, 0.7, 1, 1.5]  # 페이지 하나하나 넘기는 시간 랜덤으로..(너무 빠르게 파싱하는걸 금지)
		# ##############################게시판 Searching 옵션
		connected = 0
		#webdriverWaitFor = 'pagnation'
		webdriverWaitFor = 'tbltype_list'
		while connected == 0:
			try:
					self.driver.get(self.url)
					WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, webdriverWaitFor)))
					time.sleep(self.waitTime)
					connected = 1
			except:
					print('서버와 통신이 불안정 합니다. 재접속을 시도합니다.')
					labelstatus.setText('서버와 통신이 불안정 합니다. 재접속을 시도합니다.')
					time.sleep(3)
					self.driver.refresh()
		print('Page_' + str(1) + ' --> Searching...')
		for i in range(startpage, endpage):
			pageconnected = 0
			# if i > 1:
			# 	webdriverWaitFor_page = 'pagnation'
			# else:
			# 	webdriverWaitFor_page = 'tbltype_list'
			webdriverWaitFor_page = 'tbltype_list'
			while pageconnected == 0:
				try:
					#self.driver.get(self.url)
					WebDriverWait(self.driver, 3).until(EC.presence_of_element_located((By.CLASS_NAME, webdriverWaitFor_page)))
					pageconnected = 1
					time.sleep(self.waitTime)
				except TimeoutException:
					# self.driver.execute_script("document.getElementById('currPage').value = '%s'; document.getElementById('srchFrm').submit();" % (i + 1))
					self.driver.refresh()
					print('서버와 통신이 불안정 합니다. 재시도 합니다. Inner')
					labelstatus.setText('서버와 통신이 불안정 합니다. 재접속을 시도합니다.')
					#self.driver.get(self.url)
					#WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.ID, 'iNavPaging')))
					# self.questionCount.append(self.questionCountIndiv)
					#print('서버 재연결 시도완료 Inner')
			
			source = self.driver.page_source
			page = BeautifulSoup(source, "html.parser")  # beautiful soup으로 html형식으로 파싱
			table = page.find(class_="tbltype_list")
			trs = table.find_all('tr', class_ = '')
			trFiltered = []
			tdinboardTable = []
			for u in range(0, len(trs)):
				if "공지" not in str(trs[u]):
					trFiltered.append(trs[u])
			for t in range(0, len(trFiltered)):
				td = trFiltered[t].select('td:nth-of-type(5)')
				if len(td) > 0:
					tdinboardTable.append(td[0].getText())
			for j in range(0, len(tdinboardTable)):
				tdtext = tdinboardTable[j]
				tdsplit = tdtext.split('/')
				if len(tdsplit) == 3:
					date = int(tdsplit[0] + tdsplit[1] + tdsplit[2])  # 20180117
					if date >= self.enddate and date <= self.startdate:
						self.questionCount.append(date)
					elif date < self.enddate:
						print("여기11")
						# print('Searching Finished - From ' + str(self.startdate) + ' to ' + str(self.enddate) + '************************')
						return self.questionCount  # 게시글을 최신날짜 --> 과거날짜로 서칭하면서 기준과거 조건보다 더 과거가 나오면 게시글 카운팅을 종료한다.
				else:
					pass
			'''
			page = (i % 10) +1
			pagination = '//*[@id="srchFrm"]/div/div[6]/span/a[%s]' % page
			if i % 10 == 0 :
				if i == 10 :
					pagination = '//*[@id="srchFrm"]/div/div[6]/button[1]'
				else:
					pagination = '//*[@id="srchFrm"]/div/div[6]/button[3]'
			'''
			#self.driver.find_element_by_xpath(pagination).click()

			if len(page.select('#srchFrm > div > div.pagnation')) > 0:
				pagenum = len(page.select('#srchFrm > div > div.pagnation > span')[0].find_all('a'))
			else:
				pagenum = 0

			print("pagenum", pagenum)
			if len(tdinboardTable)>0:
				if(pagenum > 1):
					self.driver.execute_script("document.getElementById('currPage').value = '%s'; document.getElementById('srchFrm').submit();" % (i + 1))
				else:
					print("여기33")
					return self.questionCount
			else:
				print(tdinboardTable)
				print("여기22")
				return self.questionCount
			#self.driver.find_element_by_xpath(pagination).send_keys(webdriver.common.keys.Keys.SPACE)
			#processing pause : hk.kim-18.01.28
			while True:
				is_pause = check_stop_class.get_is_pause()
				if is_pause == 0:
					break
			print('Page_' + str(i + 1) + ' --> Searching...')
			labelstatus.setText('Page_' + str(i + 1) + ' --> Searching...')


	def dataResult(self, calcBoardResult, teacherName, subjectName):
		questionCountNum = []
		startDate = datetime.date(int(str(self.startdate)[0:4]), int(str(self.startdate)[4:6].lstrip('0')), int(str(self.startdate)[6:].lstrip('0')))
		endDate = datetime.date(int(str(self.enddate)[0:4]), int(str(self.enddate)[4:6].lstrip('0')), int(str(self.enddate)[6:].lstrip('0')))
		date_diff = startDate - endDate
		if startDate == endDate:
			date_diff = 0
		else:
			date_diff = str(startDate - endDate).split(" ")[0]

		duration = int(date_diff) + 1
		
		for z in range(0, duration):
			dates = endDate + datetime.timedelta(z)
			datesFormat = int(str(dates).split('-')[0] + str(dates).split('-')[1] + str(dates).split('-')[2])
			count = calcBoardResult.count(datesFormat)
			questionCountNum.append(str(datesFormat) + ':' + str(count))
		resultForReturn = []
		for z in range(0, len(questionCountNum)):
			dataresultString = subjectName + ':' + teacherName + ':' + questionCountNum[z]
			resultForReturn.append(dataresultString)
		return resultForReturn

	def xlsxWrite(self, workBook, finalresultForExcel):
		self.workbook = workBook
		worksheet = self.workbook.add_worksheet('대성마이맥')
		format = self.workbook.add_format()
		format.set_bg_color('#FF6600')
		worksheet.set_column(0, 4, 12)
		worksheet.write(0, 0, '날짜', format)
		worksheet.write(0, 1, '사이트', format)
		worksheet.write(0, 2, '과목', format)
		worksheet.write(0, 3, '선생님', format)
		worksheet.write(0, 4, '게시물수', format)
		self.finalresultForExcel = finalresultForExcel
		for y in range(0, len(self.finalresultForExcel)):
			self.split = self.finalresultForExcel[y].split(':')
			self.subject = self.split[0]
			self.teacher = self.split[1]
			self.date = self.split[2]
			self.dateformat = self.date[0:4] + '-' + self.date[4:6] + '-' + self.date[6:]
			self.count = self.split[3]
			worksheet.write(y + 1, 0, self.dateformat)
			worksheet.write(y + 1, 1, '대성마이맥')
			worksheet.write(y + 1, 2, self.subject)
			worksheet.write(y + 1, 3, self.teacher)
			worksheet.write(y + 1, 4, self.count)