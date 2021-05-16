#-*-coding: utf-8
import os
import re
import time
import xlsxwriter
import datetime
from collections import OrderedDict
from urllib.request import urlopen
from urllib.parse import urlparse, parse_qs, parse_qsl
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

print('시스템 준비중..')

# parts = urlparse("/teacher_v2/main.asp?tec_cd=megakdw&amp;dom_cd=1&amp;HomeCd=148")
# print(parse_qs(parts.query))

class GoMegastudy:
	def __init__(self, driver):
		url = "http://www.megastudy.net/teacher_v2/teacher_main.asp"

		process = False
		while process == False:
			try:
				driver.get(url)
				# WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="container"]/div[1]/div[1]/ul/li[1]/ul[1]/li[1]/a')))
				WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="teachers"]')))
				teacher_list_open_button =  driver.find_element_by_xpath('//*[@id="megaGnb"]/div[1]/div[1]/span[1]/a[2]')
				teacher_list_open_button.click()
				WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="teachers"]/div/ul')))
				process = True
			except:
				print("pass!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
				# pass
		source = driver.page_source
		self.soup = BeautifulSoup(source, "html.parser")
		driver.close()

	def tpage(self):
		return self.soup


class SetMegastudy:  # 초반 선생정보 가져오는 Setting 클래스, CalcMegastudy에 앞서 실행되어야함.
	def __init__(self, department, tpage):
		super().__init__()
		# 과목(Department분류) divMenuTecList1~8 국어/수학/영어/한국사/사회/과학/대학별고사/제2외국어.한문
		self.department = department
		self.subj_wrap = tpage.find('ul', class_ = "subj_wrap")
		self.subj = self.subj_wrap.select('li.subj')[int(self.department)]
		self.subject_teachers_ATAG = []
		self.subject_teachers = self.subj.find_all('ul')
		for ul in self.subject_teachers:
			self.subject_teachers_ATAG = self.subject_teachers_ATAG + ul.find_all('a')
		self.subject_name_ATAG = self.subj.find_all("h3")[0].find("a")
		if len(self.subject_teachers) == 0:
			return False

	def getFullList(self):  # '과목선생님 : 이름  개인주소 : http://~~~~~' 요소 형식으로 리스트 리턴
		result = []
		subjectName = self.subject_name_ATAG.getText()
		for teacher_ATAG in self.subject_teachers_ATAG:
			teacherName = teacher_ATAG.getText()
			if subjectName == "과학":
				subjectName = "과학탐구"
			elif subjectName == "사회":
				subjectName = "사회탐구"

			url = teacher_ATAG.get("href")
			url_parsed = parse_qs(urlparse(url).query)
			
			if self.checkIfIsAvailTeacher(url_parsed) == False:
				continue

			resultdata = subjectName + ': ' + teacherName
			result.append(resultdata)

		resultList = list(OrderedDict.fromkeys(result))
		return resultList
		

	def getSubjectList(self):  # '과목선생님 : 이름  개인주소 : http://~~~~~' 요소 형식으로 리스트 리턴
		result = []
		# for i in range(0, len(self.subject_teachers)):
		# 	attr_onmousedown = self.subject_teachers[i].get('onmousedown')
		# 	attr_href = str(self.subject_teachers[i].get('href'))
		# 	if attr_onmousedown and len(attr_href) > 10:
		# 		datasplit = re.compile('[^ ㄱ-ㅣ가-힣ㅣ2|/]+')
		# 		replaced = datasplit.sub('(', attr_onmousedown).split('(')
		# 		split = replaced[3].split('/')
		# 		tSubject = split[1]
		# 		if tSubject == "과학":
		# 			tSubject = "과학탐구"
		# 		elif tSubject == "사회":
		# 			tSubject = "사회탐구"
		# 		result.append(tSubject)
		for teacherATAG in self.subject_teachers_ATAG:
			url = teacherATAG.get("href")
			url_parsed = parse_qs(urlparse(url).query)
			if self.checkIfIsAvailTeacher(url_parsed) == False:
				continue
			subjectName = self.subject_name_ATAG.getText()
			if subjectName == "과학":
				subjectName = "과학탐구"
			elif subjectName == "사회":
				subjectName = "사회탐구"
			result.append(subjectName)
		resultList = list(OrderedDict.fromkeys(result))
		return resultList

	def getIDList(self):
		result = []
		for teacherATAG in self.subject_teachers_ATAG:
			url = teacherATAG.get("href")
			url_parsed = parse_qs(urlparse(url).query)
			if self.checkIfIsAvailTeacher(url_parsed) == False:
				continue

			result.append(url_parsed["tec_cd"][0])
		print(result)
		resultList = list(OrderedDict.fromkeys(result))
		return resultList

	def getNameList(self):
		result = []
		# for i in range(0, len(self.subject_teachers)):
		# 	attr_onmousedown = self.subject_teachers[i].get('onmousedown')
		# 	attr_href = str(self.subject_teachers[i].get('href'))
		# 	if attr_onmousedown and len(attr_href) > 10 and len(attr_href.split('=')) == 4:
		# 		datasplit = re.compile('[^ ㄱ-ㅣ가-힣ㅣ2|/]+')
		# 		replaced = datasplit.sub('(', attr_onmousedown).split('(')
		# 		split = replaced[3].split('/')
		# 		tName = split[2]
		# 		result.append(tName)
		for teacherATAG in self.subject_teachers_ATAG:
			url = teacherATAG.get("href")
			url_parsed = parse_qs(urlparse(url).query)
			if self.checkIfIsAvailTeacher(url_parsed) == False:
				continue
			teacherName = teacherATAG.getText()
			result.append(teacherName)
		resultList = list(OrderedDict.fromkeys(result))
		return resultList

	def getAddressList(self):
		result = []
		# for i in range(0, len(self.subject_teachers)):
		# 	attr_onmousedown = self.subject_teachers[i].get('onmousedown')
		# 	attr_href = str(self.subject_teachers[i].get('href'))
		# 	if attr_onmousedown and len(attr_href) > 10 and len(attr_href.split('=')) == 4:
		# 		result.append(attr_href)
		for teacherATAG in self.subject_teachers_ATAG:
			url = teacherATAG.get("href")
			url_parsed = parse_qs(urlparse(url).query)
			if self.checkIfIsAvailTeacher(url_parsed) == False:
				continue
			result.append(url)
		resultList = list(OrderedDict.fromkeys(result))
		return resultList

	def getBoardAddressList(self, idArray):
		result = []
		for i in range(0, len(idArray)):
			boardaddressbeforeID = 'http://www.megastudy.net/teacher_v2/bbs/bbs_list.asp?tec_cd='
			boardaddressafterID = '&LeftMenuCd=3&brd_kbn=qnabbs&LeftSubCd=1'
			boardaddress = boardaddressbeforeID + idArray[i] + boardaddressafterID
			result.append(boardaddress)
		return result

	def getIndivBoardAddress(self, id):
			boardaddressbeforeID = 'http://www.megastudy.net/teacher_v2/bbs/bbs_list.asp?tec_cd='
			boardaddressafterID = '&LeftMenuCd=3&brd_kbn=qnabbs&LeftSubCd=1#'
			boardaddress = boardaddressbeforeID + id + boardaddressafterID
			# print(boardaddress)
			return boardaddress
	
	def checkIfIsAvailTeacher(self, url_parsed):
		tec_cd = None
		dom_cd = None
		HomeCd = None
		if "tec_cd" in url_parsed:
			tec_cd = url_parsed["tec_cd"]
		if "dom_cd" in url_parsed:
			dom_cd = url_parsed["dom_cd"]
		if "HomeCd" in url_parsed:
			HomeCd = url_parsed["HomeCd"]
		
		if tec_cd == None or dom_cd == None or HomeCd == None:
			return False
		else:
			return True


class CalcMegastudy:  # 선생,과목, 게시판주소등을 토대로 본격적으로 긁어오고 결과를 엑셀로 출력하는 것까지 관련된 클래스
	def __init__(self, id, startdate, enddate, waitTime, chromedriver):
		self.boardaddressbeforeID = 'http://www.megastudy.net/teacher_v2/bbs/bbs_list.asp?tec_cd='
		self.boardaddressafterID = '&LeftMenuCd=3&brd_kbn=qnabbs&LeftSubCd=1#'
		self.id = id
		self.url = self.boardaddressbeforeID + self.id + self.boardaddressafterID
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
		endpage = 99999
		# url = self.url  # getIndivBoardAddress에 선생님List의 Index를 넣을것
		# waitTime = [0.2, 0.7, 1, 1.5]  # 페이지 하나하나 넘기는 시간 랜덤으로..(너무 빠르게 파싱하는걸 금지)
		# ##############################게시판 Searching 옵션
		fncListMove = ""
		connected = 0
		while connected == 0:
			try:
				print('try1')
				self.driver.get(self.url)
				print('driver.get(url)')
				WebDriverWait(self.driver, 40).until(EC.presence_of_element_located((By.ID, 'paging_wrap')))
				print('WebdriverWait')
				time.sleep(self.waitTime)
				print('timesleep')
				connected = 1
				print('connected')

				source_ = self.driver.page_source
				soup_ = BeautifulSoup(source_, "html.parser")
				pagerAtag = soup_.select('#paging_wrap > a')
				print("pagerAtag", pagerAtag)
				if len(pagerAtag) == 2:
					fncListMove = None
				else:
					loc_ = str(pagerAtag[1].get('href')).find('page=')
					print("loc_", loc_)
					print("pagers_",str(pagerAtag[1].get('onclick')))
					pagers_ = str(pagerAtag[1].get('href'))
					fncListMove = str(pagers_[:loc_ + 5]) + "%s')"

			except TimeoutException:
				print('except1')
				#self.driver.get(self.url)
				print('서버와 통신이 불안정 합니다. 재접속을 시도합니다.')
				labelstatus.setText('서버와 통신이 불안정 합니다. 재접속을 시도합니다.')
		
		print('Page_' + str(1) + ' --> Searching...')
		for i in range(startpage, endpage):
			pageconnected = 0
			while pageconnected == 0:
				try:
					WebDriverWait(self.driver, 60).until(EC.presence_of_element_located((By.ID, 'paging_wrap')))
					print('페이지 가져오는 중')
					labelstatus.setText('페이지 가져오는 중')
					pageconnected = 1
					time.sleep(self.waitTime)
				except TimeoutException:
					self.driver.execute_script('fncQnAList()')
					self.driver.execute_script(fncListMove % i )
					print('서버와 통신이 불안정 합니다. 재시도 합니다. Inner')
					labelstatus.setText('서버와 통신이 불안정 합니다. 재접속을 시도합니다..')

			source = self.driver.page_source
			page = BeautifulSoup(source, "html.parser")  # beautiful soup으로 html형식으로 파싱
			table = page.find(class_="commonBoardList")
			tdinboardTable = table.find_all('td', class_="number")
			for j in range(0, len(tdinboardTable)):
				tdtext = tdinboardTable[j].getText()
				tdsplit = tdtext.split('-')
				if len(tdsplit) == 3:
					date = int(tdsplit[0] + tdsplit[1] + tdsplit[2])  # 20180117
					if date >= self.enddate and date <= self.startdate:
						self.questionCount.append(date)
					elif date < self.enddate:
						print('Searching Finished - From ' + str(self.startdate) + ' to ' + str(self.enddate) + '************************')
						return self.questionCount
						# self.questionCount.append(self.questionCountIndiv)
						# return self.questionCount  # 게시글을 최신날짜 --> 과거날짜로 서칭하면서 기준과거 조건보다 더 과거가 나오면 게시글 카운팅을 종료한다.
				else:
					pass
			
			'''
			if i <= 10:  # 페이지 클릭 넘버링 처음한번만 1-10 그 이후  2-11, 2-11...로 반복해주기위한 구문
				pagenumber = i
			else:
				if i >= 21 and i % 10 == 1:
					A = A + 1
				pagenumber = (i - 10 * A - 9)
			self.driver.find_element_by_xpath('//*[@id="iNavPaging"]/a[%s]' % pagenumber).click()
			'''

			source = self.driver.page_source
			soup = BeautifulSoup(source, "html.parser")
			pager = soup.select('#paging_wrap > a')
			if len(pager) == 2:
				return self.questionCount
			else:
				loc = str(pager[1].get('href')).find('page=')
				pagers = str(pager[1].get('href'))
				fncListMove = str(pagers[:loc + 5]) + "%s')"
				javascriptclick = str(pagers[:loc + 5]) + "%s')"
				print("프린트!!" + javascriptclick % (i + 1))
				self.driver.execute_script(javascriptclick % (i + 1))

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
			# print(self.subject + ' ', self.name + ' 선생님 :', self.questionCountNum[z])
		return resultForReturn


	def xlsxWrite(self, workBook, finalresultForExcel):
		self.workbook = workBook
		worksheet = self.workbook.add_worksheet('메가스터디')
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
			self.count = self.split[3]
			self.dateformat = self.date[0:4] + '-' + self.date[4:6] + '-' + self.date[6:]
			worksheet.write(y + 1, 0, self.dateformat)
			worksheet.write(y + 1, 1, '메가스터디')
			worksheet.write(y + 1, 2, self.subject)
			worksheet.write(y + 1, 3, self.teacher)
			worksheet.write(y + 1, 4, self.count)