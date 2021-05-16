#-*-coding: utf-8
import os
import re
import time
import xlsxwriter
import datetime
from collections import OrderedDict
from urllib.request import urlopen
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

print('시스템 준비중.....')

class GoSkyedu:
	def __init__(self, driver):
		url = "https://skyedu.conects.com/teachers/"

		process = False
		tries = 0;
		while process == False:
			try:
				print(driver)
				driver.get(url)
				WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="hgroup"]/div[2]/div/ul/li[1]/div/div/div/dl[1]/dd[1]/ul/li[1]/a')))
				process = True
			except:
				tries = tries + 1
				print("pass!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
				print(tries)
				# pass
		source = driver.page_source
		self.soup = BeautifulSoup(source, "html.parser")
		driver.close()

	def tpage(self):
		try:
			# self.teacher_list_div = self.soup.select('#hgroup > div.st-conects-snb > div > ul > li.on > div > div > div')
			self.teacher_list_div = self.soup.select('#hgroup > div.st-conects-snb > div > ul > li:nth-of-type(1) > div > div > div')
			self.dtdda = self.teacher_list_div[0].find_all(['dt','a'])

		except:
			# self.labelstatus.setText("스카이에듀 선생님 페이지 로딩 실패")
			self.teacherDict = {}
			return self.teacherDict
			
		self.dtfinder = []
		self.hangul = re.compile('[^ ㄱ-ㅣ가-힣ㅣ0-9]+')

		for i in self.dtdda:
			if i.has_attr('href') == True:
				self.href = i.get("href")
				self.tName = self.hangul.sub("", i.getText()).strip()
				self.tID = self.href.split("?")[1]
				if "nondangi" in  self.href:
					self.dtfinder.append('#:'+self.tName + ':' + self.tID)
				else:
					self.dtfinder.append('@:'+self.tName + ':' + self.tID)
			else:
				self.dtfinder.append(i.getText())
		self.teacherDict = {}
		self.subjectlist = []
		self.teacherlist = []
		self.index = -1
		for j in self.dtfinder:
			if j[0] != '@' and j[0] != '#':
				self.teacherDict[j] = []
				self.subjectlist.append(j)
				self.index = self.index + 1
			else:
				self.teacherlist.append(j + '--' + str(self.index))
				self.teacherDict[self.subjectlist[self.index]].append(j)

		return self.teacherDict


class SetSkyedu:  # 초반 선생정보 가져오는 Setting 클래스
	def __init__(self, department, tpage):
		self.department = department
		self.teacherDict = tpage
		self.subject_teachers = self.teacherDict[department]
		self.nondangi_teachers = []
		for i in range(0, len(self.subject_teachers)):
			if self.subject_teachers[i].split(':')[0] == "#":
				self.nondangi_teachers.append(self.subject_teachers[i].split(':')[2])
	def getFullList(self):  # '과목선생님 : 이름  개인주소 : http://~~~~~' 요소 형식으로 리스트 리턴
		result = []
		for i in range(0, len(self.subject_teachers)):
			subjectCode = {'1':'국어', '2':'수학', '3':'영어', '10':'한국사', '4':'사회', '5':'과학', '6':'대학별고사', '7':'제2외국어', '9':'월간대치동', '11':'내신전문'}
			codeSplit = self.department
			subjectName = codeSplit
			if subjectName == "과학":
				subjectName = "과학탐구"
			elif subjectName == "사회":
				subjectName = "사회탐구"
			tNameBefore = self.subject_teachers[i].split(':')[1]
			hangul = re.compile('[^ ㄱ-ㅣ가-힣]+')
			tName = hangul.sub("", tNameBefore)
			result.append(subjectName + ': ' + tName)
		return result

	def getSubjectList(self):
		result = []
		for i in range(0, len(self.subject_teachers)):
			subjectName = self.department
			result.append(subjectName)
		return result

	def getIDList(self):
		result = []
		for i in range(0, len(self.subject_teachers)):
			teacherID = self.subject_teachers[i].split(':')[0]+':'+self.subject_teachers[i].split(':')[2]
			result.append(teacherID)
		return result

	def getNameList(self):
		result = []
		idList = []
		for i in range(0, len(self.subject_teachers)):
			result.append(self.subject_teachers[i].split(':')[1])
		return result

	def getAddressList(self):
		result = []
		for i in range(0, len(self.subject_teachers)):
			if self.subject_teachers[i].split(':')[0] == "#": #논단기
				boardaddress = 'http://nondangi.skyedu.com/teacher/qna/list.asp?%s' % self.subject_teachers[i].split(':')[2]
			elif self.subject_teachers[i].split(':')[0] == "@": #일반
				boardaddress = 'https://skyedu.conects.com/teachers/teacher_qna/?%s' % self.subject_teachers[i].split(':')[2]
			result.append(boardaddress)
		return result


	def getBoardAddressList(self, IDList):
		boardaddress = self.getAddressList()
		return boardaddress

	def getIndivBoardAddress(self, ID):
		# if ID in self.nondangi_teachers: #논단기
		# 	boardaddress = 'http://nondangi.skyedu.com/teacher/qna/list.asp?%s' % ID
		# elif ID not in self.nondangi_teachers: #일반
		# 	boardaddress = 'https://skyedu.conects.com/teachers/teacher_qna/?%s' % ID
		if ID.split(':')[0] == "#": #논단기
			boardaddress = 'http://nondangi.skyedu.com/teacher/qna/list.asp?%s' % ID.split(':')[1]
		elif ID.split(':')[0] == "@": #일반
			boardaddress = 'https://skyedu.conects.com/teachers/teacher_qna/?%s' % ID.split(':')[1]
		return boardaddress


class CalcSkyedu:  # 선생,과목, 게시판주소등을 토대로 본격적으로 긁어오고 결과를 엑셀로 출력하는 것까지 관련된 클래스
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
		
		if "nondangi" in self.url:
			webdriverWaitFor = "paging-container"
		else:
			webdriverWaitFor = "board-paging"
		while connected == 0:
			try:
					self.driver.get(self.url)
					WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, webdriverWaitFor)))
					time.sleep(self.waitTime)
					connected = 1
			except:
					if "nondangi" not in self.url:
						self.boardsource = self.driver.page_source
						self.boardpage = BeautifulSoup(self.boardsource, "html.parser")
						if len(self.boardpage.select('#bbs_content > table > tbody')[0].find_all(['tr'])) == 1:
							self.questionCount.append(0)
							return self.questionCount
						else:
							print('서버와 통신이 불안정 합니다. 재접속을 시도합니다.')
							labelstatus.setText('서버와 통신이 불안정 합니다. 재접속을 시도합니다.')
							time.sleep(1)
					else:
						print('서버와 통신이 불안정 합니다. 재접속을 시도합니다.')
						labelstatus.setText('서버와 통신이 불안정 합니다. 재접속을 시도합니다.')
						time.sleep(1)
		print('Page_' + str(1) + ' --> Searching...')
		for i in range(startpage, endpage):
			pageconnected = 0
			while pageconnected == 0:
				try:
					#self.driver.get(self.url)
					WebDriverWait(self.driver, 3).until(EC.presence_of_element_located((By.CLASS_NAME, webdriverWaitFor)))
					pageconnected = 1
					time.sleep(self.waitTime)
				except TimeoutException:
					print('서버와 통신이 불안정 합니다. 재시도 합니다. Inner; 주소가 변경되었는지 확인 필요')
					labelstatus.setText('서버와 통신이 불안정 합니다. 재접속을 시도합니다.')
					#self.driver.get(self.url)
					#WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.ID, 'iNavPaging')))
					# self.questionCount.append(self.questionCountIndiv)
					#print('서버 재연결 시도완료 Inner')
			
			source = self.driver.page_source
			page = BeautifulSoup(source, "html.parser")  # beautiful soup으로 html형식으로 파싱
			if "nondangi" not in self.url:
				firstbbsdata = page.select('#bbs_content > table > tbody > tr')[-1].select('td:nth-of-type(1)')[0].getText()
			if "nondangi" in self.url:
				internal_id = ""
				page_num = len( page.select('#wrap > div > div.container > div.teacher-sub-container.teacher-qna > div.quick-layout-container.max-container > div.layout-container > div > div')[0].find_all(['a'])) - 2
			else:
				internal_id = page.select('#search')[0].get("onkeydown").split(",")[5]
				page_num = len(page.select('#bbs_content > div.board-paging')[0].find_all(['a']))
			
			javascriptClick = ""
			classtype = ""
			# 대학별 논단기
			if 'nondangi' in self.url:
				table = page.select('#wrap > div > div.container > div.teacher-sub-container.teacher-qna > div.quick-layout-container.max-container > div.layout-container > div > table > tbody')
				trs = table[0].find_all('tr')
				javascriptClick = "document.getElementById('gotoPage').value = '%s'; document.getElementById('searchForm').submit();console.log(%s);"
				classtype = "nondangi"
			# 월간 대치동
			# elif 'monthly' in self.url:
			# 	table = page.select('#_content > div.board-contents > table > tbody')
			# 	trs = table[0].find_all('tr')
			# 	javascriptClick = "return goPage(%s);"
			# 	classtype = "monthly"
			# elif '2018' in self.url:
			# 	table = page.select('#contents-container > div > div.layout-contents > div > div.board-list > table > tbody')
			# 	trs = table[0].find_all('tr', class_ = None)
			# 	javascriptClick = "return goPage(%s);"
			# 	classtype = "general"
			# else:
			# 	table = page.find(class_="table-container")
			# 	trs = table.find_all('tr', class_ = None)
			# 	javascriptClick = "document.getElementById('page').value = '%s'; document.getElementById('Lform').submit();"
			# 	classtype = ""
			else:
				table = page.select('#bbs_content > table > tbody')
				trs = table[0].find_all('tr', class_ = None)
				javascriptClick = "return board_list(%s, 'list', 'skyedu_teacher_qna','regdate','',%s,'','','');"
				classtype = "general"

			tdinboardTable = []
			for t in range(0, len(trs)):
				#notice = trs[t].select('td:nth-of-type(1)')
				#td = trs[t].select('td:nth-of-type(5)')
				if classtype == "monthly":
					notice = trs[t].select('td:nth-of-type(2)')
					td = trs[t].select('td:nth-of-type(6)')
					if len(td) > 0 and len(str(notice[0])) > 0 :
						tdinboardTable.append(td[0].getText())
					elif len(notice) == 0: #페이지에 아무것도 없을때
						tdinboardTable.append('None')

				elif classtype == "general":
					notice0 = trs[t].select('td:nth-of-type(1)')
					if notice0[0].getText() != "공지" and notice0[0].getText() != "":
						notice = trs[t].select('td:nth-of-type(2)')
						td = trs[t].select('td:nth-of-type(6)')
						if len(td) > 0 and len(str(notice[0])) > 0 :
							tdinboardTable.append(td[0].getText())
						elif len(notice) == 0: #페이지에 아무것도 없을때
							tdinboardTable.append('None')
					else:
						pass
					
				elif classtype == "nondangi":
					notice = trs[t].select('td:nth-of-type(2)')
					td = trs[t].select('td:nth-of-type(5)')
					if len(td) > 0 and len(str(notice[0].getText())) > 0 :
						tdinboardTable.append(td[0].getText())
					elif len(notice) == 0: #페이지에 아무것도 없을때
						tdinboardTable.append('None')

			print("tdinboardTable",tdinboardTable)
			for j in range(0, len(tdinboardTable)):
				tdtext = tdinboardTable[j]
				tdsplit = tdtext.split('-')
				if classtype == "monthly":
					tdsplit = tdtext.split('.')
				if len(tdsplit) == 3:
					date = int(tdsplit[0] + tdsplit[1] + tdsplit[2])  # 20180117
					if date >= self.enddate and date <= self.startdate:
						self.questionCount.append(date)
					elif date < self.enddate:
						return self.questionCount  # 게시글을 최신날짜 --> 과거날짜로 서칭하면서 기준과거 조건보다 더 과거가 나오면 게시글 카운팅을 종료한다.
				else:
					if tdtext == 'None':
						return self.questionCount
					else:
						pass
			print("page:", i + 1, "internal_id:",internal_id)
			if page_num > 1:
				self.driver.execute_script(javascriptClick % ((i + 1), internal_id))
				#논단기의 경우 페이지 이동이 ajax로 이루어지기 때문에 ajax호출이 완전히 끝나서 페이지가 빠뀔때까지 기다리도록 처리
				if "nondangi" not in self.url:
					source_after_click = self.driver.page_source
					page_after_click = BeautifulSoup(source_after_click, "html.parser")  # beautiful soup으로 html형식으로 파싱
					#firstbbsdata_after_click = page_after_click.select('#bbs_content > table > tbody > tr:nth-of-type(1) > td:nth-of-type(1)')[0].getText()
					firstbbsdata_after_click = page_after_click.select('#bbs_content > table > tbody > tr')[-1].select('td:nth-of-type(1)')[0].getText()
					print(firstbbsdata_after_click)
					if firstbbsdata == firstbbsdata_after_click:
						print("wait ajax++++++++++++++++", firstbbsdata , firstbbsdata_after_click)
						check_ajax = 0
						while check_ajax == 0:
							time.sleep(0.5)
							source_after_click = self.driver.page_source
							page_after_click = BeautifulSoup(source_after_click, "html.parser")  # beautiful soup으로 html형식으로 파싱
							#firstbbsdata_after_click = page_after_click.select('#bbs_content > table > tbody > tr:nth-of-type(1) > td:nth-of-type(1)')[0].getText()
							firstbbsdata_after_click = page_after_click.select('#bbs_content > table > tbody > tr')[-1].select('td:nth-of-type(1)')[0].getText()
							print("wait ajax---------------------")
							if firstbbsdata != firstbbsdata_after_click:
								check_ajax = 1
						print("wait ajax****************", firstbbsdata , firstbbsdata_after_click)
				#논단기의 경우 페이지 이동이 ajax로 이루어지기 때문에 ajax호출이 완전히 끝나서 페이지가 빠뀔때까지 기다리도록 처리
			else:
				return self.questionCount
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
		worksheet = self.workbook.add_worksheet('스카이에듀')
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
			worksheet.write(y + 1, 1, '스카이에듀')
			worksheet.write(y + 1, 2, self.subject)
			worksheet.write(y + 1, 3, self.teacher)
			worksheet.write(y + 1, 4, self.count)

