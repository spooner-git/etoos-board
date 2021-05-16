import time
from bs4 import BeautifulSoup
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

print('시스템 준비중....')


class Company:
    def __init__(self, url):
        self.chrome_driver = ""
        self.soup = ""
        self.url = url

    def set_url(self, company_url):
        self.url = company_url

    def set_driver(self, driver):
        self.chrome_driver = driver

    def set_soup(self, soup):
        self.soup = soup

    def get_url(self):
        return self.url

    def get_driver(self):
        return self.chrome_driver

    def get_soup(self):
        return self.soup


class Teacher:
    def __init__(self, teacher_subject, teacher_name, teacher_code):
        self.subject = teacher_subject
        self.name = teacher_name
        self.code = teacher_code

    def set_subject(self, teacher_subject):
        self.subject = teacher_subject

    def set_name(self, teacher_name):
        self.name = teacher_name

    def set_code(self, teacher_code):
        self.code = teacher_code

    def get_subject(self):
        return self.subject

    def get_name(self):
        return self.name

    def get_code(self):
        return self.code

    def get_full_info(self):
        # return str(self.subject)+' '+str(self.name)+' : '+str(self.code)
        if self.subject == '수시논술':
            self.subject = '대학별고사'
        return str(self.subject) + ': ' + str(self.name)


class CounterData:
    num = 0
    date = ''


class Ebs(Company):
    def login(self, user_id, user_pw, labelstatus):
        labelstatus.setText('Login 시도중')
        print('login 시도')
        id = self.chrome_driver.find_element_by_name('userid')
        id.send_keys(user_id)
        password = self.chrome_driver.find_element_by_name('passwd')
        password.click()
        # password = self.chrome_driver.find_element_by_id('passwd')
        password.send_keys(user_pw)
        self.chrome_driver.find_element_by_xpath(
            '//*[@id="reNcontents"]/form/div[2]/div/fieldset/div[1]/button').click()
        print('login 완료')
        labelstatus.setText('Login 완료')

    def get_lecture_list(self):
        lecture_array = []
        lecture_list = self.soup.select(
            '#gnbWrap > div.sectionAllmenu.layerWrap > div.allMenuheader > div')
            # '#reNlnb > form > div.wrap-teacher-all > div > div.teacher-all > div > div.wrap-tabs > ul > li > a')
        for lecture in lecture_list:
            lecture_array.append(lecture.text.strip())
        return lecture_array

    def get_teacher_list(self, lecture_array):
        teacher_array = []
        i = 0
        # 국어,수학,영어,한국사
        #
        # teacher_list = self.soup.select('#reNlnb > form > div.wrap-teacher-all > div > div.teacher-all > '
        #                                 'div > div.wrap-tabs > div > div > ul')
        teacher_list = self.soup.select('#gnbWrap > div.sectionAllmenu.layerWrap > div.allMenuBody > div')

        for teacher in teacher_list:
            teacher_info = teacher.select('div > ul > li')
            # print(str(teacher_info))
            for data in teacher_info:
                data = data.select('a')[0]
                temp = data.get('href')
                temp2 = temp.split('=')
                teacher_code = temp2[1].split('&')
                teacher_name = data.text.split(' ')
                #강사 중복 제거 - hk.kim 18.01.31
                duplication_check = 0
                for teacher_duplication_check in teacher_array:
                    if teacher_duplication_check.get_code() == teacher_code[0]:
                        duplication_check = 1
                        break
                if duplication_check == 0:
                    teacher_data = Teacher(lecture_array[i], teacher_name[0], teacher_code[0])
                    teacher_array.append(teacher_data)
            i += 1
        #사탐,과탐,직탐,제2외국어,대학별고사 따로
        # selector 수정 - hk.kim 20.05.17
        # teacher_list2 = self.soup.select('#reNlnb > form > div.wrap-teacher-all > div > div.teacher-all > '
        #                                  'div > div.wrap-tabs > div.contents.type ')
        # # reNlnb > form > div.wrap-teacher-all > div > div.teacher-all > div > div.wrap-tabs > div.contents.type.row3 > div > div:nth-child(1) > div:nth-child(1)
        # for teacher in teacher_list2:
        #     teacher_info = teacher.select('div > div > ul > li > a')
        #
        #     for data in teacher_info:
        #         temp = data.get('href')
        #         temp2 = temp.split('=')
        #         teacher_code = temp2[1].split('&')
        #         teacher_name = data.text.split(' ')
        #         #강사 중복 제거 - hk.kim 18.01.31
        #         duplication_check = 0
        #         for teacher_duplication_check in teacher_array:
        #             if teacher_duplication_check.get_code() == teacher_code[0]:
        #                 duplication_check = 1
        #                 break
        #         if duplication_check == 0:
        #             teacher_data = Teacher(lecture_array[i], teacher_name[0], teacher_code[0])
        #             teacher_array.append(teacher_data)
        #     if teacher_info:
        #         i += 1
        return teacher_array

    def get_bbs_count(self, delay_time, end_date, start_date, check_stop_class,
                      labelstatus):  # processing pause : hk.kim-18.01.28
        end_point = 0
        counter = []
        counter_data = CounterData()
        bbs_page = 4
        page_counter = 1
        bbs_page_checker = 0
        counter_data.num = 0
        counter_data.date = ''
        before_num = []
        current_num = []
        if page_counter == 1:
            labelstatus.setText('Page_' + str(page_counter) + ' --> Searching...')
        print(str(start_date))
        print(str(end_date))
        while True:
            # reNcontainer > div > div > div > div.bordList > table > tbody > tr:nth-child(1)
            # bbs_lines = self.soup.select('#frm > table > tbody > tr')
            print('test1')
            bbs_lines = self.soup.select('#reNcontainer > div > div > div > div.bordList > table > tbody > tr.str_Q')
            # print(str(bbs_lines))
            bbs_number = None
            # print(str(before_num))
            # print(str(current_num))
            current_num = []
            bbs_main_check = 0
            for bbs_line in bbs_lines:
                # 마지막 페이지를 넘어가려고 하는 경우 처리
                bbs_main_check = 0
                bbs_number = bbs_line.select('td:nth-of-type(1)')
                if bbs_number is None:
                    bbs_main_check = 1
                else:
                    for bbs_number_info in bbs_number:
                        current_num.append(bbs_number_info.text)
                        if bbs_number_info.text in before_num:
                            bbs_main_check = 1

                if bbs_main_check == 0:
                    bbs_date = bbs_line.select('td:nth-of-type(7)')
                    for date in bbs_date:
                        date_text = date.text[2:]
                        # print(str(date_text))
                        if date_text <= start_date:
                            if date_text >= end_date:
                                if counter_data.num == 0:
                                    dateMac = str(date_text).replace(".", "")
                                    date_info = '20' + dateMac[0:2] + dateMac[2:4] + dateMac[4:6]  # 20180130
                                    counter_data.date = date_info
                                    counter_data.num += 1
                                else:
                                    compare_date = counter_data.date[2:4] + '.' + counter_data.date[
                                                                                  4:6] + '.' + counter_data.date[6:8]
                                    if str(compare_date) == str(date_text):
                                        counter_data.num += 1
                                    else:
                                        counter.append(counter_data)
                                        counter_data = CounterData()
                                        counter_data.num = 1
                                        dateMac = str(date_text).replace(".", "")
                                        date_info = '20' + dateMac[0:2] + dateMac[2:4] + dateMac[4:6]  # 20180130
                                        counter_data.date = date_info
                            else:
                                if counter_data.num != 0:
                                    if end_point == 0:
                                        counter.append(counter_data)
                                        counter_data = CounterData()
                                        counter_data.num = 0
                                        counter_data.date = ''
                                end_point = 1
                                break
            #게시글이 아예 없는 경우

            before_num = current_num
            if not bbs_number:
                end_point = 1
            if bbs_main_check == 1:
                end_point = 1
            # print(str(counter_data))
            # print(str(end_point))
            if end_point == 1:
                break
            else:
                #게시판 연동 확인 - hk.kim 18.01.31
                pageconnected = 0
                while pageconnected == 0:
                    try:
                        WebDriverWait(self.chrome_driver, 3).until(
                            EC.presence_of_element_located((By.ID, 'bordPaging')))
                        pageconnected = 1
                        time.sleep(2)
                    except TimeoutException:
                        print('서버와 통신이 불안정 합니다. 재시도 합니다. Inner')
                        labelstatus.setText('서버와 통신이 불안정 합니다. 재접속을 시도합니다.')

                if end_point == 1:
                    break
                else:
                    pageconnected = 0
                    while pageconnected == 0:
                        try:
                            link = self.chrome_driver.find_element_by_xpath(
                                '//*[@id="bordPaging"]/ul/li[' + str(bbs_page) + ']/span')
                            link.click()
                            WebDriverWait(self.chrome_driver, 3).until(
                                EC.presence_of_element_located((By.ID, 'bordPaging')))
                            pageconnected = 1
                            time.sleep(delay_time)
                        except TimeoutException:
                            print('서버와 통신이 불안정 합니다. 재시도 합니다. Inner')
                            labelstatus.setText('서버와 통신이 불안정 합니다. 재접속을 시도합니다.')
                        except NoSuchElementException:
                            end_point = 1
                            break
                    if end_point == 1:
                        break
                    # print('last_test5')
                    page_counter += 1
                    labelstatus.setText('Page_' + str(page_counter) + ' --> Searching...')
                    #self.chrome_driver.implicitly_wait(delay_time)
                    html = self.chrome_driver.page_source
                    # print('last_test6')
                    self.soup = BeautifulSoup(html, 'html.parser')
                    # print('last_test7')
            # print(str(bbs_page))
            if bbs_page == 2:
                bbs_page += 2
            elif bbs_page == 12:
                # if bbs_page_checker == 0:
                #     bbs_page = 4
                #     bbs_page_checker += 1
                # else:
                bbs_page += 1
            elif bbs_page == 13:
                bbs_page = 4
            else:
                bbs_page += 1
                #processing pause : hk.kim-18.01.28

            # print('last_test8')
            while True:
                is_pause = check_stop_class.get_is_pause()
                if is_pause == 0:
                    break
            # print('last_test9')

        # print('last_test10')
        if counter_data.num != 0:
            counter.append(counter_data)
            counter_data = CounterData()
            counter_data.num = 0
            counter_data.date = ''

        # print('last_test11')
        return counter

    def remove_teacher_by_name(self, teacher_array, remove_teacher_name_list):
        # 실제 값을 복사
        teacher_list = teacher_array[:]
        for teachers in teacher_list:
            for teacher_name in remove_teacher_name_list:
                if teachers.get_name() == teacher_name:
                    teacher_array.remove(teachers)

        return teacher_array

    def add_teacher_by_name(self, teacher_array, add_teacher_name_list):
        # 실제 값을 복사
        teacher_list = teacher_array[:]
        teacher_array = []
        for teachers in teacher_list:
            for teacher_name in add_teacher_name_list:
                if teachers.get_name() == teacher_name:
                    teacher_array.append(teachers)

        return teacher_array

    def remove_teacher_by_subject(self, teacher_array, remove_subject_name_list):
        # 실제 값을 복사
        teacher_list = teacher_array[:]
        for teachers in teacher_list:
            for subject_name in remove_subject_name_list:
                if teachers.get_subject() == subject_name:
                    teacher_array.remove(teachers)

        return teacher_array

    def add_teacher_by_subject(self, teacher_array, add_subject_name_list):
        # 실제 값을 복사
        teacher_list = teacher_array[:]
        teacher_array = []
        for teachers in teacher_list:
            for subject_name in add_subject_name_list:
                if teachers.get_subject() == subject_name:
                    teacher_array.append(teachers)

        return teacher_array

    def go_to_url_page(self, url, qna):
        if qna == 1:
            pageconnected = 0
            while pageconnected == 0:
                try:
                    self.chrome_driver.get(url)
                    WebDriverWait(self.chrome_driver, 50).until(
                        EC.presence_of_element_located((By.ID, 'bordPaging')))
                    pageconnected = 1
                    time.sleep(2)
                except TimeoutException:
                    print('서버와 통신이 불안정 합니다. 재시도 합니다. Inner')
        else:
            self.chrome_driver.get(url)

        html = self.chrome_driver.page_source
        self.soup = BeautifulSoup(html, 'html.parser')

    def xlsxWrite(workBook, finalresultForExcel):
        workbook = workBook
        worksheet = workbook.add_worksheet('EBS')
        format = workbook.add_format()
        format.set_bg_color('#FF6600')
        worksheet.set_column(0, 4, 12)
        worksheet.write(0, 0, '날짜', format)
        worksheet.write(0, 1, '사이트', format)
        worksheet.write(0, 2, '과목', format)
        worksheet.write(0, 3, '선생님', format)
        worksheet.write(0, 4, '게시물수', format)
        finalresultForExcel = finalresultForExcel
        for y in range(0, len(finalresultForExcel)):
            split = finalresultForExcel[y].split(':')
            subject = split[0]
            if subject == '수시논술':
                subject = '대학별고사'
            teacher = split[1]
            date = split[2]
            count = split[3]
            dateformat = date[0:4] + '-' + date[4:6] + '-' + date[6:]
            worksheet.write(y + 1, 0, dateformat)
            worksheet.write(y + 1, 1, 'EBS')
            worksheet.write(y + 1, 2, subject)
            worksheet.write(y + 1, 3, teacher)
            worksheet.write(y + 1, 4, count)
