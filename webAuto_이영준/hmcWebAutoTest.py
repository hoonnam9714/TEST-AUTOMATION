# 학습을 위해 처음부터 코딩해보자 - YoungjunLee 2021.11.22
import time
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait          #웹페이지 로딩 대기
from selenium.webdriver.support import expected_conditions as EC 
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.common.exceptions import JavascriptException
from xpath import *
from commonConst import *
from commonFunc import *
from hmcTestData import *
import pyautogui

# fileOpen
hmcTestData1 = openExcelFile(excelPath, 'testCase_multi.xlsx', "hmcTestData", 1)
# print(hmcTestData1.emNo)

# 오픈된 파일 행수만큼 Loop
for x in range(0, len(hmcTestData1)):
    hmcTestDataOne: hmcTestData = hmcTestData1[x]

    # print(hmcTestDataOne)
    # webDriver open 
    browser = openWebDriver(chromeDriverPath, "hmcQaUrl", hmcTestDataOne.emNo)    
    # 고객명 처리
    if hmcTestDataOne.custNm :
        browser.implicitly_wait(10)
        time.sleep(0.5)
        
        if hmcTestDataOne.custNm and hmcTestDataOne.custNm != "임의설계": #고객명이 존재할 경우
            time.sleep(0.5)
            browser.find_element(By.CSS_SELECTOR, ".search > input").send_keys(hmcTestDataOne.custNm) #엑셀에서 가져온 고객명 맵핑
            browser.find_element(By.CSS_SELECTOR, ".search > input").send_keys(Keys.ENTER)                    #고객명 엔터
            time.sleep(0.5)
            browser.find_element_by_xpath(hmc_xpath1).click()      #신규설계 버튼
        else:
            time.sleep(0.5)
            browser.find_element_by_xpath(hmc_xpath0).click()      #고객없이설계(임의설계)
            time.sleep(0.5)

    if hmcTestDataOne.prdCd:
        time.sleep(2)
        browser.implicitly_wait(10)
        time.sleep(0.5)
        goodCode = str(hmcTestDataOne.prdCd) #엑셀상품코드
        elems = browser.find_elements_by_tag_name("li") #화면의 li 태그를 모두 가져옴
        for elem in elems:    #반복하면서 elems 하나씩 뽑아오기
            if str(elem.get_attribute("id")):
                result = str(elem.get_attribute("id"))
                result1 = result[0:4] #앞4자리 자르기
                result2 = result[4:8] #뒷3자리 자르기
                if result1 == goodCode[0:4] and int(result2) <= int(goodCode[4:8]): #상품코드 앞4자리가 동일하고, 뒷3자리가 작거나 같으면 맵핑
                    elem = browser.find_element_by_xpath("//*[@id="+result+"]")
                    browser.execute_script("arguments[0].click();", elem)
                    break

    if hmcTestDataOne.contrRrno:
        time.sleep(0.5)
        pyautogui.moveTo(958,688) #지정한 위치로 마우스를 이동(팝업확인클릭)
        pyautogui.click()
        time.sleep(0.5)
        try:
            if hmcTestDataOne.contrRrno:
                browser.find_element(By.XPATH, hmc_xpath2_3).send_keys(hmcTestDataOne.contrRrno) #엑셀에서 가져온 계약자생년월일 맵핑
        except NoSuchElementException:
            pass
        
        try:
            if hmcTestDataOne.insdRrno:
                browser.find_element(By.XPATH, hmc_xpath2_3_4).send_keys(hmcTestDataOne.insdRrno) #엑셀에서 가져온 주피생년월일 맵핑
        except NoSuchElementException:
            pass
        
        try:
            if hmcTestDataOne.insd1Rrno:
                browser.find_element(By.XPATH, hmc_xpath2_3_4_1).send_keys(hmcTestDataOne.insd1Rrno) #엑셀에서 가져온 종피1생년월일 맵핑
        except NoSuchElementException:
            pass

        try:
            if hmcTestDataOne.insd2Rrno:
                browser.find_element(By.XPATH, hmc_xpath2_3_4_2).send_keys(hmcTestDataOne.insd2Rrno) #엑셀에서 가져온 종피2생년월일 맵핑
        except NoSuchElementException:
            pass

        try:
            if hmcTestDataOne.insd3Rrno:
                browser.find_element(By.XPATH, hmc_xpath2_3_4_3).send_keys(hmcTestDataOne.insd3Rrno) #엑셀에서 가져온 종피3생년월일 맵핑
        except NoSuchElementException:
            pass

    # elif "성별" in ws.cell(row=1,column=y).value:
    #     time.sleep(0.5)
    #     try:
    #         if ws.cell(row=x,column=y).value:
    #             if "계약자성별" in ws.cell(row=1,column=y).value:
    #                 if "남" in ws.cell(row=x,column=y).value:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_1).click()      #엑셀에서 가져온 계약자성별(남자) 맵핑
    #                 else:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_2).click()      #엑셀에서 가져온 계약자성별(여자) 맵핑
    #             elif "주피성별" in ws.cell(row=1,column=y).value:
    #                 if "남" in ws.cell(row=x,column=y).value:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_5).click()      #엑셀에서 가져온 주피성별(남자) 맵핑
    #                 else:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_6).click()      #엑셀에서 가져온 주피성별(여자) 맵핑
    #             elif "종피1성별" in ws.cell(row=1,column=y).value:
    #                 if "남" in ws.cell(row=x,column=y).value:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_6_1).click()      #엑셀에서 가져온 종피1성별(남자) 맵핑
    #                 else:
    #                     browser.find_element(By.ID, "woman03")
    #                     # browser.find_element_by_xpath(hmc_xpath2_3_6_2).click()      #엑셀에서 가져온 종피1성별(여자) 맵핑
    #             elif "종피2성별" in ws.cell(row=1,column=y).value:
    #                 if "남" in ws.cell(row=x,column=y).value:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_6_3).click()      #엑셀에서 가져온 종피2성별(남자) 맵핑
    #                 else:
    #                     element = browser.find_element(By.XPATH, hmc_xpath2_3_6_4)
    #                     browser.execute_script("arguments[0].click();", element)                 #강제클릭으로 엘리먼트 에러 해결
    #             elif "종피3성별" in ws.cell(row=1,column=y).value:
    #                 if "남" in ws.cell(row=x,column=y).value:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_6_5).click()      #엑셀에서 가져온 종피3성별(남자) 맵핑
    #                 else:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_6_6).click()      #엑셀에서 가져온 종피3성별(여자) 맵핑
    #     except NoSuchElementException:
    #         pass
    # elif "동일" in ws.cell(row=1,column=y).value:
    #     time.sleep(0.5)
    #     try:
    #         if ws.cell(row=x,column=y).value:
    #             if "Y" in ws.cell(row=x,column=y).value:
    #                 browser.find_element_by_xpath(hmc_xpath2_3_3).click()      #엑셀에서 가져온 계약자와 동일 맵핑
    #     except NoSuchElementException:
    #         pass
    # elif "직종" in ws.cell(row=1,column=y).value:
    #     time.sleep(0.5)
    #     try:
    #         if ws.cell(row=x,column=y).value:
    #             if "주피직종" in ws.cell(row=1,column=y).value:
    #                 if "비위험" in ws.cell(row=x,column=y).value:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_7).click()      
    #                 elif "위험1" in ws.cell(row=x,column=y).value:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_7_1).click()      
    #                 elif "위험2" in ws.cell(row=x,column=y).value:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_7_2).click()      
    #                 elif "위험3" in ws.cell(row=x,column=y).value:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_7_3).click()      
    #                 elif "위험4" in ws.cell(row=x,column=y).value:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_7_4).click()     
    #             elif "종피1직종" in ws.cell(row=1,column=y).value:
    #                 if "비위험" in ws.cell(row=x,column=y).value:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_8_1).click()      
    #                 elif "위험1" in ws.cell(row=x,column=y).value:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_8_2).click()      
    #                 elif "위험2" in ws.cell(row=x,column=y).value:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_8_3).click()      
    #                 elif "위험3" in ws.cell(row=x,column=y).value:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_8_4).click()      
    #                 elif "위험4" in ws.cell(row=x,column=y).value:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_8_5).click()     
    #             elif "종피2직종" in ws.cell(row=1,column=y).value:
    #                 if "비위험" in ws.cell(row=x,column=y).value:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_9_1).click()      
    #                 elif "위험1" in ws.cell(row=x,column=y).value:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_9_2).click()      
    #                 elif "위험2" in ws.cell(row=x,column=y).value:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_9_3).click()      
    #                 elif "위험3" in ws.cell(row=x,column=y).value:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_9_4).click()      
    #                 elif "위험4" in ws.cell(row=x,column=y).value:
    #                     element = browser.find_element(By.XPATH, hmc_xpath2_3_9_5)
    #                     browser.execute_script("arguments[0].click();", element)                 #강제클릭으로 엘리먼트 에러 해결
    #             elif "종피3직종" in ws.cell(row=1,column=y).value:
    #                 if "비위험" in ws.cell(row=x,column=y).value:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_9_6).click()      
    #                 elif "위험1" in ws.cell(row=x,column=y).value:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_9_7).click()      
    #                 elif "위험2" in ws.cell(row=x,column=y).value:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_9_8).click()      
    #                 elif "위험3" in ws.cell(row=x,column=y).value:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_9_9).click()      
    #                 elif "위험4" in ws.cell(row=x,column=y).value:
    #                     browser.find_element_by_xpath(hmc_xpath2_3_9_0).click()     
    #     except NoSuchElementException:
    #         pass
    # elif "보험종류1" == ws.cell(row=1,column=y).value:
    #     time.sleep(1.5)
    #     pyautogui.moveTo(958,688) #지정한 위치로 마우스를 이동(팝업확인클릭)
    #     pyautogui.click()
    #     time.sleep(0.5)
    #     try:
    #         elem = browser.find_element_by_xpath(hmc_xpath2_4)      #보험종류변경
    #         browser.execute_script("arguments[0].click();", elem)
    #     except:
    #         print("보험종류 변경없음")
    #     time.sleep(0.5)
        
    #     elems1 = browser.find_elements_by_class_name("insureSelect") #class_name "insureSelect"(보험종류)을 가지는 모든 엘리먼트 가져오기
    #     count=0
    #     countDetail = []
    #     #보험종류 및 보험종류별 상세 갯수
    #     for e1 in elems1:    
    #         if e1.text:
    #             w = str(e1.text)
    #             countDetail.append(str(w.count('\n')))  #보험종류별 상세 갯수
    #             count = count+1   #보험종류 갯수
    
    #     for i in range(0,count): #보험종류 만큼 반복(col)
    #         for j in range(0,int(countDetail[i])): #보험종류별 상세 만큼 반복(row)
    #             elem = browser.find_element_by_xpath(hmc_xpath2_5_0+'['+str(i+1)+']/p/select/option'+'['+str(j+1)+']')
    #             time.sleep(0.5)
    #             if ws.cell(row=x+j,column=y+i).value:  #보험종류1이 존재하면 맵핑 후 다음으로 진행
    #                 elemText = str(ws.cell(row=x,column=y+i).value)
    #                 elem1 = browser.find_element_by_xpath(hmc_xpath2_5_0+'['+str(i+1)+']/p/select')
    #                 elem1.send_keys(elemText) #엑셀 보험종류로 select 값 셋팅
    #     time.sleep(0.5)
    #     try:
    #         browser.find_element_by_xpath(hmc_xpath2_6).click()      #적용 버튼    
    #     except:
    #         print("보험종류 변경 후 적용버튼이 존재하지 않아 skip")
    #     time.sleep(0.5)
    # elif "건강" in ws.cell(row=1,column=y).value:
    #     browser.implicitly_wait(10)
    #     try:    
    #         if "Y" == ws.cell(row=x,column=y).value:
    #             elem = browser.find_element_by_id("chk01")      #주피 건강체 체크박스
    #             browser.execute_script("arguments[0].click();", elem)
    #             time.sleep(0.5)
    #             elem = browser.find_element_by_id("chk02")      #종피 건강체 체크박스
    #             browser.execute_script("arguments[0].click();", elem)
    #             time.sleep(0.5)
    #     except:
    #         pass
    #     browser.find_element_by_xpath(hmc_xpath3).click()        #다음 버튼    
    # elif "플랜선택" in ws.cell(row=1,column=y).value:
    #     browser.implicitly_wait(10)
    #     time.sleep(3)

    #     #######################################
    #     # 플랜설계 팝업 닫기 
    #     #######################################
    #     try:
    #         elem = browser.find_element_by_xpath('//*[@id="container"]/div[3]/div[2]/div/div/button/span') #직접설계
    #         browser.execute_script("arguments[0].click();", elem)
    #     except NoSuchElementException:
    #         pass

    #     if ws.cell(row=x,column=y).value:
    #         try:
    #             for i in range(1,10):
    #                 time.sleep(0.5)
    #                 if ws.cell(row=x,column=y).value == browser.find_element_by_xpath(hmc_xpath2_8+'['+str(i)+']/a/p').text:
    #                     time.sleep(0.5)
    #                     browser.find_element_by_xpath(hmc_xpath2_8+'['+str(i)+']/a/p').click()      #플랜선택
    #                     break
    #         except:
    #             pass
    # elif "주계약" in ws.cell(row=1,column=y).value:
    #     browser.implicitly_wait(10)
    #     if ws.cell(row=x,column=y).value and "연금" not in ws.cell(row=x,column=5).value: #주계약 선택이 있고, 연금/저축 상품이 아니면
    #         if "가입" in ws.cell(row=x,column=y).value:
    #             browser.find_element_by_xpath(hmc_xpath5_4).click()        #가입금액
    #         else:
    #             browser.find_element_by_xpath(hmc_xpath5_5).click()        #합계보험료
    #         time.sleep(0.5)
    #         browser.find_element_by_xpath(hmc_xpath5).send_keys(ws.cell(row=x,column=y+1).value)      #가입금액 or 합계보험료
    #         time.sleep(0.5)
    #         for i in range(2,11): #연금설계탭-연금개시나이 ~ 연금선지급탭-선지급기간 9개 항목 맵핑
    #             if ws.cell(row=x,column=y+i).value:
    #                 selectClick(i)
    #                 time.sleep(0.5)
    #         for i in range(16,19): #보험기간/납입기간/납입주기 3개 항목 맵핑
    #             if ws.cell(row=x,column=y+i).value:    
    #                 selectClick(i)
    #                 time.sleep(0.5)
    #     elif "연금" in ws.cell(row=x,column=5).value or "저축" in ws.cell(row=x,column=5).value: #연금/저축
    #         if ws.cell(row=x,column=y+1).value: #주계약보험료가 존재하면
    #             browser.find_element_by_xpath(hmc_xpath5).send_keys(ws.cell(row=x,column=y+1).value)      #주계약보험료
    #         if ws.cell(row=x,column=y+11).value: #연금보험료(거치)가 존재하면
    #             browser.find_element_by_xpath(hmc_xpath5_5_1).send_keys(ws.cell(row=x,column=y+11).value)    
    #         if ws.cell(row=x,column=y+12).value: #연금보험료(적립)가 존재하면
    #             browser.find_element_by_xpath(hmc_xpath5_5_2).send_keys(ws.cell(row=x,column=y+12).value)
            
    #         for i in range(13,19): #연금지급형태 ~ 납입주기 6개 항목 맵핑
    #             if ws.cell(row=x,column=y+i).value:
    #                 selectClick(i)
    #                 time.sleep(0.5)
    #     else: #플랜선택한 종신의 경우 다음단계로 스킵
    #         continue
    #     browser.implicitly_wait(10)
    #     try:
    #         if ws.cell(row=x,column=y+19).value:   #환급특약제외
    #             checkBoxClick(hmc_xpath5_8, y+19, '#chk02')
    #         if ws.cell(row=x,column=y+20).value:   #연금납입면제특약제외
    #             checkBoxClick(hmc_xpath5_8, y+20, '#chk01')
    #     except JavascriptException:
    #         pass
    #     element = browser.find_element(By.CSS_SELECTOR, ".b_plus:nth-child(3)")  #특약추가버튼
    #     browser.execute_script("arguments[0].click();", element)                 #강제클릭으로 엘리먼트 에러 해결
    #     time.sleep(0.5)
    
    #     browser.implicitly_wait(10)
    #     browser.find_element_by_class_name("btn_white_normal").click()        #특약전체선택1
    #     time.sleep(0.5)
    #     browser.find_element_by_class_name("btn_point_pop").click()        #특약전체추가1
    #     time.sleep(0.5)
        
    # elif "특약코드" in ws.cell(row=1,column=y).value:
    #     browser.implicitly_wait(10)
    #     if ws.cell(row=x,column=y).value:
    #         optionCode = str(ws.cell(row=x,column=y).value) #특약코드
    #         time.sleep(0.5)

    #         pyautogui.click()
    #         pyautogui.hotkey("pagedown") #스크롤 다운
    #         time.sleep(0.5)

    #         # 5의 배수 만큼 반복
    #         if k%5==0:
    #             for k in range(1,k//5+1):
    #                 pyautogui.hotkey("pagedown") #스크롤 다운
    #             time.sleep(0.5)
    #         k=k+1
            
    #         try:
    #             result = browser.find_element_by_xpath('//*[@id='+optionCode+']/div/p[1]/font')
    #             result = result.text[0:1] #'[' 값 리턴
    #         except:
    #             result = ""

    #         if ws.cell(row=x,column=y+2).value and (result is None or result == ""): #특약금액이 존재하고 특약명 앞에 "["가 없는 경우
    #             #특약 상세 열기
    #             element = browser.find_element_by_xpath('//*[@id='+optionCode+']/div[1]/a') #특약 선택
    #             browser.execute_script("arguments[0].click();", element) #강제클릭으로 엘리먼트 에러 해결
    #             time.sleep(0.5)
    #             element = browser.find_element_by_xpath('//*[@id='+optionCode+']/div[2]/div/div[1]/div/input')
    #             time.sleep(0.5)

    #             # create action chain object
    #             action = ActionChains(browser)
    #             # perform the operation
    #             action.move_to_element(element).click().perform()

    #             pyautogui.hotkey("ctrl","a") #전체선택
    #             time.sleep(0.5)
    #             pyautogui.hotkey("del") #전체선택
    #             time.sleep(0.5)
    #             browser.find_element_by_xpath('//*[@id='+optionCode+']/div[2]/div/div[1]/div/input').send_keys(ws.cell(row=x,column=y+2).value) #특약금액

            
    #         if ws.cell(row=x,column=y+2).value and result == "[": #특약금액이 존재하고 특약명 앞에 "["가 있는 경우
    #             #특약 상세 열기
    #             element = browser.find_element_by_xpath('//*[@id='+optionCode+']/div/a') #특약 선택
    #             browser.execute_script("arguments[0].click();", element) #강제클릭으로 엘리먼트 에러 해결
    #             time.sleep(0.5)
    #             element = browser.find_element_by_xpath('//*[@id='+optionCode+']/div[2]/div/div/div/input')
    #             time.sleep(0.5)

    #             # create action chain object
    #             action = ActionChains(browser)
    #             # perform the operation
    #             action.move_to_element(element).click().perform()

    #             time.sleep(0.5)
    #             pyautogui.hotkey("ctrl","a") #전체선택
    #             time.sleep(0.5)
    #             pyautogui.hotkey("del") #전체선택   
    #             time.sleep(0.5)
    #             browser.find_element_by_xpath('//*[@id='+optionCode+']/div[2]/div/div/div/input').send_keys(ws.cell(row=x,column=y+2).value) #특약금액
            
    #         try:
    #             if ws.cell(row=x,column=y+3).value: #보험기간이 존재하면
    #                 selectClick(3)
    #             if ws.cell(row=x,column=y+4).value: #납입기간이 존재하면
    #                 selectClick(4)
    #         except ElementClickInterceptedException:
    #             pass            
    #         except InvalidSelectorException:
    #             pass
            
    #         time.sleep(0.5)
    #         #특약 상세 닫기
    #         element = browser.find_element_by_xpath('//*[@id='+optionCode+']/div/a') #특약 선택
    #         browser.execute_script("arguments[0].click();", element) #강제클릭으로 엘리먼트 에러 해결
    #         time.sleep(0.5)
    # #마지막 col일 경우 상세결과보기 진행
    # elif "비고" in ws.cell(row=1,column=y).value:
    #     browser.implicitly_wait(10)
    #     WebDriverWait(browser, 30).until(EC.presence_of_element_located((By.XPATH, hmc_xpath6)))
    #     browser.find_element_by_xpath(hmc_xpath6).click()       #상세결과보기
    #     time.sleep(5)
    #     browser.implicitly_wait(10)
    #     try:
    #         elem = WebDriverWait(browser, 30).until(EC.presence_of_element_located((By.XPATH, hmc_xpath7))) #계약자명으로 엘리먼트 로딩 확인
    #     except:
    #         pass
    #     try:
    #         elem = WebDriverWait(browser, 30).until(EC.presence_of_element_located((By.XPATH, hmc_xpath8))) #설계저장으로 엘리먼트 로딩 확인
    #         elem.click()
    #     except:
    #         pass
    #     browser.implicitly_wait(10)
    #     WebDriverWait(browser, 30).until(EC.presence_of_element_located((By.XPATH, hmc_xpath9)))
    #     time.sleep(1)
    #     browser.find_element_by_xpath(hmc_xpath9).click()       #상품설명서 미리보기
    #     time.sleep(20)
    #     pyautogui.click()
    #     pyautogui.hotkey("ctrl","s") #PDF 최초 저장
    #     # uiImgFind("OZ_IMG_DOWNLOAD_2.png")  #오즈 다운로드 이미지 파일
    #     time.sleep(1)
    #     pyautogui.moveTo(126,992) #지정한 위치로 마우스를 이동
    #     pyautogui.click()
    #     time.sleep(1)
    #     pyautogui.moveTo(524,59) #지정한 위치로 마우스를 이동
    #     pyautogui.click()
    #     pyautogui.hotkey("ctrl","s") #PDF 파일명 변경을 위한 저장
    #     # uiImgFind("file_download_2.png")  # 우측 상단 PDF 다운로드 이미지 파일
    #     time.sleep(1)
    #     pyautogui.click()
    #     time.sleep(1)
    #     pyperclip.copy("C:/Users/Administrator/Desktop/PythonWorkspace") #상품설명서 저장 경로 클립보드에 복사
    #     pyautogui.hotkey("ctrl","v") #붙여넣기
    #     time.sleep(1)
    #     pyautogui.press('enter') #엔터키
    #     time.sleep(1)

    #     current_datetime = datetime.datetime.now()
    #     dateformat = "%Y%m%d%H%M%S"
    #     tmp = current_datetime.strftime(dateformat) #현재시간
    #     filename = str(ws.cell(row=x,column=5).value)+'_'+str(ws.cell(row=x,column=3).value)+'_'+tmp

    #     pyautogui.moveTo(474,703) #지정한 위치로 마우스를 이동
    #     pyautogui.click()
    #     time.sleep(1)
    #     pyperclip.copy(filename) #상품명+고객명 클립보드에 복사

    #     pyautogui.hotkey("ctrl","v") #붙여넣기
    #     time.sleep(1)
    #     pyautogui.press('enter') #엔터키
    #     time.sleep(1)
    # ###########################################
    # #브라우저 종료
    # ###########################################
    sleep(5)
    
    browser.quit()

# print(len(hmcTestData1))
# print(hmcTestData1[0].emNo.value)

# broswer.close()
