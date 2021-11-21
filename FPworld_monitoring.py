import pyautogui
import pyperclip
import time
import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By

###############################
# 메인frame 전환 및 초기화 함수
###############################
def mainFrameInit(arg1,arg2): #agr1: 업무명, arg2: 현재시간
    browser.switch_to.default_content() #메인화면으로 frame전환
    screenshot('windowContainer1_subWindow0_body',arg1,arg2) # 스크린샷
    #메인로고 더블클릭(frame 초기화)
    browser.find_element(By.XPATH, '//*[@id="btnMainLogo"]').click()
    browser.find_element(By.XPATH, '//*[@id="btnMainLogo"]').click()

###############################
# 스크린샷 함수
###############################
def screenshot(arg1, arg2, arg3): #arg1:ID, arg2:업무명, arg3:현재시간
    time.sleep(0.5)
    elem = browser.find_element(By.ID, arg1)
    elem.screenshot('/Users/chaekyunghoon/desktop/PythonWorkSpace/img/'+arg2+'_'+ str(arg3) +'.png') #스크린샷

###############################
# 메뉴 선택 함수
###############################
def menuChoice(arg1, arg2): #arg1:대분류, arg2:소분류
    browser.implicitly_wait(10)
    time.sleep(0.5)
    browser.find_element(By.ID, arg1).click() #대분류
    time.sleep(0.5)
    browser.find_element(By.ID, arg2).click() #소분류
    time.sleep(1)
    browser.implicitly_wait(10)
    browser.switch_to.frame('windowContainer1_subWindow0_iframe') #frame 전환
    time.sleep(0.5)

###############################
# 카카오톡 전송
###############################
def kakaoSend():
    browser.quit() #브라우저 종료
    time.sleep(0.5)
    pyautogui.moveTo(30,457) #카카오톡 메시지 박스
    time.sleep(0.5)
    pyautogui.click()
    time.sleep(0.5)
    pyperclip.copy("FPworld 모니터링 실패_"+str(currentTime)) #클립보드에 복사
    pyautogui.hotkey("command","v") #붙여넣기
    time.sleep(0.5)
    pyautogui.moveTo(346,482) #카카오톡 전송
    time.sleep(0.5)
    pyautogui.click()
    time.sleep(0.5)

###############################
# 메인
###############################
if __name__ == "__main__":
    browser = webdriver.Chrome(executable_path='/Users/chaekyunghoon/desktop/PythonWorkSpace/chromedriver')
    url = "https://hmp.hanwhalife.com/online/fp" #HMC QA 테스트 URL
    browser.get(url) #FP월드 URL
    browser.maximize_window()
    browser.find_element(By.ID, 'ibxUserId').send_keys("2140046") 
    browser.find_element(By.ID, 'ibxPassword').send_keys("hwtrust02!")
    browser.find_element(By.ID, 'rdoTitl_input_0').click() #스마트인증 라디오 버튼
    time.sleep(0.5)
    browser.find_element(By.XPATH, '//*[@id="bt_login"]/a').click() #로그인 버튼
    time.sleep(0.5)

    ###############################
    # 팝업 종료 처리
    ###############################
    browser.implicitly_wait(10)
    handles = browser.window_handles #메인페이지 및 팝업 갯수

    # 메인페이지가 아니면 팝업(브라우저) 종료
    for handle in handles:
        if handle != handles[0]:
            browser.switch_to.window(handle)
            browser.close()

    time.sleep(0.5)
    browser.switch_to.window(browser.window_handles[0]) #메인화면으로 전환
    elem = browser.find_element(By.XPATH, '//*[@id="smallpop_close"]') #내부팝업(?) 종료
    browser.execute_script("arguments[0].click();", elem)

    ###############################
    # 모니터링 반복 시작
    ###############################
    failcount = 0 #실패카운트
    while True:
        try:
            current_datetime = datetime.datetime.now()
            dateformat = "%Y%m%d%H%M%S"
            currentTime = current_datetime.strftime(dateformat) #현재시간 ex)20211103153201

            ###############################
            # 예외발생시 종료 및 카카오톡 발송
            ###############################
            if failcount > 0: 
                kakaoSend() #카카오톡 발송
                break
            
            ###############################
            # 정상일 경우
            ###############################
            if (int('090000') < int(currentTime[8:14]) or int('90000') > int(currentTime[8:14])): #16시 이후거나 08시 이전이면 실행
            # if (int('180000') < int(currentTime[8:14]) or int('90000') > int(currentTime[8:14])) and int(currentTime[11:14]) == int('000'): #16시 이후거나 08시 이전이면 10분 단위로 실행
                screenshot('wq_uuid_287','메인',currentTime) # 메인화면 스크린샷

                ###############################
                # 고객 조회
                ###############################
                menuChoice('genMenuDepth1_0_menuNm','genMenuDepth2_0_genMenuDepth3_0_menuNm') #메뉴선택
                browser.find_element(By.XPATH, '//*[@id="ipt_srchAdmnCustNo"]').send_keys('1') #고객번호
                time.sleep(0.5)
                browser.find_element(By.XPATH, '//*[@id="btnSrch"]/a').click() #조회
                mainFrameInit('고객',currentTime)

                ###############################
                # 보장분석 조회
                ###############################
                menuChoice('genMenuDepth1_2_menuNm','genMenuDepth2_0_genMenuDepth3_0_menuNm') #메뉴선택
                browser.find_element(By.XPATH, '//*[@id="ibxCustName"]').send_keys('채경훈') #고객명
                time.sleep(0.5)
                browser.find_element(By.XPATH, '//*[@id="btnSuch"]/a').click() #조회
                time.sleep(0.5)
                browser.switch_to.frame('mabnf001pvw10_iframe') #frame 전환(신용정보원 or 당사)
                time.sleep(0.5)
                browser.find_element(By.XPATH, '//*[@id="btnPlus"]/a/h2').click() #신용정보원
                time.sleep(0.5)
                browser.switch_to.default_content() #메인화면으로 전환
                time.sleep(0.5)
                browser.switch_to.frame('windowContainer1_subWindow0_iframe') #frame 전환(고객명, 조회)
                time.sleep(0.5)
                browser.switch_to.frame('mabnf001pvw13_iframe') #frame 전환(새로불러오기 or 기존불러오기)
                time.sleep(0.5)
                browser.find_element(By.XPATH, '//*[@id="btnNewDataPrcs"]/a/h2').click() #신용정보원 새로불러오기
                time.sleep(0.5)
                mainFrameInit('보장분석',currentTime)

                ###############################
                # 변액 기준가/수익률 조회
                ###############################
                menuChoice('genMenuDepth1_2_menuNm','genMenuDepth2_5_genMenuDepth3_0_menuNm') #메뉴선택
                browser.find_element(By.XPATH, '//*[@id="btnSearch"]/a').click() #조회
                time.sleep(0.5)
                mainFrameInit('변액',currentTime)

                ###############################
                # 가입설계 조회
                ###############################
                menuChoice('genMenuDepth1_3_menuNm','genMenuDepth2_0_genMenuDepth3_15_menuNm') #메뉴선택
                browser.find_element(By.ID, 'ibxCustName').send_keys('채경훈') #고객명 입력
                pyautogui.press('enter') #엔터키(조회)
                time.sleep(2)
                mainFrameInit('가입설계',currentTime)

                ###############################
                # 속보 조회
                ###############################
                menuChoice('genMenuDepth1_4_menuNm','genMenuDepth2_0_genMenuDepth3_0_menuNm') #메뉴선택
                time.sleep(1)
                mainFrameInit('속보',currentTime)

                ###############################
                # 증권번호별 계약내용 조회
                ###############################
                menuChoice('genMenuDepth1_4_menuNm','genMenuDepth2_3_genMenuDepth3_2_menuNm') #메뉴선택
                time.sleep(0.5)
                browser.find_element(By.ID, 'ibxPolyNo').send_keys('522725186') #증번입력
                pyautogui.press('enter') #엔터키(조회)
                time.sleep(0.5)    
                mainFrameInit('계약',currentTime)

                ###############################
                # 제지급 처리관리 조회
                ###############################
                menuChoice('genMenuDepth1_4_menuNm','genMenuDepth2_4_genMenuDepth3_0_menuNm') #메뉴선택
                time.sleep(0.5)
                browser.find_element(By.ID, 'ibxPolyNo').send_keys('522725186') #증번입력
                pyautogui.press('enter') #엔터키(조회)
                time.sleep(1)    
                pyautogui.press('enter') #엔터키(조회)
                time.sleep(0.5)
                mainFrameInit('제지급',currentTime)
        except Exception as e:
            print("### 예외발생: " + str(e) + '_' + str(currentTime))
            failcount = failcount+1
            print("### failcount : " + str(failcount))
