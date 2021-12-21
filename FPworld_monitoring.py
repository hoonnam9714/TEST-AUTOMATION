import pyautogui
import pyperclip
import time
import datetime
import smtplib
from account import *
from selenium import webdriver
from selenium.webdriver.common.by import By
from FPworld_account import *
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from email.message import EmailMessage

###############################
# 메인frame 전환 및 초기화 함수
###############################
def mainFrameInit(arg1,arg2): #agr1: 업무명, arg2: 현재시간
    browser.switch_to.default_content() #메인화면으로 frame전환
    time.sleep(1)
    screenshot('windowContainer1_subWindow0_body',arg1,arg2) # 스크린샷
    time.sleep(1)
    #메인로고 더블클릭(frame 초기화)
    browser.find_element(By.XPATH, '//*[@id="btnMainLogo"]').click()
    browser.find_element(By.XPATH, '//*[@id="btnMainLogo"]').click()

###############################
# 스크린샷 함수
###############################v
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
def kakaoSend(arg1):
    browser.quit() #브라우저 종료
    time.sleep(0.5)
    pyautogui.moveTo(30,457) #카카오톡 메시지 박스
    time.sleep(0.5)
    pyautogui.click()
    time.sleep(0.5)
    pyperclip.copy("FPworld 모니터링 실패_" + arg1 + "_" + str(currentTime)) #클립보드에 복사
    pyautogui.hotkey("command","v") #붙여넣기
    time.sleep(0.5)
    pyautogui.moveTo(346,482) #카카오톡 전송
    time.sleep(0.5)
    pyautogui.click()
    time.sleep(0.5)

###############################
# 메일 전송
###############################
def mailSend(arg1):
    browser.quit() #브라우저 종료
    msg = EmailMessage()
    msg["Subject"] = "[헬스체크]FP월드 시스템 모니터링" #제목
    msg["From"] = "dept2020000025@hanwha.com" #보내는 사람
    #여러명에게 메일을 보낼 때
    msg["To"] = "hoonnam9714@gmail.com"
    # msg["To"] = "hoonnam9714@hanwha.com, kevinkim@hanwha.com"

    msg.set_content(" FP월드 가입설계 시스템 모니터링 중 장애가 발생 했습니다. 담당자는 FP월드 시스템 점검 바랍니다. \n (해당 메일은 시스템 모니터링 프로그램에서 자동으로 발송됩니다.)\n 감사합니다.") #본문

    # msg.add_attachment()
    # with open(filename+".pdf", "rb") as f:
    #     msg.add_attachment(f.read(), maintype = "application", subtype= "pdf", filename = f.name) #파일을 읽어와서 타입에 맞게 설정(PDF)

    with smtplib.SMTP("smtp.gmail.com",587) as smtp:
        smtp.ehlo() #연결이 잘 수립되는지 확인
        smtp.starttls() #모든 내용이 암호화되어 전송
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD) #로그인
        smtp.send_message(msg)

###############################
# 페이지 로딩 대기
###############################
def pageWait(arg1):
    elem = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, arg1))) #10초동안 해당 엘리먼트가 존재하는지 대기

###############################
# 메인
###############################
if __name__ == "__main__":
    # 사용자 입력
    sabun = pyautogui.prompt("로그인 사번을 입력해주세요.","입력") #사용자 입력 사번
    password = pyautogui.password("로그인 비밀번호를 입력해주세요.","입력") #사용자 입력 패스워드
    
    browser = webdriver.Chrome(executable_path='/Users/chaekyunghoon/desktop/PythonWorkSpace/chromedriver')
    url = "https://hmp.hanwhalife.com/online/fp" #HMC QA 테스트 URL
    browser.get(url) #FP월드 URL
    browser.maximize_window()
    time.sleep(0.5)
    browser.find_element(By.ID, 'ibxUserId').send_keys(sabun) 
    time.sleep(0.5)
    browser.find_element(By.ID, 'ibxPassword').send_keys(password)
    time.sleep(0.5)
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
    gubun = "전체"
    while True:
        try:
            current_datetime = datetime.datetime.now()
            dateformat = "%Y%m%d%H%M%S"
            currentTime = current_datetime.strftime(dateformat) #현재시간 ex)20211103153201

            ###############################
            # 예외발생시 종료 및 카카오톡 발송
            ###############################
            if failcount > 5: 
                # kakaoSend(gubun) #카카오톡 발송
                mailSend(gubun)
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
                gubun = "고객"
                browser.find_element(By.XPATH, '//*[@id="ipt_srchAdmnCustNo"]').send_keys('1') #고객번호
                time.sleep(2)
                browser.find_element(By.XPATH, '//*[@id="btnSrch"]/a').click() #조회
                mainFrameInit('고객',currentTime)

                ###############################
                # 보장분석 조회
                ###############################
                menuChoice('genMenuDepth1_2_menuNm','genMenuDepth2_0_genMenuDepth3_0_menuNm') #메뉴선택
                gubun = "보장분석"
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
                # menuChoice('genMenuDepth1_2_menuNm','genMenuDepth2_5_genMenuDepth3_0_menuNm') #메뉴선택
                # gubun = "변액"
                # time.sleep(0.5)
                # browser.implicitly_wait(10)
                # pageWait('//*[@id="btnSearch"]/a') #메인 조회 페이지 로딩 대기
                # time.sleep(2)
                # browser.find_element(By.XPATH, '//*[@id="btnSearch"]/a').click() #조회
                # time.sleep(0.5)
                # mainFrameInit('변액',currentTime)

                ###############################
                # 가입설계 조회
                ###############################
                menuChoice('genMenuDepth1_3_menuNm','genMenuDepth2_0_genMenuDepth3_15_menuNm') #메뉴선택
                gubun = "가입설계"
                time.sleep(2)
                browser.find_element(By.ID, 'ibxCustName').send_keys('채경훈') #고객명 입력
                time.sleep(2)
                pyautogui.press('enter') #엔터키(조회)
                time.sleep(2)
                mainFrameInit('가입설계',currentTime)

                ###############################
                # 속보 조회
                ###############################
                menuChoice('genMenuDepth1_4_menuNm','genMenuDepth2_0_genMenuDepth3_0_menuNm') #메뉴선택
                gubun = "속보"
                time.sleep(2)
                mainFrameInit('속보',currentTime)

                ###############################
                # 증권번호별 계약내용 조회
                ###############################
                menuChoice('genMenuDepth1_4_menuNm','genMenuDepth2_3_genMenuDepth3_2_menuNm') #메뉴선택
                gubun = "계약내용"
                time.sleep(2)
                browser.find_element(By.ID, 'ibxPolyNo').send_keys('522725186') #증번입력
                time.sleep(2)
                pyautogui.press('enter') #엔터키(조회)
                time.sleep(2)
                mainFrameInit('계약',currentTime)

                ###############################
                # 제지급 처리관리 조회
                ###############################
                menuChoice('genMenuDepth1_4_menuNm','genMenuDepth2_4_genMenuDepth3_0_menuNm') #메뉴선택
                gubun = "제지급"
                time.sleep(2)
                browser.find_element(By.ID, 'ibxPolyNo').send_keys('522725186') #증번입력
                time.sleep(2)
                pyautogui.press('enter') #엔터키(조회)
                time.sleep(1)    
                pyautogui.press('enter') #엔터키(조회)
                time.sleep(2)
                mainFrameInit('제지급',currentTime)
        except Exception as e:
            print("### 예외발생: " + str(e) + '_' + str(currentTime))
            failcount = failcount+1
            print("### failcount : " + str(failcount))
