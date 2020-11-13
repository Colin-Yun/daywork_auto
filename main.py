from selenium import webdriver

chrome_path = "C:\Program Files\Google\Chrome\Application\chromedriver\chromedriver.exe"


'''************************************************
* @Function Name : set_chrome_opt
************************************************'''
def set_chrome_opt(opt1):
    opt = webdriver.ChromeOptions()

    # OPT1 : VISIBLE or INVISBLE
    if opt1 == 1:           #INVISIBLE
        opt.add_argument('headless')
        opt.add_argument('disable-gpu')
    else:
        pass

    return opt


'''************************************************
* @Function Name : acc_web
************************************************'''
def acc_web():
    import time
    _id = 'bchyun'
    _pw = 's060913!'

    driver = webdriver.Chrome(chrome_path, chrome_options=set_chrome_opt(0))

    url = 'http://kpm.kyungshin.co.kr/usr/'
    driver.get(url)

    driver.switch_to_frame('main')
    driver.find_element_by_name('id').send_keys(_id)
    driver.find_element_by_name('pw').send_keys(_pw)
    driver.find_element_by_xpath('//*[@id="3"]/tbody/tr[1]/td[4]/a/img').click()

    driver.switch_to_alert().accept()       #크롬 adobe alert 확인
    driver.switch_to_frame('main')

    driver.find_element_by_xpath('/html/body/table/tbody/tr/td/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[8]/a/img').click()       #업무활동클릭
    driver.implicitly_wait(5)
    time.sleep(5)

    driver.switch_to_frame('Frame')
    driver.switch_to_frame('detailFrm3')

    driver.find_element_by_xpath('/html/body/table/tbody/tr/td[2]/table[2]/tbody/tr/td[2]/a[2]/img').click()    #등록클릭
    driver.switch_to_window(driver.window_handles[1])

    '''===================================='''
    ''' >>> 업무1 입력 '''
    '''===================================='''
    driver.find_element_by_xpath('//*[@id="dailyWorkTable_p"]/tbody/tr[3]/td[3]/img').click() #업무구분 순번1 -> [선택] 선택
    driver.switch_to_frame('layerFrm')  #업무구분설정팜업 Frame 이동
    driver.find_element_by_xpath('/html/body/table/tbody/tr[3]/td[2]/form/table[3]/tbody/tr[2]/td[2]/table/tbody/tr[2]/td[2]/input').click() #일상행정업무 선택
    driver.find_element_by_xpath('/html/body/table/tbody/tr[3]/td[2]/form/table[5]/tbody/tr[2]/td[2]/table/tbody/tr[3]/td/table/tbody/tr[2]/td[2]/table/tbody/tr[1]/td[2]/input').click() #소프트웨어 설계/검증 선택
    driver.find_element_by_xpath('/html/body/table/tbody/tr[3]/td[2]/form/table[6]/tbody/tr/td/a[1]/img').click()   #확인

    work_history ="""- SU2b 사양 설계 등등\n- KY 리프로그램 테스트 등등\n- 업무자동화툴 개발"""

    driver.switch_to_default_content() #업무구분설정팜업 Frame 재이동
    driver.find_element_by_xpath('//*[@id="dailyWorkTable_p"]/tbody/tr[4]/td[2]/textarea').send_keys(work_history) # 업무내용기입


    '''===================================='''
    ''' >>> 업무2 입력 '''
    '''===================================='''
    driver.find_element_by_xpath('//*[@id="dailyWorkTable_p"]/tbody/tr[5]/td[3]/img').click() #업무구분 순번2 -> [선택] 선택
    driver.switch_to_frame('layerFrm')  #업무구분설정팜업 Frame 이동
    driver.find_element_by_xpath('/html/body/table/tbody/tr[3]/td[2]/form/table[3]/tbody/tr[2]/td[2]/table/tbody/tr[2]/td[2]/input').click() #일상행정업무 선택
    driver.find_element_by_xpath('/html/body/table/tbody/tr[3]/td[2]/form/table[5]/tbody/tr[2]/td[2]/table/tbody/tr[3]/td/table/tbody/tr[2]/td[2]/table/tbody/tr[1]/td[2]/input').click() #소프트웨어 설계/검증 선택
    driver.find_element_by_xpath('/html/body/table/tbody/tr[3]/td[2]/form/table[6]/tbody/tr/td/a[1]/img').click()   #확인

    work_history ="""- RJ PE 사양 설계 등등\n- KY 리프로그램 테스트 등등\n- 업무자동화툴 개발"""

    driver.switch_to_default_content() #업무구분설정팜업 Frame 재이동
    driver.find_element_by_xpath('//*[@id="dailyWorkTable_p"]/tbody/tr[6]/td[2]/textarea').send_keys(work_history) # 업무내용기입


    time.sleep(10)
    return



'''************************************************
* @Function Name : main
************************************************'''
def main():
    acc_web()

    return

if __name__ == '__main__':
    main()
