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
year_dic = {'2018': 0, '2019': 1, '2020':2, '2021': 3, '2022':4}
mon_dic = {'01': 0, '02': 1, '03': 2, '04': 3, '05': 4,
           '06': 5, '07': 6, '08': 7, '09': 8, '10': 9, '11': 10, '12': 11}

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
    ''' >>> 업무일자 선택'''
    '''===================================='''
    # >>>> 업무일자선택
    driver.find_element_by_xpath('/html/body/form/table/tbody/tr[3]/td[2]/table[1]/tbody/tr/td/table[2]/tbody/tr/td[2]/input').click()    #날짜선택
    driver.switch_to_alert().accept()       #작성주의사항 확인

    from selenium.webdriver.support.ui import Select
    work_rec_date = '2021-08-02'
    work_rec_date_li = work_rec_date.split('-')

    year = year_dic[work_rec_date_li[0]]
    mon = mon_dic[work_rec_date_li[1]]

    # >>>>> 연도 선택 포맷팅 및 선택
    form_year_xpath = '//*[@id="ui-datepicker-div"]/div[1]/div/select[1]'
    #form_year_xpath = form_year_xpath.format(year)
    select = Select(driver.find_element_by_xpath(form_year_xpath))  # 연도 선택
    print(year)
    select.select_by_index(year)

    # >>>>> 달 선택 포맷팅 및 선택
    form_mon_xpath = '//*[@id="ui-datepicker-div"]/div[1]/div/select[2]'
    #form_mon_xpath = form_mon_xpath.format(mon)
    select = Select(driver.find_element_by_xpath(form_mon_xpath))  # 달 선택
    print(mon)
    select.select_by_index(mon)

    # >>>> 달력 일자 선택
    week_no, day = gt_calendar(work_rec_date)                                                           # 해당일자의 월주번호와 요일산출하여 xpath 인덱스 반환
    if day == 1 or day == 7: # 일요일과 토요일 스킵

        day_xpath_form = '//*[@id="ui-datepicker-div"]/table/tbody/tr[{0}]/td[{1}]'
        day_xpath_form = day_xpath_form.format(week_no, day)                                                # 반환된 xpath 인덱스로 xpath 포맷팅
        print(day_xpath_form)
        driver.find_element_by_xpath(day_xpath_form).click()    #일자선택

        driver.switch_to_default_content()

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
* @Function Name : rec_work_hist
************************************************'''
def rec_work_hist(driver, date, work_desc):
    return


'''************************************************
* @Function Name : gt_calendar
************************************************'''
from datetime import datetime, timedelta
import calendar
def gt_calendar(date):
    date_split = date.split('-')

    # date_split[0] : year
    # date_split[1] : month
    # date_split[2] : day
    week_no = gt_week_no(int(date_split[0]), int(date_split[1]), int(date_split[2]))
    day = calendar.weekday(int(date_split[0]), int(date_split[1]), int(date_split[2]))

    if day > 5:
        day = 1
    else:
        day = day + 2

    return week_no, day


'''************************************************
* @Function Name : gt_date
************************************************'''
def gt_date(y, m, d):
    '''y: year(4 digits)
     m: month(2 digits)
     d: day(2 digits'''
    s = f'{y:04d}-{m:02d}-{d:02d}'
    return datetime.strptime(s, '%Y-%m-%d')


'''************************************************
* @Function Name : gt_week_no
************************************************'''
def gt_week_no(y, m, d):
    target = gt_date(y, m, d)
    firstday = target.replace(day=1)
    if firstday.weekday() == 6:
        origin = firstday
    elif firstday.weekday() < 3:
        origin = firstday - timedelta(days=firstday.weekday() + 1)
    else:
        origin = firstday + timedelta(days=6-firstday.weekday())
    return (target - origin).days // 7 + 1


'''************************************************
* @Function Name : main
************************************************'''
def main():
    #gt_calendar('2020-11-16')
    acc_web()

    return

if __name__ == '__main__':
    main()
