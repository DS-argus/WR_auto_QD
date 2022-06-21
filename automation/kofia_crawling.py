import pandas as pd

from automation.SSLpatch import no_ssl_verification
from automation.kofia_codes import code

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import time
from datetime import date


def SEIBRO_crawler():
    # selenium option settings
    options = webdriver.ChromeOptions()
    options.add_argument('headless')
    options.add_argument('window-size=1920x1080')
    options.add_argument("disable-gpu")

    # escaping SSL error
    with no_ssl_verification():
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),
                                  options=options)

        # url 자체가 최근 일자로 조회되는 듯 함
        url = "https://seibro.or.kr/websquare/control.jsp?w2xPath=/IPORTAL/user/derivCombi/BIP_CNTS07003V.xml&menuNo=193#"
        driver.get(url)

        driver.maximize_window()

        comp_list = []
        num_issued_list = []
        sum_issued_list = []

        row1 = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, "// *[ @ id = 'row122']")))
        row2 = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, "// *[ @ id = 'row125']")))
        driver.implicitly_wait(5)

        #넓게보기 클릭
        button = driver.find_element(by=By.XPATH, value="//*[@id='btn_wide_img']")
        button.click()

        for i in range(len(row1)):

            comp = row1[i].find_elements(by=By.CSS_SELECTOR,
                                         value="td")[0].find_element(by=By.CSS_SELECTOR,
                                                                     value="nobr")

            comp_name = comp.text

            # 이름 다를 경우 변경
            if comp_name == "케이비증권":
                comp_name = "KB증권"
            elif comp_name == "아이비케이투자증권":
                comp_name = "IBK투자증권"
            elif comp_name == "합계":
                break

            if comp_name in comp_list:
               continue

            if comp_name not in code.keys():
               continue

            num_issued = row1[i].find_elements(by=By.CSS_SELECTOR,
                                                value="td")[26].find_element(by=By.CSS_SELECTOR,
                                                                            value="nobr")

            sum_issued = row2[i].find_elements(by=By.CSS_SELECTOR,
                                                value="td")[25].find_element(by=By.CSS_SELECTOR,
                                                                            value="nobr")

            comp_list.append(comp_name)
            num_issued_list.append(num_issued.text)
            sum_issued_list.append(sum_issued.text)

    driver.quit()

    df_issue = pd.DataFrame(index=comp_list, columns=['발행종목수', '발행잔액'])
    df_issue['발행종목수'] = num_issued_list
    df_issue['발행잔액'] = sum_issued_list

    return df_issue


def NCR_crawler():

    # selenium option settings
    options = webdriver.ChromeOptions()
    options.add_argument('headless')
    options.add_argument('window-size=1920x1080')
    options.add_argument("disable-gpu")

    # escaping SSL error
    with no_ssl_verification():
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),
                                  options=options)

        # url 자체가 최근 분기말로 조회되는 듯 함
        url = "https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/compann/DISNetCapRate.xml&divisionId=MDIS02003013000000&serviceId=SDIS02003013000#!"
        driver.get(url)

        driver.maximize_window()
        element = WebDriverWait(driver, 15).until(EC.presence_of_all_elements_located((By.XPATH, "// *[ @ id = 'gRow1']")))
        
        comp_list = []
        ncr_list = []
        verical_ordinate = 160

        for i in range(20):  # 스크롤 다운 횟수(넉넉하게, 아래에서 조건만족하면 끝낼 것)
            for j in range(12):
                comp = element[j].find_elements(by=By.CSS_SELECTOR,
                                                value="td")[1].find_element(by=By.CSS_SELECTOR,
                                                                            value="nobr")

                comp_name = comp.text

                # 이름 다를 경우 변경
                if comp_name == "아이비케이투자증권":
                    comp_name = "IBK투자증권"
                elif comp_name == "현대차증권주식회사":
                    comp_name = "현대차증권"



               # 이미 크롤링한 회사면 다음줄로 넘어감
                if comp_name in comp_list:
                    continue

                if comp_name not in code.keys():
                    continue

                ncr = element[j].find_elements(by=By.CSS_SELECTOR,
                                               value="td")[6].find_element(by=By.CSS_SELECTOR,
                                                                           value="nobr")

                comp_list.append(comp_name)
                ncr_list.append(ncr.text)

            if len(comp_list) == len(code.keys()):
                break

            scroll = driver.find_element(by=By.CSS_SELECTOR, value="#grdMain_scrollY_div")
            driver.execute_script("arguments[0].scrollTop = arguments[1]", scroll, verical_ordinate)
            verical_ordinate += 160
            time.sleep(1)

        driver.quit()

        df_ncr = pd.DataFrame(index=comp_list, columns=['순자본비율'])
        df_ncr['순자본비율'] = ncr_list

        return df_ncr


def FS_crawler(last_quarter):

    # 직전분기말 날짜 입력
    last = date.strftime(last_quarter, "%Y%m%d")

    comps = ",".join(code.values())

    url_main = "https://dis.kofia.or.kr/websquare/index.jsp?w2xPath=/wq/cmpann/DISCompCmpSrchList.xml&divisionId=MDIS02004000000000&serviceId=&"

    url_date = "standardDt=" + last

    url_comps = "&uCompList=" + comps

    url_items = "&uCompNm=/v8ASQBCAEvSLMeQyZ2tjAAsAEsAQsmdrYwALABOAEjSLMeQyZ2tjAAsrVC89MmdrYwALLMAwuDJ%0Ana2MACy6VLmsziDJna2MACy7+LeYxdDBS8mdrYwALMC8wTHJna2MACzC4MYByZ2tjAAswuDVXK4IxzXSLMeQACzHIMVI0MDJna2MACzQpMbAyZ2tjAAs1ViwmK4IxzXSLMeQACzVWMd00izHkMmdrYwALNVcrW3SLMeQyZ2tjAAs1VzWVNIsx5DJna2MACzWBLMAzCjJna2M&uItemList=%27000040%27,%27000038%27&uItemNm=/v/HkLz4zR2sxAAsvYDMRM0drMQ%3D&targetGb=A"

    # selenium option settings
    options = webdriver.ChromeOptions()
    options.add_argument('headless')
    options.add_argument('window-size=1920x1080')
    options.add_argument("disable-gpu")

    # escaping SSL error
    with no_ssl_verification():
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),
                                  options=options)

        url = url_main + url_date + url_comps + url_items
        driver.get(url)

        driver.maximize_window()

        element = WebDriverWait(driver, 15).until(EC.presence_of_all_elements_located((By.XPATH, "//*[@id='row2']")))

        comp_list = []
        lia_list = []
        equity_list = []
        verical_ordinate = 20

        for i in range(15):     # 스크롤 다운 횟수(넉넉하게, 아래에서 조건만족하면 끝낼 것)
            for j in range(5):
                comp = element[j].find_elements(by=By.CSS_SELECTOR,
                                         value="td")[0].find_element(by=By.CSS_SELECTOR,
                                         value="nobr")

                # 이미 크롤링한 회사면 다음줄로 넘어감
                if comp.text in comp_list:
                    continue
    
                liabilities = element[j].find_elements(by=By.CSS_SELECTOR,
                                         value="td")[1].find_element(by=By.CSS_SELECTOR,
                                         value="nobr")
    
                equities = element[j].find_elements(by=By.CSS_SELECTOR,
                                         value="td")[2].find_element(by=By.CSS_SELECTOR,
                                         value="nobr")
    
                comp_list.append(comp.text)
                lia_list.append(liabilities.text)
                equity_list.append(equities.text)

            if len(comp_list) == len(code.keys()):
                break

            scroll = driver.find_element(by=By.CSS_SELECTOR, value="#grdCmpr_scrollY_div")
            driver.execute_script("arguments[0].scrollTop = arguments[1]", scroll, verical_ordinate)
            verical_ordinate += 20
            time.sleep(1)

        driver.quit()

        df_fs = pd.DataFrame(index=comp_list, columns=['부채총계', '자본총계'])
        df_fs['부채총계'] = lia_list
        df_fs['자본총계'] = equity_list

        return df_fs


if __name__ == "__main__":

    print(FS_crawler(date(2022, 3, 31)))

    print(NCR_crawler())

    print(SEIBRO_crawler())
