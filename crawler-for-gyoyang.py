from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys

import os
import pandas as pd
from openpyxl import load_workbook

import time

SLEEP_TIME = 2

chrome_options = Options()

driver = webdriver.Chrome("driver/chromedriver", chrome_options=chrome_options)
driver.implicitly_wait(3)
driver.get('https://portal.hanyang.ac.kr/sso/lgin.do')

driver.find_element_by_class_name("goPortal").click()
driver.implicitly_wait(3)

hanyang_id = ""
hanyang_pw = ""

# LOGIN
driver.find_element_by_id("userId").send_keys(hanyang_id)
time.sleep(0.5)
driver.find_element_by_id("password").send_keys(hanyang_pw)
time.sleep(0.5)
driver.find_element_by_xpath("""//*[@id="hyinContents"]/div[1]/form/div/fieldset/p[3]/a""").click()

driver.implicitly_wait(3)
driver.find_element_by_id("btn_cancel").click()

yeongyeok_list_2019 = ["A1", "C1", "C3", "C4", "C5", "C6", "C7", "E1", "E2", "E3", "E4", "E5", "BA"]

yeongyeok_list_2020 = ["G1", "G2", "G3", "G4", "G5", "G6", "G7", "G8"]

year_list = ["2019", "2020"]

# Move to Target Page
driver.get("https://portal.hanyang.ac.kr/port.do#!UDMyMDA0NiRAXmhha3NhLyRAXiRAXk0zMTkyNjIkQF7qsJXsnZjtj4nqsIDqsrDqs7zqsoDsg4kkQF5NMDAzNzk0JEBeZjZjZmU2NjA0MzUyMDBkZWNkMjMxZmI4OTVhZDg0NmI1MTU3YmU4YTYyZTUwYTVmZWIxY2U1MGJkMWFmMmI4Nw==")
driver.implicitly_wait(3)

tot_raw_data = []

for year in year_list:
    raw_data = []

    year_select = Select(driver.find_element_by_id("cbYear"))
    year_select.select_by_value(year)

    term_select = Select(driver.find_element_by_id("cbTerm"))
    term_select.select_by_value("10") # 10 - 1학기, 20 - 2학기

    gupgSeq_select = Select(driver.find_element_by_id("cbGupgSeq"))
    gupgSeq_select.select_by_value("01") # 01 - 중간강의평가

    campus_select = Select(driver.find_element_by_id("cbCampus"))
    campus_select.select_by_value("Y")

    search_condition_select = Select(driver.find_element_by_id("cbSearch"))
    search_condition_select.select_by_value("3") # 2 - 전공, 3 - 교양

    yeongyeok_select = Select(driver.find_element_by_id("cbYeongyeok"))

    if year == "2019":
        yeongyeok_list = yeongyeok_list_2019
    else:
        yeongyeok_list = yeongyeok_list_2020

    time.sleep(0.5)

    for yeongyeok in yeongyeok_list:
        yeongyeok_select.select_by_value(yeongyeok)
        # driver.find_element_by_id("btn_search2").click()
        yeongyeok_click_element = driver.find_element_by_id("btn_search3")

        driver.execute_script("arguments[0].click();", yeongyeok_click_element)

        driver.implicitly_wait(3)
        time.sleep(SLEEP_TIME - 0.5)

        # Crawl data
        table = driver.find_element_by_id("gdMain")
        tbody = table.find_element_by_tag_name("tbody")
        trs = tbody.find_elements_by_tag_name("tr")

        for tr in trs:
            # tr.find_element_by_id("jonghapScore").click()
            click_element = tr.find_element_by_id("jonghapScore")
            driver.execute_script("arguments[0].click();", click_element)

            time.sleep(SLEEP_TIME - 1)

            raw_popup_table = driver.find_element_by_id("suce0100_pop_Form") # type(popup_table.text) = str

            popup_table = raw_popup_table.text.split('\n')[1:]

            if len(popup_table) > 1:
                haksu_num = tr.find_element_by_id("haksuNo").text
                isu_num = tr.find_element_by_id("isuGbNm").text
                isu_grade = tr.find_element_by_id("isuGradeNm").text
                gwamok_name = tr.find_element_by_id("gwamokNm").text
                gyogangsa_name = tr.find_element_by_id("gyogangsaNm").text

                for element in popup_table:
                    s = element.split(' ')[1:]
                    question = ' '.join(s[:-1])
                    answer = float(s[-1])
                    raw_data.append([haksu_num, isu_num, isu_grade, gwamok_name, gyogangsa_name, question, answer])

            driver.find_element_by_tag_name("body").send_keys(Keys.ESCAPE)
            time.sleep(0.5)

        print(f"Fin {year} {yeongyeok}")
        time.sleep(0.5)

    tot_raw_data.append(raw_data)


dir_path = "some_path/"
file_name = "hanyang_midterm_survey_gyoyang.xlsx"

if not os.path.isdir(dir_path):
    os.mkdir(dir_path)

writer = pd.ExcelWriter(path=dir_path + file_name, engine="openpyxl")

df1 = pd.DataFrame(tot_raw_data[0])
df1.to_excel(writer, sheet_name="2019", index=False, index_label=False)

df2 = pd.DataFrame(tot_raw_data[1])
df2.to_excel(writer, sheet_name="2020", index=False, index_label=False)

writer.save()
writer.close()

print("Fin!")