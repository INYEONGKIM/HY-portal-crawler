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


hakgwa_list_2019 = ["Y0000415", "Y0000416", "Y0000417", "Y0000981", "Y0001155", "Y0000389", "Y0001083", "Y0001084",
                    "Y0000386", "Y0000387", "Y0000388", "Y0000982", "Y0000983", "Y0001030", "Y0001251", "Y0001252",
                    "Y0001253", "Y0001254", "Y0001166", "Y0001167", "Y0001168", "Y0001169", "Y0001172", "Y0001171",
                    "Y0001170", "Y0000596", "Y0001174", "Y0001175", "Y0001176", "Y0001177", "Y0001178", "Y0001179",
                    "Y0000452", "Y0000467", "Y0000453", "Y0000454", "Y0000478", "Y0001180", "Y0000455", "Y0000475",
                    "Y0001181", "Y0001066", "Y0000488", "Y0000489", "Y0001244", "Y0000517", "Y0000519", "Y0000968",
                    "Y0000984", "Y0000990", "Y0000991", "Y0000992", "Y0000993", "Y0000994", "Y0001246", "Y0001247",
                    "Y0000355", "Y0000356", "Y0001111", "Y0001112", "Y0001113", "Y0001110", "Y0000358", "Y0001190",
                    "Y0000709"]

hakgwa_list_2020 = ["Y0000415", "Y0000416", "Y0000417", "Y0001155", "Y0000389", "Y0001083", "Y0001084", "Y0000386",
                    "Y0000387", "Y0000388", "Y0000982", "Y0000983", "Y0001030", "Y0001251", "Y0001252", "Y0001253",
                    "Y0001254", "Y0001166", "Y0001167", "Y0001169", "Y0001172", "Y0001171", "Y0001170", "Y0000596",
                    "Y0001174", "Y0001175", "Y0001176", "Y0001177", "Y0001178", "Y0001179", "Y0000452", "Y0000467",
                    "Y0000453", "Y0000454", "Y0001180", "Y0000455", "Y0001181", "Y0001066", "Y0000488", "Y0000489",
                    "Y0001244", "Y0000517", "Y0000519", "Y0000968", "Y0000984", "Y0000990", "Y0000991", "Y0000992",
                    "Y0000993", "Y0000994", "Y0001246", "Y0001247", "Y0001111", "Y0001112", "Y0001113", "Y0001110",
                    "Y0000358", "Y0001190", "Y0000709"]

year_list = ["2019", "2020"]

# Move to Target Page
driver.get("https://portal.hanyang.ac.kr/port.do#!UDMyMDA0NiRAXmhha3NhLyRAXiRAXk0zMTkyNjIkQF7qsJXsnZjtj4nqsIDqsrD"
           "qs7zqsoDsg4kkQF5NMDAzNzk0JEBeZjZjZmU2NjA0MzUyMDBkZWNkMjMxZmI4OTVhZDg0NmI1MTU3YmU4YTYyZTUwYTVmZWIxY2U1"
           "MGJkMWFmMmI4Nw==")
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
    search_condition_select.select_by_value("2")

    hakgwa_select = Select(driver.find_element_by_id("cbHakgwa"))

    if year == "2019":
        hakgwa_list = hakgwa_list_2019
    else:
        hakgwa_list = hakgwa_list_2020

    for hakgwa in hakgwa_list:
        hakgwa_select.select_by_value(hakgwa)
        # driver.find_element_by_id("btn_search2").click()
        hakgwa_click_element = driver.find_element_by_id("btn_search2")

        driver.execute_script("arguments[0].click();", hakgwa_click_element)

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
                daehak_name = tr.find_element_by_id("slgDaehakNm").text
                haksu_num = tr.find_element_by_id("haksuNo").text
                isu_num = tr.find_element_by_id("isuGbNm").text
                isu_grade = tr.find_element_by_id("isuGradeNm").text
                gwamok_name = tr.find_element_by_id("gwamokNm").text
                gyogangsa_name = tr.find_element_by_id("gyogangsaNm").text

                for element in popup_table:
                    s = element.split(' ')[1:]
                    question = ' '.join(s[:-1])
                    answer = float(s[-1])
                    raw_data.append([daehak_name, haksu_num, isu_num, isu_grade, gwamok_name, gyogangsa_name, question, answer])

            driver.find_element_by_tag_name("body").send_keys(Keys.ESCAPE)
            # driver.implicitly_wait(3)
            time.sleep(0.5)

        print(f"Fin {year} {hakgwa}")
        time.sleep(0.5)

    tot_raw_data.append(raw_data)


dir_path = "some_path/"
file_name = "hanyang_midterm_survey.xlsx"

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