import requests
import pandas
import threading
import socket
import smtplib
import datetime
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import re
import time

URL = "https://www.mvdis.gov.tw/m3-emv-trn/exm/locations#anchor"
CANCELLED_URL = "https://www.mvdis.gov.tw/m3-emv-trn/exm/query"
CHROME_DRIVER_LOCATION = "YOUR_CHROME_DRIVER"
# TEST_DATE = "1110729"
MY_EMAIL = "YOUR_EMAIL_ADDRESS"
PASSWORD = "YOUR_EMAIL_PASSWORD"
recipient = "RECIPIENT_EMAIL_ADDRESS"

# You may change the test date from today date
DAYS_FROM_NOW = 2
time_in_eng = ""
date_to_choose = ""

# Set Test Date to search under 民國 Calender
test_date = datetime.datetime.now()
test_year_minguo = int(test_date.strftime("%Y")) - 1911
print(test_year_minguo)
test_month = int(test_date.strftime("%m"))
test_day = int(test_date.strftime("%d")) + DAYS_FROM_NOW
test_date_in_minguo = f"{test_year_minguo}{test_month:02d}{test_day:02d}"
print(test_date_in_minguo)
test_date_calender_year = test_date.strftime("%Y%m%d-%H:%M")
# print(test_date_in_minguo)

def check_available_date():
    global time_in_eng, date_to_choose
    try:
        # Connect
        chrome_options = Options()
        chrome_options.add_experimental_option("detach", True)
        # chrome_options.add_argument("enable-automation")
        service_obj = Service(CHROME_DRIVER_LOCATION)
        driver = webdriver.Chrome(service=service_obj, options=chrome_options)
        driver.get(URL)

        # Select type of test
        type_of_test_select = Select(driver.find_element(By.ID, "licenseTypeCode"))
        type_of_test_select.select_by_value("A")

        # Input Date
        driver.find_element(By.ID, "expectExamDateStr").send_keys(test_date_in_minguo)

        # Select Region
        place_of_test_select_1 = Select(driver.find_element(By.ID, "dmvNoLv1"))
        place_of_test_select_1.select_by_value("40")

        # Select DMV Station
        place_of_test_select_2 = Select(driver.find_element(By.ID, "dmvNo"))
        place_of_test_select_2.select_by_value("40")

        # Click Search Button
        driver.find_element(By.XPATH, '//*[@id="form1"]/div/a/img').click()

        # Get Content
        driver.get("https://www.mvdis.gov.tw/m3-emv-trn/exm/locations#anchor")
        search_result = driver.find_elements(By.CLASS_NAME, "align_c")
        result_date = ""
        result_time = ""
        seat_available = 0

        # Find first available date
        for i in range(0,len(search_result)):
            # print(f"{i} - {search_result[i].text}")
            if search_result[i].text == ".. 報名 SignUp":
                result_date = search_result[i - 3]
                result_time = search_result[i - 2]
                seat_available = search_result[i - 1]
                row_num = round((i-7)/4)
                date_to_choose = f'//*[@id="trnTable"]/tbody/tr[{row_num}]/td[4]/a'
                break

        # Extract available month and day
        # print(result_date.text)
        result_date = result_date.text.split()[0].replace("年"," ").replace("月"," ").replace("日"," ").split()
        available_month = int(result_date[1])
        available_day = int(result_date[2])
        # print(available_month)
        # print(available_day)

        # Check AM or PM
        available_time = result_time.text.replace("場次", " ").split()
        available_time = available_time[0]
        if available_time == "上午":
            time_in_eng = "AM"
        elif available_time == "下午":
            time_in_eng = "PM"
        # print(available_time)

        # Check number of available seats
        seat_available = seat_available.text
        # print(seat_available)

        # Read last date form log
        with open("log.txt", mode="r") as file_read:
            booked_schedule = file_read.readlines()[-1]
        # print(booked_schedule)
        booked_list = booked_schedule.split("/")
        booked_month = int(booked_list[1])
        booked_day = int(booked_list[2])
        booked_time = booked_list[3]
        # print(booked_month)
        # print(booked_day)
        # print(booked_time)
        print(f"Booked - Available\n"
              f"{booked_month}      -      {available_month}\n"
              f"{booked_day}     -      {available_day}\n")

        # Match log and recent available date
        if available_month <= booked_month:
            if available_day < booked_day:
                with smtplib.SMTP("smtp.exmail.qq.com") as connection:
                    connection.starttls()
                    connection.login(MY_EMAIL, PASSWORD)
                    # sock_out = socket.socket(socket.AF_INET, socket.SOCK_STREAM)  # TCP Connection
                    # sock_out.connect((remote_ip, remote_port))
                    # sock_out.send("ASD")
                    connection.sendmail(from_addr=MY_EMAIL, to_addrs=recipient,
                                        msg=f"Subject: More Recent Driving Test Date Available\n\n"
                                            f"Current booked date: {booked_month}/{booked_day}.\n"
                                            f"New available date: {available_month}/{available_day}.\n"
                                            f"Available position: {seat_available}.")

                # Keep change log
                with open("log.txt", mode="a") as file:
                    file.writelines(f"{test_date_calender_year}/"
                                    f"{available_month}/"
                                    f"{available_day}/"
                                    f"{available_time}/"
                                    f"{seat_available}\n")
                time.sleep(3)
                cancelled_booked()
                time.sleep(3)
                # Click Confirm Button
                driver.find_element(By.XPATH, date_to_choose).click()
                time.sleep(3)

                # Click Agreement Button
                driver.find_element(By.XPATH, '/html/body/div[9]/div[2]/a').click()
                time.sleep(3)

                # Input Applicant Information
                driver.find_element(By.ID, "idNo").send_keys("YOUR_ID")
                driver.find_element(By.ID, "birthdayStr").send_keys("YOUR_BIRTHDAY")
                driver.find_element(By.ID, "name").send_keys("YOUR_NAME")
                driver.find_element(By.ID, "contactTel").send_keys("YOU_MOBILE")
                driver.find_element(By.ID, "email").send_keys("YOUR_EMAIL")
                time.sleep(3)

                # Click Confirm Button
                # I used the qq exchange server, you may change the email server as you wish
                driver.find_element(By.XPATH, '//*[@id="form1"]/div').click()
                driver.find_element(By.XPATH, '//*[@id="form1"]/table/tbody/tr[6]/td/a[1]').click()
                time.sleep(3)
                with smtplib.SMTP("smtp.exmail.qq.com") as connection:
                    connection.starttls()
                    connection.login(MY_EMAIL, PASSWORD)
                    # sock_out = socket.socket(socket.AF_INET, socket.SOCK_STREAM)  # TCP Connection
                    # sock_out.connect((remote_ip, remote_port))
                    # sock_out.send("ASD")
                    connection.sendmail(from_addr=MY_EMAIL, to_addrs=recipient,
                                        msg=f"Subject: New Booking Date\n\n"
                                            f"New booked date has been changed."
                                            f"New date: {available_month}/{available_day}.\n"
                                            f"New time: {time_in_eng}.")
                booking_confirmation()

    except Exception as e:
        print(e, 'DMV Taiwan')

def cancelled_booked():
    # Connect
    chrome_options = Options()
    chrome_options.add_experimental_option("detach", True)
    # chrome_options.add_argument("enable-automation")
    service_obj = Service(CHROME_DRIVER_LOCATION)
    connection = webdriver.Chrome(service=service_obj, options=chrome_options)
    connection.get(CANCELLED_URL)

    # Input ID Number
    connection.find_element(By.ID, "idNo").send_keys("YOUR_ID")

    # Input Birthday
    connection.find_element(By.ID, "birthdayStr").send_keys("YOUR_BIRTHDAY")

    # Check booked date
    connection.find_element(By.XPATH, '//*[@id="form1"]/div/div/div/div/a/img').click()

    # Cancelled booked
    connection.find_element(By.XPATH, '/html/body/table/tbody/tr[2]/td/div[3]/table/tbody/tr[2]/td[5]/a').click()

    # Pop up confirmation
    confirm_button = connection.switch_to.alert
    confirm_button.accept()

    # This is for the pre-booked schedule, you need to cancelled the existing appointment, in order to make a new ones
    # Send Email informed booking cancelled
    with smtplib.SMTP("smtp.exmail.qq.com") as connection:
        connection.starttls()
        connection.login(MY_EMAIL, PASSWORD)
        # sock_out = socket.socket(socket.AF_INET, socket.SOCK_STREAM)  # TCP Connection
        # sock_out.connect((remote_ip, remote_port))
        # sock_out.send("ASD")
        connection.sendmail(from_addr=MY_EMAIL, to_addrs=recipient,
                            msg=f"Subject: Driving License Date Cancelled\n\nYour Booking at the Motor Vehical Test Registration has been cancelled")
        # sock_out.close()


def booking_confirmation():
    # Connect
    chrome_options = Options()
    chrome_options.add_experimental_option("detach", True)
    # chrome_options.add_argument("enable-automation")
    service_obj = Service(CHROME_DRIVER_LOCATION)
    connection2 = webdriver.Chrome(service=service_obj, options=chrome_options)
    connection2.get(CANCELLED_URL)

    # Input ID Number
    connection2.find_element(By.ID, "idNo").send_keys("YOUR_ID")

    # Input Birthday
    connection2.find_element(By.ID, "birthdayStr").send_keys("YOUR_BIRTHDAY")

    # Check booked date
    connection2.find_element(By.XPATH, '//*[@id="form1"]/div/div/div/div/a/img').click()

    # Get Content
    confirmation_result = connection2.find_elements(By.CLASS_NAME, "align_c")
    confirm_line_1 = confirmation_result[10].text.split()[0].replace("年"," ").replace("月"," ").replace("日"," ").split()
    confirm_line_2 = confirmation_result[11].text.replace("場次", " ").split()
    confirm_line_2 = confirm_line_2[0]
    confirmed_calender_date = f"{int(confirm_line_1[0])+1911}-0{confirm_line_1[1]}-{confirm_line_1[2]}"
    # print(confirmed_calender_date)
    if confirm_line_2 == '上午':
        am_pm = "AM"
    else:
        am_pm = "PM"
    # print(confirm_line_1)
    # print(confirm_line_2)
    # print(am_pm)

    # Send Email informed confirmed schedule
    with smtplib.SMTP("smtp.exmail.qq.com") as connection:
        connection.starttls()
        connection.login(MY_EMAIL, PASSWORD)
        # sock_out = socket.socket(socket.AF_INET, socket.SOCK_STREAM)  # TCP Connection
        # sock_out.connect((remote_ip, remote_port))
        # sock_out.send("ASD")
        connection.sendmail(from_addr=MY_EMAIL, to_addrs=recipient,
                            msg=f"Subject: Confirmed Driving Test Date\n\n"
                                f"{confirmed_calender_date} {am_pm}"
                            )
        # sock_out.close()

check_available_date()
# cancelled_booked()
# booking_confirmation()

# a = [10,14,18,22,26,30,34,38,42,46,50,54,58,62,66,70,74,78,82,86,90,94,98,102,107]
# for i in a:
#     print(round((i-7)/4))