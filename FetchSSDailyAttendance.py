#!/usr/bin/python2.7
import datetime
import time
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
import pandas as pd
import pickle
import os
import glob
import logging
import shutil

import smtplib
import mimetypes
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.message import Message
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.text import MIMEText

DownloadRate = 600
iteration = 0

data_email_id = 'acknowledgesynchronization'
data_email_password = 'synchronizationacknowledge'
MailRate = 600

emailfrom = data_email_id + '@gmail.com'
emailto = data_email_id + '@gmail.com'

username = data_email_id
password = data_email_password

msg = MIMEMultipart()
msg["From"] = emailfrom
msg["To"] = emailto
msg.preamble = "Incremental Changes"
fileToSend = 'todays_attendance.pkl'
#driver = webdriver.Firefox('C:\\Program Files\\Mozilla Firefox')  # Optional argument, if not specified will search path.
#driver = webdriver.Firefox()  # Optional argument, if not specified will search path.
#profile = webdriver.FirefoxProfile()
#profile.accept_untrusted_certs = True

iteration = iteration + 1
print('-------------------------\n')
print("iteration = " + str(iteration) + ' @ ' + datetime.datetime.strftime(datetime.datetime.now(),'%d-%b-%Y %H:%M:%S'))
from pyvirtualdisplay import Display
from pyvirtualdisplay.smartdisplay import SmartDisplay
display = Display(visible=0, size=(1280, 1024))  
display = SmartDisplay(visible=0, size=(1280, 1024))  
display.start()
download_directory = '/home/rootcat/Downloads/AttendanceDate'
options = webdriver.ChromeOptions() 
preferences = {"download.default_directory": 'Downloads/AttendanceDate',"download.directory_upgrade":"true","default_content_setting_values.automatic_downloads":2}
options.add_experimental_option("prefs", preferences)
driver = webdriver.Chrome('/usr/lib64/chromium/chromedriver',chrome_options=options)  # Optional argument, if not specified will search path.
try:

    print("Fetching BIMS login page")
    driver.get('https://r:r@110.172.171.195:13254/bimswebsite/Common/WebPages/Login.aspx?ReturnUrl=%2fbimswebsite%2fCommon%2fWebPages%2fReports.aspx%3fStateInstanceID%3d3ac698a1-f888-4b44-8f74-2a486af07a66');
    element = WebDriverWait(driver, 10).until(lambda x: x.find_element_by_id("ctl00_ContentPlaceHolder1_txtUser"))

    #time.sleep(5) # Let the user actually see something!
    print("Searching username\/password")
    search_box = driver.find_element_by_id('ctl00_ContentPlaceHolder1_txtUser')
    search_box.send_keys('bh0011gb8104')
    password_box = driver.find_element_by_id('txtPassword')
    password_box.send_keys('20@Million')
    submit_button = driver.find_element_by_id('ctl00_ContentPlaceHolder1_btnSubmit')
    submit_button.click()

    print("Waiting for page load after username password")
    element = WebDriverWait(driver, 10).until(lambda x: x.find_element_by_id("ctl00_ContentPlaceHolder1_ddlArea_AttendanceGetDates"))
    area_drop_down = driver.find_element_by_id('ctl00_ContentPlaceHolder1_ddlArea_AttendanceGetDates')
    area_drop_down.click()
    area_drop_down.send_keys(Keys.DOWN + Keys.ENTER)
    print("Selected Centre")
    time.sleep(5)
    element = WebDriverWait(driver, 10).until(lambda x: x.find_element_by_id("ctl00_ContentPlaceHolder1_ddlArea_AttendanceGetDates"))
    center_drop_down = driver.find_element_by_id('ctl00_ContentPlaceHolder1_ddlCentre_AttendanceGetDates')
    center_drop_down.send_keys(Keys.DOWN + Keys.ENTER)
    print("Selected Area")
    time.sleep(5)
    element = WebDriverWait(driver, 10).until(lambda x: x.find_element_by_id("ctl00_ContentPlaceHolder1_ddlArea_AttendanceGetDates"))
    dept_drop_down = driver.find_element_by_id('ctl00_ContentPlaceHolder1_ddlDept_AttendanceGetDates')
    dept_drop_down.send_keys(Keys.DOWN + Keys.ENTER + Keys.ENTER)

    from_date_box = driver.find_element_by_id('ctl00_ContentPlaceHolder1_txtFromDate_AttendanceGetDates')
    from_date_box.send_keys(datetime.datetime.strftime(datetime.datetime.now() - datetime.timedelta(days=7),'%Y/%b/%d'))
    to_date_box = driver.find_element_by_id('ctl00_ContentPlaceHolder1_txtToDate_AttendanceGetDates')
    to_date_box.send_keys(datetime.datetime.strftime(datetime.datetime.now(),'%Y/%b/%d'))
    print("Entered Start and end date")

    fetch_command = driver.find_element_by_id('ctl00_ContentPlaceHolder1_btnFetch_AttendanceGetDates')
    fetch_command.click()
    print("Waiting for download")




    time.sleep(180)
    img = display.waitgrab()
    img.save('/tmp/successful_fetch.png')
    driver.quit()
    print("Now trying to mail")

    try:
        #filenames = sorted(filter(os.path.isfile, os.listdir('C:\\Users\\nxa19154\\Downloads\\')), key=os.path.getmtime)
        #filenames = sorted(filter(os.path.isfile, os.listdir(download_directory)), 
        #            key=lambda x: os.path.getmtime(x),reverse=True)

        filenames = filter(os.path.isfile, glob.glob(download_directory + "/Report*.xls"))
        print(filenames)
        filenames = sorted(filenames,key=lambda x: os.path.getmtime(x),reverse=True)
        print('---------------------------\n')
        print(filenames)
        for filename in filenames:
            # print(os.path.join(directoy, filename))
            (mode, ino, dev, nlink, uid, gid, size, atime, mtime, ctime) = os.stat(os.path.join(download_directory,filename))
            print(filename + " was last modified: %s" % time.ctime(mtime))
            now = datetime.datetime.now()
            if datetime.datetime.strptime(time.ctime(mtime), "%a %b %d %H:%M:%S %Y") > now.replace(hour=0, minute=0, second=0):
                print("selected file\n")
                print(datetime.datetime.strftime(datetime.datetime.now(),'%d-%b-%Y %H:%M:%S') + ' : ' + filename + " was last modified: %s" % time.ctime(mtime))
                dfs = pd.read_html(os.path.join(download_directory,filename),header=0)
                df = pd.concat(dfs)

                df['DutyDate'] = df['DutyDate'].apply(lambda x: datetime.datetime.strptime(x,'%d %b %Y'))
                df.to_excel('debug_after_time_conversion.xlsx')
                todays_attendance_df = df[df['DutyDate'] >= now.replace(hour=0, minute=0, second=0 , microsecond=0)]

                todays_attendance_df.to_excel('debug_todays_attendance.xlsx')
                todays_attendance_df.to_pickle(fileToSend)
                msg["Subject"] = "SENDING_INCREMENTAL_UPDATE:DATE: " + now.strftime("%Y-%m-%d")

                fp = open(fileToSend, "rb")
                attachment = MIMEBase('application', 'octet-stream')
                attachment.set_payload(fp.read())
                fp.close()
                encoders.encode_base64(attachment)
                attachment.add_header("Content-Disposition", "attachment", filename=fileToSend)
                msg.attach(attachment)


                server = smtplib.SMTP("smtp.gmail.com:587",timeout=10)
                server.starttls()
                server.login(username,password)
                server.sendmail(emailfrom, emailto, msg.as_string())
                server.quit()
                src = os.path.join(download_directory,filename)
                shutil.copy2(src, '/tmp/')
                for f in glob.glob(download_directory + "/Report*.xls"):
                    os.remove(os.path.join(download_directory,f))
                break
    except Exception as e:
        print("Failed Somewhere to mail")
        print(str(e))
        
    print("Going to sleep after mailing")
    #time.sleep(MailRate)
        #time.sleep(DownloadRate) # Let the user actually see something!
except Exception as e:
    print("Something went wrong: in fetching")
    print(str(e))
    img = display.waitgrab()
    img.save('/tmp/failed_fetch.png')
    driver.quit()
    #time.sleep(DownloadRate) # Let the user actually see something!
