from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
#from selenium.webdriver.common.proxy import Proxy,ProxyType
import time
import cookielib
import requests
import csv
import xlsxwriter
from xlutils.copy import copy
from xlrd import open_workbook

input_file_name = raw_input("Enter The Input file Name (with csv Extention ): ")
output_file_name = raw_input("Enter The file Name (with xls Extention ) : ")
#print output_file_name
workbook = xlsxwriter.Workbook(output_file_name)
worksheet = workbook.add_worksheet()
workbook.close()
book_ro = open_workbook(output_file_name)
book = copy(book_ro)
sheet1 = book.get_sheet(0)
count=0
roww=0
coll=0
#page_content=''
print 'Launching Chrome..'
#prox = Proxy()
#prox.proxy_type = ProxyType.MANUAL
#prox.http_proxy = "127.0.0.1:9667"
#prox.socks_proxy = "127.0.0.1:9667"
#prox.ssl_proxy = "127.0.0.1:9667"
#capabilities = webdriver.DesiredCapabilities.CHROME
#prox.add_to_capabilities(capabilities)
options = webdriver.ChromeOptions()
options.add_argument('--ignore-certificate-errors')
options.add_argument('--ignore-ssl-errors')
capa = DesiredCapabilities.CHROME
capa["pageLoadStrategy"] = "none"
browser = webdriver.Chrome(executable_path='C:\Users\lenovo\Desktop\python\chromedriver.exe',chrome_options=options,desired_capabilities=capa)
#print 'Waiting for 2 mins...'
#time.sleep(90)
print 'Entering to Hoovers...'
with open(input_file_name, "r") as f:
    reader=csv.reader(f)
    for row in reader:
        site = row[0]
        checker={'value': 1}
        #print checker['value']
        attempt_count={'value': 1}
        count+=1
        def page_l():
            if attempt_count['value']<3:
                try:
                    time.sleep(2)
                    browser.get(site)
                    wait = WebDriverWait(browser, 12)
                    wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="content"]/div[1]/div[5]/div/div/div/div[1]/table/tbody')))
                    browser.execute_script("window.stop();")
                except TimeoutException:
                    attempt_count['value']+=1
                    page_l()
            else:
                pass
        if count<400:
            try:
                time.sleep(2)
                browser.get(site)
           #time.sleep(5)
                wait = WebDriverWait(browser, 12)
                wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="content"]/div[1]/div[5]/div/div/div/div[1]/table/tbody')))
                browser.execute_script("window.stop();")
            except TimeoutException:
                page_l()
            #continue
            #browser.get(site)
            #wait = WebDriverWait(browser, 15)
            #wait.until(EC.visibility_of_element_located((By.XPATH, "/html/body/div/div[2]/section[1]")))
            #browser.execute_script("window.stop();")
        #time.sleep(3)
            el_count={'value': 1}
            el_count1={'value': 1}
            def element_fun():
                try:
                    elements=browser.find_element_by_xpath('//*[@id="content"]/div[1]/div[5]/div/div/div/div[1]/table/tbody')
                    checker['value']=0
                #print checker['value']
                #page_content=browser.find_element_by_xpath("/html/body/div/div[2]/section[1]").get_attribute("outerHTML")
                except NoSuchElementException:
                    if el_count['value']<2:
                        el_count['value']+=1
                        print '~~~~~~~~Waiting For 10 Seconds~~~~~~~~~~'
                    #browser.get(site)
                    #time.sleep(5)
                    #wait = WebDriverWait(browser, 15)
                    #wait.until(EC.visibility_of_element_located((By.XPATH, "/html/body/div/div[2]/section[1]")))
                    #browser.execute_script("window.stop();")
                        page_l()
                        element_fun()
                    elif (el_count['value']==2) and (el_count1['value']==1):
                        print '~~~~~~~~Retrying~~~~~~~~~~'
                        el_count['value']=1
                        el_count1['value']+=1
                        try:
                            time.sleep(2)
                            browser.get(site)
                       #time.sleep(5)
                            wait = WebDriverWait(browser, 15)
                            wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="content"]/div[1]/div[5]/div/div/div/div[1]/table/tbody')))
                            browser.execute_script("window.stop();")
                            element_fun()
                        except TimeoutException:
                            pass
            try:        #page_content=browser.find_element_by_xpath("/html/body/div/div[2]/section[1]").get_attribute("outerHTML")
                element_fun()
            except TimeoutException:
            #continue
            #browser.get(site)
            #wait = WebDriverWait(browser, 15)
            #wait.until(EC.visibility_of_element_located((By.XPATH, "/html/body/div/div[2]/section[1]")))
            #browser.execute_script("window.stop();")
                page_l()
                element_fun()
            if checker['value']==0:
                print str(count)+' '+site
                try:
                    elems1=browser.find_elements_by_xpath('//*[@id="content"]/div[1]/div[5]/div/div/div/div[1]/table/tbody/tr/td[1]/a')
                    for elem1 in elems1:
                        people_link=elem1.get_attribute("href")
                        sheet1.write(roww,coll,site)
                        sheet1.write(roww,coll+1,people_link)
                        roww+=1
                        book.save(output_file_name)
                except:
                    continue
            else:
                print str(count)+' *** '+site+' *** Element Not Found'
                pass
        else:
            browser.close()
            time.sleep(2)
            count=1
            browser = webdriver.Chrome(executable_path='C:\Users\lenovo\Desktop\python\chromedriver.exe',chrome_options=options,desired_capabilities=capa)
            continue
print 'Closing Chrome..'
browser.close()
