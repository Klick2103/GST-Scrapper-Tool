from email.headerregistry import Address
from re import I
import time
import pyautogui as a
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

#opening edge
edgeBrowser = webdriver.Edge("G:\Program Files\VS Code\Web driver\msedgedriver.exe")
edgeBrowser.get('https://services.gst.gov.in/services/searchtp') 
edgeBrowser.maximize_window()
time.sleep(5)
FailSafeException=True

#reading excel
read=pd.read_excel('G:\Program Files\VS Code\Explore\Book1.xlsx')

GST_Number=[]
Reg_Status=[]
Type=[]
Address=[]

#looping
for i in read.index:
    entry=read.loc[i]
    GSTIN=edgeBrowser.find_element(By.NAME,'for_gstin')
    GSTIN.send_keys(entry['GST Number'])

#to enter manually captcha 
    a.press('tab')
    time.sleep(2)

#wait unitl 6 digit capctha is given by user
    WebDriverWait(edgeBrowser, 20).until(lambda driver: len(edgeBrowser.find_element(By.ID,"fo-captcha").get_attribute("value")) == 6)
    submit=edgeBrowser.find_element(By.ID,'lotsearch')
    submit.click()

#wait if user types invalid captcha
    WebDriverWait(edgeBrowser, 60).until(EC.presence_of_element_located((By.XPATH,'//*[@id="lottable"]/div[2]/div[2]/div/div[2]/p[2]')))
   
#Bot will store the required fields    
    GST_Number.append(entry['GST Number'])
    Reg_Status.append(edgeBrowser.find_elements(By.XPATH,'//*[@id="lottable"]/div[2]/div[2]/div/div[2]/p[2]')[0].text)   
    Type.append(edgeBrowser.find_elements(By.XPATH,'//*[@id="lottable"]/div[2]/div[2]/div/div[3]/p[2]')[0].text)
    Address.append(edgeBrowser.find_elements(By.XPATH,'//*[@id="lottable"]/div[2]/div[3]/div/div[3]/p[2]')[0].text)
    time.sleep(5)

#Stored fields will be writtem in given file
    df=pd.DataFrame({'GST_Number':GST_Number,'Reg_Status':Reg_Status,'Type':Type,'Address':Address})
    writer=pd.ExcelWriter('G:\Program Files\VS Code\Explore\Output.xlsx',engine='openpyxl')
    df.to_excel(writer,sheet_name="Result")
    writer.save()
edgeBrowser.quit()
