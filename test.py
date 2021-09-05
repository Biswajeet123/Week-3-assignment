#FB screenshot
from selenium import webdriver

import time

from selenium.webdriver.common.keys import Keys

from selenium.webdriver.common.by import By

#from bs4 import BeautifulSoup

from selenium.webdriver.chrome.options import Options

from PIL import Image

import pandas as pd


df = pd.read_excel("C:/Users/goenka.yg/Procter and Gamble/Baby Care CMK (Fractal) - Documents/3. Working Files/Digital Dashboard/Processing data/Screenshots_subbrands_code/fblinks.xlsx")

username = "swaro.bs@pg.com"

password = "bswaro123"

DRIVER_PATH = 'C:/Users/goenka.yg/Downloads/chromedriver.exe'

option = Options()

option.add_argument("--disable-infobars")

option.add_argument("start-maximized")

option.add_argument("--disable-extensions")

# Pass the argument 1 to allow and 2 to block

option.add_experimental_option("prefs", {

 

    "profile.default_content_setting_values.notifications": 1

 

})


driver = webdriver.Chrome(chrome_options=option,executable_path=DRIVER_PATH)
#https://fb.me/1HIwX9mRztZbAQO
rm = ":/."
 
#len(df['Link to Asset'])
for i in range(len(df['link'])):

 

    driver.get(df['link'][i])
    file_name = df['link'][i]
    for char in rm:
        file_name=file_name.replace(char,"")
    #print(file_name)
    if i == 0:
        driver.find_element_by_id("username").send_keys('goenka.yg')
       
        driver.find_element_by_id("password").send_keys('Fractal=yg73')
       
        driver.find_element_by_class_name("ladda-button").send_keys(Keys.RETURN)
       
        time.sleep(5)
        
        driver.find_element_by_id("email").send_keys(username)

        driver.find_element_by_id("pass").send_keys(password)

        time.sleep(5)

        driver.find_element_by_name("login").send_keys(Keys.RETURN)

        time.sleep(10)

        driver.execute_script("window.scrollTo(0,100);")
        
        driver.get_screenshot_as_file(file_name+".png")

 

        x = 288
        
        y = 84

 

        width = 288+318

        height = 84+518

        im = Image.open(file_name+".png")

        im = im.crop((int(x), int(y), int(width), int(height)))

        im.save(file_name+".png") 

    else:
        
        time.sleep(10)
        
        driver.execute_script("window.scrollTo(0,100);")
        
        driver.get_screenshot_as_file(file_name+".png")
 

        x = 288


        y = 84


        width = 288+318
 

        height = 84+518


        im = Image.open(file_name+".png")
 

        im = im.crop((int(x), int(y), int(width), int(height))) 

        im.save(file_name+".png")                


        time.sleep(10)

driver.quit()


#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------
#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
#MOAT automation

from selenium import webdriver

import time

from selenium.webdriver.common.keys import Keys

from selenium.webdriver.common.action_chains import ActionChains

from selenium.webdriver.common.by import By

from bs4 import BeautifulSoup

from selenium.webdriver.chrome.options import Options

from selenium.webdriver.support.ui import Select

from PIL import Image

import pandas as pd

import pyautogui


username = "david.i@pg.com"

password = "Fractal@1234"

DRIVER_PATH = './chromedriver.exe'

s_date = "Jul 31,2021" #Write the date in MM/DD/YYYY format

e_date = "Aug 29,2021"

 

#keep the username password same, do not change  and set the driver path according to your username & folder name
#s_date and e_date to be taken from DMO file Insights data to and from column and keep the date in MM/DD/YYYY format only.

 

data_list = [2,4,6,8,10]

dateicon_position = [3,3,2,2,2]

csv_list = ["Brazil-FB-Display","Brazil-FB-Video","Brazil-IG-Display","Brazil-IG-Video","Brazil-IG-Stories-Video"]

month_name = "_3_July_2021-1_Aug_2021"

#Do not change any list element and variable names
#month_name to be kept as '_D_MMMM_YYYY-D_MMMM_YYYY' format

option = Options()

option.add_argument("--disable-infobars")

option.add_argument("start-maximized")

option.add_argument("--disable-extensions")

# Pass the argument 1 to allow and 2 to block

option.add_experimental_option("prefs", {

    "profile.default_content_setting_values.notifications": 1

})

 

driver = webdriver.Chrome(chrome_options=option,executable_path=DRIVER_PATH)

 

driver.get("https://analytics.moat.com/login")

 

driver.find_element_by_id("username").send_keys('goenka.yg') #type your PG username inside send_keys
       
driver.find_element_by_id("password").send_keys('Fractal=yg73') #type your PG password inside send_keys
       
driver.find_element_by_class_name("ladda-button").send_keys(Keys.RETURN)
       
time.sleep(5)

 

driver.find_element_by_name("username").send_keys(username)

 

driver.find_element_by_class_name("moat-button").send_keys(Keys.RETURN)



time.sleep(5)


driver.find_element_by_name("password").send_keys(password)


driver.find_element_by_class_name("moat-button").send_keys(Keys.RETURN)


for i in range(len(data_list)):



    time.sleep(25)

 

    elm = driver.find_element_by_xpath("(//div[@class='truncated-text-with-copy__text '])[position()=" + str(data_list[i]) + "]")

 

    elm.click()

 

   
    time.sleep(10)

 


    export = driver.find_element_by_xpath("//*/span[contains(text(), 'Exports')]")

 

    export.click()

 

    new_export = driver.find_element_by_xpath("//*/div[contains(text(), 'Open the new export wizard!')]")
    new_export.click()

 


    tab_list = driver.window_handles
    time.sleep(10)
    driver.switch_to.window(tab_list[1])

 


    driver.execute_script("window.scrollTo(0,document.body.scrollHeight)") 

    date_icon = driver.find_element_by_xpath("//span[@class = 'info-pop-up date-range-picker-info-popup']")

    date_icon.click()   
    
    time.sleep(5)


    start_date = driver.find_element_by_xpath("(//input[@class = 'input-main-wrapper__input '])[position()=1]")

    start_date.send_keys(Keys.CONTROL + "a")
    start_date.send_keys(Keys.DELETE)

 
    start_date.send_keys(s_date)
 

    end_date = driver.find_element_by_xpath("(//input[@class = 'input-main-wrapper__input '])[position()=2]")

 
    end_date.send_keys(Keys.CONTROL + "a")
    end_date.send_keys(Keys.DELETE)

 

    end_date.send_keys(e_date)

 

    #driver.find_element_by_class_name("moat-button moat-button--small").send_keys(Keys.RETURN)
    
    driver.execute_script("window.scrollTo(0,150)")
    time.sleep(10)
    driver.find_element_by_xpath("//*/button[contains(text(), 'Apply')]").click()

 

    add_filter = driver.find_element_by_xpath("//*/span[contains(text(), 'Add Filter')]").click()
    select_filter = driver.find_element_by_xpath("(//div[@class= 'moat-select__control css-0'])")
    select_filter.click()
    dropacc = driver.find_element_by_xpath("//*/div[contains(text(), 'Account')]")
    dropacc.click()

 

    brand_drop = driver.find_element_by_xpath("(//div[@class = 'moat-select__value-container css-0'])")
    brand_drop.click()
    driver.execute_script("window.scrollTo(0,150)")
    time.sleep(10)
    drop_text = driver.find_element_by_xpath("//input[@id = 'react-select-5-input']")
    drop_text.send_keys("Pampers")
    time.sleep(15)

 

    action = ActionChains(driver)
    action.send_keys(Keys.RETURN)
    action.perform()
    #1197844896947987: HUB P&G - 
    #brand_drop.send_keys(Keys.RETURN)
    
    #driver.find_element_by_xpath("//*[contains(text(), '1197844896947987: HUB P&G - Pampers')]").click()
    #brand_drop.send_keys("pampers")
    #brand_drop.send_keys(Keys.ENTER)
    
    #brand_drop = driver.find_element_by_xpath("//div[@class = 'moat-select__value-container css-0']").send_keys("")
    driver.find_element_by_xpath("//*/button[contains(text(), 'Next')]").click()
    #driver.find_element_by_xpath("(//button[@class = 'moat-button'])[position()=8]").send_keys(Keys.RETURN)

 
    Date_check = driver.find_element_by_xpath("//*/div[contains(text(), 'Date')]")
    Date_check.click()
    Campaign_check = driver.find_element_by_xpath("//*/div[contains(text(), 'Campaign')]")
    Campaign_check.click()
    AdSet_check = driver.find_element_by_xpath("//*/div[contains(text(), 'Ad Set')]")
    AdSet_check.click()
    Ad_check = driver.find_element_by_xpath("//*/div[contains(text(), 'Ad')]")
    Ad_check.click()
    Country_check = driver.find_element_by_xpath("//*/div[contains(text(), 'Country')]")
    Country_check.click()
    SelectAll_check = driver.find_element_by_xpath("//*/div[contains(text(), 'Select All')]")
    SelectAll_check.click()
    driver.find_element_by_xpath("//*/button[contains(text(), 'Next')]").click()
    driver.find_element_by_xpath("//*/button[contains(text(), 'Run Export')]").click()
    time.sleep(10)
    driver.close()
    tab_list = driver.window_handles
    driver.switch_to.window(tab_list[0])
    driver.back()
    driver.back()
     
time.sleep(10)

driver.get("https://analytics.moat.com/accounts/3279/exports/on-demand")

time.sleep(100)

driver.get("https://analytics.moat.com/accounts/3279/exports/on-demand")

time.sleep(50)

for i in range(1,6):

    time.sleep(20)

    #driver.find_element_by_xpath("(//*[local-name()='svg'][@class='overflow-container--icon'])[position() = " + str(i) +"]").click()
    driver.find_element_by_xpath("(//div[@class = 'overflow-container--icon'])[position() = " + str(i) +"]").click()
 
    time.sleep(10)

    driver.find_element_by_xpath("//*[contains(text(), 'Download')]").click()

time.sleep(20)   

 

driver.quit()



#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------
#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
#MOAT automation
