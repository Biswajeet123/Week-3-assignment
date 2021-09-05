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
