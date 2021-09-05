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
#amazon automation

import requests
from glob import glob
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
from time import sleep
from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import re
from selenium.webdriver.chrome.options import Options
from PIL import Image
from requests_html import HTMLSession
import warnings

 

warnings.filterwarnings('ignore')
xl = pd.read_excel("./live star ratings.xlsx",sheet_name = None) #read all sheet names in the excel file
xl = list(xl.keys()) #sheet names in excel file
HEADERS = {'User-Agent':
           'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/44.0.2403.157 Safari/537.36',
                           'Accept-Language': 'en-US, en;q=0.5'}

 


#for amazon with out poe pox
with requests.Session() as s:
    amazon = pd.read_excel("./live star ratings.xlsx",sheet_name = xl[1])
    amazon_df = pd.DataFrame([],columns = ["brand_from_ppt","Volume","Ratings","weighted_avg_star_rating"])
    for i in range(amazon.shape[0]):
        time.sleep(10)

 

        link = amazon.link[i]
        #print(link)
        product_search = amazon.product_search[i].split(",")
        filter_product = amazon.not_contain[i].split("{")
        page = s.get(link,headers = HEADERS,verify = False)
        soup = BeautifulSoup(page.content, features="html.parser")

 

        rr_link = soup.findAll("div", {"class":"a-row a-size-small"})
        rr = []
        for j in rr_link:
            rr.append(j.text)
        #print(len(rr))

 


        product_link = soup.findAll("span", {"class":"a-size-base-plus a-color-base a-text-normal"})
        product_name = []
        for k in product_link:
            product_name.append(k.text)
        #print(len(product_name))

 

        #if len(rr) == len(product_name):
            #print(i)
            
        #product_name = list(map(lambda x: x.split('\n')[-1],product_name))
        total_reviews = list(map(lambda x: int(x.strip().split(" ")[-1].replace(',','')),rr))
        total_reviews = [i for i in total_reviews if i != 0]
        total_ratings = list(map(lambda x: float(x.strip().split(" ")[0]),rr))
        total_ratings = [i for i in total_ratings if i != 0.0]
        
        raw_data = pd.DataFrame({'Product name':pd.Series(product_name),'total_reviews':pd.Series(total_reviews),'total_ratings':pd.Series(total_ratings)})
        raw_data = raw_data.dropna()
        #raw_data['total_reviews'] = raw_data['reviews and ratings'].apply(lambda x: int(x.strip().split(" ")[-1].replace(',','')))
        #raw_data['total_ratings'] = raw_data['reviews and ratings'].apply(lambda x: float(x.strip().split(" ")[0]))
        
        if product_search[0] == "pampers cruisers":
            df = raw_data[(raw_data['Product name'].str.contains(product_search[0],case = False) == True) & (raw_data['Product name'].str.contains('360') == False) & (raw_data['Product name'].str.contains(product_search[1],case = False) == True) & (raw_data['Product name'].apply(lambda x: any(filter_product in x for filter_product in filter_product)) == False)]
            #df = df[df.total_reviews == df.total_reviews.max()]
            brand_from_ppt = amazon["Brand from ppt"][i]
            weighted_avg_star_rating = round((df.total_reviews*df.total_ratings).sum()/(df.total_reviews.sum()),2)
            Volume = df.total_reviews.sum()
            Ratings = round(df.total_ratings.mean(),2)
            temp = pd.DataFrame({"brand_from_ppt":[brand_from_ppt],"Volume":[Volume],"Ratings":[Ratings],"weighted_avg_star_rating":[weighted_avg_star_rating]})
            amazon_df = amazon_df.append(temp)
        elif product_search[0] == "cruisers":
            df = raw_data[(raw_data['Product name'].str.contains(product_search[0],case = False) == True) & (raw_data['Product name'].str.contains('360') == True) & (raw_data['Product name'].str.contains(product_search[1],case = False) == True) & (raw_data['Product name'].apply(lambda x: any(filter_product in x for filter_product in filter_product)) == False)]
            #df = df[df.total_reviews == df.total_reviews.max()]
            brand_from_ppt = amazon["Brand from ppt"][i]
            weighted_avg_star_rating = round((df.total_reviews*df.total_ratings).sum()/(df.total_reviews.sum()),2)
            Volume = df.total_reviews.sum()
            Ratings = round(df.total_ratings.mean(),2)
            temp = pd.DataFrame({"brand_from_ppt":[brand_from_ppt],"Volume":[Volume],"Ratings":[Ratings],"weighted_avg_star_rating":[weighted_avg_star_rating]})
            amazon_df = amazon_df.append(temp)
        else:
            if product_search[1] != '':
                df = raw_data[(raw_data['Product name'].str.contains(product_search[0],case = False) == True) & (raw_data['Product name'].str.contains(product_search[1],case = False) == True) & (raw_data['Product name'].apply(lambda x: any(filter_product in x for filter_product in filter_product)) == False)]
                #df = df[df.total_reviews == df.total_reviews.max()]
                brand_from_ppt = amazon["Brand from ppt"][i]
                weighted_avg_star_rating = round((df.total_reviews*df.total_ratings).sum()/(df.total_reviews.sum()),2)
                Volume = df.total_reviews.sum()
                Ratings = round(df.total_ratings.mean(),2)
                temp = pd.DataFrame({"brand_from_ppt":[brand_from_ppt],"Volume":[Volume],"Ratings":[Ratings],"weighted_avg_star_rating":[weighted_avg_star_rating]})
                amazon_df = amazon_df.append(temp)
            else:
                df = raw_data[(raw_data['Product name'].str.contains(product_search[0],case = False) == True) & (raw_data['Product name'].apply(lambda x: any(filter_product in x for filter_product in filter_product)) == False)]
                #df = df[df.total_reviews == df.total_reviews.max()]
                brand_from_ppt = amazon["Brand from ppt"][i]
                weighted_avg_star_rating = round((df.total_reviews*df.total_ratings).sum()/(df.total_reviews.sum()),2)
                Volume = df.total_reviews.sum()
                Ratings = round(df.total_ratings.mean(),2)
                temp = pd.DataFrame({"brand_from_ppt":[brand_from_ppt],"Volume":[Volume],"Ratings":[Ratings],"weighted_avg_star_rating":[weighted_avg_star_rating]})
                amazon_df = amazon_df.append(temp)

 


#for amazon with POE POX

 

with requests.Session() as s1:
    
    amazon_total = pd.DataFrame([],columns = ["brand_from_ppt","Volume","Ratings","weighted_avg_star_rating"])
    for t in range(3,len(xl)):
        amazon1 = pd.read_excel("./live star ratings.xlsx",sheet_name = xl[t])
        amazon_df1 = pd.DataFrame([],columns = ["brand_from_ppt","Volume","Ratings","weighted_avg_star_rating"])
        for i in range(amazon1.shape[0]):
            time.sleep(10)

 

            link = amazon1.link[i]
            #print(link)
            product_search = amazon1.product_search[i].split(",")
            filter_product = amazon1.not_contain[i].split("{")
            page = s1.get(link,headers = HEADERS,verify = False)
            soup = BeautifulSoup(page.content, features="html.parser")

 

            rr_link = soup.findAll("div", {"class":"a-row a-size-small"})
            rr = []
            for j in rr_link:
                rr.append(j.text)
            #print(len(rr))

 


            product_link = soup.findAll("span", {"class":"a-size-base-plus a-color-base a-text-normal"})
            product_name = []
            for k in product_link:
                product_name.append(k.text)
            #print(len(product_name))

 

            #if len(rr) == len(product_name):
                #print(i)
                
            #product_name = list(map(lambda x: x.split('\n')[-1],product_name))
            total_reviews = list(map(lambda x: int(x.strip().split(" ")[-1].replace(',','')),rr))
            total_reviews = [i for i in total_reviews if i != 0]
            total_ratings = list(map(lambda x: float(x.strip().split(" ")[0]),rr))
            total_ratings = [i for i in total_ratings if i != 0.0]
            
            raw_data1 = pd.DataFrame({'Product name':pd.Series(product_name),'total_reviews':pd.Series(total_reviews),'total_ratings':pd.Series(total_ratings)})
            raw_data1 = raw_data1.dropna()
            #raw_data1['total_reviews'] = raw_data1['reviews and ratings'].apply(lambda x: int(x.strip().split(" ")[-1].replace(',','')))
            #raw_data1['total_ratings'] = raw_data1['reviews and ratings'].apply(lambda x: float(x.strip().split(" ")[0]))

 

            df1 = raw_data1[(raw_data1['Product name'].str.contains(product_search[0].strip(),case = False) == True) & (raw_data1['Product name'].str.contains(product_search[1].strip(),case = False) == True) & (raw_data1['Product name'].apply(lambda x: any(filter_product in x for filter_product in filter_product)) == False)]
            brand_from_ppt = xl[t]
            weighted_avg_star_rating = round((df1.total_reviews*df1.total_ratings).sum()/(df1.total_reviews.sum()),2)
            Volume = df1.total_reviews.sum()
            Ratings = round(df1.total_ratings.mean(),2)
            temp = pd.DataFrame({"brand_from_ppt":[brand_from_ppt],"Volume":[Volume],"Ratings":[Ratings],"weighted_avg_star_rating":[weighted_avg_star_rating]})
            amazon_df1 = amazon_df1.append(temp)
        
        brand_from_ppt = xl[t]
        weighted_avg_star_rating = round((amazon_df1.Volume*amazon_df1.Ratings).sum()/(amazon_df1.Volume.sum()),2)
        Volume = amazon_df1.Volume.sum()
        Ratings = round(amazon_df1.Ratings.mean(),2)
        amazon_df2 = pd.DataFrame({"brand_from_ppt":[brand_from_ppt],"Volume":[Volume],"Ratings":[Ratings],"weighted_avg_star_rating":[weighted_avg_star_rating]})
        amazon_total = amazon_total.append(amazon_df2)

 

amazon_df = amazon_df.append(amazon_total)
amazon_df = amazon_df.reset_index(drop=True)
amazon_df.columns = ["Brand from ppt","Volume","Ratings","Weighted Average Star Rating"]

 

amazon_df

#for walmart---------------------------------------------------------------------------------------------------------------------------------

DRIVER_PATH = './chromedriver.exe'
option = Options()
option.add_argument("--disable-infobars")

option.add_argument("start-maximized")


option.add_argument("--disable-extensions")

# Pass the argument 1 to allow and 2 to block

option.add_experimental_option("prefs", {


    "profile.default_content_setting_values.notifications": 1

})

 

driver = webdriver.Chrome(chrome_options=option,executable_path=DRIVER_PATH)
walmart = pd.read_excel("./live star ratings.xlsx",sheet_name = xl[2])
walmart_df = pd.DataFrame([],columns = ["brand_from_ppt","Volume","Ratings","weighted_avg_star_rating"])
for i in range(walmart.shape[0]):
    link = walmart.link[i]
    product_search = walmart.product_search[i].split(",")
    filter_product = walmart.not_contain[i].split("{")
    
    if i == 0:
        driver.get(link)
        driver.find_element_by_id("username").send_keys('swaro.bs')

 

        driver.find_element_by_id("password").send_keys('Fractal48')
        driver.find_element_by_class_name("ladda-button").send_keys(Keys.RETURN)
        time.sleep(60) #press-hold
        x = driver.find_elements_by_xpath("//div[@class='search-result-product-title gridview']")
        y = driver.find_elements_by_xpath("//span[@class='average-rating']")
        z = driver.find_elements_by_xpath("//span[@class='stars-reviews-count']")
        product_name = []
        for span in x:
            product_name.append(span.text)
        product_rating = []
        for span in y:
            product_rating.append(span.text)
        product_reviews = []
        for span in z:
            product_reviews.append(span.text)
            
        product_name = list(map(lambda x: x.split('\n')[-1],product_name))
        product_reviews = list(map(lambda x: int(x.split('\n')[0]),product_reviews))
        product_reviews = [i for i in product_reviews if i != 0]
        product_rating = list(map(lambda x: float(x.split('\n')[1]),product_rating))
        product_rating = [i for i in product_rating if i != 0.0]
        
        raw_data1 = pd.DataFrame({'Product name':pd.Series(product_name),'total_reviews':pd.Series(product_reviews),'total_ratings':pd.Series(product_rating)})
        raw_data1 = raw_data1.dropna()
        #raw_data1["Product name"] = raw_data1["Product name"].apply(lambda x: x.split('\n')[-1])
        #raw_data1["total_reviews"] = raw_data1["total_reviews"].apply(lambda x: int(x.split('\n')[0]))
        #raw_data1["total_ratings"] = raw_data1["total_ratings"].apply(lambda x: float(x.split('\n')[1]))
        df = raw_data1[(raw_data1['Product name'].str.contains(product_search[0].strip(),case = False) == True) & (raw_data1['Product name'].str.contains(product_search[1].strip(),case = False) == True) & (raw_data1['Product name'].apply(lambda x: any(filter_product in x for filter_product in filter_product)) == False)]
        #df = df[df.total_reviews == df.total_reviews.max()]
        #df['brand_from_ppt'] = walmart["Brand from ppt"][i]
        brand_from_ppt = walmart["Brand from ppt"][i]
        weighted_avg_star_rating = round((df.total_reviews*df.total_ratings).sum()/(df.total_reviews.sum()),2)
        Volume = df.total_reviews.sum()
        Ratings = round(df.total_ratings.mean(),2)
        temp = pd.DataFrame({"brand_from_ppt":[brand_from_ppt],"Volume":[Volume],"Ratings":[Ratings],"weighted_avg_star_rating":[weighted_avg_star_rating]})
        walmart_df = walmart_df.append(temp)
    else:
        driver.get(link)
        time.sleep(30)
        x = driver.find_elements_by_xpath("//div[@class='search-result-product-title gridview']")
        y = driver.find_elements_by_xpath("//span[@class='average-rating']")
        z = driver.find_elements_by_xpath("//span[@class='stars-reviews-count']")
        product_name = []
        for span in x:
            product_name.append(span.text)
        product_rating = []
        for span in y:
            product_rating.append(span.text)
        product_reviews = []
        for span in z:
            product_reviews.append(span.text)
            
        product_name = list(map(lambda x: x.split('\n')[-1],product_name))
        product_reviews = list(map(lambda x: int(x.split('\n')[0]),product_reviews))
        product_reviews = [i for i in product_reviews if i != 0]
        product_rating = list(map(lambda x: float(x.split('\n')[1]),product_rating))
        product_rating = [i for i in product_rating if i != 0.0]
        
        raw_data1 = pd.DataFrame({'Product name':pd.Series(product_name),'total_reviews':pd.Series(product_reviews),'total_ratings':pd.Series(product_rating)})
        raw_data1 = raw_data1.dropna()
        #raw_data1["Product name"] = raw_data1["Product name"].apply(lambda x: x.split('\n')[-1])
        #raw_data1["total_reviews"] = raw_data1["total_reviews"].apply(lambda x: int(x.split('\n')[0]))
        #raw_data1["total_ratings"] = raw_data1["total_ratings"].apply(lambda x: float(x.split('\n')[1]))
        df = raw_data1[(raw_data1['Product name'].str.contains(product_search[0].strip(),case = False) == True) & (raw_data1['Product name'].str.contains(product_search[1].strip(),case = False) == True) & (raw_data1['Product name'].apply(lambda x: any(filter_product in x for filter_product in filter_product)) == False)]
        brand_from_ppt = walmart["Brand from ppt"][i]
        weighted_avg_star_rating = round((df.total_reviews*df.total_ratings).sum()/(df.total_reviews.sum()),2)
        Volume = df.total_reviews.sum()
        Ratings = round(df.total_ratings.mean(),2)
        temp = pd.DataFrame({"brand_from_ppt":[brand_from_ppt],"Volume":[Volume],"Ratings":[Ratings],"weighted_avg_star_rating":[weighted_avg_star_rating]})
        walmart_df = walmart_df.append(temp)
driver.close()

 

walmart_df.columns = ["Brand from ppt","Volume","Ratings","Weighted Average Star Rating"]
live_star = amazon_df.append(walmart_df)
live_star = live_star.reset_index(drop=True)
live_star

 

live_star.to_excel("./live_star.xlsx",index = False)

#------------------------------------------------------------------------------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------------------------------------------------------------------------------

 

#for revuze automation

 

import requests
import json
from datetime import datetime, date
import calendar
import pandas as pd
from dateutil.relativedelta import relativedelta
today = datetime.date(datetime.now()) #gives today
today_year = today.year # gives current year
today_month = today.month -1 # gives current month
start_date = date(today_year,today_month, 1)
end_date = date(today_year,today_month,calendar.monthrange(today_year, today_month)[1])
print(start_date)
print(end_date)

 

start_date_range = []
for i in range(1,13): #loop for startdate of last "N" months
    sd = start_date + relativedelta(months=-i)
    start_date_range.append(sd)
end_date_range = []
for i in range(1,13): #loop for enddate of last "N" months
    ed = end_date + relativedelta(months=-i)
    end_date_range.append(ed)
date_table = pd.DataFrame({"startdate":start_date_range,"enddate":end_date_range})
date_table.loc[-1] = [start_date,end_date]  # adding a row
date_table.index = date_table.index + 1  # shifting index
date_table = date_table.sort_index()  # sorting by index
date_table = date_table.iloc[0:-1,:]
date_table['startdate_month'] = date_table.startdate.apply(lambda x: x.strftime("%b")) #month MMM format
date_table['enddate_month'] = date_table.enddate.apply(lambda x: x.strftime("%b"))
date_table['startdate_year'] = date_table.startdate.apply(lambda x: x.year)
date_table['enddate_year'] = date_table.enddate.apply(lambda x: x.year)
date_table['enddate_day'] = date_table.enddate.apply(lambda x: x.day)
date_table['file_name'] = "P1M All groups_M"+(date_table.index+1).map(str)+".xlsx" #file name to be saved as
date_table

 

url = "https://gateway.revuze.it/api/groups"

 

#change Authorization and project_uuid after logging in to revuze portal, take from groups api request header

 

HEADERS = {'User-Agent':
           'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/44.0.2403.157 Safari/537.36',
                           'Accept-Language': 'en-US, en;q=0.5',
          'Authorization':'Bearer eyJhbGciOiJSUzI1NiIsImtpZCI6Ijk1NDRiYWI5Y2QwMmRiYTNlMTNkMGMwN2JmNjcwZWY1IiwidHlwIjoiSldUIn0.eyJuYmYiOjE2MzA1NjA0NzksImV4cCI6MTYzMTYxMTY3OSwiaXNzIjoiaHR0cDovL3Nzby5yZXZ1emUuaXQiLCJhdWQiOlsiaHR0cDovL3Nzby5yZXZ1emUuaXQvcmVzb3VyY2VzIiwiZ2F0ZXdheV9wcm9kdWN0aW9uIl0sImNsaWVudF9pZCI6ImRhc2hib2FyZF91aSIsInN1YiI6IjA5Y2YyNTVmLTEyOWYtNDQ5ZS1iYzYwLThjYjk0OTA3YWE3MyIsImF1dGhfdGltZSI6MTYzMDU2MDQ3OCwiaWRwIjoibG9jYWwiLCJzY29wZSI6WyJvcGVuaWQiLCJwcm9maWxlIiwiZ2F0ZXdheV9wcm9kdWN0aW9uIiwib2ZmbGluZV9hY2Nlc3MiXSwiYW1yIjpbInB3ZCJdfQ.DJUv3mJRfSsDqo99SBM-0znG5W_OK7g5gBS2K5JOZ8FcsYgPX1Ll8IXrmtAmoMyOFgB4xjelPD023qrASNrxrkg-ngaDQnt07pLuyeOjf8OpDuqfUo9RWGVhAFCperfMBgyGbGg5snDWE1W--1lEjV56kjDa1x9bXjWleIfZaOZIjq71LXj_2Z2vjKCnym6cUii4SP0sDU27A30vqz4mVxcr2jLaUDLCTkvFqtPJ9DFWiNtMyI3a7Huz0cIrlVFDeE62dSwCgob-rtAI4Az1JoVFs4fbFl18Bksyiim-0KzYQiLcCCj5OyyyK3UF-D2ngpDXC7Q9k7fsk9hwC8IW8w',
          'project_uuid':'e7b31d49-6a58-11e8-9211-baa4b8fd5ab8'}

 

#for P1M "N" months data
for i in range(date_table.shape[0]):
    fromDate = str(date_table.startdate[i]) #fromdate to be passed as filter in api payload
    toDate = str(date_table.enddate[i]) #todate to be passed as filter in api payload
    payload = {"filters":{"hierarchies":[],"sources":[],"star_ratings":[],"brands":[],"entities":[],"topics":[],"mega_topics":[],"review_text":[],"groups":[],"fromDate":fromDate,"toDate":toDate,"selectedRange":"Custom Range","price":[],"promotion":"","custom_filters":[],"quote_text":[]},"objectPaging":{"numberOfRows":10,"page":1,"searchText":"","sort":""},"onlyUniqueReviews":True}
    r = requests.post(url,json=payload,headers = HEADERS)
    if r.status_code == 200:
        df = json.loads(r.text)
        total = int(df['total']) #get the total number of rows then pass again in the payload with only one page
        payload1 = {"filters":{"hierarchies":[],"sources":[],"star_ratings":[],"brands":[],"entities":[],"topics":[],"mega_topics":[],"review_text":[],"groups":[],"fromDate":fromDate,"toDate":toDate,"selectedRange":"Custom Range","price":[],"promotion":"","custom_filters":[],"quote_text":[]},"objectPaging":{"numberOfRows":total,"page":1,"searchText":"","sort":""},"onlyUniqueReviews":True}
        r1 = requests.post(url,json=payload1,headers = HEADERS)
        if r1.status_code == 200:
            df1 = json.loads(r1.text)
            df1 = pd.DataFrame(df1['data']) #change json to dataframe
            df1.to_excel("./P1M data/"+date_table.file_name[i],index=False) #save the output files into P1M data folder with filename

 

print("      ")
print("---------------------------output files generated-------------------------------------")
print("      ")

 

file_name_all = ['P1M_all.xlsx','P3M_all.xlsx','P6M_all.xlsx','P12M_all.xlsx']
file_name_promo = ['P1M_promo.xlsx','P3M_promo.xlsx','P6M_promo.xlsx','P12M_promo.xlsx']
file_name_non_promo = ['P1M_non_promo.xlsx','P3M_non_promo.xlsx','P6M_non_promo.xlsx','P12M_non_promo.xlsx']
sd_p1m = str(end_date + relativedelta(months=-1)+relativedelta(days=1))
sd_p3m = str(end_date + relativedelta(months=-3)+relativedelta(days=1))
sd_p6m = str(end_date + relativedelta(months=-6)+relativedelta(days=1))
sd_p12m = str(end_date + relativedelta(months=-12)+relativedelta(days=1))
s_date = [sd_p1m,sd_p3m,sd_p6m,sd_p12m]
promo_df = pd.DataFrame({"file_name_all":file_name_all,"file_name_promo":file_name_promo,"file_name_non_promo":file_name_non_promo,"s_date":s_date,"e_date":end_date})
promo_df

 

#promo data
for i in range(promo_df.shape[0]):
    fromDate = promo_df.s_date[i]
    toDate = str(promo_df.e_date[i])
    payload_promo = {"filters":{"hierarchies":[],"sources":[],"star_ratings":[],"brands":[],"entities":[],"topics":[],"mega_topics":[],"review_text":[],"groups":[],"fromDate":fromDate,"toDate":toDate,"selectedRange":"Custom Range","price":[],"promotion":True,"custom_filters":[],"quote_text":[]},"objectPaging":{"numberOfRows":10,"page":1,"searchText":"","sort":""},"onlyUniqueReviews":True}
    r = requests.post(url,json=payload_promo,headers = HEADERS)
    if r.status_code == 200:
        df = json.loads(r.text)
        total = int(df['total'])
        payload_promo1 = {"filters":{"hierarchies":[],"sources":[],"star_ratings":[],"brands":[],"entities":[],"topics":[],"mega_topics":[],"review_text":[],"groups":[],"fromDate":fromDate,"toDate":toDate,"selectedRange":"Custom Range","price":[],"promotion":True,"custom_filters":[],"quote_text":[]},"objectPaging":{"numberOfRows":total,"page":1,"searchText":"","sort":""},"onlyUniqueReviews":True}
        r1 = requests.post(url,json=payload_promo1,headers = HEADERS)
        if r1.status_code == 200:
            df1 = json.loads(r1.text)
            df1 = pd.DataFrame(df1['data']) #change json to dataframe
            df1.to_excel("./Promotion data/"+promo_df.file_name_promo[i],index=False)
print("      ")
print("---------------------------output files generated-------------------------------------")
print("      ")

 

#non promo data
for i in range(promo_df.shape[0]):
    fromDate = promo_df.s_date[i]
    toDate = str(promo_df.e_date[i])
    payload_promo = {"filters":{"hierarchies":[],"sources":[],"star_ratings":[],"brands":[],"entities":[],"topics":[],"mega_topics":[],"review_text":[],"groups":[],"fromDate":fromDate,"toDate":toDate,"selectedRange":"Custom Range","price":[],"promotion":False,"custom_filters":[],"quote_text":[]},"objectPaging":{"numberOfRows":10,"page":1,"searchText":"","sort":""},"onlyUniqueReviews":True}
    r = requests.post(url,json=payload_promo,headers = HEADERS)
    if r.status_code == 200:
        df = json.loads(r.text)
        total = int(df['total'])
        payload_promo1 = {"filters":{"hierarchies":[],"sources":[],"star_ratings":[],"brands":[],"entities":[],"topics":[],"mega_topics":[],"review_text":[],"groups":[],"fromDate":fromDate,"toDate":toDate,"selectedRange":"Custom Range","price":[],"promotion":False,"custom_filters":[],"quote_text":[]},"objectPaging":{"numberOfRows":total,"page":1,"searchText":"","sort":""},"onlyUniqueReviews":True}
        r1 = requests.post(url,json=payload_promo1,headers = HEADERS)
        if r1.status_code == 200:
            df1 = json.loads(r1.text)
            df1 = pd.DataFrame(df1['data']) #change json to dataframe
            df1.to_excel("./Non Promotion data/"+promo_df.file_name_non_promo[i],index=False)
print("      ")
print("---------------------------output files generated-------------------------------------")
print("      ")

 

#Combined data
for i in range(promo_df.shape[0]):
    fromDate = promo_df.s_date[i]
    toDate = str(promo_df.e_date[i])
    payload_promo = {"filters":{"hierarchies":[],"sources":[],"star_ratings":[],"brands":[],"entities":[],"topics":[],"mega_topics":[],"review_text":[],"groups":[],"fromDate":fromDate,"toDate":toDate,"selectedRange":"Custom Range","price":[],"promotion":"","custom_filters":[],"quote_text":[]},"objectPaging":{"numberOfRows":10,"page":1,"searchText":"","sort":""},"onlyUniqueReviews":True}
    r = requests.post(url,json=payload_promo,headers = HEADERS)
    if r.status_code == 200:
        df = json.loads(r.text)
        total = int(df['total'])
        payload_promo1 = {"filters":{"hierarchies":[],"sources":[],"star_ratings":[],"brands":[],"entities":[],"topics":[],"mega_topics":[],"review_text":[],"groups":[],"fromDate":fromDate,"toDate":toDate,"selectedRange":"Custom Range","price":[],"promotion":"","custom_filters":[],"quote_text":[]},"objectPaging":{"numberOfRows":total,"page":1,"searchText":"","sort":""},"onlyUniqueReviews":True}
        r1 = requests.post(url,json=payload_promo1,headers = HEADERS)
        if r1.status_code == 200:
            df1 = json.loads(r1.text)
            df1 = pd.DataFrame(df1['data']) #change json to dataframe
            df1.to_excel("./Combined data/"+promo_df.file_name_all[i],index=False)
print("      ")
print("---------------------------output files generated-------------------------------------")
print("      ")

#------------------------------------------------------------------------------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------------------------------------------------------------------------------

 

#for revuze automation P12M

 

import requests
import json
from datetime import datetime, date
import calendar
import pandas as pd
from dateutil.relativedelta import relativedelta
today = datetime.date(datetime.now()) #gives today
today_year = today.year # gives current year
today_month = today.month # gives current month
start_date = date(today_year-1,today_month,1)
end_date = date(today_year,today_month-1,calendar.monthrange(today_year, today_month-1)[1])
print(start_date)
print(end_date)

 

start_date_range = []
for i in range(1,25): #loop for startdate of last 24 months
    sd = start_date + relativedelta(months=-i)
    start_date_range.append(sd)
end_date_range = []
for i in range(1,25): #loop for enddate of last 24 months
    ed = end_date + relativedelta(months=-i)
    end_date_range.append(ed)
date_table = pd.DataFrame({"startdate":start_date_range,"enddate":end_date_range})
date_table.loc[-1] = [start_date,end_date]  # adding a row
date_table.index = date_table.index + 1  # shifting index
date_table = date_table.sort_index()  # sorting by index
date_table = date_table.iloc[0:-1,:]
date_table['startdate_month'] = date_table.startdate.apply(lambda x: x.strftime("%b")) #month MMM format
date_table['enddate_month'] = date_table.enddate.apply(lambda x: x.strftime("%b"))
date_table['startdate_year'] = date_table.startdate.apply(lambda x: x.year)
date_table['enddate_year'] = date_table.enddate.apply(lambda x: x.year)
date_table['enddate_day'] = date_table.enddate.apply(lambda x: x.day)
date_table['file_name'] = "P12M All groups_M"+(date_table.index+1).map(str)+".xlsx" #file name to be saved as
date_table

 

url = "https://gateway.revuze.it/api/groups"

 

#change Authorization and project_uuid after logging in to revuze portal, take from groups api request header

 

HEADERS = {'User-Agent':
           'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/44.0.2403.157 Safari/537.36',
                           'Accept-Language': 'en-US, en;q=0.5',
          'Authorization':'Bearer eyJhbGciOiJSUzI1NiIsImtpZCI6Ijk1NDRiYWI5Y2QwMmRiYTNlMTNkMGMwN2JmNjcwZWY1IiwidHlwIjoiSldUIn0.eyJuYmYiOjE2MzAzMDI5MjgsImV4cCI6MTYzMTM1NDEyOCwiaXNzIjoiaHR0cDovL3Nzby5yZXZ1emUuaXQiLCJhdWQiOlsiaHR0cDovL3Nzby5yZXZ1emUuaXQvcmVzb3VyY2VzIiwiZ2F0ZXdheV9wcm9kdWN0aW9uIl0sImNsaWVudF9pZCI6ImRhc2hib2FyZF91aSIsInN1YiI6IjA5Y2YyNTVmLTEyOWYtNDQ5ZS1iYzYwLThjYjk0OTA3YWE3MyIsImF1dGhfdGltZSI6MTYzMDMwMjkyNywiaWRwIjoibG9jYWwiLCJzY29wZSI6WyJvcGVuaWQiLCJwcm9maWxlIiwiZ2F0ZXdheV9wcm9kdWN0aW9uIiwib2ZmbGluZV9hY2Nlc3MiXSwiYW1yIjpbInB3ZCJdfQ.Y2zHSWJIrDFyPuIbbRnsGJrvcl5R_4JkSRlo4gZh9N0zjrBBucH3lA_DozZpwllkCLvE72ubJQgV4Bv1lFhFcbimUgGF0dJYN_440-XpRiwhggWiU7w0g1vKYZ29WLEWlO8HQNCmcVO78x-0BTiVySfQglHKYsYpFoH4k_jtRtXd7Q5fJcQCSSqLxL-c_8v_3ybjh-Cp1ci853b0l_AhMwjHopc2k1lFiOV44Gh3LZz_AsU9WRrYYVgLg56lsUcm4PX24lrWzpLf44LbyT6ORgQK2uy7d9KdroScCmm24ffV1I6FCYLcrxLZZdQA3_kUbqVT2s-44l3M7DYJPrdfgQ',
          'project_uuid':'e7b31d49-6a58-11e8-9211-baa4b8fd5ab8'}

 


for i in range(date_table.shape[0]):
    fromDate = str(date_table.startdate[i]) #fromdate to be passed as filter in api payload
    toDate = str(date_table.enddate[i]) #todate to be passed as filter in api payload
    payload = {"filters":{"hierarchies":[],"sources":[],"star_ratings":[],"brands":[],"entities":[],"topics":[],"mega_topics":[],"review_text":[],"groups":[],"fromDate":fromDate,"toDate":toDate,"selectedRange":"Custom Range","price":[],"promotion":"","custom_filters":[],"quote_text":[]},"objectPaging":{"numberOfRows":10,"page":1,"searchText":"","sort":""},"onlyUniqueReviews":True}
    r = requests.post(url,json=payload,headers = HEADERS)
    if r.status_code == 200:
        df = json.loads(r.text)
        total = int(df['total']) #get the total number of rows then pass again in the payload with only one page
        payload1 = {"filters":{"hierarchies":[],"sources":[],"star_ratings":[],"brands":[],"entities":[],"topics":[],"mega_topics":[],"review_text":[],"groups":[],"fromDate":fromDate,"toDate":toDate,"selectedRange":"Custom Range","price":[],"promotion":"","custom_filters":[],"quote_text":[]},"objectPaging":{"numberOfRows":total,"page":1,"searchText":"","sort":""},"onlyUniqueReviews":True}
        r1 = requests.post(url,json=payload1,headers = HEADERS)
        if r1.status_code == 200:
            df1 = json.loads(r1.text)
            df1 = pd.DataFrame(df1['data']) #change json to dataframe
            df1.to_excel("./P12M data/"+date_table.file_name[i],index=False) #save the output files into P12M data folder with filename
        
        
print("      ")
print("---------------------------output files generated-------------------------------------")
print("      ")
