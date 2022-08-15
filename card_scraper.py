import os
from selenium import webdriver
from selenium.webdriver.common.by import By
import pyautogui as pag
import pandas as pd
from openpyxl import Workbook

product = str(input('Choose a pokemon product: '))
os.environ['PATH'] += r"C:/Users/vkpha/Drivers"
driver = webdriver.Chrome()

driver.implicitly_wait(3)
#Pulls up site pokemon section
driver.get('https://shop.tcgplayer.com/price-guide/pokemon/base-set')
pokemon_section = driver.find_element(By.LINK_TEXT, 'Pokemon')

#Goes to Pokemon Advanced Search
pokemon_section.click()

product_name = driver.find_element(By.ID, 'ProductName')
product_name.send_keys(product)

#Find all rarities
rarities = driver.find_elements(By.ID, 'Rarity')
for rarity in rarities:
    rarity.click()


#Search
search = driver.find_element(By.CSS_SELECTOR, 'input[value=\'Search\']')
search.click()
'''
#Refresh - Site Can't Be reached
try:
    reload = driver.find_element(By.ID, 'reload-button')
    reload.click()
except NoSuchElementException:
    print('No Reload')
'''
#Sorting
#sorting = driver.find_element(By.CSS_SELECTOR, 'select[aria-label=\'Sort products by:\']')
#sorting.click()

'''
#To Highest Price
for i in range(3):
    pag.press('down')
pag.press('Enter')
'''

scraped = []
search_results = driver.find_element(By.CLASS_NAME, 'search-results')
mult_search_result = search_results.find_elements(By.CLASS_NAME, 'search-result')
for search_result in mult_search_result:
    scraped.append(search_result.text)

df = pd.DataFrame()
df['product'] = scraped[0::1]
df.to_excel('result.xlsx', index=False)