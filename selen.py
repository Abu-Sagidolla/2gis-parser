# -*- coding: utf-8 -*-
import csv
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep
from selenium.common.exceptions import NoSuchElementException,ElementNotInteractableException,ElementClickInterceptedException
import re
import pyautogui as pg
from bs4 import BeautifulSoup as bs
import openpyxl


driver = webdriver.Chrome("yandexdriver.exe")
m = 'https://2gis.ru/moscow'
#term = 'Дмитрия Донского бульвар, 11'

numb = []
NUMBERS = []
names = []
sites = []
des = []
count = 0

def system(term):
    
    driver.get(m)
    search = driver.find_element_by_xpath('//*[@id="root"]/div/div/div[1]/div[1]/div[2]/div/div/div[1]/div/div/div/div/div[2]/form/div/input')
    search.send_keys(term)
    #search.send_keys(Keys.ENTER)
    pg.press('enter')
    sleep(1.8)
    try:
        search = driver.find_element_by_xpath('//*[@id="root"]/div/div/div[1]/div[1]/div[2]/div[2]/div/div/div/div/div[2]/div[2]/div/div/div/div/div[1]/div/div[2]/div[1]/div/div/div[1]/div[2]/div/a')
        search.click()
    except NoSuchElementException:
        pg.press('tab',presses=12,interval=0.3)
        pg.press('enter')
        sleep(0.9)
        search = driver.find_element_by_xpath('//*[@id="root"]/div/div/div[1]/div[1]/div[2]/div[2]/div/div/div/div/div[2]/div[2]/div/div/div/div/div[1]/div/div[2]/div[1]/div[3]/div/div[1]/div[2]/div')
        search.click()


    info_urls = driver.current_url
    print(info_urls)
    
    sleep(0.8)
    pg.press('tab',presses=28,interval=0.2)
    pg.press('down',presses=100,interval=0.07)
    search = driver.find_elements_by_class_name('_vhuumw')
    spans = driver.find_elements_by_class_name('_6vzrncr')
    temple = []

    for i in search:
        temple.append(i.get_attribute('href'))
    html = driver.page_source.encode('utf-8')
    
    
    html = html.decode('utf-8')
    
    soup = bs(html,'html.parser')

    a = soup.findAll('a',attrs={'class':'_vhuumw'})
    for j in a:
        if not re.search('x',str(j)) and not re.search('филиал',str(j)) and not re.search('Реклама',str(j)  or not re.search('Заказать',str(j))):
            names.append(j.text)
            #urls.append(j['href'])
    print(names)

    global clear
    clear = [i for i in temple if i != None and re.search('firm',str(i))]
        
    #print(clear)
    #print(len(clear))
    return 'parsed )))'

#system(term)

def grabber(url):
    
    driver.get(url)
    

    try:
        try:
            botton = driver.find_element_by_xpath('//*[@id="root"]/div/div/div[1]/div[1]/div[2]/div/div/div[2]/div/div/div[2]/div[2]/div/div[1]/div/div/div[1]/div/div[2]/div[2]/div[2]/div[1]/div/div[3]/div[2]/div/button')
            botton.click()
        except ElementClickInterceptedException:
            pass   
    except NoSuchElementException:
        print('nothing')
        pass

    temple = []
    html = driver.page_source.encode('utf-8')
    html = html.decode('utf-8')
    soup = bs(html,'html.parser')
    a = soup.findAll('a',attrs={'class':'_ke2cp9k'})
    nums = [i['href'] for i in a if re.search('tel',i['href'])]
    for i in nums:
        temple.append(i[4:])
        numbs[counts]
        count += 1    
    a = soup.findAll('a',attrs={'class':'_vhuumw'})
    site = [i.text for i in a if re.search('.com',i.text) or re.search('.ru',i.text) or re.search('www',i.text)]
    print(nums)
    if len(site) < 1:
        sites.append('')
    else:
        for i in site:
            sites.append(i)
    a = soup.findAll('span',attrs={'class':'_oqoid'})
    try :
        des.append(a[1].text)
    except IndexError:
        des.append(a[0].text)
        pass    
           
        
def main():
    global k
    k = 'Список спарсенных улиц'
    for j in clear:
        print(grabber(j))   

s = []
    

files = open('streets.txt','r',encoding='utf-8')
for lines in files:
    system(lines)
    for u in clear:
        print(grabber(u))
    main()
    for i in names:
        s.append(lines)
response = [i for i in zip(s,names,numb,sites,des)]    
wb = openpyxl.Workbook()
wb.create_sheet(title = 'Первый лист', index = 0)
sheet = wb['Первый лист']
sheet.append(['Код','Физлицо','Водитель','Группа','Наименования','Юридическое наименование','Юридический адрес','Фактический адрес','Паспорт','Должность','Организация','Область','Район','ИИН','КПП','ОКПО','ОРГН','ОКВЭД','РабТел1','РабТел2','МобТел1','МобТел2','Факс','Email','Вебсайт','БИК','Банк','К/с','Р/с','Скидка','Заметки','Тип организации'])
for i in range(len(response)):
    sheet.append(['','','','Телемаркетинг',response[i][1],'','',response[i][0],'','','','','','','','','','',response[i][2],'','','','','',response[i][3],'','','','','','',response[i][4]])    
    wb.save('ParsedList.xlsx')    
driver.close()    





