import requests
from bs4 import BeautifulSoup
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import glob
import os
import openpyxl


url=input('Введите ссылку')
headers = {
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 YaBrowser/24.4.0.0 Safari/537.36'
}


def make_description(item_title):
    with open('description_text.txt', 'r', encoding='utf-8-sig') as file:
        for i in file:
            description_text=i.replace('item_title', item_title)
    return description_text


with open('title_text.txt', 'r', encoding='utf-8-sig') as file:
    for i in file:
        title_text = i

def get_links():
    """
    Собирает ссылки на товары
    """
    links = []
    response = requests.get(url, headers=headers).text
    soup = BeautifulSoup(response, 'lxml')
    pages = int(soup.find('ul', 'paging-list').find_all('li')[-2].text)
    for page in range(pages):
        if page == 0:
            response = response
        else:
            if '?SortingAttribute' not in url:
                link = f"{url}/{page}"
            else:
                link = url.split('?SortingAttribute')
                link = f"{link[0]}/{page}?SortingAttribute{link[1]}"
            response = requests.get(link, headers=headers).text
        soup = BeautifulSoup(response, 'lxml')
        blocks = soup.find_all('div', class_='ident')
        for block in blocks:
            item_link = block.find('a', 'js-tracking')['href']
            links.append(item_link)
        time.sleep(1)
    return links

def get_info():
    file = []
    path = os.getcwd()
    for filename in glob.glob(os.path.join(path, '*.xlsx')):
        file.append(filename)
    filename = file[0]
    book = openpyxl.load_workbook(filename)
    sheet = book.active
    links=get_links()
    with webdriver.Chrome() as browser:
        row = 2
        count=0
        for item_link in links:
            # print(item_link)
            browser.get(item_link)
            time.sleep(1)
            title = browser.find_element(By.CSS_SELECTOR, 'h1').text
            description = make_description(title)
            title = title_text+' '+browser.find_element(By.CSS_SELECTOR, 'h1').text
            features = browser.find_element(By.CSS_SELECTOR, 'div.feature-box').find_elements(By.TAG_NAME, 'li')
            for feature in features:
                description += '\n' + feature.text
            price = browser.find_element(By.CLASS_NAME, 'js-product-price-box').find_element(By.CSS_SELECTOR,
                                                                                             'span.kor-product-sale-price-value').text.strip(
                '€').strip()
            article=browser.find_element(By.CSS_SELECTOR, 'span.artnr').text.strip('Товар: ')
            images=[]
            images_block=browser.find_element(By.CSS_SELECTOR, 'div.product-image').find_elements(By.CSS_SELECTOR, 'div.slick-slide')
            for image in images_block:
                try:
                    image=image.find_element(By.TAG_NAME,'img')
                    image_link=image.get_attribute('src')
                except:
                    break
                try:
                    image_link=image_link.replace('0160', '0640')
                except:
                    image_link=image.get_attribute('data-lazy')
                    image_link='https:'+image_link.replace('0160', '0640')
                if len(images)>=10:
                    break
                else:
                    images.append(image_link)
            images=' | '.join(images)
            yotube_request='https://www.youtube.com/results?search_query='+title.replace(' ','+').replace('&','%26')
            browser.get(yotube_request)
            time.sleep(1)
            video_youtube_link=''
            video_youtube_links=browser.find_elements(By.ID, 'video-title')
            for youtube_link in video_youtube_links:
                youtube_link_title=youtube_link.text
                if title.lower() in youtube_link_title.lower():
                    video_youtube_link=youtube_link.get_attribute('href')
                    break
            sheet.cell(column=3, row=row).value = article
            sheet.cell(column=4, row=row).value = video_youtube_link
            sheet.cell(column=5, row=row).value = price
            sheet.cell(column=13, row=row).value = title
            sheet.cell(column=15, row=row).value = description
            sheet.cell(column=20, row=row).value = images
            row += 1
            book.save(filename)
            count+=1
            print(f'Осталось собрать ссылок {len(links)-count}')

get_info()