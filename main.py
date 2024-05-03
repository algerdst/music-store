import requests
from bs4 import BeautifulSoup
import time


url='https://www.musicstore.com/ru_RU/RUB/-/ST-/cat-GITARRE-GITEGSTRAT'


headers={
    'user-agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 YaBrowser/24.4.0.0 Safari/537.36'
}

response=requests.get(url, headers=headers).text
soup=BeautifulSoup(response, 'lxml')
pages=int(soup.find('ul','paging-list').find_all('li')[-2].text)

for page in range(1, pages+1):
    if page==1:
        response=response
    else:
        link=f"{url}/{page}"
        response=requests.get(link, headers=headers).text
    soup=BeautifulSoup(response, 'lxml')
    blocks=soup.find_all('div', class_='ident')
    for block in blocks:
        item_link=block.find('a', 'js-tracking')['href']
        name=block.find('a', 'name').text
        print(item_link, ' ', name)
    time.sleep(1)