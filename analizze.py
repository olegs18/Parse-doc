# pip install scrapeasy
# pip install requests
from scrapeasy import Website, Page
import os

def init():
    # Инициализация веб-сайта
    web = Website("https://tikocash.com/solange/index.php/2022/04/13/how-do-you-control-irrational-fear-and-overthinking/")

    # Получение ссылок всех подсайтов
    links = web.getSubpagesLinks()

    # Поиск медиа
    images = web.getImages()

    # Скачать медиа
    web.download("img", "fahrschule/images")

    # В результате мы получим список всех доменов, которые где-то ссылаются
    domains = web.getLinks(intern=False, extern=False, domain=True)

    # Получение связанных доменов (внешние ссылки)
    domains = web.getLinks(intern=False, extern=True, domain=False)

    # Загрузка видео
    # Инициализация страницы
    w3=Page("https://www.w3schools.com/html/html5_video.asp")
    w3.download("video", "w3/videos")
    video_links = w3.getVideos()

    # Загрузка файлов других типов (например, pdf или ics)
    calendar_links = Page("https://www.icu.uzh.ch/events/id/207").get("ics")
    Page("http://mathcourses.ch/mat182.html").download("pdf", "mathcourses/pdf-files")

# w3=Page("https://legalway.org/admin-posluga/n-01418/", False) 
# HTML = w3.getHTML()
# print(HTML)
# if HTML:
#     with open(r"html.txt", "w", encoding="utf-8") as file:
#         file.write(HTML)
    

# os.system('curl -o output.txt --cookie "greeting=hello" -k https://legalway.org/admin-posluga/n-01418/')
def get_data(p):
    url = f'https://legalway.org/admin-posluga/n-{p}/'

    from bs4 import BeautifulSoup
    import requests
    import json
    page = requests.get(url)

    soup = BeautifulSoup(page.text, "html.parser")
    # allNews = soup.find_all('div', class_={'border-bottom', 'py-2'})
    allNews = soup.find_all('div', class_=lambda x: x and 'border-bottom' in x and 'py-2' in x and 'text-center' not in x)
    filteredNews = []
    for data in allNews:
        name = ''
        if data.find('h5', class_='fw-bold') is not None:
            name = data.text
        ls = name.strip().split('\n')
        n = ls.pop(0)
        filteredNews.append({n: ' '.join(ls) })
    # Сохраняем список в JSON файл
    with open(f'{p}.json', 'w') as json_file:
        json.dump(filteredNews, json_file)
# print(filteredNews)
def get_all_data_523():
    p = 'all_523'
    url = f'https://zakon.rada.gov.ua/laws/show/523-2014-%D1%80.frame'

    from bs4 import BeautifulSoup
    import requests
    import json
    page = requests.get(url)

    soup = BeautifulSoup(page.text, "html.parser")
    # allNews = soup.find_all('div', class_={'border-bottom', 'py-2'})
    allNews = soup.find('table', cellpadding='1').find_all('tr', valign='top')
    print(allNews)
    filteredNews = []
    for data in allNews:
        name = n = name2 = ''
        el = data.select('td:nth-child(2)')
        if el is not None:
            if len(el) > 0:
                n = el[0].text.strip()
                if n == '' or '{' in n:
                    continue
                else:
                    print([n])
            else:
                continue
        else:
            continue
        el = data.select('td:nth-child(3)')
        if el is not None:
            if len(el) > 0:
                name = el[0].text.strip()
        el = data.select('td:nth-child(5)')
        if el is not None:
            if len(el) > 0:
                name2 = el[0].text.strip()
        filteredNews.append([n, name, name2])
        # print(n, name)
    # Сохраняем список в JSON файл
    with open(f'{p}.json', 'w') as json_file:
        json.dump(filteredNews, json_file)
# 
# get_all_data_523()

def get_data_523():
    import json
    with open('all_523.json') as json_data:
        d = json.load(json_data)
        json_data.close()
    return d
# get_data_523()