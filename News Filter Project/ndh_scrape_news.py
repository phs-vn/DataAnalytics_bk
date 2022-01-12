import time
import pandas as pd
from selenium.webdriver.chrome.service import Service
from phs import *


class NoNewsFound(Exception):
    pass


class PageFailToLoad(Exception):
    pass


t = 2
max_wait_time = 10
ignored_exceptions = (
    ValueError,
    IndexError,
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    ElementNotInteractableException,
    PageFailToLoad,
)

PATH = r'D:\DataAnalytics\News Filter Project\chromedriver_win32\chromedriver.exe'
chrome_options = Options()

driver = webdriver.Chrome(options=chrome_options, service=Service(PATH))
driver.get('https://ndh.vn/doanh-nghiep')
scroll_pause_time = 1
screen_height = driver.execute_script("return window.screen.height;")
i = 1

type_save = []
title_save = []
url_save = []
content = []
time_save = []
df = pd.DataFrame()

# infinite scrolling
while True:
    # scroll one screen height each time
    driver.execute_script("window.scrollTo(0, {screen_height}*{i});".format(screen_height=screen_height, i=i))
    i += 1
    time.sleep(scroll_pause_time)
    # update scroll height each time after scrolled, as the scroll height can change after we scrolled the page
    scroll_height = driver.execute_script("return document.body.scrollHeight;")
    # Break the loop when the height we need to scroll to is larger than the total scroll height
    if screen_height * i > scroll_height:
        driver.find_element(By.XPATH, '//*[@id="btnLoadmore"]/a').click()
        elements = driver.find_elements(By.XPATH, '//*[@id="content-list-news"]/article')
        if len(elements) > 500:
            break
        # try:
        #     elements = driver.find_elements(By.XPATH, '//*[@id="content-list-news"]/article')
        for ele in elements:
            titles = (ele.text.split('\n'))[0]
            news_type = ((ele.text.split('\n'))[1]).split('/')[0]
            time_news = ((ele.text.split('\n'))[1]).split('/')[1]
            next_url = ele.find_element(By.TAG_NAME, 'a').get_attribute('href')

            next_driver = webdriver.Chrome(options=chrome_options, service=Service(PATH))
            next_driver.get(next_url)
            next_tags = next_driver.find_elements(By.XPATH, '/html/body/section[4]/div[2]/article')
            paragraphs = [each_next_tag.text for each_next_tag in next_tags]
            para_listToStr = ' '.join([str(elem) for elem in paragraphs])

            title_save.append(titles)
            type_save.append(news_type)
            time_save.append(time_news)
            url_save.append(next_url)
            content.append(para_listToStr)
            time.sleep(t)

            next_driver.quit()
        dictionary = {
            'Thời gian': time_save,
            'Loại tin tức': type_save,
            'Tiêu đề': title_save,
            'Link': url_save,
            'Nội dung': content
        }
        df = pd.DataFrame(dictionary)
        # except ignored_exceptions:
        #     break
else:
    driver.quit()

df.to_pickle("./News Filter Project/output_data/vietstock_data.pkl")
