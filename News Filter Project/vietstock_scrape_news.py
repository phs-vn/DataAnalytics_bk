import pandas as pd
from selenium.webdriver.chrome.service import Service
from phs import *


PATH = r'D:\DataAnalytics\News Filter Project\chromedriver_win32\chromedriver.exe'
chrome_options = Options()
# chrome_options.add_argument("--headless")

driver = webdriver.Chrome(options=chrome_options, service=Service(PATH))
driver.get('https://vietstock.vn/doanh-nghiep.htm')

type_save = []
title_save = []
url_save = []
content = []
time_save = []
df = pd.DataFrame()
elements = []

while df.shape[0] < 300:
    element = driver.find_elements(By.XPATH, '//*[@id="channel-container"]/*/div')

    for ele in elements:
        news_type = (ele.text.split('\n'))[0]
        titles = (ele.text.split('\n'))[1]
        time_news = (ele.text.split('\n'))[2]
        next_url = ele.find_element(By.TAG_NAME, 'a').get_attribute('href')

        next_driver = webdriver.Chrome(options=chrome_options, service=Service(PATH))
        next_driver.get(next_url)
        next_tags = next_driver.find_elements(By.XPATH, '//*[@id="vst_detail"]/p')
        paragraphs = [each_next_tag.text for each_next_tag in next_tags]
        para_listToStr = ' '.join([str(elem) for elem in paragraphs])

        type_save.append(news_type)
        title_save.append(titles)
        time_save.append(time_news)
        url_save.append(next_url)
        content.append(para_listToStr)

        next_driver.quit()

    dictionary = {
        'Thời gian': time_save,
        'Loại tin tức': type_save,
        'Tiêu đề': title_save,
        'Link': url_save,
        'Nội dung': content
    }
    df = pd.DataFrame(dictionary)

    time.sleep(5)
    nextpage_button = driver.find_element(By.XPATH, '//*[@id="page-next "]/a')
    nextpage_button.click()
else:
    driver.quit()

