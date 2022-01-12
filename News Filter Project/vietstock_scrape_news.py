from selenium.webdriver.chrome.service import Service
from request_phs.stock import *


class NoNewsFound(Exception):
    pass


class PageFailToLoad(Exception):
    pass


mr_data = pd.read_excel(r'D:\DataAnalytics\News Filter Project\Intranet list_07.01.2022 final.xlsx',
                        sheet_name='Sheet1', usecols="B")
margin_list = mr_data['Mã CK'].tolist()

# mr_data = internal.mlist()

t = 2
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
# chrome_options.add_argument("--headless")

driver = webdriver.Chrome(options=chrome_options, service=Service(PATH))
driver.get('https://vietstock.vn/tai-chinh.htm')

type_save = []
title_save = []
url_save = []
content = []
time_save = []
df = pd.DataFrame()
wait = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions)
wait_2 = WebDriverWait(driver, 90, ignored_exceptions=ignored_exceptions)

while len(title_save) < 30:
    # elements = driver.find_elements(By.XPATH, '//*[@id="channel-container"]/*/div')
    elements = wait.until(
        EC.presence_of_all_elements_located((By.XPATH, '//*[@id="channel-container"]/*/div'))
    )

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
        check = [mr in para_listToStr for mr in margin_list]
        if any(check):
            type_save.append(news_type)
            title_save.append(titles)
            time_save.append(time_news)
            url_save.append(next_url)
            content.append(para_listToStr)
        time.sleep(t)
        next_driver.quit()
    time.sleep(8)
    next_page = wait_2.until(EC.presence_of_element_located((By.XPATH, '//*[@id="page-next "]/a')))
    next_page.click()

    time.sleep(t)

    # nextpage_button = driver.find_element(By.XPATH, '//*[@id="page-next "]/a')
    # nextpage_button.click()
else:
    driver.quit()

dictionary = {
    'Thời gian': time_save,
    'Loại tin tức': type_save,
    'Tiêu đề': title_save,
    'Link': url_save,
    'Nội dung': content
}
df = pd.DataFrame(dictionary)
df.to_pickle("./News Filter Project/output_data/vietstock_tai-chinh_data.pickle")
