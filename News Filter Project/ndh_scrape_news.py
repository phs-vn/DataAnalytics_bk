from selenium.webdriver.chrome.service import Service
from phs import *


class NoNewsFound(Exception):
    pass


class PageFailToLoad(Exception):
    pass


scroll_pause_time = 1
max_wait_time = 10
i = 1
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
chrome_options.add_argument("--headless")

driver = webdriver.Chrome(options=chrome_options, service=Service(PATH))
driver.get('https://ndh.vn/tai-chinh')

screen_height = driver.execute_script("return window.screen.height;")

mr_data = pd.read_excel(r'D:\DataAnalytics\News Filter Project\Intranet list_07.01.2022 final.xlsx',
                        sheet_name='Sheet1', usecols="B")
margin_list = mr_data['Mã CK'].tolist()
margin_list = [' {0}'.format(elem) for elem in margin_list]

type_save = []
title_save = []
url_save = []
content = []
df = pd.DataFrame()

# infinite scrolling
while True:
    elements = driver.find_elements(By.XPATH, '//*[@id="content-list-news"]/article')
    # scroll one screen height each time
    driver.execute_script("window.scrollTo(0, {screen_height}*{i});".format(screen_height=screen_height, i=i))
    time.sleep(scroll_pause_time)
    # update scroll height each time after scrolled, as the scroll height can change after we scrolled the page
    scroll_height = driver.execute_script("return document.body.scrollHeight;")
    # Break the loop when the height we need to scroll to is larger than the total scroll height

    if screen_height * i > scroll_height:
        driver.find_element(By.XPATH, '//*[@id="btnLoadmore"]/a').click()
    if len(elements) >= 250:
        break
    i += 1

for ele in elements:
    title = (ele.text.split('\n'))[0]
    news_type = ((ele.text.split('\n'))[1]).split('/')[0]
    next_url = ele.find_element(By.TAG_NAME, 'a').get_attribute('href')

    next_driver = webdriver.Chrome(options=chrome_options, service=Service(PATH))
    next_driver.get(next_url)
    next_tags = next_driver.find_elements(By.XPATH, '/html/body/section[4]/div[2]/article')
    paragraphs = [each_next_tag.text for each_next_tag in next_tags]
    para_listToStr = ' '.join([str(elem) for elem in paragraphs])
    check = [mr in para_listToStr for mr in margin_list]

    if any(check):
        print("Add title: ", title, ' + ', 'Length: ', len(content))
        title_save.append(title)
        type_save.append(news_type)
        url_save.append(next_url)
        content.append(para_listToStr)
    if len(content) > 40:
        break
    time.sleep(t)

dictionary = {
    'Link': url_save,
    'Loại tin tức': type_save,
    'Tiêu đề': title_save,
    'Nội dung': content
}
df = pd.DataFrame(dictionary)

# df.to_pickle("./News Filter Project/output_data/ndh_tai-chinh_data.pickle")
