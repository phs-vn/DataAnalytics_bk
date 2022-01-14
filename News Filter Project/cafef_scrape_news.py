from selenium.webdriver.chrome.service import Service
from phs import *


class NoNewsFound(Exception):
    pass


class PageFailToLoad(Exception):
    pass


ignored_exceptions = (
    ValueError,
    IndexError,
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    ElementNotInteractableException,
    PageFailToLoad,
)
t = 2

mr_data = pd.read_excel(r'D:\DataAnalytics\News Filter Project\Intranet list_07.01.2022 final.xlsx',
                        sheet_name='Sheet1', usecols="B")
margin_list = mr_data['Mã CK'].tolist()
margin_list = [' {0}'.format(elem) for elem in margin_list]

chrome_options = Options()
chrome_options.add_argument("--headless")

title_save = []
url_save = []
content = []
desc_save = []

PATH = r'D:\DataAnalytics\News Filter Project\chromedriver_win32\chromedriver.exe'
driver = webdriver.Chrome(options=chrome_options, service=Service(PATH))
wait = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions)

df = pd.DataFrame()
i = 1

"""
    /31/: thị trường chứng khoán
    /36/: doanh nghiệp
    /35/: bất động sản
    /34/: tài chính ngân hàng
"""
while True:
    url = 'https://cafef.vn/timeline/34/trang-' + str(i) + '.chn'
    driver.get(url)

    elements = wait.until(
        EC.presence_of_all_elements_located((By.XPATH, '/html/body/li'))
    )
    for ele in elements:
        title = (ele.text.split('\n'))[0]
        description = (ele.text.split('\n'))[2]

        next_url = ele.find_element(By.TAG_NAME, 'a').get_attribute('href')
        next_driver = webdriver.Chrome(options=chrome_options, service=Service(PATH))
        next_driver.get(next_url)

        next_elements = next_driver.find_elements(By.XPATH, '//*[@id="mainContent"]/p')
        paragraphs = [each_next_tag.text for each_next_tag in next_elements]
        para_listToStr = ' '.join([str(elem) for elem in paragraphs])
        check = [mr in para_listToStr for mr in margin_list]

        if any(check):
            print("Add title: ", title)
            url_save.append(next_url)
            title_save.append(title)
            desc_save.append(description)
            content.append(para_listToStr)
        time.sleep(t)
    i = i + 1
    time.sleep(t)

    if len(content) == 30:
        break

dictionary = {
    'News link': url_save,
    'Tiêu đề': title_save,
    'Mô tả': desc_save,
    'Nội dung': content
}
df = pd.DataFrame(dictionary)

# df.to_pickle("./News Filter Project/output_data/cafef_tai-chinh-ngan-hang_data.pickle")
