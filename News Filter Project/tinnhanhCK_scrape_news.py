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
    PageFailToLoad
)

i = 1
t = 2

PATH = r'D:\DataAnalytics\News Filter Project\chromedriver_win32\chromedriver.exe'
chrome_options = Options()
chrome_options.add_argument("--headless")

mr_data = pd.read_excel(r'D:\DataAnalytics\News Filter Project\Intranet list_07.01.2022 final.xlsx',
                        sheet_name='Sheet1', usecols="B")
margin_list = mr_data['Mã CK'].tolist()
margin_list = [' {0}'.format(elem) for elem in margin_list]

driver = webdriver.Chrome(options=chrome_options, service=Service(PATH))
driver.get('https://tinnhanhchungkhoan.vn/doanh-nghiep/')

title_save = []
url_save = []
content = []

df = pd.DataFrame()

while True:
    elements = driver.find_elements(By.XPATH, '/html/body/div[1]/div/div[1]/div/div/div[1]/div[2]/div[1]/article')
    driver.execute_script(
        "return arguments[0].scrollIntoView(true);",
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="viewmore"]'))))
    driver.find_element(By.XPATH, '//*[@id="viewmore"]').click()
    if len(elements) >= 350:
        break

for ele in elements:
    print('Index: ', elements.index(ele) + 1)
    try:
        next_url = ele.find_element(By.TAG_NAME, 'a').get_attribute('href')
        print('URL: ', next_url)
    except ignored_exceptions:
        break
    next_driver = webdriver.Chrome(options=chrome_options, service=Service(PATH))
    next_driver.get(next_url)

    title = next_driver.find_element(
        By.CSS_SELECTOR,
        '#article-container > div > div.main-column.article > h1').text
    description = next_driver.find_element(
        By.CSS_SELECTOR,
        '#article-container > div > div.main-column.article > div.article__sapo.cms-desc').text
    body = next_driver.find_elements(
        By.CSS_SELECTOR,
        '#article-container > div > div.main-column.article > div.article__body.cms-body > p')
    paragraphs = [each_body.text for each_body in body]
    para_listToStr = ' '.join([str(elem) for elem in paragraphs])
    check = [mr in para_listToStr for mr in margin_list]
    if any(check):
        title_save.append(title)
        url_save.append(next_url)
        content.append(description + ' ' + para_listToStr)
        print('Content Length: ', len(content))
    if len(content) > 200:
        break
    time.sleep(t)


dictionary = {
    'Link': url_save,
    'Tiêu đề': title_save,
    'Nội dung': content
}
df = pd.DataFrame(dictionary)
# df.to_pickle("./News Filter Project/output_data/tinnhanhCK_doanh-nghiep_data.pickle")
