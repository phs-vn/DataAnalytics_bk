from request_phs.stock import *


class NoNewsFound(Exception):
    pass


class PageFailToLoad(Exception):
    pass


PATH = r'D:\DataAnalytics\chromedriver_win32\chromedriver.exe'
ignored_exceptions = (
    ValueError,
    IndexError,
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    ElementNotInteractableException,
    PageFailToLoad,
    NoSuchWindowException
)

t = 2
margin_list = internal.mlist()


def run():
    idx = 0
    chrome_options = Options()

    urls = []

    title_save = []
    url_save = []
    content = []

    driver = webdriver.Chrome(options=chrome_options, service=Service(PATH))
    driver.get('https://tinnhanhchungkhoan.vn/doanh-nghiep/')
    wait = WebDriverWait(driver, 30, ignored_exceptions=ignored_exceptions)

    while True:
        elements = driver.find_elements(By.XPATH, '/html/body/div[1]/div/div[1]/div/div/div[1]/div[2]/div[1]/article')
        print(len(elements))
        driver.execute_script(
            "return arguments[0].scrollIntoView(true);",
            wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="viewmore"]'))))
        driver.find_element(By.XPATH, '//*[@id="viewmore"]').click()
        if len(elements) >= 350:
            break

    for ele in elements:
        lst_url = ele.find_element(By.TAG_NAME, 'a').get_attribute('href')
        urls.append(lst_url)

    while len(content) < 150:
        try:
            print('Index:', idx)
            next_url = urls[idx]
            chrome_options.add_argument("--headless")
            next_driver = webdriver.Chrome(options=chrome_options, service=Service(PATH))
            next_driver.get(next_url)

            time.sleep(t)

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
                print('Url:', next_url, '+', 'Length:', len(content))
        except ignored_exceptions:
            pass
        idx += 1
    driver.quit()

    dictionary = {
        'Link': url_save,
        'Tiêu đề': title_save,
        'Nội dung': content
    }
    df = pd.DataFrame(dictionary)
    df.to_pickle(r"D:\DataAnalytics\News Filter Project\output_data\tinnhanhCK_doanh-nghiep_data_2.pickle")
