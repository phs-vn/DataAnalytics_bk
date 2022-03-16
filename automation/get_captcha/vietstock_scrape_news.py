from request.stock import *


class NoNewsFound(Exception):
    pass


class PageFailToLoad(Exception):
    pass


# PATH = r'D:\DataAnalytics\dependency\chromedriver.exe'
PATH = join(dirname(dirname(realpath(__file__))), 'dependency', 'chromedriver')
ignored_exceptions = (
    ValueError,
    IndexError,
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    ElementNotInteractableException,
    PageFailToLoad,
    ElementClickInterceptedException
)
margin_list = internal.mlist()
t = 2

type_save = []
title_save = []
url_save = []
content = []
time_save = []


def run():
    chrome_options = Options()
    # chrome_options.add_argument('--headless')
    driver = webdriver.Chrome(options=chrome_options, service=Service(PATH))
    driver.get('https://vietstock.vn/doanh-nghiep.htm')
    wait = WebDriverWait(driver, 5, ignored_exceptions=ignored_exceptions)

    while len(title_save) < 150:
        # elements = driver.find_elements(By.XPATH, '//*[@id="channel-container"]/*/div')
        elements_1 = wait.until(
            EC.presence_of_all_elements_located((By.XPATH, '//*[@id="channel-container"]/*/div'))
        )

        for ele in elements_1:
            news_type = (ele.text.split('\n'))[0]
            titles = (ele.text.split('\n'))[1]
            time_news = (ele.text.split('\n'))[2]
            next_url = ele.find_element(By.TAG_NAME, 'a').get_attribute('href')
            next_driver = webdriver.Chrome(options=chrome_options, service=Service(PATH))
            next_driver.get(next_url)

            wait_2 = WebDriverWait(next_driver, 5, ignored_exceptions=ignored_exceptions)
            # next_tags = next_driver.find_elements(By.XPATH, '//*[@id="vst_detail"]/p')
            try:
                elements_2 = wait_2.until(
                    EC.presence_of_all_elements_located((By.XPATH, '//*[@id="vst_detail"]/p'))
                )
                paragraphs = list(filter(None, [each_next_tag.text for each_next_tag in elements_2]))
                para_listToStr = ' '.join([str(elem) for elem in paragraphs])
                check = [mr in para_listToStr for mr in margin_list]
                if any(check):
                    type_save.append(news_type)
                    title_save.append(titles)
                    time_save.append(time_news)
                    url_save.append(next_url)
                    content.append(para_listToStr)
                    print('Title: ', titles, '\n', 'Length:', len(title_save))
                else:
                    pass
                next_driver.quit()
                time.sleep(t)
            except ignored_exceptions:
                next_driver.refresh()

        while True:
            try:
                next_page = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#page-next\  > a')))
                next_page.click()
                time.sleep(t)
                break
            except ignored_exceptions:
                continue
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
    df.to_pickle(r"D:\DataAnalytics\news_analysis\output_data\vietstock\vietstock_data_3"
                 r"\vietstock_doanh-nghiep_data-3.pickle")
