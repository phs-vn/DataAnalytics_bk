from request.stock import *


class NoNewsFound(Exception):
    pass


class PageFailToLoad(Exception):
    pass


PATH = r'D:\DataAnalytics\dependency\chromedriver.exe'
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


def run():
    chrome_options = Options()
    # chrome_options.add_argument("--headless")

    driver = webdriver.Chrome(options=chrome_options, service=Service(PATH))
    driver.get('https://vietstock.vn/kinh-te.htm')

    type_save = []
    title_save = []
    url_save = []
    content = []
    time_save = []
    wait = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions)
    wait_2 = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions)

    while len(title_save) < 130:
        # elements = driver.find_elements(By.XPATH, '//*[@id="channel-container"]/*/div')
        elements = wait.until(
            EC.presence_of_all_elements_located((By.XPATH, '//*[@id="channel-container"]/*/div'))
        )
        try:
            for ele in elements:
                news_type = (ele.text.split('\n'))[0]
                titles = (ele.text.split('\n'))[1]
                time_news = (ele.text.split('\n'))[2]
                next_url = ele.find_element(By.TAG_NAME, 'a').get_attribute('href')

                chrome_options.add_argument("--headless")
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
                    print('Title: ', titles, ' + ', 'Length:', len(title_save))
                time.sleep(t)
                next_driver.quit()
        except ignored_exceptions:
            continue
        time.sleep(4)
        try:
            next_page = wait_2.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#page-next\  > a')))
            next_page.click()
        except ignored_exceptions:
            time.sleep(5)
            next_page = wait_2.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#page-next\  > a')))
            next_page.click()

        time.sleep(t)
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
    df.to_pickle(r"D:\DataAnalytics\news_analysis\output_data\vietstock_kinh-te_data-2.pickle")
