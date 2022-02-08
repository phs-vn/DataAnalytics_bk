from request.stock import *


class NoNewsFound(Exception):
    pass


class PageFailToLoad(Exception):
    pass


PATH = join(dirname(dirname(realpath(__file__))), 'dependency', 'chromedriver')
ignored_exceptions = (
    ValueError,
    IndexError,
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    ElementNotInteractableException,
    PageFailToLoad,
)
t = 1
margin_list = internal.mlist()


def run():
    chrome_options = Options()
    chrome_options.add_argument("--headless")

    title_save = []
    url_save = []
    content = []
    desc_save = []

    driver = webdriver.Chrome(options=chrome_options, service=Service(PATH))
    wait = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions)

    """
        /31/: thị trường chứng khoán
        /36/: doanh nghiệp
        /35/: bất động sản
        /34/: tài chính ngân hàng
    """
    i = 1
    while True:
        url = 'https://cafef.vn/timeline/31/trang-' + str(i) + '.chn'
        driver.get(url)

        elements = wait.until(
            EC.presence_of_all_elements_located((By.XPATH, '/html/body/li'))
        )

        for ele in elements:
            try:
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
                    url_save.append(next_url)
                    title_save.append(title)
                    desc_save.append(description)
                    content.append(para_listToStr)
                    print('URL:', next_url, '+', 'Length:', len(content))
                time.sleep(t)
            except ignored_exceptions:
                pass
        i = i + 1
        time.sleep(t)

        if len(content) >= 1000:
            break

    dictionary = {
        'News link': url_save,
        'Tiêu đề': title_save,
        'Mô tả': desc_save,
        'Nội dung': content
    }
    df = pd.DataFrame(dictionary)

    df.to_pickle(r"D:\DataAnalytics\news_analysis\output_data\cafef_data_3\cafef_thi-truong-chung-khoan_data_3.pickle")
