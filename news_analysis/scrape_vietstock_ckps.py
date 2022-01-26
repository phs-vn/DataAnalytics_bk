"""
- Date format: dd/mm/yyyy

"""

from request import *


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
    ElementClickInterceptedException,
    NoNewsFound,
)

t = 0.5


def get_price(date_input: str, ticker: str):
    now = dt.datetime.now()

    if '-' in date_input:
        date_input = date_input.replace('-', '/')

    if dt.datetime.strptime(date_input, '%d/%m/%Y') <= now:
        ticker = ticker.upper()

        url = 'https://finance.vietstock.vn/chung-khoan-phai-sinh/thong-ke-giao-dich.htm?'
        chrome_options = Options()
        driver = webdriver.Chrome(options=chrome_options, service=Service(PATH))
        driver.get(url)
        wait = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions)

        from_date = driver.find_element(By.XPATH, '//*[@id="txtFromDate"]/input')
        from_date.clear()
        from_date.send_keys(date_input)

        to_date = driver.find_element(By.XPATH, '//*[@id="txtToDate"]/input')
        to_date.clear()
        to_date.send_keys(date_input)
        time.sleep(t)

        ma_ck = driver.find_element(
            By.XPATH,
            '//*[@id="trading-result"]/div/div[1]/div[1]/div/div[1]/div/div[3]/div/span'
        )
        ma_ck.click()
        input_field = driver.find_element(By.XPATH, '/html/body/span/span/span[1]/input')
        input_field.send_keys(ticker)
        input_field.send_keys(Keys.ENTER)

        while True:
            try:
                btn_xem = wait.until(EC.presence_of_element_located((
                    By.XPATH,
                    '//*[@id="trading-result"]/div/div[1]/div[1]/div/div[2]/button'
                )))
                btn_xem.click()
                break
            except ignored_exceptions:
                continue

        time.sleep(t)

        gia_dong_cua = driver.find_element(
            By.XPATH,
            '//*[@id="statistic-price"]/table/tbody/tr/td[6]/span'
        ).text
        gia_dong_cua = float(gia_dong_cua.replace(',', ''))
        driver.quit()

        return gia_dong_cua
    else:
        pass

# get_price('27-12-2021', 'VN30F2203')
