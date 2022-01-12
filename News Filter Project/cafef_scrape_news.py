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


def scroll_down(driver):
    """A method for scrolling the page."""

    # Get scroll height.
    last_height = driver.execute_script("return document.body.scrollHeight")

    while True:

        # Scroll down to the bottom.
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

        # Wait to load the page.
        time.sleep(2)

        # Calculate new scroll height and compare with last scroll height.
        new_height = driver.execute_script("return document.body.scrollHeight")

        if new_height == last_height:
            break

        last_height = new_height


PATH = r'D:\DataAnalytics\News Filter Project\chromedriver_win32\chromedriver.exe'
chrome_options = Options()

driver = webdriver.Chrome(options=chrome_options, service=Service(PATH))
driver.get('https://cafef.vn/thi-truong-chung-khoan.chn')
wait = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions)

scroll_down(driver)
next_page = wait.until(
    EC.presence_of_element_located((
        By.XPATH, '//*[@id="form1"]/div[2]/div[4]/div[2]/div[1]/div[3]/div[4]/a'
    )))
next_page.click()