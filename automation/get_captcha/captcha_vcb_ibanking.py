from automation.get_captcha import *


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
    ElementClickInterceptedException
)

# PATH = join(dirname(dirname(realpath(__file__))), 'dependency', 'chromedriver')
PATH = 'D:/DataAnalytics/dependency/chromedriver.exe'
chrome_options = Options()
driver = webdriver.Chrome(options=chrome_options, service=Service(PATH))
wait = WebDriverWait(driver, 5, ignored_exceptions=ignored_exceptions)


# test get captcha, loc it based on condition and put it in the input box on web
def run():
    i=1
    driver.get('https://www.vietcombank.com.vn/IBanking20/default.aspx?&cc=578665076A9FCED7FFF2D95FD191F61E')
    while True:
        xpath = '//*[@alt="Chuỗi bảo mật"]'
        CAPTCHA = wait.until(EC.presence_of_element_located((By.XPATH,xpath)))

        imgPATH = join(fr'C:\Users\namtran\Share Folder\Get Captcha\vcb_ibanking_dataset\captcha_{i}.png')
        CAPTCHA.screenshot(imgPATH)

        predictedCAPTCHA = pytesseract.image_to_string(imgPATH).replace('\n', '').upper()
        highRiskChars = '08DGILMORSTZ'
        if len(predictedCAPTCHA)==6 and len(set(predictedCAPTCHA)&set(highRiskChars))==0 and re.search("[A-Z0-9]",predictedCAPTCHA):
            print(f'CAPTCHA_{i}:', predictedCAPTCHA)
            xpath = '//*[@placeholder="Nhập số bên"]'
            captchaInput = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
            captchaInput.clear()
            captchaInput.send_keys(predictedCAPTCHA)
            time.sleep(2)
        driver.refresh()
        time.sleep(1)
        i += 1
        if i > 100:
            break

