from request_phs.stock import *


class NoNewsFound(Exception):
    pass


class PageFailToLoad(Exception):
    pass


# PATH = join(dirname(dirname(realpath(__file__))), 'chromedriver_win32', 'chromedriver.exe')
PATH = r'D:\DataAnalytics\chromedriver_win32\chromedriver.exe'

ignored_exceptions = (
    ValueError,
    IndexError,
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    ElementNotInteractableException,
    PageFailToLoad,
    ElementClickInterceptedException,
    NoNewsFound
)

file_stock = pd.read_pickle(r'D:\DataAnalytics\News Filter Project\list_stock.pickle').reset_index()
stocks = file_stock['ticker'].tolist()

t = 1

for stock in stocks:
    time_lst = []
    price_1_lst = []
    price_2_lst = []
    price_3_lst = []
    kl_lo_lst = []
    kl_tich_luy_lst = []
    ty_trong_lst = []

    chrome_options = Options()
    # chrome_options.add_argument("--start-maximized")
    try:
        url = 'https://finance.vietstock.vn/' + stock + '/thong-ke-giao-dich.htm'
        driver = webdriver.Chrome(options=chrome_options, service=Service(PATH))
        driver.get(url)

        # Đăng nhập
        login_element = driver.find_element(By.XPATH, '/html/body/div[2]/div[6]/div/div[2]/div[2]/a[3]')
        login_element.click()
        email_element = driver.find_element(By.XPATH, '//*[@id="txtEmailLogin"]')
        email_element.clear()
        email_element.send_keys('namtran@phs.vn')
        password_element = driver.find_element(By.XPATH, '//*[@id="txtPassword"]')
        password_element.clear()
        password_element.send_keys('123456789')
        login_element = driver.find_element(By.XPATH, '//*[@id="btnLoginAccount"]')
        login_element.click()
        time.sleep(1)

        wait = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions)
        print(stock)
        press = 1
        while True:
            try:
                # lấy cột Thời gian
                get_times = wait.until(EC.presence_of_all_elements_located((
                    By.XPATH,
                    '//*[@id="deal-content"]/div/div/div[2]/div/table/tbody/tr/td[1]'
                )))
                t_lst = [dt.datetime.strptime(t.text, "%H:%M:%S").time() for t in get_times]
                time_lst = time_lst + t_lst

                # lấy cột Giá
                # vì cột giá có format là "g1 g2(g3%)" => sẽ có 3 giá trị nằm trong cột này
                price_1 = wait.until(EC.presence_of_all_elements_located((
                    By.XPATH,
                    '//*[@id="deal-content"]/div/div/div[2]/div/table/tbody/tr/td[2]/span/span[1]'
                )))
                p_1 = [int(p.text.replace(',', '')) for p in price_1]
                price_1_lst = price_1_lst + p_1
                price_2 = wait.until(EC.presence_of_all_elements_located((
                    By.XPATH,
                    '//*[@id="deal-content"]/div/div/div[2]/div/table/tbody/tr/td[2]/span/span[2]'
                )))
                p_2 = [int(p.text.replace(',', '')) for p in price_2]
                price_2_lst = price_2_lst + p_2
                price_3 = wait.until(EC.presence_of_all_elements_located((
                    By.XPATH,
                    '//*[@id="deal-content"]/div/div/div[2]/div/table/tbody/tr/td[2]/span/span[4]'
                )))
                p_3 = [float(p.text.replace('%', '')) for p in price_3]
                price_3_lst = price_3_lst + p_3

                # lấy cột KL lô
                get_kl_lo = wait.until(EC.presence_of_all_elements_located((
                    By.XPATH,
                    '//*[@id="deal-content"]/div/div/div[2]/div/table/tbody/tr/td[3]'
                )))
                kllo_lst = [int(kl_lo.text.replace(',', '')) for kl_lo in get_kl_lo]
                kl_lo_lst = kl_lo_lst + kllo_lst

                # lấy cột KL tích lũy
                get_kl_tich_luy = wait.until(EC.presence_of_all_elements_located((
                    By.XPATH,
                    '//*[@id="deal-content"]/div/div/div[2]/div/table/tbody/tr/td[4]'
                )))
                kl_tl_lst = [int(kl_tich_luy.text.replace(',', '')) for kl_tich_luy in get_kl_tich_luy]
                kl_tich_luy_lst = kl_tich_luy_lst + kl_tl_lst

                # lấy cột tỷ trọng
                get_ty_trong = wait.until(EC.presence_of_all_elements_located((
                    By.XPATH,
                    '//*[@id="deal-content"]/div/div/div[2]/div/table/tbody/tr/td[5]'
                )))
                if '-' in [ty_trong.text for ty_trong in get_ty_trong]:
                    tt_lst = [int(ty_trong.text.replace('-', '0')) for ty_trong in get_ty_trong]
                else:
                    tt_lst = [round(float(ty_trong.text.replace('%', '')), 2) for ty_trong in get_ty_trong]
                ty_trong_lst = ty_trong_lst + tt_lst

                while True:
                    next_button = wait.until(EC.presence_of_element_located((
                        By.XPATH,
                        '//*[@id="btn-page-next"]'
                    )))
                    try:
                        next_button.click()
                        time.sleep(0.1)
                        break
                    except ignored_exceptions:
                        continue
            except ignored_exceptions:
                continue
            press += 1
            page = int(driver.find_element(
                By.XPATH,
                '//*[@id="deal-content"]/div/div/div[2]/div/div/div/span[1]/span[2]'
            ).text)
            print(press)
            if press > page:
                break
        time.sleep(t)
        driver.quit()

        print('Stock:', stock, '-', time_lst[-1], '-', ty_trong_lst[-1])
        print("------------------------")
        dictionary = {
            'Thời gian': time_lst,
            'Giá 1': price_1_lst,
            'Giá 2': price_2_lst,
            'Giá 3': price_3_lst,
            'KL lô': kl_lo_lst,
            'KL tích lũy': kl_tich_luy_lst,
            'Tỷ trọng': ty_trong_lst
        }
        df = pd.DataFrame(dictionary)
        check = df.duplicated().all(axis=0)
        if check:
            df.to_pickle(fr"D:\DataAnalytics\News Filter Project\output_stock_data\{stock}-duplicated.pickle")
        else:
            df.to_pickle(fr"D:\DataAnalytics\News Filter Project\output_stock_data\{stock}.pickle")
        time.sleep(2)
    except ignored_exceptions:
        pass

# check = pd.read_pickle(r'D:\DataAnalytics\News Filter Project\output_stock_data\CII.pickle')
# check = pd.read_pickle(r'D:\DataAnalytics\News Filter Project\output_stock_data\AGG-duplicated.pickle')
