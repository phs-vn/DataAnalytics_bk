from request_phs.stock import *

print('Experimental v2 is used')
wait_time = 5

class NoNewsFound(Exception):
    pass

class vsd:

    def __init__(self):

        self.PATH = join(dirname(dirname(realpath(__file__))), 'phs', 'chromedriver')

        self.ignored_exceptions \
            = (ValueError,
               IndexError,
               NoSuchElementException,
               StaleElementReferenceException,
               TimeoutException,
               ElementNotInteractableException)

        self.fixedmp_path = r'\\192.168.10.101\phs-storge-2018\RiskManagementDept\RMD_Data\Luu tru van ban\RMC Meeting 2018\00. Meeting minutes\Data'
        self.fixedmp_series = pd.read_excel(join(self.fixedmp_path,'Fixed Max Price.xlsx'), usecols=['Stock'], squeeze=True)
        self.fixedmp_list = self.fixedmp_series.tolist()

        self.margin_path = r'\\192.168.10.101\phs-storge-2018\RiskManagementDept\RMD_Data\Luu tru van ban\Information Public\Intranet\Intranet'
        self.margin_date = [name.split('_')[-1][:10] for name in listdir(self.margin_path)]
        self.margin_year = [date[-4:] for date in self.margin_date]
        self.margin_month = [date[3:5] for date in self.margin_date]
        self.margin_day = [date[:2] for date in self.margin_date]
        self.margin_date = [x+y+z for x,y,z in zip(self.margin_year,self.margin_month,self.margin_day)]
        self.margin_date.sort()
        self.last_date = self.margin_date[-1]
        self.margin_year = self.last_date[:4]
        self.margin_month = self.last_date[4:6]
        self.margin_day = self.last_date[-2:]
        self.margin_file = f'Intranet list_{self.margin_day}.{self.margin_month}.{self.margin_year} final.xlsx'
        self.margin_list = pd.read_excel(join(self.margin_path,self.margin_file), usecols=['Mã CK'], squeeze=True).tolist()


    def tinTCPH(self, num_hours:int=48, fixed_mp:bool=True) -> pd.DataFrame:

        """
        This function returns a DataFrame and export an excel file containing
        news update published in 'https://vsd.vn/vi/alo/-f-_bsBS4BBXga52z2eexg'
        (tin từ tổ chức phát hành).

        :param num_hours: number of hours in the past that's in our concern
        :param fixed_mp: care only about stock in "fixed max price" (True)
         or not (False)

        :return: summary report of news update
        """

        start_time = time.time()

        now = datetime.now()
        fromtime = now

        url = 'https://vsd.vn/vi/alo/-f-_bsBS4BBXga52z2eexg'
        driver = webdriver.Chrome(executable_path=self.PATH)
        driver.get(url)
        driver.maximize_window()

        keywords = ['Cổ tức', 'Tạm ứng cổ tức',
                    'Tạm ứng', 'Chi trả', 'Bằng tiền',
                    'Cổ phiếu', 'Quyền mua',
                    'Chuyển dữ liệu đăng ký',
                    'cổ tức', 'tạm ứng cổ tức',
                    'tạm ứng', 'chi trả', 'bằng tiền',
                    'cổ phiếu', 'quyền mua',
                    'chuyển dữ liệu đăng ký']

        output_table = pd.DataFrame()

        bmk_time = btime(now.strftime('%Y-%m-%d %H:%M:%S'),-num_hours)
        while fromtime >= datetime.strptime(bmk_time,'%Y-%m-%d %H:%M:%S'):

            news_time = []
            news_headlines = []
            news_urls = []

            news_ratios = []
            news_recordsdate = []
            news_paymentdate = []

            tags \
                = WebDriverWait(driver, 1,
                                ignored_exceptions=self.ignored_exceptions) \
                .until(
                expected_conditions.presence_of_all_elements_located(
                    (By.XPATH, '//*[@id="d_list_news"]/ul/li')))

            for tag_ in tags:
                tags \
                    = WebDriverWait(driver, 5,
                                    ignored_exceptions=self.ignored_exceptions) \
                    .until(
                    expected_conditions.presence_of_all_elements_located(
                        (By.XPATH, '//*[@id="d_list_news"]/ul/li')))
                try:
                    h3_tag \
                        = WebDriverWait(tag_, 5,
                                        ignored_exceptions=self.ignored_exceptions)\
                        .until(
                        expected_conditions.presence_of_element_located(
                            (By.TAG_NAME, 'h3')))
                except self.ignored_exceptions:
                    h3_tag = tag_.find_element_by_tag_name('h3')

                if fixed_mp is True:
                    if h3_tag.text[:3] in self.fixedmp_list:
                        txt = h3_tag.text
                    else:
                        continue
                else:
                    txt = h3_tag.text

                check = [word in txt for word in keywords]

                if any(check):

                    news_headlines += [txt]

                    try:
                        sub_url \
                            = WebDriverWait(h3_tag, 5,
                                  ignored_exceptions=self.ignored_exceptions)\
                            .until(
                            expected_conditions.presence_of_element_located(
                                (By.TAG_NAME, 'a'))).get_attribute('href')

                    except self.ignored_exceptions:
                        sub_url = h3_tag.find_element_by_tag_name('a')\
                            .get_attribute('href')

                    news_urls += [f'=HYPERLINK("{sub_url}","Link")']

                    news_time_ = tag_.find_element_by_tag_name('div').text
                    news_time_ \
                        = datetime.strptime(news_time_[-21:],
                                            '%d/%m/%Y - %H:%M:%S')
                    news_time += [news_time_]


                    sub_driver = webdriver.Chrome(executable_path=self.PATH)
                    sub_driver.get(sub_url)

                    try:
                        info_heads \
                            = WebDriverWait(sub_driver, 5,
                                            ignored_exceptions=self.ignored_exceptions) \
                            .until(
                            expected_conditions.presence_of_all_elements_located(
                                (By.XPATH,
                                 "//div[substring(@class,string-length(@class)"
                                 "-string-length('item-info')+1)='item-info']")))
                    except self.ignored_exceptions:
                        time.sleep(5)
                        info_heads = sub_driver.find_elements_by_xpath(
                            "//div[substring(@class,string-length(@class)"
                            "-string-length('item-info')+1)='item-info']")

                    if len(info_heads) == 0:
                        news_dict = dict()
                    else:
                        # get info from heading table
                        info_heads_text = [head.text[:-1]
                                           for head in info_heads]
                        info_tails \
                            = WebDriverWait(sub_driver, 5,
                                  ignored_exceptions=self.ignored_exceptions)\
                            .until(
                            expected_conditions.presence_of_all_elements_located(
                                (By.XPATH,
                                 "//div[substring(@class,string-length(@class)"
                                 "-string-length('item-info-main')+1)='item-info-main']")))

                        info_tails_text = [tail.text for tail in info_tails]
                        news_dict = {h:t for h,t in
                                     zip(info_heads_text, info_tails_text)}

                    try:
                        content_elem \
                            = WebDriverWait(sub_driver,5,
                                            ignored_exceptions=self.ignored_exceptions) \
                            .until(
                            expected_conditions.presence_of_element_located(
                                (By.XPATH,"//div[@style='text-align: justify;']")))
                    except self.ignored_exceptions:
                        # một vài tin đăng có cấu trúc thay đổi
                        time.sleep(5)
                        content_elem = sub_driver.find_element_by_xpath(
                            "//*[@id='wrapper']/main/div/div/div/div[1]/div[2]/div/div/p[2]")

                    content = content_elem.text
                    content_split = content.split('\n')

                    def f(full_list):
                        result = ''
                        if len(full_list) != 0:
                            for i in range(len(full_list)-1):
                                result += full_list[i] + '\n'
                            result += full_list[-1]
                        return result


                    dividend_kws = ['Chi trả cổ tức', 'Tỷ lệ thực hiện']
                    mask_dividend = []
                    for row in content_split:
                        mask_dividend += [any([keyword in row
                                               for keyword in dividend_kws])]

                    dividends \
                        = list(np.array(content_split)[np.array(mask_dividend)])
                    dividends = f(dividends)
                    news_dict['Tỷ lệ thực hiện'] = dividends


                    payment_kws = ['Thời gian thanh toán',
                                   'Ngày thanh toán',
                                   'Thời gian thực hiện']
                    mask_payment = []
                    for row in content_split:
                        mask_payment \
                            += [any([keyword in row for keyword in payment_kws])]

                    payments \
                        = list(np.array(content_split)[np.array(mask_payment)])
                    payments = f(payments)
                    news_dict['Thời gian thanh toán'] = payments

                    try:
                        news_ratios += [news_dict['Tỷ lệ thực hiện']]
                    except KeyError:
                        news_ratios += ['']
                    try:
                        news_recordsdate += [news_dict['Ngày đăng ký cuối cùng']]
                    except KeyError:
                        news_recordsdate += ['']
                    try:
                        news_paymentdate += [news_dict['Thời gian thanh toán']]
                    except KeyError:
                        news_paymentdate += ['']

                    sub_driver.quit()

                else:
                    pass

            output_table = pd.concat([output_table,
                                      pd.DataFrame(
                                          {'Thời gian': news_time,
                                           'Tiêu đề': news_headlines,
                                           'Tỷ lệ thực hiện': news_ratios,
                                           'Ngày đăng ký cuối cùng': news_recordsdate,
                                           'Ngày thanh toán': news_paymentdate,
                                           'Link': news_urls}
                                      )], ignore_index=True)

            # Turn Page
            time.sleep(1)
            nextpage_button \
                = WebDriverWait(driver,5,ignored_exceptions=self.ignored_exceptions)\
                .until(
                expected_conditions.presence_of_all_elements_located(
                    (By.XPATH,
                     "//button[substring(@onclick,1,string-length('changePage'))"
                     "='changePage']")))[-2]

            nextpage_button.click()

            # Check time
            last_tag \
                = WebDriverWait(driver,5,ignored_exceptions=self.ignored_exceptions) \
                .until(
                expected_conditions.presence_of_all_elements_located(
                    (By.XPATH,
                     '//*[@id="d_list_news"]/ul/li')))[-1]

            fromtime = last_tag.find_element_by_tag_name('div').text
            fromtime = datetime.strptime(fromtime[-21:],
                                         '%d/%m/%Y - %H:%M:%S')

        driver.quit()

        if output_table.empty is True:
            raise NoNewsFound(f'Không có tin trong {num_hours} giờ vừa qua')

        def f(frame):
            series = frame['Tiêu đề'].str.split(': ')
            table = pd.DataFrame(series.tolist(),
                                 columns=['cophieu','tieude'])
            return table

        output_table[['Mã cổ phiếu','Lý do, mục đích']] = output_table.transform(f)

        def f(frame):
            from_exc = frame['Lý do, mục đích'].str.findall('[A-Z]+').str.get(-2)
            from_exc = from_exc.apply(
                lambda x: x if x in ['HOSE', 'HNX', 'UPCOM'] else '')
            to_exc = frame['Lý do, mục đích'].str.findall('[A-Z]+').str.get(-1)
            to_exc = to_exc.apply(
                lambda x: x if x in ['HOSE', 'HNX', 'UPCOM'] else '')
            result = pd.DataFrame({'Chuyển Từ Sàn': from_exc,
                                   'Chuyển Đến Sàn': to_exc})
            return result


        output_table[['Chuyển từ sàn', 'Chuyển đến sàn']] \
            = output_table.transform(f)

        def f(elem):

            if elem != '':
                try:
                    ryear = elem[-4:]
                    rmonth = elem[-7:-5]
                    rday = elem[-10:-8]

                    result = bdate(f'{ryear}-{rmonth}-{rday}',-1)
                    year = result[:4]
                    month = result[5:7]
                    day = result[-2:]

                    result = f'{day}/{month}/{year}'
                except ValueError:
                    result = ''
            else:
                result = ''

            return result

        output_table['Ngày giao dịch không hưởng quyền'] \
            = output_table['Ngày đăng ký cuối cùng'].map(f)

        output_table.drop(['Tiêu đề'], axis=1, inplace=True)

        check_margin = lambda ticker: 'Yes' if ticker in (self.margin_list + ['HOSE','HNX','UPCOM']) else 'No'
        output_table['Đang cho vay'] = output_table['Mã cổ phiếu'].map(check_margin)

        output_table = output_table[['Thời gian',
                                     'Mã cổ phiếu',
                                     'Đang cho vay',
                                     'Lý do, mục đích',
                                     'Tỷ lệ thực hiện',
                                     'Ngày đăng ký cuối cùng',
                                     'Ngày giao dịch không hưởng quyền',
                                     'Ngày thanh toán',
                                     'Chuyển từ sàn',
                                     'Chuyển đến sàn',
                                     'Link']]

        print(f'Finished ::: Total execution time: {int(time.time() - start_time)}s\n')

        return output_table


    def tinTVBT(self, num_hours:int=48) -> pd.DataFrame:

        """
        This function returns a DataFrame and export an excel file containing
        news update published in 'https://vsd.vn/vi/tin-thi-truong-phai-sinh'
        (tin từ thành viên bù trừ).

        :param num_hours: number of hours in the past that's in our concern

        :return: summary report of news update
        """

        start_time = time.time()

        now = datetime.now()
        fromtime = now

        url = 'https://vsd.vn/vi/tin-thi-truong-phai-sinh'
        driver = webdriver.Chrome(executable_path=self.PATH)
        driver.get(url)

        driver.maximize_window()

        keywords = ['tỷ lệ ký quỹ ban đầu',
                    'Tỷ lệ ký quỹ ban đầu',
                    'hợp đồng tương lai',
                    'Hợp đồng tương lai',
                    'HĐTL']

        output_table = pd.DataFrame()

        bmk_time = btime(now.strftime('%Y-%m-%d %H:%M:%S'), -num_hours)
        while fromtime >= datetime.strptime(bmk_time, '%Y-%m-%d %H:%M:%S'):

            news_time = []
            news_headlines = []
            news_urls = []

            time.sleep(1)

            tags \
                = WebDriverWait(driver,5,ignored_exceptions=self.ignored_exceptions) \
                .until(
                expected_conditions.presence_of_all_elements_located(
                    (By.XPATH, '//*[@id="tab1"]/ul/li')))

            for tag_ in tags:

                tags \
                    = WebDriverWait(driver,5,ignored_exceptions=self.ignored_exceptions) \
                    .until(
                    expected_conditions.presence_of_all_elements_located(
                        (By.XPATH, '//*[@id="tab1"]/ul/li')))
                try:
                    h3_tag \
                        = WebDriverWait(tag_,5,ignored_exceptions=self.ignored_exceptions)\
                        .until(
                        expected_conditions.presence_of_element_located(
                            (By.TAG_NAME, 'h3')))
                except self.ignored_exceptions:
                    h3_tag = tag_.find_element_by_tag_name('h3')


                if h3_tag.text[:3].isupper():
                    continue
                else:
                    txt = h3_tag.text
                    check = [word in txt for word in keywords]
                    if any(check):
                        news_headlines += [txt]

                        sub_url = h3_tag.find_element_by_tag_name('a') \
                            .get_attribute('href')
                        news_urls += [f'=HYPERLINK("{sub_url}","Link")']

                        news_time_ = tag_.find_element_by_tag_name('div').text
                        news_time_ \
                            = datetime.strptime(news_time_[-21:],
                                                '%d/%m/%Y - %H:%M:%S')
                        news_time += [news_time_]

            output_table = pd.concat([output_table,
                                      pd.DataFrame(
                                          {'Thời gian': news_time,
                                           'Lý do, Mục đích': news_headlines,
                                           'Link': news_urls}
                                      )], ignore_index=True)

            # Turn Page
            time.sleep(1)
            nextpage_button \
                = WebDriverWait(driver, 5,
                                ignored_exceptions=self.ignored_exceptions) \
                .until(
                expected_conditions.presence_of_all_elements_located(
                    (By.XPATH,
                     "//*[@id='d_number_of_page']/button")))[-2]
            nextpage_button.click()

            # Check time
            last_tag \
                = WebDriverWait(driver, 5,
                                ignored_exceptions=self.ignored_exceptions) \
                .until(
                expected_conditions.presence_of_all_elements_located(
                    (By.XPATH,
                     '//*[@id="tab1"]/ul/li')))[-1]

            fromtime = last_tag.find_element_by_tag_name('div').text
            fromtime = datetime.strptime(fromtime[-21:],
                                         '%d/%m/%Y - %H:%M:%S')

        driver.quit()

        if output_table.empty is True:
            raise NoNewsFound(f'Không có tin trong {num_hours} giờ vừa qua')

        print(f'Finished ::: Total execution time: {int(time.time()-start_time)}s\n')

        return output_table


vsd = vsd()


class hnx:

    def __init__(self):

        self.PATH = join(dirname(dirname(realpath(__file__))), 'phs',
                         'chromedriver')

        self.ignored_exceptions \
            = (ValueError,
               IndexError,
               NoSuchElementException,
               StaleElementReferenceException,
               TimeoutException,
               ElementNotInteractableException)

        self.fixedmp_path = r'\\192.168.10.101\phs-storge-2018\RiskManagementDept\RMD_Data\Luu tru van ban\RMC Meeting 2018\00. Meeting minutes\Data'
        self.fixedmp_series = pd.read_excel(join(self.fixedmp_path,'Fixed Max Price.xlsx'), usecols=['Stock'], squeeze=True)
        self.fixedmp_list = self.fixedmp_series.tolist()

        self.margin_path = r'\\192.168.10.101\phs-storge-2018\RiskManagementDept\RMD_Data\Luu tru van ban\Information Public\Intranet\Intranet'
        self.margin_date = [name.split('_')[-1][:10] for name in listdir(self.margin_path)]
        self.margin_year = [date[-4:] for date in self.margin_date]
        self.margin_month = [date[3:5] for date in self.margin_date]
        self.margin_day = [date[:2] for date in self.margin_date]
        self.margin_date = [x+y+z for x,y,z in zip(self.margin_year,self.margin_month,self.margin_day)]
        self.margin_date.sort()
        self.last_date = self.margin_date[-1]
        self.margin_year = self.last_date[:4]
        self.margin_month = self.last_date[4:6]
        self.margin_day = self.last_date[-2:]
        self.margin_file = f'Intranet list_{self.margin_day}.{self.margin_month}.{self.margin_year} final.xlsx'
        self.margin_list = pd.read_excel(join(self.margin_path,self.margin_file),
                                         usecols=['Mã CK'], squeeze=True).tolist()


    def tinTCPH(self, num_hours:int=48, fixed_mp:bool=True):

        """
        This function returns a DataFrame and export an excel file containing
        news update published in 'https://www.hnx.vn/thong-tin-cong-bo-ny-tcph.html'
        (tin từ tổ chức phát hành).

        :param num_hours: number of hours in the past that's in our concern
        :param fixed_mp: care only about stock in "fixed max price" (True)
         or not (False)

        :return: summary report of news update
        """

        start_time = time.time()

        now = datetime.now()
        fromtime = now

        url = 'https://www.hnx.vn/thong-tin-cong-bo-ny-tcph.html'
        driver = webdriver.Chrome(executable_path=self.PATH)
        driver.get(url)

        driver.maximize_window()

        key_words = ['Cổ tức', 'Tạm ứng cổ tức',
                     'Tạm ứng', 'Chi trả',
                     'Bằng tiền', 'Quyền mua',
                     'cổ tức', 'tạm ứng cổ tức',
                     'tạm ứng', 'chi trả',
                     'bằng tiền', 'quyền mua']

        links = []
        box_text = []
        firms = []
        titles = []
        times = []
        tickers = []

        bmk_time = btime(now.strftime('%Y-%m-%d %H:%M:%S'), -num_hours)
        while fromtime >= datetime.strptime(bmk_time, '%Y-%m-%d %H:%M:%S'):

            def f():
                # wait for the element to appear, avoid stale element reference
                ticker_elems \
                    = WebDriverWait(driver,5,ignored_exceptions=self.ignored_exceptions) \
                    .until(
                    expected_conditions.presence_of_all_elements_located(
                        (By.XPATH, "//*[@id='_tableDatas']/tbody/*/td[3]/a")))
                title_elems \
                    = WebDriverWait(driver,5,ignored_exceptions=self.ignored_exceptions) \
                    .until(
                    expected_conditions.presence_of_all_elements_located(
                        (By.XPATH, "//*[@id='_tableDatas']/tbody/*/td[5]/a")))
                return ticker_elems, title_elems

            try:
                ticker_elems, title_elems = f()
                tickers_inpage = [t.text for t in ticker_elems if t.text != '']
                titles_inpage = [t.text for t in title_elems if t.text != '']
                title_elems = [t for t in title_elems if t.text != '']
            except self.ignored_exceptions:
                time.sleep(15)
                ticker_elems, title_elems = f()
                tickers_inpage = [t.text for t in ticker_elems if t.text != '']
                titles_inpage = [t.text for t in title_elems if t.text != '']
                title_elems = [t for t in title_elems if t.text != '']


            ticker_title_inpage = [f'{ticker} {title}' for ticker,title
                                   in zip(tickers_inpage, titles_inpage)]

            for ticker_title in ticker_title_inpage:

                sub_check = [key_word in ticker_title for key_word in key_words]

                if any(sub_check):
                    pass
                else:
                    continue

                title_no = ticker_title_inpage.index(ticker_title)+1

                ticker_ = driver.find_elements_by_xpath\
                    (f"//*[@id='_tableDatas']/tbody/tr[{title_no}]/td[3]")[-1].text

                if fixed_mp is True:
                    if ticker_ in self.fixedmp_list:
                        pass
                    else:
                        continue

                title_ = driver.find_elements_by_xpath\
                    (f"//*[@id='_tableDatas']/tbody/tr[{title_no}]/td[5]")[-1].text
                firm_ = driver.find_elements_by_xpath\
                    (f"//*[@id='_tableDatas']/tbody/tr[{title_no}]/td[4]")[-1].text
                time_ = driver.find_elements_by_xpath\
                    (f"//*[@id='_tableDatas']/tbody/tr[{title_no}]/td[2]")[-1].text

                titles += [title_]
                firms += [firm_]
                times += [time_]
                tickers += [ticker_]

                # open popup window:
                click_obj = title_elems[title_no-1]
                click_obj.click()

                time.sleep(1)

                # evaluate popup content
                popup_content \
                    = WebDriverWait(driver,5,ignored_exceptions=self.ignored_exceptions) \
                    .until(
                    expected_conditions.presence_of_element_located(
                        (By.XPATH, "//div[@class='Box-Noidung']")))
                content = popup_content.text
                if content in ['.', '', ' ']: # nếu có link, ko có nội dung
                    box_text += ['']
                    link_elems = driver.find_elements_by_xpath(
                        "//div[@class='divLstFileAttach']/p/a")
                    links += [[link.get_attribute('href') for link in link_elems]]
                else: # nếu có nội dung, ko có link
                    links += ['']
                    box_text += [content]

                # exit pop-up windows
                WebDriverWait(driver,5,ignored_exceptions=self.ignored_exceptions).until(
                    expected_conditions.element_to_be_clickable
                    ((By.XPATH,"//*[@id='divViewDetailArticles']/*/input"))).click()

                time.sleep(1)

            # check time
            fromtime = driver.find_elements_by_xpath(
                "//*[@id='_tableDatas']/tbody/tr[10]/td[2]")[-1].text
            fromtime = datetime.strptime(fromtime, '%d/%m/%Y %H:%M')

            # Next-page click (qua trang)
            try:
                 WebDriverWait(driver, 5,
                                    ignored_exceptions=self.ignored_exceptions) \
                    .until(
                    expected_conditions.presence_of_all_elements_located(
                        (By.XPATH, "//*[@id='d_number_of_page']/li")))[-2].click()
            except self.ignored_exceptions:
                time.sleep(1)
                driver.find_elements_by_xpath("//*[@id='d_number_of_page']/li")[-2].click()

        driver.quit()

        # export to DataFrame
        df = pd.DataFrame(list(zip(times, tickers, firms,
                                   titles, box_text, links)),
                          columns=['Thời gian', 'Mã CK', 'Tên TCPH',
                                   'Tiêu đề', 'Nội dung',
                                   'File đính kèm'])


        if df.empty is True:
            raise NoNewsFound(f'Không có tin trong {num_hours} giờ vừa qua')

        df['Nội dung'] = df['Nội dung'].str.split('\n')
        df['Lý do / mục đích'] = ''
        df['Tỷ lệ thực hiện'] = ''
        df['Ngày đăng ký cuối cùng'] = ''
        df['Ngày giao dịch không hưởng quyền'] = ''
        df['Thời gian thực hiện'] = ''
        for table_row in range(df.shape[0]):
            content_rows = df['Nội dung'].iloc[table_row]

            reasons = np.array(content_rows)[['*' in row for row in content_rows]]
            for reason in reasons:
                df['Lý do / mục đích'].iloc[table_row] \
                    += reason.lstrip().lstrip('*') + '\n'

            ratios = np.array(content_rows)[['Tỷ lệ thực hiện'
                                             in row for row in content_rows]]
            for ratio in ratios:
                df['Tỷ lệ thực hiện'].iloc[table_row] \
                    += ratio.lstrip().lstrip('-') + '\n'

            record_dates = np.array(content_rows)[['Ngày đăng ký cuối cùng'
                                                   in row for row in content_rows]]

            for record_date in record_dates:
                df['Ngày đăng ký cuối cùng'].iloc[table_row] \
                    += record_date + '\n'

                rtime = df['Ngày đăng ký cuối cùng'].iloc[table_row]

                try:
                    ryear = rtime[-5:-1]
                    rmonth = rtime[-8:-6]
                    rday = rtime[-11:-9]
                    result = bdate(f'{ryear}-{rmonth}-{rday}',-1)
                    year = result[:4]
                    month = result[5:7]
                    day = result[-2:]
                    df['Ngày giao dịch không hưởng quyền'].iloc[table_row] \
                        += f'{day}/{month}/{year}' + '\n'
                except ValueError:
                    pass

            payment_dates = np.array(content_rows)[['Thời gian thực hiện'
                                                    in row for row in content_rows]]
            for payment_date in payment_dates:
                df['Thời gian thực hiện'].iloc[table_row] \
                    += payment_date.lstrip().lstrip('-') + '\n'

            def f(full_list):
                if len(full_list) == 1:
                    result = f'=HYPERLINK("{full_list[0]}","Link")'
                elif len(full_list) > 1:
                    result = ''
                    result += f'=HYPERLINK("{full_list[0]}","Link")'
                    for i in range(1, len(full_list)):
                        result += f'&" "&HYPERLINK("{full_list[i]}","Link")'
                else:
                    result = None
                return result

            ls = df.loc[df.index[table_row],'File đính kèm']
            df.loc[df.index[table_row],'File đính kèm'] = f(ls)

        df.drop(['Nội dung'], axis=1, inplace=True)

        check_margin = lambda ticker: 'Yes' if ticker in (self.margin_list + ['HOSE','HNX','UPCOM']) else 'No'
        df['Đang cho vay'] = df['Mã CK'].map(check_margin)

        df = df[['Thời gian',
                 'Mã CK',
                 'Đang cho vay',
                 'Tên TCPH',
                 'Tiêu đề',
                 'Lý do / mục đích',
                 'Tỷ lệ thực hiện',
                 'Ngày đăng ký cuối cùng',
                 'Ngày giao dịch không hưởng quyền',
                 'Thời gian thực hiện',
                 'File đính kèm']]

        print(f'Finished ::: Total execution time: {int(time.time()-start_time)}s\n')

        return df


    def tintuso(self, num_hours:int=48):

        """
        This function returns a DataFrame and export an excel file containing
        news update published in 'https://www.hnx.vn/thong-tin-cong-bo-ny-hnx.html'
        (tin từ sở). If upcom = True, include news update published in
        'https://www.hnx.vn/vi-vn/thong-tin-cong-bo-up-hnx.html' (tin từ sở)

        :param num_hours: number of hours in the past that's in our concern

        :return: summary report of news update
        """

        start_time = time.time()

        now = datetime.now()
        from_time = now

        driver = webdriver.Chrome(executable_path=self.PATH)

        url = 'https://www.hnx.vn/thong-tin-cong-bo-ny-hnx.html'
        driver.get(url)

        driver.maximize_window()

        key_words = ['tạm ngừng giao dịch',
                     'vào diện' , 'hủy niêm yết', 'ra khỏi diện',
                     'hủy bỏ niêm yết',
                     'kiểm soát', 'cảnh cáo', 'cảnh báo',
                     'không đủ điều kiện giao dịch ký quỹ',
                     'ra khỏi', 'vào',
                     'hạn chế giao dịch',
                     'chuyển giao dịch']

        links = []
        box_text = []
        titles = []
        times = []
        tickers = []

        bmk_time = btime(now.strftime('%Y-%m-%d %H:%M:%S'), -num_hours)
        while from_time >= datetime.strptime(bmk_time, '%Y-%m-%d %H:%M:%S'):

            def f():
                # wait for the element to appear, avoid stale element reference
                ticker_elems \
                    = WebDriverWait(driver,5,ignored_exceptions=self.ignored_exceptions) \
                    .until(
                    expected_conditions.presence_of_all_elements_located(
                        (By.XPATH, "//*[@id='_tableDatas']/tbody/*/td[3]/a")))
                title_elems \
                    = WebDriverWait(driver,5,ignored_exceptions=self.ignored_exceptions) \
                    .until(
                    expected_conditions.presence_of_all_elements_located(
                        (By.XPATH, "//*[@id='_tableDatas']/tbody/*/td[4]/a")))
                return ticker_elems, title_elems

            try:
                ticker_elems, title_elems = f()
                tickers_inpage = [t.text for t in ticker_elems if t.text != '']
                titles_inpage = [t.text for t in title_elems if t.text != '']
                title_elems = [t for t in title_elems if t.text != '']
            except self.ignored_exceptions:
                time.sleep(15)
                ticker_elems, title_elems = f()
                tickers_inpage = [t.text for t in ticker_elems if t.text != '']
                titles_inpage = [t.text for t in title_elems if t.text != '']
                title_elems = [t for t in title_elems if t.text != '']

            ticker_title_inpage = [f'{ticker} {title}' for ticker, title
                                   in zip(tickers_inpage, titles_inpage)]

            for ticker_title in ticker_title_inpage:

                check = [key_word in ticker_title for key_word in key_words]

                if any(check):
                    pass
                else:
                    continue

                title_no = ticker_title_inpage.index(ticker_title) + 1
                titles += [driver.find_elements_by_xpath \
                               (f"//*[@id='_tableDatas']/tbody/tr[{title_no}]/td[4]")[0].text]
                times += [driver.find_elements_by_xpath \
                              (f"//*[@id='_tableDatas']/tbody/tr[{title_no}]/td[2]")[0].text]
                tickers += [driver.find_elements_by_xpath \
                                (f"//*[@id='_tableDatas']/tbody/tr[{title_no}]/td[3]")[0].text]

                # open popup window:
                click_obj = title_elems[title_no-1]
                click_obj.click()
                time.sleep(1)

                # evaluate popup content
                popup_content = WebDriverWait(driver,5,
                                              ignored_exceptions=self.ignored_exceptions) \
                    .until(expected_conditions.presence_of_element_located(
                    (By.XPATH,"//div[@class='Box-Noidung']")))

                content = popup_content.text
                box_text += [content]
                link_elems = driver.find_elements_by_xpath(
                    "//div[@class='divLstFileAttach']/p/a")
                links += [[link.get_attribute('href') for link in link_elems]]

                # exit popup window
                WebDriverWait(driver,5,ignored_exceptions=self.ignored_exceptions).until(
                    expected_conditions.element_to_be_clickable
                    ((By.XPATH,"//*[@id='divViewDetailArticles']/*/input"))).click()

                time.sleep(1)

                # check time
                from_time = driver.find_element_by_xpath(
                    "//*[@id='_tableDatas']/tbody/tr[10]/td[2]").text
                from_time = datetime.strptime(from_time, '%d/%m/%Y %H:%M')


            # next page
            WebDriverWait(driver, 5,
                          ignored_exceptions=self.ignored_exceptions).until(
                expected_conditions.element_to_be_clickable
                ((By.XPATH, "//*[@id='next']"))).click()

        driver.quit()

        # export to DataFrame
        df = pd.DataFrame(list(zip(times, tickers, titles, box_text, links)),
                          columns=['Thời gian', 'Mã CK',
                                   'Tiêu đề', 'Nội dung',
                                   'File đính kèm'])

        check_margin = lambda ticker: 'Yes' if ticker in (self.margin_list + ['HOSE','HNX','UPCOM']) else 'No'
        df.insert(2, 'Đang cho vay', df['Mã CK'].map(check_margin))

        if df.empty is True:
            raise NoNewsFound(f'Không có tin trong {num_hours} giờ vừa qua')

        def f(full_list):
            if len(full_list) == 1:
                result = f'=HYPERLINK("{full_list[0]}","Link")'
            elif len(full_list) > 1:
                result = ''
                result += f'=HYPERLINK("{full_list[0]}","Link")'
                for i in range(1,len(full_list)):
                    result += f'&" "&HYPERLINK("{full_list[i]}","Link")'
            else:
                result = None
            return result

        df['File đính kèm'] = df['File đính kèm'].map(f)

        print(f'Finished ::: Total execution time: {int(time.time() - start_time)}s\n')

        return df


hnx = hnx()


class hose:

    def __init__(self):

        self.PATH = join(dirname(dirname(realpath(__file__))),'phs','chromedriver')

        self.ignored_exceptions \
            = (ValueError,
               IndexError,
               NoSuchElementException,
               StaleElementReferenceException,
               TimeoutException,
               ElementNotInteractableException)

        self.fixedmp_path = r'\\192.168.10.101\phs-storge-2018\RiskManagementDept\RMD_Data\Luu tru van ban\RMC Meeting 2018\00. Meeting minutes\Data'
        self.fixedmp_series = pd.read_excel(join(self.fixedmp_path,'Fixed Max Price.xlsx'), usecols=['Stock'], squeeze=True)
        self.fixedmp_list = self.fixedmp_series.tolist()

        self.margin_path = r'\\192.168.10.101\phs-storge-2018\RiskManagementDept\RMD_Data\Luu tru van ban\Information Public\Intranet\Intranet'
        self.margin_date = [name.split('_')[-1][:10] for name in listdir(self.margin_path)]
        self.margin_year = [date[-4:] for date in self.margin_date]
        self.margin_month = [date[3:5] for date in self.margin_date]
        self.margin_day = [date[:2] for date in self.margin_date]
        self.margin_date = [x+y+z for x,y,z in zip(self.margin_year,self.margin_month,self.margin_day)]
        self.margin_date.sort()
        self.last_date = self.margin_date[-1]
        self.margin_year = self.last_date[:4]
        self.margin_month = self.last_date[4:6]
        self.margin_day = self.last_date[-2:]
        self.margin_file = f'Intranet list_{self.margin_day}.{self.margin_month}.{self.margin_year} final.xlsx'
        self.margin_list = pd.read_excel(join(self.margin_path,self.margin_file),
                                         usecols=['Mã CK'], squeeze=True).tolist()


    def tintonghop(self, num_hours:int=48, fixed_mp:bool=True):

        """
        This function returns a DataFrame and export an excel file containing
        news update published in
        'https://www.hsx.vn/Modules/Cms/Web/NewsByCat/822d8a8c-fd19-4358-9fc9-d0b27a666611?fid=0318d64750264e31b5d57c619ed6b338'
        (tin từ sở).

        :param num_hours: number of hours in the past that's in our concern
        :param fixed_mp: care only about stock in "fixed max price" (True)
        or not (False)

        :return: summary report of news update
        """

        start_time = time.time()

        now = datetime.now()
        from_time = now

        url = 'https://www.hsx.vn/Modules/Cms/Web/NewsByCat/822d8a8c-fd19-4358-9fc9-d0b27a666611?fid=0318d64750264e31b5d57c619ed6b338'
        driver = webdriver.Chrome(executable_path=self.PATH)
        driver.get(url)

        driver.maximize_window()


        keywords = ['tạm ngừng giao dịch',
                    'vào diện' ,'hủy niêm yết','ra khỏi diện',
                    'kiểm soát', 'cảnh cáo', 'cảnh báo',
                    'hủy đăng ký', 'hủy bỏ',
                    'không đủ điều kiện giao dịch ký quỹ',
                    'hạn chế giao dịch', 'chuyển giao dịch',
                    'giao dịch đầu tiên cổ phiếu niêm yết',
                    'giao dịch đầu tiên của cổ phiếu niêm yết']

        excl_keywords = ['ban', 'Ban', 'bầu', 'Nhắc nhở', 'họp']

        output_table = pd.DataFrame()
        bmk_time = btime(now.strftime('%Y-%m-%d %H:%M:%S'), -num_hours)
        while from_time >= datetime.strptime(bmk_time, '%Y-%m-%d %H:%M:%S'):

            time.sleep(5)

            headline_tags \
                = WebDriverWait(driver, 5,
                                ignored_exceptions=self.ignored_exceptions) \
                .until(expected_conditions.presence_of_all_elements_located(
                    (By.XPATH, '*//td[3]/a')))[1:]

            headline_text = [t.text for t in headline_tags]
            headline_url = [t.get_attribute('href') for t in headline_tags]

            time_tags \
                = WebDriverWait(driver, 5,
                                ignored_exceptions=self.ignored_exceptions) \
                .until(expected_conditions.presence_of_all_elements_located(
                    (By.XPATH, '*//td[2]')))[2:]

            time_text = [t.text for t in time_tags]

            def f(s):
                if s.endswith('SA'):
                    s = s.rstrip(' SA')
                    return datetime.strptime(s,'%d/%m/%Y %H:%M:%S')
                if s.endswith('CH'):
                    s = s.rstrip(' CH')
                    return datetime.strptime(s,'%d/%m/%Y %H:%M:%S') + timedelta(hours=12)

            time_text = [f(t) for t in time_text]

            for num in range(len(headline_text)):

                sub_title = headline_text[num]
                sub_url = headline_url[num]
                sub_time = time_text[num]

                sub_contents = []
                sub_pdfs = []

                check_1 = [word in sub_title for word in keywords]
                check_2 = [word not in sub_title for word in excl_keywords]

                if any(check_1) and all(check_2):

                    sub_driver = webdriver.Chrome(executable_path=self.PATH)
                    sub_driver.get(sub_url)

                    time.sleep(1)
                    try:
                        content = sub_driver.find_element_by_xpath(
                            '/html/body/div[2]/div[1]/div[1]/div[1]/*//p[1]').text
                    except self.ignored_exceptions:
                        content = WebDriverWait(sub_driver, 5,
                                                ignored_exceptions
                                                =self.ignored_exceptions) \
                            .until(
                            expected_conditions.presence_of_element_located(
                                (By.XPATH,'/html/body/div[2]/div[1]/div[1]/div[1]/*//p[1]'))).text

                    content_rows = content.split('\n')

                    sub_content = [row for row in content_rows]

                    time.sleep(1)
                    try:
                        pdf_elements = sub_driver.find_elements_by_xpath('*//td[2]/a')
                    except self.ignored_exceptions:
                        pdf_elements = WebDriverWait(sub_driver, 5,
                                                     ignored_exceptions
                                                     =self.ignored_exceptions) \
                            .until(
                            expected_conditions.presence_of_all_elements_located(
                                (By.XPATH, '*//td[2]/a')))

                    pdf_files = [t.get_attribute('href') for t in pdf_elements]

                    def f1(full_list):
                        result = ''
                        if len(full_list) != 0:
                            for i in range(len(full_list)-1):
                                result += full_list[i] + '\n'
                            result += full_list[-1]
                        return result

                    def f2(full_list):
                        if len(full_list) == 1:
                            result = f'=HYPERLINK("{full_list[0]}","Link")'
                        elif len(full_list) > 1:
                            result = ''
                            result += f'=HYPERLINK("{full_list[0]}","Link")'
                            for i in range(1, len(full_list)):
                                result += f'&" "&HYPERLINK("{full_list[i]}","Link")'
                        else:
                            result = None
                        return result

                    sub_contents += [f1(sub_content)]
                    sub_pdfs += [f2(pdf_files)]

                    sub_driver.quit()

                # report and merge to parent table
                frames = pd.DataFrame({'Thời gian': sub_time,
                                       'Tiêu đề': sub_title,
                                       'Nội dung': sub_contents,
                                       'File đính kèm': sub_pdfs,
                                       'Link': f'=HYPERLINK("{sub_url}","Link")'})

                output_table = pd.concat([output_table, frames],
                                         ignore_index=True)

            from_time = time_text[-1]

            # Next Page:
            time.sleep(1)
            try:
                driver.find_elements_by_xpath(
                    '//*[@id="DbGridPager_2"]/a')[-2].click()
            except self.ignored_exceptions:
                WebDriverWait(driver, 5,
                              ignored_exceptions=self.ignored_exceptions) \
                    .until(
                    expected_conditions.presence_of_all_elements_located(
                        (By.XPATH, '//*[@id="DbGridPager_2"]/a')))[-2].click()

        driver.quit()

        url = 'https://www.hsx.vn/Modules/Cms/Web/NewsByCat/dca0933e-a578-4eaf-8b29-beb4575052c5?fid=6d1f1d5e9e6c4fb593077d461e5155e7'
        driver = webdriver.Chrome(executable_path=self.PATH)
        driver.get(url)

        from_time = now

        driver.maximize_window()

        while from_time >= datetime.strptime(bmk_time, '%Y-%m-%d %H:%M:%S'):

            time.sleep(5)

            headline_tags \
                = WebDriverWait(driver, 5,
                                ignored_exceptions=self.ignored_exceptions) \
                .until(
                expected_conditions.presence_of_all_elements_located(
                    (By.XPATH, '*//td[3]//*')))[1:]

            headline_text = [t.text for t in headline_tags]

            time_tags \
                = WebDriverWait(driver, 5,
                                ignored_exceptions=self.ignored_exceptions) \
                .until(
                expected_conditions.presence_of_all_elements_located(
                    (By.XPATH, '*//td[2]')))[2:]

            time_text = [t.text for t in time_tags]

            def f(s):
                if s.endswith('SA'):
                    s = s.rstrip(' SA')
                    return datetime.strptime(s,'%d/%m/%Y %H:%M:%S')
                if s.endswith('CH'):
                    s = s.rstrip(' CH')
                    return datetime.strptime(s,'%d/%m/%Y %H:%M:%S') + timedelta(hours=12)

            time_text = [f(t) for t in time_text]

            for num in range(len(headline_text)):

                sub_object = headline_tags[num]
                sub_title = headline_text[num]
                sub_time = time_text[num]

                sub_content = []
                sub_pdfs = []

                check_1 = [word in sub_title for word in keywords]
                check_2 = [word not in sub_title for word in excl_keywords]

                if any(check_1) and all(check_2):

                    # open popup windows
                    sub_object.click()

                    time.sleep(1)
                    try:
                        content = driver.find_element_by_xpath(
                            '/html/body/div[7]/div/div/div/div/div/div[2]/div[2]').text
                    except self.ignored_exceptions:
                        content = WebDriverWait(driver, 5,
                                                ignored_exceptions
                                                =self.ignored_exceptions) \
                            .until(
                            expected_conditions.presence_of_element_located(
                                (By.XPATH, '/html/body/div[7]/div/div/div/div/div/div[2]/div[2]'))).text
                    sub_content = [content]

                    time.sleep(1)

                    year_text = str(from_time.year)
                    try:
                        pdf_elements \
                            = driver.find_elements_by_partial_link_text(year_text)
                    except self.ignored_exceptions:
                        pdf_elements = WebDriverWait(driver, 5,
                                                     ignored_exceptions
                                                     =self.ignored_exceptions) \
                            .until(
                            expected_conditions.presence_of_all_elements_located(
                                (By.PARTIAL_LINK_TEXT, year_text)))

                    pdf_files = [t.get_attribute('href') for t in pdf_elements]

                    def f1(full_list):
                        if len(full_list) == 1:
                            result = f'=HYPERLINK("{full_list[0]}","Link")'
                        elif len(full_list) > 1:
                            result = ''
                            result += f'=HYPERLINK("{full_list[0]}","Link")'
                            for i in range(1, len(full_list)):
                                result += f'&" "&HYPERLINK("{full_list[i]}","Link")'
                        else:
                            result = None
                        return result

                    sub_pdfs += [f1(pdf_files)]

                    # close popup windows
                    WebDriverWait(driver, 5,
                                  ignored_exceptions=self.ignored_exceptions).until(
                        expected_conditions.element_to_be_clickable
                        ((By.XPATH, '//*[@title="Close"]'))).click()

                    time.sleep(1)

                # report and merge to parent table
                frames = pd.DataFrame({'Thời gian': sub_time,
                                       'Tiêu đề': sub_title,
                                       'Nội dung': sub_content,
                                       'File đính kèm': sub_pdfs})
                output_table = pd.concat([output_table, frames],
                                         ignore_index=True)

            from_time = time_text[-1]

            # Next Page:
            time.sleep(1)
            try:
                driver.find_elements_by_xpath(
                    '//*[@id="DbGridPager_2"]/a')[-2].click()
            except self.ignored_exceptions:
                WebDriverWait(driver, 5,
                              ignored_exceptions=self.ignored_exceptions) \
                    .until(
                    expected_conditions.presence_of_all_elements_located(
                        (By.XPATH, '//*[@id="DbGridPager_2"]/a')))[-2].click()

        driver.quit()

        if output_table.empty is True:
            raise NoNewsFound(f'Không có tin trong {num_hours} giờ vừa qua')

        # select out tickers from headline
        output_table.insert(1, 'Mã cổ phiếu',
                            output_table['Tiêu đề'].str.split(': ').str.get(0))

        # check if ticker is in margin list
        check_margin = lambda ticker: 'Yes' if ticker in (self.margin_list + ['HOSE','HNX','UPCOM']) else 'No'
        output_table.insert(2, 'Đang cho vay', output_table['Mã cổ phiếu'].map(check_margin))

        if fixed_mp is True:
            output_table = output_table[output_table['Mã cổ phiếu'].isin(self.fixedmp_list+['HOSE'])]

        print(f'Finished ::: Total execution time: {int(time.time()-start_time)}s\n')

        return output_table


hose = hose()