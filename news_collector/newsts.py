from request.stock import *


class NoNewsFound(Exception):
  pass


class vsd:

  def __init__(self):

    self.PATH = join(dirname(dirname(realpath(__file__))),'dependency',
                     'chromedriver')

    self.ignored_exceptions \
      = (ValueError,
         IndexError,
         NoSuchElementException,
         StaleElementReferenceException,
         TimeoutException,
         ElementNotInteractableException)

  def tinTCPH(self,num_hours: int = 48) -> pd.DataFrame:

    """
    This function returns a DataFrame and export an excel file containing
    news update published in 'https://vsd.vn/vi/alo/-f-_bsBS4BBXga52z2eexg'
    (tin từ tổ chức phát hành).

    :param num_hours: number of hours in the past that's in our concern

    :return: summary report of news update
    """

    start_time = time.time()

    now = dt.datetime.now()
    fromtime = now

    url = 'https://vsd.vn/vi/alo/-f-_bsBS4BBXga52z2eexg'
    driver = webdriver.Chrome(executable_path=self.PATH)
    driver.get(url)
    driver.maximize_window()

    keywords = ['Hủy đăng ký chứng khoán',
                'chứng nhận đăng ký chứng khoán lần đầu']

    output_table = pd.DataFrame()

    bmk_time = btime(now.strftime('%Y-%m-%d %H:%M:%S'),-num_hours)
    while fromtime >= dt.datetime.strptime(bmk_time,'%Y-%m-%d %H:%M:%S'):

      news_time = []
      news_headlines = []
      news_urls = []

      news_date_actives = []
      news_depository = []
      news_noidung = []

      time.sleep(1)

      tags = WebDriverWait(driver,5,ignored_exceptions=self.ignored_exceptions) \
        .until(EC.presence_of_all_elements_located(
        (By.XPATH,'//*[@id="d_list_news"]/ul/li')
      ))

      for tag_ in tags:
        tags \
          = WebDriverWait(driver,5,ignored_exceptions=self.ignored_exceptions) \
          .until(EC.presence_of_all_elements_located(
          (By.XPATH,'//*[@id="d_list_news"]/ul/li')
        ))
        try:
          h3_tag \
            = WebDriverWait(tag_,5,ignored_exceptions=self.ignored_exceptions) \
            .until(EC.presence_of_element_located(
            (By.TAG_NAME,'h3')
          ))
        except self.ignored_exceptions:
          h3_tag = tag_.find_element_by_tag_name('h3')

        txt = h3_tag.text
        check = [word in txt for word in keywords]

        if any(check):

          news_headlines += [txt]

          try:
            sub_url \
              = WebDriverWait(h3_tag,5,
                              ignored_exceptions=self.ignored_exceptions) \
              .until(
              EC.presence_of_element_located(
                (By.TAG_NAME,'a'))).get_attribute('href')

          except self.ignored_exceptions:
            sub_url = h3_tag.find_element_by_tag_name('a') \
              .get_attribute('href')

          news_urls += [f'=HYPERLINK("{sub_url}","Link")']

          news_time_ = tag_.find_element_by_tag_name('div').text
          news_time_ \
            = dt.datetime.strptime(news_time_[-21:],
                                   '%d/%m/%Y - %H:%M:%S')
          news_time += [news_time_]

          sub_driver = webdriver.Chrome(executable_path=self.PATH)
          sub_driver.get(sub_url)

          try:
            info_heads \
              = WebDriverWait(sub_driver,5,
                              ignored_exceptions=self.ignored_exceptions) \
              .until(
              EC.presence_of_all_elements_located(
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
              = WebDriverWait(sub_driver,5,
                              ignored_exceptions=self.ignored_exceptions) \
              .until(
              EC.presence_of_all_elements_located(
                (By.XPATH,
                 "//div[substring(@class,string-length(@class)"
                 "-string-length('item-info-main')+1)='item-info-main']")))

            info_tails_text = [tail.text for tail in info_tails]
            news_dict = {h:t for h,t in
                         zip(info_heads_text,info_tails_text)}

          # Cho tin: "Hủy đăng ký chứng khoán"
          try:
            content_elem \
              = WebDriverWait(sub_driver,5,
                              ignored_exceptions=self.ignored_exceptions) \
              .until(
              EC.presence_of_element_located(
                (By.XPATH,"//div[@style='text-align: justify;']")))
          except self.ignored_exceptions:
            # một vài tin đăng có cấu trúc thay đổi
            time.sleep(5)
            content_elem = sub_driver.find_element_by_xpath(
              "//*[@id='wrapper']/main/div/div/div/div[1]/div[2]/div/div/p[2]")

          content = content_elem.text
          content_split = content.split('\n')

          date_kw = 'lập và chuyển danh sách người sở hữu chứng khoán'

          mask_date = [date_kw in row for row in content_split]

          exchange_deposit = ''
          try:
            noidung \
              = np.array(content_split)[np.array(mask_date)][0]
          except IndexError:
            noidung = ''

          if noidung == '':
            try:
              # Cho tin: "cấp giấy chứng nhận đăng ký chứng khoán lần đầu"
              time.sleep(5)
              content_elem = sub_driver.find_element_by_xpath(
                '//*[@id="wrapper"]/main/div/div/div/div[1]/div[2]/div/div/p[2]/b')
              noidung = content_elem.text
              content_elem = sub_driver.find_element_by_xpath(
                '//*[@id="wrapper"]/main/div/div/div/div[1]/div[2]/div/div/div[14]')
              exchange_deposit = content_elem.text
            except self.ignored_exceptions:
              noidung = ''
              exchange_deposit = ''

          news_dict['Nội dung'] = noidung
          news_dict['Đăng ký trên sàn'] = exchange_deposit

          try:
            news_date_actives += [news_dict['Ngày hiệu lực hủy đăng ký']]
          except KeyError:
            news_date_actives += ['']
          try:
            news_depository += [news_dict['Đăng ký trên sàn']]
          except KeyError:
            news_depository += ['']
          try:
            news_noidung += [news_dict['Nội dung']]
          except KeyError:
            news_noidung += ['']

          sub_driver.quit()

        else:
          pass

      output_table = pd.concat([output_table,
                                pd.DataFrame(
                                  {'Thời gian'                :news_time,
                                   'Tiêu đề'                  :news_headlines,
                                   'Ngày hiệu lực hủy đăng ký':news_date_actives,
                                   'Đăng ký trên sàn'         :news_depository,
                                   'Nội dung'                 :news_noidung,
                                   'Link'                     :news_urls}
                                )],ignore_index=True)

      # Turn Page
      time.sleep(1)
      nextpage_button \
        = WebDriverWait(driver,5,ignored_exceptions=self.ignored_exceptions) \
        .until(
        EC.presence_of_all_elements_located(
          (By.XPATH,
           "//button[substring(@onclick,1,string-length('changePage'))"
           "='changePage']")))[-2]

      nextpage_button.click()

      # Check time
      last_tag \
        = WebDriverWait(driver,5,ignored_exceptions=self.ignored_exceptions) \
        .until(
        EC.presence_of_all_elements_located(
          (By.XPATH,
           '//*[@id="d_list_news"]/ul/li')))[-1]

      fromtime = last_tag.find_element_by_tag_name('div').text
      fromtime = dt.datetime.strptime(fromtime[-21:],
                                      '%d/%m/%Y - %H:%M:%S')

    driver.quit()

    if output_table.empty is True:
      raise NoNewsFound(f'Không có tin trong {num_hours} giờ vừa qua')

    def f(frame):
      series = frame['Tiêu đề'].str.split(': ')
      table = pd.DataFrame(series.tolist(),
                           columns=['cophieu','tieude'])
      return table

    output_table[['Mã cổ phiếu','Lý do']] = output_table.transform(f)
    output_table.drop(['Tiêu đề'],axis=1,inplace=True)

    def f(sentence):
      for exch in ['HOSE','HNX','UPCOM']:
        if exch in sentence:
          result = exch
        else:
          result = ''
      return result

    output_table['Đăng ký trên sàn'] = output_table['Đăng ký trên sàn'].map(f)

    output_table = output_table[['Thời gian',
                                 'Mã cổ phiếu',
                                 'Lý do',
                                 'Nội dung',
                                 'Đăng ký trên sàn',
                                 'Ngày hiệu lực hủy đăng ký',
                                 'Link']]

    print(f'Finished ::: Total execution time: {int(time.time()-start_time)}s\n')

    return output_table

  def tinTVLK(self,num_hours: int = 48) -> pd.DataFrame:

    """
    This function returns a DataFrame and export an excel file containing
    news update published in 'https://www.vsd.vn/vi/alc/4'
    (tin từ thành viên lưu ký).

    :param num_hours: number of hours in the past that's in our concern

    :return: summary report of news update
    """

    start_time = time.time()

    now = dt.datetime.now()
    fromtime = now

    url = 'https://www.vsd.vn/vi/alc/4'
    driver = webdriver.Chrome(executable_path=self.PATH)
    driver.get(url)
    driver.maximize_window()

    keywords = ['ngày hạch toán của cổ phiếu',
                'ngày hạch toán cổ phiếu',
                'ngày hạch toán của chứng quyền',
                'ngày hạch toán chứng quyền']
    excluded_words = ['bổ sung']

    output_table = pd.DataFrame()
    bmk_time = btime(now.strftime('%Y-%m-%d %H:%M:%S'),-num_hours)
    while fromtime >= dt.datetime.strptime(bmk_time,'%Y-%m-%d %H:%M:%S'):

      news_time = []
      news_headlines = []
      news_urls = []
      news_tradedate = []

      time.sleep(1)

      tags \
        = WebDriverWait(driver,5,ignored_exceptions=self.ignored_exceptions) \
        .until(
        EC.presence_of_all_elements_located(
          (By.XPATH,"//*[@id='d_list_news']/ul/li")))

      for tag_ in tags:

        tags \
          = WebDriverWait(driver,5,
                          ignored_exceptions=self.ignored_exceptions) \
          .until(
          EC.presence_of_all_elements_located(
            (By.XPATH,"//*[@id='d_list_news']/ul/li")))

        try:
          h3_tag \
            = WebDriverWait(tag_,5,
                            ignored_exceptions=self.ignored_exceptions) \
            .until(
            EC.presence_of_element_located(
              (By.TAG_NAME,'h3')))
        except self.ignored_exceptions:
          h3_tag = tag_.find_element_by_tag_name('h3')

        txt = h3_tag.text
        check_1 = [word not in txt for word in excluded_words]
        check_2 = [word in txt for word in keywords]

        if all(check_1) and any(check_2):

          news_headlines += [txt]

          sub_url = h3_tag.find_element_by_tag_name('a') \
            .get_attribute('href')
          news_urls += [f'=HYPERLINK("{sub_url}","Link")']

          news_time_ = tag_.find_element_by_tag_name('div').text
          news_time_ \
            = dt.datetime.strptime(news_time_[-21:],
                                   '%d/%m/%Y - %H:%M:%S')
          news_time += [news_time_]

          sub_driver = webdriver.Chrome(executable_path=self.PATH)
          sub_driver.get(sub_url)

          info_heads \
            = sub_driver.find_elements_by_xpath(
            "//div[substring(@class,string-length(@class)"
            "-string-length('item-info')+1)='item-info']")

          info_heads_text = [head.text[:-1]
                             for head in info_heads]

          info_tails \
            = sub_driver.find_elements_by_xpath(
            "//div[substring(@class,string-length(@class)"
            "-string-length('item-info-main')+1)='item-info-main']")
          info_tails_text = [tail.text for tail in info_tails]

          news_dict = {h:t for h,t in
                       zip(info_heads_text,info_tails_text)}

          news_head = [t for t in info_heads_text
                       if t.startswith('Ngày giao dịch')][0]
          try:
            news_tradedate += [news_dict[news_head]]
          except KeyError:
            news_tradedate += ['']

          sub_driver.quit()

      output_table = pd.concat([output_table,
                                pd.DataFrame(
                                  {'Thời gian'                :news_time,
                                   'Tiêu đề'                  :news_headlines,
                                   'Ngày giao dịch chính thức':news_tradedate,
                                   'Link'                     :news_urls}
                                )],ignore_index=True)

      def f(frame):
        series = frame['Tiêu đề'].str.split(': ')
        table = pd.DataFrame(series.tolist(),
                             columns=['cophieu/chungquyen','tieude'])
        return table

      if output_table['Tiêu đề'].empty:
        pass
      else:
        output_table[['Mã cổ phiếu / chứng quyền','Lý do, mục đích']] \
          = output_table.transform(f)

      # Turn Page
      nextpage_button \
        = WebDriverWait(driver,5,
                        ignored_exceptions=self.ignored_exceptions) \
        .until(
        EC.presence_of_all_elements_located(
          (By.XPATH,
           "//*[@id='d_number_of_page']/button")))[-2]

      nextpage_button.click()

      # Check time
      last_tag \
        = WebDriverWait(driver,5,
                        ignored_exceptions=self.ignored_exceptions) \
        .until(
        EC.presence_of_all_elements_located(
          (By.XPATH,
           '//*[@id="d_list_news"]/ul/li')))[-1]

      fromtime = last_tag.find_element_by_tag_name('div').text
      fromtime = dt.datetime.strptime(fromtime[-21:],
                                      '%d/%m/%Y - %H:%M:%S')

    driver.quit()

    if output_table.empty is True:
      raise NoNewsFound(f'Không có tin trong {num_hours} giờ vừa qua')

    output_table.drop(['Tiêu đề'],axis=1,inplace=True)

    output_table = output_table[['Thời gian',
                                 'Mã cổ phiếu / chứng quyền',
                                 'Lý do, mục đích',
                                 'Ngày giao dịch chính thức',
                                 'Link']]

    print(f'Finished ::: Total execution time: {int(time.time()-start_time)}s\n')

    return output_table


vsd = vsd()


class hnx:

  def __init__(self):

    self.PATH = join(dirname(dirname(realpath(__file__))),'dependency',
                     'chromedriver')

    self.ignored_exceptions \
      = (ValueError,
         IndexError,
         NoSuchElementException,
         StaleElementReferenceException,
         TimeoutException,
         ElementNotInteractableException)

  def tintuso(self,num_hours: int = 48):

    """
    This function returns a DataFrame and export an excel file containing
    news update published in
    'https://www.hnx.vn/thong-tin-cong-bo-ny-hnx.html' (tin từ sở) and
    'https://www.hnx.vn/vi-vn/thong-tin-cong-bo-up-hnx.html' (tin từ sở)

    :param num_hours: number of hours in the past that's in our concern

    :return: summary report of news update
    """

    start_time = time.time()

    now = dt.datetime.now()
    from_time = now

    driver = webdriver.Chrome(executable_path=self.PATH)

    url = 'https://www.hnx.vn/thong-tin-cong-bo-ny-hnx.html'
    driver.get(url)
    driver.maximize_window()

    key_words = ['ngày giao dịch đầu tiên cổ phiếu niêm yết',
                 'ngày giao dịch đầu tiên của cổ phiếu niêm yết',
                 'ngày giao dịch đầu tiên cổ phiếu chuyển giao dịch',
                 'ngày giao dịch đầu tiên của cổ phiếu chuyển giao dịch',
                 'ngày hủy niêm yết',
                 'ngày giao dịch đầu tiên cổ phiếu đăng ký giao dịch',
                 'ngày giao dịch đầu tiên của cổ phiếu đăng ký giao dịch',
                 'ngày hủy đăng ký giao dịch cổ phiếu',
                 'ngày hủy đăng ký giao dịch của cổ phiếu',
                 'ngày hủy ĐKGD cổ phiếu',
                 'ngày hủy ĐKGD của cổ phiếu']

    excluded_words = ['bổ sung']

    links = []
    box_text = []
    titles = []
    times = []
    tickers = []

    bmk_time = btime(now.strftime('%Y-%m-%d %H:%M:%S'),-num_hours)
    while from_time >= dt.datetime.strptime(bmk_time,'%Y-%m-%d %H:%M:%S'):

      def f():
        # wait for the element to appear, avoid stale element reference
        ticker_elems \
          = WebDriverWait(driver,5,ignored_exceptions=self.ignored_exceptions) \
          .until(
          EC.presence_of_all_elements_located(
            (By.XPATH,"//*[@id='_tableDatas']/tbody/*/td[3]/a")))
        title_elems \
          = WebDriverWait(driver,5,ignored_exceptions=self.ignored_exceptions) \
          .until(
          EC.presence_of_all_elements_located(
            (By.XPATH,"//*[@id='_tableDatas']/tbody/*/td[4]/a")))
        return ticker_elems,title_elems

      try:
        ticker_elems,title_elems = f()
        tickers_inpage = [t.text for t in ticker_elems if t.text != '']
        titles_inpage = [t.text for t in title_elems if t.text != '']
        title_elems = [t for t in title_elems if t.text != '']
      except self.ignored_exceptions:
        time.sleep(15)
        ticker_elems,title_elems = f()
        tickers_inpage = [t.text for t in ticker_elems if t.text != '']
        titles_inpage = [t.text for t in title_elems if t.text != '']
        title_elems = [t for t in title_elems if t.text != '']

      ticker_title_inpage = [f'{ticker} {title}' for ticker,title
                             in zip(tickers_inpage,titles_inpage)]

      for ticker_title in ticker_title_inpage:

        subcheck_1 = [key_word in ticker_title for key_word in key_words]
        subcheck_2 = [excl_word not in ticker_title for excl_word in excluded_words]

        if any(subcheck_1) and all(subcheck_2):
          pass
        else:
          continue

        title_no = ticker_title_inpage.index(ticker_title)+1
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
          .until(EC.presence_of_element_located(
          (By.XPATH,"//div[@class='Box-Noidung']")))

        content = popup_content.text
        box_text += [content]
        link_elems = driver.find_elements_by_xpath(
          "//div[@class='divLstFileAttach']/p/a")
        links += [[link.get_attribute('href') for link in link_elems]]

        # close popup windows
        WebDriverWait(driver,5,
                      ignored_exceptions=self.ignored_exceptions) \
          .until(EC.element_to_be_clickable(
          (By.XPATH,'//*[@id="divViewDetailArticles"]/*/input'))).click()

        time.sleep(1)

        # check time
        from_time = driver.find_element_by_xpath(
          "//*[@id='_tableDatas']/tbody/tr[10]/td[2]").text
        from_time = dt.datetime.strptime(from_time,'%d/%m/%Y %H:%M')

      WebDriverWait(driver,5,ignored_exceptions=self.ignored_exceptions) \
        .until(EC.element_to_be_clickable(
        (By.XPATH,"//*[@id='next']"))).click()

      time.sleep(1)

    url = 'https://www.hnx.vn/vi-vn/thong-tin-cong-bo-up-hnx.html'
    driver.get(url)

    driver.maximize_window()

    fromtime = now

    bmk_time = btime(now.strftime('%Y-%m-%d %H:%M:%S'),-num_hours)
    while fromtime >= dt.datetime.strptime(bmk_time,'%Y-%m-%d %H:%M:%S'):

      def f():
        # wait for the element to appear, avoid stale element reference
        ticker_elems \
          = WebDriverWait(driver,5,ignored_exceptions=self.ignored_exceptions) \
          .until(
          EC.presence_of_all_elements_located(
            (By.XPATH,"//*[@id='_tableDatas']/tbody/*/td[3]/a")))
        title_elems \
          = WebDriverWait(driver,5,ignored_exceptions=self.ignored_exceptions) \
          .until(
          EC.presence_of_all_elements_located(
            (By.XPATH,
             "//*[@id='_tableDatas']/tbody/*/td[4]/a")))
        return ticker_elems,title_elems

      try:
        ticker_elems,title_elems = f()
        tickers_inpage = [t.text for t in ticker_elems if t.text != '']
        titles_inpage = [t.text for t in title_elems if t.text != '']
        title_elems = [t for t in title_elems if t.text != '']
      except self.ignored_exceptions:
        time.sleep(15)
        ticker_elems,title_elems = f()
        tickers_inpage = [t.text for t in ticker_elems if t.text != '']
        titles_inpage = [t.text for t in title_elems if t.text != '']
        title_elems = [t for t in title_elems if t.text != '']

      ticker_title_inpage = [f'{ticker} {title}' for ticker,title
                             in zip(tickers_inpage,titles_inpage)]

      for ticker_title in ticker_title_inpage:
        sub_check = [key_word in ticker_title for key_word in key_words]
        if any(sub_check):
          pass
        else:
          continue

        title_no = ticker_title_inpage.index(ticker_title)+1
        titles += [driver.find_elements_by_xpath
                   (f"//*[@id='_tableDatas']/tbody/tr[{title_no}]/td[4]")[0].text]
        times += [driver.find_elements_by_xpath
                  (f"//*[@id='_tableDatas']/tbody/tr[{title_no}]/td[2]")[0].text]
        tickers += [driver.find_elements_by_xpath
                    (f"//*[@id='_tableDatas']/tbody/tr[{title_no}]/td[3]")[0].text]

        # open popup window:
        click_obj = title_elems[title_no-1]
        click_obj.click()
        time.sleep(1)

        popup_content \
          = WebDriverWait(driver,5,ignored_exceptions=self.ignored_exceptions) \
          .until(
          EC.presence_of_element_located(
            (By.XPATH,"//div[@class='Box-Noidung']")))

        content = popup_content.text
        if content in ['.','',' ']:  # nếu có link, ko có nội dung
          box_text += ['']
          link_elems = driver.find_elements_by_xpath(
            "//div[@class='divLstFileAttach']/p/a")
          links += [[link.get_attribute('href') for link in link_elems]]
        else:  # nếu có nội dung, ko có link
          links += ['']
          box_text += [content]

        WebDriverWait(driver,5,ignored_exceptions=self.ignored_exceptions) \
          .until(
          EC.element_to_be_clickable(
            (By.XPATH,"//*[@id='divViewDetailArticles']/*/input"))).click()

        time.sleep(1)

        # check time
        fromtime = driver.find_element_by_xpath(
          "//*[@id='_tableDatas']/tbody/tr[10]/td[2]").text
        fromtime = dt.datetime.strptime(fromtime,'%d/%m/%Y %H:%M')

      WebDriverWait(driver,5,ignored_exceptions=self.ignored_exceptions) \
        .until(EC.element_to_be_clickable(
        (By.XPATH,"//*[@id='next']"))).click()

      time.sleep(1)

    driver.quit()

    # export to DataFrame
    df = pd.DataFrame(list(zip(times,tickers,titles,box_text,links)),
                      columns=['Thời gian','Mã CK',
                               'Tiêu đề','Nội dung',
                               'File đính kèm'])

    def f(x):

      year = int(x.split('/')[2][:4])
      month = int(x.split('/')[1][:2])
      day = int(x.split('/')[0][:2])
      hour = int(x.split(':')[0][-2:])
      minute = int(x.split(':')[1][:2])
      second = 0

      return dt.datetime(year,month,day,hour,minute,second)

    df['Thời gian'] = df['Thời gian'].map(f)
    df.sort_values('Thời gian',ascending=False,inplace=True)

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
    print(f'Finished ::: Total execution time: {int(time.time()-start_time)}s\n')

    return df


hnx = hnx()


class hose:

  def __init__(self):

    self.PATH = join(dirname(dirname(realpath(__file__))),'dependency','chromedriver')

    self.ignored_exceptions \
      = (ValueError,
         IndexError,
         NoSuchElementException,
         StaleElementReferenceException,
         TimeoutException,
         ElementNotInteractableException)

  def tinTCNY(self,num_hours: int = 48):

    """
    This function returns a DataFrame and export an excel file containing
    news update published in
    'https://www.hsx.vn/Modules/Cms/Web/NewsByCat/dca0933e-a578-4eaf-8b29-beb4575052c5?fid=6d1f1d5e9e6c4fb593077d461e5155e7'
    (tin về hoạt động của tổ chức niêm yết)

    :param num_hours: number of hours in the past that's in our concern

    :return: summary report of news update
    """

    start_time = time.time()

    url = 'https://www.hsx.vn/Modules/Cms/Web/NewsByCat/dca0933e-a578-4eaf-8b29-beb4575052c5?fid=6d1f1d5e9e6c4fb593077d461e5155e7'
    driver = webdriver.Chrome(executable_path=self.PATH)
    driver.get(url)

    driver.maximize_window()

    now = dt.datetime.now()
    from_time = now

    keywords = ['niêm yết và ngày giao dịch đầu tiên',
                'hủy niêm yết','hủy đăng ký',
                'chuyển giao dịch']

    excl_keywords = ['bổ sung']

    output_table = pd.DataFrame()
    bmk_time = btime(now.strftime('%Y-%m-%d %H:%M:%S'),-num_hours)
    while from_time >= dt.datetime.strptime(bmk_time,'%Y-%m-%d %H:%M:%S'):

      time.sleep(5)

      headline_tags \
        = WebDriverWait(driver,5,
                        ignored_exceptions=self.ignored_exceptions) \
            .until(
        EC.presence_of_all_elements_located(
          (By.XPATH,'*//td[3]//*')))[1:]

      headline_text = [t.text for t in headline_tags]

      time_tags \
        = WebDriverWait(driver,5,
                        ignored_exceptions=self.ignored_exceptions) \
            .until(
        EC.presence_of_all_elements_located(
          (By.XPATH,'*//td[2]')))[2:]

      time_text = [t.text for t in time_tags]

      def f(s):
        if s.endswith('SA'):
          s = s.rstrip(' SA')
          return dt.datetime.strptime(s,'%d/%m/%Y %H:%M:%S')
        if s.endswith('CH'):
          s = s.rstrip(' CH')
          return dt.datetime.strptime(s,'%d/%m/%Y %H:%M:%S')+dt.timedelta(hours=12)

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
          time.sleep(1)
          sub_object.click()

          try:
            content = driver.find_element_by_xpath(
              '/html/body/div[7]/div/div/div/div/div/div[2]/div[2]').text
          except self.ignored_exceptions:
            content = WebDriverWait(driver,5,
                                    ignored_exceptions
                                    =self.ignored_exceptions) \
              .until(
              EC.visibility_of_element_located(
                (By.XPATH,'/html/body/div[7]/div/div/div/div/div/div[2]/div[2]'))).text
          sub_content = [content]

          time.sleep(1)

          year_text = str(from_time.year)
          try:
            pdf_elements \
              = driver.find_elements_by_partial_link_text(year_text)
          except self.ignored_exceptions:
            pdf_elements = WebDriverWait(driver,5,
                                         ignored_exceptions
                                         =self.ignored_exceptions) \
              .until(
              EC.presence_of_all_elements_located(
                (By.PARTIAL_LINK_TEXT,year_text)))

          pdf_files = [t.get_attribute('href') for t in pdf_elements]

          def f1(full_list):
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

          sub_pdfs += [f1(pdf_files)]

          # close popup windows
          driver.find_element_by_xpath('//*[@title="Close"]').click()
          time.sleep(1)

        # report and merge to parent table
        frames = pd.DataFrame({'Thời gian'    :sub_time,
                               'Tiêu đề'      :sub_title,
                               'Nội dung'     :sub_content,
                               'File đính kèm':sub_pdfs})

        output_table = pd.concat([output_table,frames],
                                 ignore_index=True)

      time.sleep(1)
      from_time = time_text[-1]

      # Next Page:
      time.sleep(1)
      try:
        driver.find_elements_by_xpath(
          '//*[@id="DbGridPager_2"]/a')[-2].click()
      except self.ignored_exceptions:
        WebDriverWait(driver,20,
                      ignored_exceptions=self.ignored_exceptions) \
          .until(
          EC.presence_of_all_elements_located(
            (By.XPATH,'//*[@id="DbGridPager_2"]/a')))[-2].click()

    driver.quit()

    if output_table.empty is True:
      raise NoNewsFound(f'Không có tin trong {num_hours} giờ vừa qua')

    # select out tickers from headline
    output_table.insert(1,'Mã cổ phiếu',
                        output_table['Tiêu đề'].str.split(': ').str.get(0))

    print(f'Finished ::: Total execution time: {int(time.time()-start_time)}s\n')

    return output_table

  def tinCW(self,num_hours: int = 48):

    """
    This function returns a DataFrame and export an excel file containing
    news update published in
    'https://www.hsx.vn/Modules/Cms/Web/NewsByCat/95cd3266-e6d1-42a3-beb5-20ed010aea4a?fid=f91940eef3384bcdbe96ee9aa3eefa04'
    (tin chứng quyền).

    :param num_hours: number of hours in the past that's in our concern

    :return: summary report of news update
    """

    start_time = time.time()

    now = dt.datetime.now()
    from_time = now

    url = 'https://www.hsx.vn/Modules/Cms/Web/NewsByCat/95cd3266-e6d1-42a3-beb5-20ed010aea4a?fid=f91940eef3384bcdbe96ee9aa3eefa04'

    driver = webdriver.Chrome(executable_path=self.PATH)
    driver.get(url)

    driver.maximize_window()

    keywords = ['chấp thuận niêm yết chứng quyền']

    output_table = pd.DataFrame()
    bmk_time = btime(now.strftime('%Y-%m-%d %H:%M:%S'),-num_hours)
    while from_time >= dt.datetime.strptime(bmk_time,'%Y-%m-%d %H:%M:%S'):

      time.sleep(5)

      try:
        headline_tags = driver.find_elements_by_xpath('*//td[3]/a')[1:]
      except self.ignored_exceptions:
        headline_tags \
          = WebDriverWait(driver,5,
                          ignored_exceptions=self.ignored_exceptions) \
              .until(
          EC.presence_of_all_elements_located(
            (By.XPATH,'*//td[3]/*')))[1:]

      headline_text = []
      headline_url = []
      for t in headline_tags:
        headline_text += [t.text]
        headline_url += [t.get_attribute('href')]

      try:
        time_tags = driver.find_elements_by_xpath('*//td[2]')[2:]
      except self.ignored_exceptions:
        time_tags \
          = WebDriverWait(driver,5,
                          ignored_exceptions=self.ignored_exceptions) \
              .until(
          EC.presence_of_all_elements_located(
            (By.XPATH,'*//td[2]')))[2:]

      time_text = [t.text for t in time_tags]

      def f(s):
        if s.endswith('SA'):
          s = s.rstrip(' SA')
          return dt.datetime.strptime(s,'%d/%m/%Y %H:%M:%S')
        if s.endswith('CH'):
          s = s.rstrip(' CH')
          return dt.datetime.strptime(s,'%d/%m/%Y %H:%M:%S') \
                 +dt.timedelta(hours=12)

      time_text = [f(t) for t in time_text]

      for num in range(len(headline_text)):

        sub_title = headline_text[num]
        sub_url = headline_url[num]
        sub_time = time_text[num]

        sub_content = []
        sub_pdfs = []

        check = [word in sub_title for word in keywords]

        if any(check):

          sub_driver = webdriver.Chrome(executable_path=self.PATH)
          sub_driver.get(sub_url)

          time.sleep(1)
          try:
            content = sub_driver.find_element_by_xpath(
              '//*[@id="body"]//*//div[2]/p').text
          except self.ignored_exceptions:
            content = WebDriverWait(sub_driver,5,
                                    ignored_exceptions
                                    =self.ignored_exceptions) \
              .until(
              EC.presence_of_element_located(
                (By.XPATH,'//*[@id="body"]//*//div[2]/p'))).text

          time.sleep(1)
          try:
            pdf_elements = sub_driver.find_elements_by_xpath(
              '*//td[2]/a')
          except self.ignored_exceptions:
            pdf_elements = WebDriverWait(sub_driver,5,
                                         ignored_exceptions
                                         =self.ignored_exceptions) \
              .until(
              EC.presence_of_all_elements_located(
                (By.XPATH,'*//td[2]/a')))

          pdf_files = [t.get_attribute('href') for t in pdf_elements]

          def f1(full_list):
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

          sub_content += [content]
          sub_pdfs += [f1(pdf_files)]

          sub_driver.quit()

        # report and merge to parent table
        frames = pd.DataFrame({'Thời gian'    :sub_time,
                               'Tiêu đề'      :sub_title,
                               'Nội dung'     :sub_content,
                               'File đính kèm':sub_pdfs,
                               'Link'         :f'=HYPERLINK("{sub_url}","Link")'})

        output_table = pd.concat([output_table,frames],
                                 ignore_index=True)

      from_time = time_text[-1]

      # Next Page:
      WebDriverWait(driver,5,ignored_exceptions=self.ignored_exceptions) \
        .until(EC.visibility_of_all_elements_located(
        (By.XPATH,'//*[@id="DbGridPager_2"]/a')))[-2].click()

    driver.quit()

    if output_table.empty is True:
      raise NoNewsFound(f'Không có tin trong {num_hours} giờ vừa qua')

    print(f'Finished ::: Total execution time: {int(time.time()-start_time)}s\n')

    return output_table


hose = hose()
