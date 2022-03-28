import cv2.cv2
import pytesseract

from automation.finance import *


class Base:

    def __init__(
        self,
        bank:str
    ):
        """
        Base class for all Bank classes
        :param bank: bank name
        """

        self.bank = bank
        self.ignored_exceptions = (
            ValueError,
            IndexError,
            NoSuchElementException,
            StaleElementReferenceException,
            TimeoutException,
            ElementNotInteractableException,
            ElementClickInterceptedException
        )
        self.PATH = join(dirname(dirname(dirname(realpath(__file__)))),'dependency','chromedriver')
        self.user, self.password, self.URL = get_bank_authentication(self.bank).values()

        self.downloadFolder = r'C:\Users\hiepdang\Downloads'


class BIDV(Base):

    def __init__(self):
        super().__init__('BIDV')
        self.driver = webdriver.Chrome(executable_path=self.PATH)
        self.driver.maximize_window()
        self.driver.get(self.URL)
        self.wait = WebDriverWait(self.driver,5,ignored_exceptions=self.ignored_exceptions)

    def __repr__(self):
        return f'<BankObject_BIDV>'

    def TienGuiThanhToan(
        self,
        fromDate:dt.datetime,
        toDate:dt.datetime,
    ):

        """
        :param fromDate: Ngày bắt đầu lấy dữ liệu
        :param toDate: Ngầy kết thúc lấy dữ liệu
        """

        """
        Bước 1: Đăng nhập
        """
        while True:
            CAPTCHA = self.wait.until(EC.presence_of_element_located((By.ID,'idImageCap')))
            imgPATH = join(dirname(__file__),'CAPTCHA',f'{self.bank}.png')
            CAPTCHA.screenshot(imgPATH) # download CAPTCHA về dưới dạng .png
            predictedCAPTCHA = pytesseract.image_to_string(imgPATH).replace('\n','').replace(' ','')
            print(predictedCAPTCHA)
            highRiskChars = 'ACGJTZWSIacgijqzws145789[].,/|'
            if len(set(highRiskChars) & set(predictedCAPTCHA)) == 0 and len(predictedCAPTCHA) == 6: # Cases do not need refresh CAPTCHA
                # Input CAPTCHA
                captchaInput = self.driver.find_element(By.ID,'captcha')
                captchaInput.clear()
                captchaInput.send_keys(predictedCAPTCHA)
                # Input Username
                userInput = self.driver.find_element(By.ID,'username')
                userInput.clear()
                userInput.send_keys(self.user)
                # Input Password
                userInput = self.driver.find_element(By.ID,'password')
                userInput.clear()
                userInput.send_keys(self.password)
                # Click "Đăng nhập"
                loginButton = self.driver.find_element(By.ID,'btLogin')
                loginButton.click()
                break
            self.driver.find_element(By.CLASS_NAME,'btnRefresh').click() # Cases need refresh CAPTCHA

        """
        Bước 2: Lấy dữ liệu
        """
        # Click Menu bar
        self.wait.until(EC.presence_of_element_located((By.ID,'menu-toggle-22'))).click()
        # Click "Tài khoản"
        self.wait.until(EC.presence_of_element_located((By.LINK_TEXT,'Vấn tin'))).click()
        # Click "Tiền gửi thanh toán"
        self.wait.until(EC.presence_of_element_located((By.LINK_TEXT,'Tiền gửi thanh toán'))).click()
        # Dọn dẹp folder trước khi download
        for file in listdir(self.downloadFolder):
            if 'EBK_BC_TIENGOITHANHTOAN' in file:
                os.remove(join(self.downloadFolder,file))
        # Bắt đầu download files
        frames = []
        for d in pd.date_range(fromDate,toDate):
            dateString = d.strftime('%d/%m/%Y')
            dateInput = self.wait.until(EC.presence_of_element_located((By.ID,'searchDate')))
            dateInput.clear()
            dateInput.send_keys(dateString)
            exportButton = self.wait.until(EC.presence_of_element_located((By.CLASS_NAME,'export')))
            exportButton.click()
            while True: # chờ đến khi download xong
                lastBIDVFile = first(
                    listdir(self.downloadFolder),
                    lambda x: ('EBK_BC_TIENGOITHANHTOAN' in x) and ('crdownload' not in x)
                )
                if lastBIDVFile is not None:
                    break
                else:
                    time.sleep(1)

            # Read file to DataFrame
            filePATH = join(self.downloadFolder,lastBIDVFile)
            dateData = pd.read_excel(filePATH,usecols='C:N',skiprows=7,skipfooter=1,engine='xlrd')
            dateData = dateData.dropna(how='all',axis=1)
            dateData.insert(0,'Date',d)
            dateData.columns = [
                'Date',
                'AccountNumber',
                'AccountName',
                'Balance',
                'Blockade',
                'Availability',
                'OverdraftLimit',
                'Currency',
                'OpenDate',
                'Status',
                'OpenBranch',
            ]
            frames.append(dateData)
            # Delete downloaded file
            os.remove(filePATH)

        self.driver.quit()
        table = pd.concat(frames)

        return table


class VTB(Base):

    def __init__(self):
        super().__init__('VTB')
        self.driver = webdriver.Chrome(executable_path=self.PATH)
        self.driver.get(self.URL)
        self.driver.maximize_window()
        self.wait = WebDriverWait(self.driver,5,ignored_exceptions=self.ignored_exceptions)

    def __repr__(self):
        return f'<BankObject_VTB>'

    def TienGuiThanhToan(
        self,
        fromDate:dt.datetime,
        toDate:dt.datetime,
    ):

        """
        :param fromDate: Ngày bắt đầu lấy dữ liệu
        :param toDate: Ngầy kết thúc lấy dữ liệu
        """

        """
        Bước 1: Đăng nhập
        """
        # Username
        userInput = self.wait.until(EC.presence_of_element_located((By.XPATH,'//*[@placeholder="Tên đăng nhập"]')))
        userInput.clear()
        userInput.send_keys(self.user)
        # Password
        passwordInput = self.wait.until(EC.presence_of_element_located((By.XPATH,'//*[@placeholder="Mật khẩu"]')))
        passwordInput.clear()
        passwordInput.send_keys(self.password)
        # Click đăng nhập
        loginButton = self.wait.until(EC.presence_of_element_located((By.XPATH,'//*[@type="submit"]')))
        loginButton.click()

        """
        Bước 2: Lấy dữ liệu
        """
        # Click menu "Tài khoản"
        self.wait.until(EC.presence_of_element_located((By.LINK_TEXT,'Tài khoản'))).click()
        # Click sub-menu "Danh sách tài khoản"
        self.wait.until(EC.presence_of_element_located((By.LINK_TEXT,'Danh sách tài khoản'))).click()
        # Dọn dẹp folder trước khi download
        for file in listdir(self.downloadFolder):
            if 'danh-sach-tai-khoan-thanh-toan' in file:
                os.remove(join(self.downloadFolder,file))
        # table Element
        tableElements = self.wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME,'MuiTableBody-root')))
        tableElement = tableElements[0]
        tableElement.find_element(By.LINK_TEXT,'Xem thêm').click()

        # Create function to clear input box and send dates as string
        def sendDate(element,d):
            action = ActionChains(self.driver)
            action.click(element)
            action.key_down(Keys.CONTROL,element)
            action.send_keys_to_element(element,'a')
            action.key_up(Keys.CONTROL,element)
            action.send_keys_to_element(element,Keys.BACKSPACE)
            action.send_keys_to_element(element,d.strftime('%d/%m/%Y'))
            action.send_keys_to_element(element,Keys.ENTER)
            action.perform()

        records = []
        accountNumbers = filter(lambda x: len(x)==12,tableElement.text.split('\n'))
        for x in accountNumbers:
            self.wait.until(EC.presence_of_element_located((By.LINK_TEXT,x))).click()
            fromDateInput,toDateInput = self.wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME,'ant-picker-input')))

            for d in pd.date_range(fromDate,toDate):
                # Điền ngày
                sendDate(fromDateInput,d)
                sendDate(toDateInput,d)
                self.wait.until(EC.presence_of_element_located((By.CLASS_NAME,'btn-submit'))).click()
                # Lấy số dư
                rawResult = self.driver.find_elements(By.XPATH,'//*[@class="f-price"]/p/b')[1].text
                closingBalance,currency = rawResult.split()
                records.append((d,x,float(closingBalance.replace(',','')),currency))

            self.driver.back()
            self.wait.until(EC.presence_of_element_located((By.LINK_TEXT,'Xem thêm'))).click()

        self.driver.quit()
        balanceTable = pd.DataFrame(
            data = records,
            columns=['Date','AccountNumber','Balance','Currency']
        )

        return balanceTable


class IVB(Base):

    def __init__(self):
        super().__init__('IVB')
        self.driver = webdriver.Chrome(executable_path=self.PATH)
        self.driver.get(self.URL)
        self.driver.maximize_window()
        self.wait = WebDriverWait(self.driver,10,ignored_exceptions=self.ignored_exceptions)

    def __repr__(self):
        return f'<BankObject_IVB>'

    def TienGuiThanhToan(
        self,
        fromDate:dt.datetime,
        toDate:dt.datetime,
    ):

        """
        :param fromDate: Ngày bắt đầu lấy dữ liệu
        :param toDate: Ngầy kết thúc lấy dữ liệu
        """

        """
        Bước 1: Đăng nhập
        """
        # Username
        xpath = '//*[@placeholder="Tên truy cập"]'
        userInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        userInput.clear()
        userInput.send_keys(self.user)
        # Password
        xpath = '//*[@placeholder="Mật khẩu"]'
        passwordInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        passwordInput.clear()
        passwordInput.send_keys(self.password)
        # CAPTCHA
        xpath = '//*[@placeholder="Mã xác thực"]'
        captchaInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        CAPTCHA = input('Nhập CAPTCHA của IVB để tiếp tục:') # ko đọc được tự động
        captchaInput.clear()
        captchaInput.send_keys(CAPTCHA)
        # Click Đăng nhập
        self.wait.until(EC.presence_of_element_located((By.XPATH,'//*[@onclick="logon()"]'))).click()
        # Nếu xuất hiện màn hình xác nhận
        xpath = '//*[@onclick="forceSubmit()"]'
        possibleButtons = self.driver.find_elements(By.XPATH,xpath)
        if possibleButtons:
            possibleButtons[0].click()

        """
        Bước 2: Lấy dữ liệu
        """

        # Click tab "Tài khoản"
        xpath = '//*[@data-menu-id="1"]'
        self.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()
        # Click subtab "Sao kê tài khoản"
        self.wait.until(EC.visibility_of_element_located((By.ID,'2_2'))).click()
        # Chọn "Tài khoản thanh toán" từ dropdown list
        self.driver.switch_to.frame('mainframe')
        accountTypeInput = Select(self.wait.until(EC.presence_of_element_located((By.ID,'selectedAccType'))))
        accountTypeInput.select_by_visible_text('Tài khoản Thanh toán')

        # Điền số tài khoản:
        accountElems = self.driver.find_elements(By.XPATH,'//*[@id="account_list"]/option')
        options = [a.text for a in accountElems]
        records = []
        for option in options:
            time.sleep(1)
            account = option.split()[0]
            currency = option.split()[-1].replace(']','')
            if account not in ('1017816-066','1017816-069','1017816-068'):
                continue
            accountInput = Select(self.wait.until(EC.presence_of_element_located((By.ID,'account_list'))))
            accountInput.select_by_visible_text(option)

            # Điền ngày
            for d in pd.date_range(fromDate,toDate):
                # Từ ngày
                fromDateInput = self.driver.find_element(By.ID,'beginDate')
                fromDateInput.clear()
                fromDateInput.send_keys(d.strftime('%d/%m/%Y'))
                # Đến ngày
                toDateInput = self.driver.find_element(By.ID,'endDate')
                toDateInput.clear()
                toDateInput.send_keys(d.strftime('%d/%m/%Y'))
                # Click "Truy vấn"
                self.driver.find_element(By.ID,'btnQuery').click()
                # Lấy số dư (cuối kỳ)
                valueStrings = [e.text for e in self.driver.find_elements(By.XPATH,'//*[@class="result_head"]')]
                balanceString = first(valueStrings, lambda x: 'Số dư cuối kỳ' in x)
                try:
                    balance = float(balanceString.split()[-1].replace(',',''))
                except (ValueError,): # catch lỗi web không hiện số khi ko có dữ liệu
                    balance = 0
                records.append(tuple([d,account,balance,currency]))

        self.driver.quit()
        balanceTable = pd.DataFrame(
            records,
            columns=['Date','AccountNumber','Balance','Currency']
        )

        return balanceTable


class VCB(Base):

    def __init__(self):
        super().__init__('VCB')
        self.driver = webdriver.Chrome(executable_path=self.PATH)
        self.driver.get(self.URL)
        self.driver.maximize_window()
        self.wait = WebDriverWait(self.driver,10,ignored_exceptions=self.ignored_exceptions)

    def __repr__(self):
        return f'<BankObject_VCB>'

    def TienGuiThanhToan(
        self,
        fromDate:dt.datetime,
        toDate:dt.datetime,
    ):

        """
        :param fromDate: Ngày bắt đầu lấy dữ liệu
        :param toDate: Ngầy kết thúc lấy dữ liệu
        """

        """
        Bước 1: Đăng nhập
        """
        # CAPTCHA
        while True:
            xpath = '//*[@alt="Chuỗi bảo mật"]'
            CAPTCHA = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
            imgPATH = join(dirname(__file__),'CAPTCHA',f'{self.bank}.png')
            CAPTCHA.screenshot(imgPATH)  # download CAPTCHA về dưới dạng .png
            predictedCAPTCHA = pytesseract.image_to_string(imgPATH).replace('\n','')
            if len(predictedCAPTCHA) == 6 and len(set(predictedCAPTCHA) & set('ABCDGILMOSZ23456789.,£€')) == 0:
                xpath = '//*[@placeholder="Nhập số bên"]'
                captchaInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
                captchaInput.clear()
                captchaInput.send_keys(predictedCAPTCHA)
                break # case không cần refresh
            self.driver.refresh() # case cần refresh

        # Username
        xpath = '//*[@placeholder="Tên truy cập"]'
        userInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        userInput.clear()
        userInput.send_keys(self.user)
        # Password
        xpath = '//*[@placeholder="Mật khẩu"]'
        passwordInput = self.wait.until(EC.presence_of_element_located((By.XPATH,xpath)))
        passwordInput.clear()
        passwordInput.send_keys(self.password)

        # Click 'Đăng nhập'
        xpath = '//*[@value="Đăng nhập"]'
        self.wait.until(EC.presence_of_element_located((By.XPATH,xpath))).click()

        """
        Bước 2: Lấy dữ liệu
        """
        # Lấy danh sách đường dẫn vào tài khoản thanh toán
        xpath = '//*[@id="dstkdd-tbody"]//td/a'
        accountElems = self.wait.until(EC.presence_of_all_elements_located((By.XPATH,xpath)))
        URLs = [e.get_attribute('href') for e in accountElems]

        # Create function to clear input box and send dates as string
        def sendDate(element,d):
            action = ActionChains(self.driver)
            action.click(element)
            action.key_down(Keys.CONTROL,element)
            action.send_keys_to_element(element,'a')
            action.key_up(Keys.CONTROL,element)
            action.send_keys_to_element(element,Keys.BACKSPACE)
            action.send_keys_to_element(element,d.strftime('%d/%m/%Y'))
            action.send_keys_to_element(element,Keys.ENTER)
            action.perform()

        records = []
        for URL in URLs:
            self.driver.get(URL)
            account = self.wait.until(EC.presence_of_element_located((By.)))
            for d in pd.date_range(fromDate,toDate):
                # Điền ngày
                startDateInput = self.wait.until(EC.presence_of_element_located((By.CLASS_NAME,'startDate')))
                loc = startDateInput.location_once_scrolled_into_view
                self.driver.execute_script(f'window.scrollTo(0, {loc["y"]})')
                sendDate(startDateInput,d)
                endDateInput = self.wait.until(EC.presence_of_element_located((By.CLASS_NAME,'endDate')))
                sendDate(endDateInput,d)
                # Click "Xem sao kê"
                self.wait.until(EC.presence_of_element_located((By.ID,'TransByDate'))).click()
                # Số dư cuối kỳ
                ID = 'ctl00_Content_TransactionDetail_EndBalance'
                balanceString = self.wait.until(EC.presence_of_element_located((By.ID,ID))).text
                balance = float(balanceString.replace(',',''))




