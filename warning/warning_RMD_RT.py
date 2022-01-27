"""
Cú pháp: run(True) hoặc run(False)  |  True: Ẩn browser window, False Không ẩn browser window
"""
from request.stock import *


def __formatPrice__(x):
    if not x:  # x == ''
        x = np.nan
    elif x in ('ATO','ATC'):
        x = np.nan
    else:
        x = float(x)
    return x

def __formatVolume__(x):
    if not x:  # x == ''
        x = np.nan
    else:
        x = float(x.replace(',',''))*10  # 2,30 -> 2300, 10 -> 100
    return x


def __SendMailRetry__(func): # Decorator

    def wrapper(*args,**kwargs):
        n = 0
        while True:
            now = dt.datetime.now()
            """
            Quét HOSE trước
            Quét HNX sau
            (vẫn quét 1 lần nếu chạy ngoài giờ giao dịch)
            """
            currentWarningsHOSE, lastWarningFile = func('HOSE',*args,**kwargs) # quét HOSE trước
            print(currentWarningsHOSE)
            currentWarningsHNX, lastWarningFile = func('HNX',*args,**kwargs) # quét HNX sau
            print(currentWarningsHNX)
            currentWarnings = pd.concat([currentWarningsHOSE,currentWarningsHNX])

            # Lưu file tạm để đọc ở lần quét sau
            currentWarnings.to_pickle(join(dirname(__file__),'TempFiles',lastWarningFile))

            if currentWarnings.empty: # Khi không có Warnings
                print(f"No warnings at {now.strftime('%H:%M:%S %d.%m.%Y')}")
                time.sleep(8*60) # sleep 8 phút
                continue

            # Khi có Warnings
            if isfile(join(dirname(__file__),'TempFiles',lastWarningFile)):
                # đọc kết quả lần quét cuối cùng
                lastWarnings = pd.read_pickle(join(dirname(__file__),'TempFiles',lastWarningFile))
                mergeTable = pd.merge(currentWarnings,lastWarnings,on=['Ticker','Message'],indicator=True,)

                repeatWarnings = mergeTable.loc[mergeTable['_merge']=='both',['Ticker','Message']]
                repeatWarningsHTML = repeatWarnings.to_html(index=False,header=False)
                repeatWarningsHTML = repeatWarningsHTML.replace(
                    'border="1"',
                    'style="border-collapse:collapse; border-spacing:0px;"',
                )  # remove borders

                newWarnings = mergeTable.loc[mergeTable['_merge']=='left_only',['Ticker','Message']]
                newWarningsHTML = newWarnings.to_html(index=False,header=False)
                newWarningsHTML = newWarningsHTML.replace(
                    'border="1"',
                    'style="border-collapse:collapse; border-spacing:0px; color:red"',
                )  # remove borders, set font color to red

                html_str = f"""
                <html>
                    <head></head>
                    <body>
                        {newWarningsHTML}
                        {repeatWarningsHTML}
                        <p style="font-family:Times New Roman; font-size:90%"><i>
                            -- Generated by Reporting System
                        </i></p>
                    </body>
                </html>
                """

            else:
                newWarningsHTML = currentWarnings.to_html()
                newWarningsHTML = newWarningsHTML.replace(
                    'border="1"',
                    'border="1" style="border-collapse:collapse","border-spacing:0px","color:red";',
                )  # remove borders, set font color to red

                html_str = f"""
                <html>
                    <head></head>
                    <body>
                        {newWarningsHTML}
                        <p style="font-family:Times New Roman; font-size:90%"><i>
                            -- Generated by Reporting System
                        </i></p>
                    </body>
                </html>
                """

            outlook = Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = 'vando@phs.vn; anhnguyenthu@phs.vn; huyhuynh@phs.vn;'
            mail.CC = 'hiepdang@phs.vn'
            mail.Subject = f"Market Alert {now.strftime('%H:%M:%S %d.%m.%Y')}"
            mail.HTMLBody = html_str
            mail.Send()

            n += 1
            print('Quét lần: ',n)

            # set các điều kiện break:
            if now.time()>dt.time(15,0,0):
                tempFiles = listdir(join(dirname(__file__),'TempFiles'))
                trashFiles = [f for f in tempFiles if not f.startswith('README')]
                for file in trashFiles:
                    os.remove(join(dirname(__file__),'TempFiles',file))
                break
            if dt.time(11,30,0)<now.time()<dt.time(13,0,0):
                break
            if now.time()<dt.time(9,0,0):
                break
            time.sleep(8*60)  # nghỉ 8 phút

    return wrapper


@__SendMailRetry__
def run(
    exchange:str,
    hide_window=True  # nên set=True để ko bị pop up cửa sổ browser liên tục trong phiên
) -> pd.DataFrame:

    PATH = join(dirname(dirname(realpath(__file__))),'dependency','chromedriver')
    ignored_exceptions = (
        ValueError,
        IndexError,
        NoSuchElementException,
        StaleElementReferenceException,
        TimeoutException,
        ElementNotInteractableException
    )

    start = time.time()

    if exchange == 'HOSE':
        url = 'https://iboard.ssi.com.vn/bang-gia/hose'
        mlist = internal.mlist(['HOSE'])
    elif exchange == 'HNX':
        url = 'https://iboard.ssi.com.vn/bang-gia/hnx'
        mlist = internal.mlist(['HNX'])
    else:
        raise ValueError('Currently monitor HOSE and HNX only')

    lastWarningFile = f"{dt.datetime.now().strftime('%Y.%m.%d')}.pickle"

    # produce/get avgVolume File
    todate = dt.datetime.now().strftime('%Y-%m-%d')
    avgVolumeFile = join(dirname(__file__),'TempFiles',f"{exchange}_AvgPrice_{todate.replace('-','.')}")
    if not isfile(avgVolumeFile):
        fromdate = bdate(todate,-22)
        avgVolume = pd.Series(index=mlist,dtype=object)
        for ticker in mlist:
            avgVolume.loc[ticker] = ta.hist(ticker,fromdate,todate)['total_volume'].mean()
        avgVolume.to_pickle(avgVolumeFile)
    else:
        avgVolume = pd.read_pickle(avgVolumeFile)

    options = Options()
    if hide_window:
        options.headless = True
    driver = webdriver.Chrome(executable_path=PATH,options=options)
    wait = WebDriverWait(driver,60,ignored_exceptions=ignored_exceptions)
    driver.get(url)

    warnings = pd.DataFrame(columns=['Message'],index=pd.Index(mlist,name='Ticker'))
    for ticker in warnings.index:
        tickerElement = wait.until(EC.presence_of_element_located((By.XPATH,f'//tbody/*[@id="{ticker}"]')))
        sub_wait = WebDriverWait(tickerElement,60,ignored_exceptions=ignored_exceptions)
        Floor = __formatPrice__(sub_wait.until(EC.presence_of_element_located((By.CLASS_NAME,'floor'))).text)
        MatchPrice = __formatPrice__(sub_wait.until(EC.presence_of_element_located((By.CLASS_NAME,'matchedPrice'))).text)
        SellVolume1 = __formatVolume__(sub_wait.until(EC.presence_of_element_located((By.CLASS_NAME,'best1OfferVol'))).text)
        SellVolume2 = __formatVolume__(sub_wait.until(EC.presence_of_element_located((By.CLASS_NAME,'best2OfferVol'))).text) 
        SellVolume3 = __formatVolume__(sub_wait.until(EC.presence_of_element_located((By.CLASS_NAME,'best3OfferVol'))).text)
        FrgBuy = __formatVolume__(sub_wait.until(EC.presence_of_element_located((By.CLASS_NAME,'buyForeignQtty'))).text)
        FrgSell = __formatVolume__(sub_wait.until(EC.presence_of_element_located((By.CLASS_NAME,'sellForeignQtty'))).text)
        if MatchPrice==Floor:
            # Điều kiện 1:
            if SellVolume1 + SellVolume2 + SellVolume3 >= 0.2 * avgVolume.loc[ticker]:
                warnings.loc[ticker,'Message'] = ': Giảm sàn, Dư bán hơn 20% KLTB'
            # Điều kiện 2:
            if FrgSell - FrgBuy >= 0.2 * avgVolume.loc[ticker]:
                warnings.loc[ticker,'Message'] = ': Giảm sàn, NN bán hơn 20% KLTB'

    warnings = warnings.dropna().reset_index()

    driver.quit()

    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')

    return warnings, lastWarningFile
