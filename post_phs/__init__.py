from request_phs.stock import *
from breakeven_price.monte_carlo import monte_carlo
from text_mining import newsts as newsts
from text_mining import newsrmd

class post:

    def __init__(self):
        pass

    @staticmethod
    def breakeven(
            tickers:list=None,
            exchanges:list=None,
            alpha:float=0.05,
            standard:str='bics',
            level:int=1,
    ) -> pd.DataFrame:

        """
        This method post Monte Carlo model's results in breakeven_price
        sub-project to shared API and return the output table

        :param tickers: list of tickers, if 'all': run all tickers
        :param exchanges: list of exchanges, if 'all': run all tickers
        :param alpha: significant level of statistical tests
        :param standard: Industry Classification Standard for Credit Rating
        :param level: level of classification
        :return: pandas.DataFrame
        """

        start_time = time.time()

        destination_dir_github = join(dirname(dirname(realpath(__file__))),'breakeven_price','result_table')
        destination_dir_network = r'\\192.168.10.28\images\breakeven'
        destination_dir_rmd = r'\\192.168.10.101\phs-storge-2018\RiskManagementDept' \
                              r'\RMD_Data\Luu tru van ban\RMC Meeting 2018\00. Meeting minutes\Gia Hoa Von'

        chart_foler = join(realpath(dirname(dirname(__file__))),'breakeven_price','result_chart')

        if exchanges == 'all' or tickers == 'all':
            tickers = fa.tickers(exchange='all')
        elif exchanges is not None and exchanges != 'all':
            tickers = []
            for exchange in exchanges:
                tickers += fa.tickers(exchange=exchange)
        elif tickers is not None and tickers != 'all':
            pass

        # CREDIT RATING
        rating_path = join(realpath(dirname(dirname(__file__))),'credit_rating','result')
        # adjust result_table_gen
        gen_table = pd.read_csv(join(rating_path,'result_table_gen.csv'),index_col=['ticker'])
        gen_table = gen_table.loc[gen_table['level']==f'{standard}_l{level}']
        gen_table.drop(['standard','level','industry'],axis=1,inplace=True)
        # merge to ready-to-use result table
        bank_table = pd.read_csv(join(rating_path,'result_table_bank.csv'),index_col=['ticker'])
        ins_table = pd.read_csv(join(rating_path,'result_table_ins.csv'),index_col=['ticker'])
        sec_table = pd.read_csv(join(rating_path,'result_table_sec.csv'),index_col=['ticker'])
        rating_table = pd.concat([gen_table,bank_table,ins_table,sec_table])
        rating_table.drop(['BM_'],inplace=True)
        rating_table.fillna('Not enough data',inplace=True)

        network_table_path = join(destination_dir_network,'tables','result.csv')
        table = pd.DataFrame(columns=['ticker','Breakeven Price'])
        table.set_index(keys=['ticker'], inplace=True)
        table.to_csv(network_table_path)

        now = datetime.now()
        github_file_name = f'{now.day}.{now.month}.{now.year}.csv'
        github_table_path = join(destination_dir_github,github_file_name)
        table = pd.DataFrame(columns=['Ticker','Group','Breakeven Price','Proposed Price'])
        table.set_index(keys=['Ticker'],inplace=True)
        table.to_csv(github_table_path)

        rmd_file_name = f'{now.day}.{now.month}.{now.year}.csv'
        rmd_table_path = join(destination_dir_rmd,rmd_file_name)
        table = pd.DataFrame(
            columns=['Ticker','Group','Breakeven Price','1% at Risk','3% at Risk','5% at Risk','Proposed Price']
        )
        table.set_index(keys=['Ticker'], inplace=True)
        table.to_csv(rmd_table_path)

        for ticker in tickers:
            full_range = np.arange(start=0,stop=1.01,step=0.1)
            alpha_range = full_range[full_range>alpha]
            alpha_range = np.insert(alpha_range,0,alpha)
            for chosen_alpha in alpha_range:
                try:
                    rating = rating_table.loc[ticker,rating_table.columns[-1]]
                    breakeven_price,lv1_price,lv2_price,lv3_price = monte_carlo(ticker=ticker,alpha=chosen_alpha)
                    f = lambda x: int(adjprice(x).replace(',',''))
                    breakeven_price = f(breakeven_price)
                    if rating == 'Not enough data':
                        group = 'Not enough data'
                        proposed_price = 'Not enough data'
                    elif rating >= 75:
                        group = 'A'
                        proposed_price = f(lv3_price)
                    elif rating >= 50:
                        group = 'B'
                        proposed_price = f(lv2_price)
                    elif rating >= 25:
                        group = 'C'
                        proposed_price = f(lv1_price)
                    else:
                        group = 'D'
                        proposed_price = breakeven_price
                    with open(github_table_path,mode='a',newline='') as github_file:
                        github_writer = csv.writer(github_file,delimiter=',')
                        github_writer.writerow(
                            [ticker,group,breakeven_price,lv1_price,lv2_price,lv3_price,proposed_price]
                        )
                    with open(network_table_path,mode='a',newline='') as network_file:
                        network_writer = csv.writer(network_file,delimiter=',')
                        network_writer.writerow([ticker,breakeven_price])
                    with open(rmd_table_path,mode='a',newline='') as rmd_file:
                        rmd_writer = csv.writer(rmd_file,delimiter=',')
                        rmd_writer.writerow([ticker,group,breakeven_price,proposed_price])
                    break
                except (ValueError, KeyError, IndexError):
                    print(f'{ticker} cannot be simulated with given significance level, running with higher alpha instead')
                    print('-------')
            try:
                shutil.copy(join(chart_foler, f'{ticker}.png'),join(realpath(destination_dir_network), 'charts'))
            except FileNotFoundError:
                print(f'{ticker} cannot be graphed')

            print('===========================')

        print('Finished!')
        print("Total execution time is: %s seconds" %(time.time()-start_time))


    @staticmethod
    def credit_rating(
            standard:str='bics',
            level:int=1
    ):

        destination_path = r'\\192.168.10.28\images\creditrating'
        chart_path = os.path.join(destination_path, 'charts')
        table_path = os.path.join(destination_path, 'tables')

        # POST CREDIT RATING
        rating_path = join(realpath(dirname(dirname(__file__))), 'credit_rating', 'result')
        rating_files = [file for file in listdir(rating_path) if file.startswith('result') and 'gen' not in file]
        for file in rating_files:
            shutil.copy(join(rating_path, file), join(table_path, 'rating'))
        # adjust result_table_gen
        gen_table = pd.read_csv(join(rating_path, 'result_table_gen.csv'), index_col=['ticker'])
        gen_table = gen_table.loc[gen_table['level']==f'{standard}_l{level}']
        gen_table.drop(['standard','level','industry'], axis=1, inplace=True)
        gen_table.to_csv(join(table_path, 'rating', 'result_table_gen.csv'))
        # merge to ready-to-use result table
        bank_table = pd.read_csv(join(rating_path, 'result_table_bank.csv'), index_col=['ticker'])
        ins_table = pd.read_csv(join(rating_path, 'result_table_ins.csv'), index_col=['ticker'])
        sec_table = pd.read_csv(join(rating_path, 'result_table_sec.csv'), index_col=['ticker'])
        result_table = pd.concat([gen_table, bank_table, ins_table, sec_table])
        result_table.drop(['BM_'], inplace=True)
        result_table.to_csv(join(table_path, 'rating', 'result_summary.csv'))

        component_files = [file for file in listdir(join(rating_path, 'Compare with Industry'))]
        for file in component_files:
            shutil.copy(join(rating_path, 'Compare with Industry', file), join(chart_path, 'component'))

        result_files = [file for file in listdir(join(rating_path, 'Result'))]
        for file in result_files:
            shutil.copy(join(rating_path, 'Result', file), join(chart_path, 'rating'))

        # CREATE IMAGE
        rating_tickers = {name[:3] for name in result_files}
        component_tickers = {name[:3] for name in component_files}

        # intersect
        tickers = rating_tickers & component_tickers
        k_rating = 1.4
        for ticker in tickers:

            rating_image = Image.open(join(chart_path,'rating',f'{ticker}_result.png'))
            rating_image = rating_image.resize(
                (int(rating_image.width*k_rating),int(rating_image.height*k_rating))
            )
            component_image = Image.open(join(chart_path,'component',f'{ticker}_compare_industry.png'))
            component_image = component_image.resize(
                (int(rating_image.width),int(rating_image.width*component_image.height/component_image.width))
            )
            result_image = Image.new(
                'RGB',
                (rating_image.width,rating_image.height+component_image.height),
                (255,255,255)
            )
            result_image.paste(rating_image)
            result_image.paste(component_image, (0,rating_image.height))
            result_image.save(join(chart_path, 'result', f'{ticker}_result.png'))


    @staticmethod  # This function need to be generalized through any period of time and through multiple exchanges
    def fc_consecutive(fc_type:str, mlist:bool=True, consecutive_days:int=1,
                       exchange:str='HOSE', segment:str='all') \
            -> pd.DataFrame:

        """
        This method returns list of tickers that up ceil (fc_type='ceil')
        or down floor (fc_type='floor') in a given exchange of a given segment
        in n consecutive trading days

        :param fc_type: allow either 'ceil' (for ceil price) or 'floor' (for floor price)
        :param mlist: report margin list only (True) or not (False)
        :param consecutive_days: minimum number of ceil/floor days.
        :param exchange: allow values in fa.exchanges. Do not allow 'all'
        :param segment: allow values in fa.segments or 'all'.
        For mlist=False only, if mlist=True left as default

        :return: pd.DataFrame (columns: 'Ticker', 'Exchange', 'Consecutive Days'
        """

        start_time = time.time()

        if mlist is True:
            full_tickers = internal.mlist([exchange])
        else:
            full_tickers = fa.tickers(segment, exchange)

        result_table = pd.DataFrame(columns=['Ticker','Exchange','Consecutive Days'])
        for ticker in full_tickers:
            try:
                # xét 10 phiên gần nhất
                df = ta.hist(ticker)[['trading_date','ref','close']].tail(10)
                df.set_index('trading_date', drop=True, inplace=True)
                df = np.round(df*1000,0)
                df['floor'] = df['ref'].apply(
                    fc_price,
                    price_type=fc_type,
                    exchange=exchange
                )
                df['close'] = df['close'].map(lambda x: int(np.round(x)))
                print(ticker)
                print(df)
                n = 0
                for i in range(df.shape[0]):
                    condition1 = (df['floor'][-i-1:] == df['close'][-i-1:]).all()
                    print(condition1)
                    condition2 = (df['floor'][-i-1:] != df['ref'][-i-1:]).all()
                    # the second condition is to ignore trash tickers whose price
                    # less than 1000 VND (a single price step equivalent to
                    # more than 7%(HOSE), 10%(HNX), 15%(UPCOM))
                    if condition1 and condition2:
                        n += 1
                    else:
                        break

                if n > 0:
                    record = {'Ticker':ticker,
                              'Exchange':exchange,
                              'Consecutive Days':n}
                    result_table = result_table.append(record,ignore_index=True)
                    print(result_table)
                    result_table = result_table[
                        result_table['Consecutive Days'] >= consecutive_days]
                    print(ticker, '::: Failed')
                    print('Result:', result_table)
                else:
                    print(ticker, '::: Passed')

            except ValueError:
                print(f'{ticker} is not listed in specified exchange')

            print('-------------------------')

        result_table.set_index('Ticker', drop=True, inplace=True)

        print('Finished!')
        print("Total execution time is: %s seconds" % (time.time() - start_time))

        return result_table


    @staticmethod
    def news_ts(num_hours:int=48):

        """
        This method runs all functions in module text_mining.newsts (try till success)
        and exports all resulted DataFrames to a single excel file in the specified folder
        for daily usage of TS. This function is called in a higher-level module and
        automatically run on a daily basis

        :param num_hours: number of hours in the past that's in our concern

        :return: None
        """

        now = date.today()
        date_string = now.strftime('%Y%m%d')

        path = r'C:\Users\hiepdang\News Report\Trading Service'
        file_name = f'{date_string}_news.xlsx'
        file_path = fr'{path}\{file_name}'

        while True:
            try:
                try:
                    vsd_TCPH = newsts.vsd.tinTCPH(num_hours)
                    break
                except newsts.vsd.ignored_exceptions:
                    vsd_TCPH = pd.DataFrame()
                    continue
            except newsts.NoNewsFound:
                vsd_TCPH = pd.DataFrame()
                break

        while True:
            try:
                try:
                    vsd_TVLK = newsts.vsd.tinTVLK(num_hours)
                    break
                except newsts.vsd.ignored_exceptions:
                    vsd_TVLK = pd.DataFrame()
                    continue
            except newsts.NoNewsFound:
                vsd_TVLK = pd.DataFrame()
                break

        while True:
            try:
                try:
                    hnx_tintuso = newsts.hnx.tintuso(num_hours)
                    break
                except newsts.hnx.ignored_exceptions:
                    hnx_tintuso = pd.DataFrame()
                    continue
            except newsts.NoNewsFound:
                hnx_tintuso = pd.DataFrame()
                break

        while True:
            try:
                try:
                    hose_tinTCNY = newsts.hose.tinTCNY(num_hours)
                    break
                except newsts.hose.ignored_exceptions:
                    hose_tinTCNY = pd.DataFrame()
                    continue
            except newsts.NoNewsFound:
                hose_tinTCNY = pd.DataFrame()
                break

        while True:
            try:
                try:
                    hose_tinCW = newsts.hose.tinCW(num_hours)
                    break
                except newsts.hose.ignored_exceptions:
                    hose_tinCW = pd.DataFrame()
                    continue
            except newsts.NoNewsFound:
                hose_tinCW = pd.DataFrame()
                break

        with pd.ExcelWriter(file_path) as writer:
            vsd_TCPH.to_excel(writer, sheet_name='vsd_tinTCPH', index=False)
            vsd_TVLK.to_excel(writer, sheet_name='vsd_tinTVLK', index=False)
            hnx_tintuso.to_excel(writer, sheet_name='hnx_tintuso', index=False)
            hose_tinTCNY.to_excel(writer, sheet_name='hose_tinTCNY', index=False)
            hose_tinCW.to_excel(writer, sheet_name='hose_tinCW', index=False)

        wb = xw.Book(file_path)

        wb.sheets['vsd_tinTCPH'].autofit()
        wb.sheets['vsd_tinTCPH'].range("C:D").column_width = 50
        wb.sheets['vsd_tinTCPH'].range("A:XFD").api.WrapText = True
        wb.sheets['vsd_tinTCPH'].range("A:XFD").api.Cells.VerticalAlignment = 2

        wb.sheets['vsd_tinTVLK'].autofit()
        wb.sheets['vsd_tinTVLK'].range("C:C").column_width = 80
        wb.sheets['vsd_tinTVLK'].range("A:XFD").api.WrapText = True
        wb.sheets['vsd_tinTVLK'].range("A:XFD").api.Cells.VerticalAlignment = 2

        wb.sheets['hnx_tintuso'].autofit()
        wb.sheets['hnx_tintuso'].range("C:C").column_width = 80
        wb.sheets['hnx_tintuso'].range("A:XFD").api.WrapText = True
        wb.sheets['hnx_tintuso'].range("A:XFD").api.Cells.VerticalAlignment = 2

        wb.sheets['hose_tinTCNY'].autofit()
        wb.sheets['hose_tinTCNY'].range("C:D").column_width = 90
        wb.sheets['hose_tinTCNY'].range("A:XFD").api.WrapText = True
        wb.sheets['hose_tinTCNY'].range("A:XFD").api.Cells.VerticalAlignment = 2

        wb.sheets['hose_tinCW'].autofit()
        wb.sheets['hose_tinCW'].range("B:B").column_width = 45
        wb.sheets['hose_tinCW'].range("C:C").column_width = 90
        wb.sheets['hose_tinCW'].range("D:E").column_width = 20
        wb.sheets['hose_tinCW'].range("A:XFD").api.WrapText = True
        wb.sheets['hose_tinCW'].range("A:XFD").api.Cells.VerticalAlignment = 2

        wb.save()
        wb.close()


    @staticmethod
    def news_rmd(num_hours:int=48, fixed_mp:bool=True):

        """
        This method runs all functions in module text_mining.newsrmd (try till success)
        and exports all resulted DataFrames to a single excel file in the specified folder
        for daily usage of RMD. This function is called in a higher-level module and
        automatically run on a daily basis

        :param num_hours: number of hours in the past that's in our concern
        :param fixed_mp: care only about stock in "fixed max price" (True)
         or not (False)

        :return: None
        """


        now = datetime.now()
        time_string = now.strftime('%Y%m%d_@%H%M')

        path = r'\\192.168.10.101\phs-storge-2018' \
               r'\RiskManagementDept\RMD_Data' \
               r'\Luu tru van ban\RMC Meeting 2018' \
               r'\00. Meeting minutes\Data\News Update'
        file_name = f'{time_string}.xlsx'
        file_path = fr'{path}\{file_name}'

        while True:
            try:
                try:
                    vsd_TCPH = newsrmd.vsd.tinTCPH(num_hours,fixed_mp)
                    break
                except newsrmd.vsd.ignored_exceptions:
                    vsd_TCPH = pd.DataFrame()
                    continue
            except newsrmd.NoNewsFound:
                vsd_TCPH = pd.DataFrame()
                break

        while True:
            try:
                try:
                    vsd_TVBT = newsrmd.vsd.tinTVBT(num_hours)
                    break
                except newsrmd.vsd.ignored_exceptions:
                    vsd_TVBT = pd.DataFrame()
                    continue
            except newsrmd.NoNewsFound:
                vsd_TVBT = pd.DataFrame()
                break

        while True:
            try:
                try:
                    hnx_TCPH = newsrmd.hnx.tinTCPH(num_hours,fixed_mp)
                    break
                except newsrmd.hnx.ignored_exceptions:
                    hnx_TCPH = pd.DataFrame()
                    continue
            except newsrmd.NoNewsFound:
                hnx_TCPH = pd.DataFrame()
                break

        while True:
            try:
                try:
                    hnx_tintuso = newsrmd.hnx.tintuso(num_hours)
                    break
                except newsrmd.hnx.ignored_exceptions:
                    hnx_tintuso = pd.DataFrame()
                    continue
            except newsrmd.NoNewsFound:
                hnx_tintuso = pd.DataFrame()
                break

        while True:
            try:
                try:
                    hose_tintonghop = newsrmd.hose.tintonghop(num_hours,fixed_mp)
                    break
                except newsrmd.hose.ignored_exceptions:
                    hose_tintonghop = pd.DataFrame()
                    continue
            except newsrmd.NoNewsFound:
                hose_tintonghop = pd.DataFrame()
                break

        with pd.ExcelWriter(file_path) as writer:
            vsd_TCPH.to_excel(
                writer,
                sheet_name='vsd_tinTCPH',
                index=False
            )
            vsd_TVBT.to_excel(
                writer,
                sheet_name='vsd_tinTVBT',
                index=False
            )
            hnx_TCPH.to_excel(
                writer,
                sheet_name='hnx_tinTCPH',
                index=False
            )
            hnx_tintuso.to_excel(
                writer,
                sheet_name='hnx_tintuso',
                index=False
            )
            hose_tintonghop.to_excel(
                writer,
                sheet_name='hose_tintonghop',
                index=False
            )

        wb = xw.Book(file_path)

        wb.sheets['vsd_tinTCPH'].autofit()
        wb.sheets['vsd_tinTCPH'].range("D:E").column_width = 50
        wb.sheets['vsd_tinTCPH'].range("F:G").column_width = 17
        wb.sheets['vsd_tinTCPH'].range("H:H").column_width = 35
        wb.sheets['vsd_tinTCPH'].range("K:K").column_width = 5
        wb.sheets['vsd_tinTCPH'].range("A:XFD").api.WrapText = True
        wb.sheets['vsd_tinTCPH'].range("A:XFD").api.Cells.VerticalAlignment = 2

        wb.sheets['vsd_tinTVBT'].autofit()
        wb.sheets['vsd_tinTVBT'].range("C:C").column_width = 20
        wb.sheets['vsd_tinTVBT'].range("A:XFD").api.WrapText = True
        wb.sheets['vsd_tinTVBT'].range("A:XFD").api.Cells.VerticalAlignment = 2

        wb.sheets['hnx_tinTCPH'].autofit()
        wb.sheets['hnx_tinTCPH'].range("E:E").column_width = 80
        wb.sheets['hnx_tinTCPH'].range("A:XFD").api.WrapText = True
        wb.sheets['hnx_tinTCPH'].range("A:XFD").api.Cells.VerticalAlignment = 2

        wb.sheets['hnx_tintuso'].autofit()
        wb.sheets['hnx_tintuso'].range("D:D").column_width = 80
        wb.sheets['hnx_tintuso'].range("A:XFD").api.WrapText = True
        wb.sheets['hnx_tintuso'].range("A:XFD").api.Cells.VerticalAlignment = 2

        wb.sheets['hose_tintonghop'].autofit()
        wb.sheets['hose_tintonghop'].range("D:E").column_width = 90
        wb.sheets['hose_tintonghop'].range("A:XFD").api.WrapText = True
        wb.sheets['hose_tintonghop'].range("A:XFD").api.Cells.VerticalAlignment = 2

        wb.save()
        wb.close()


post = post()

