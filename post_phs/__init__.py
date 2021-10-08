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

        now = dt.datetime.now()
        github_file_name = f'{now.day}.{now.month}.{now.year}.csv'
        github_table_path = join(destination_dir_github,github_file_name)
        table = pd.DataFrame(columns=['Ticker','Group','BreakFeven Price','1% at Risk','3% at Risk','5% at Risk','Proposed Price'])
        table.set_index(keys=['Ticker'],inplace=True)
        table.to_csv(github_table_path)

        rmd_file_name = f'{now.day}.{now.month}.{now.year}.csv'
        rmd_table_path = join(destination_dir_rmd,rmd_file_name)
        table = pd.DataFrame(
            columns=['Ticker','Group','Breakeven Price','Proposed Price']
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
    def risk_alert(
            mlist:bool=True,
            exchange:str='HOSE',
            segment:str='all'
    ) \
            -> pd.DataFrame:

        """
        This method returns list of tickers that up ceil (fc_type='ceil')
        or down floor (fc_type='floor') in a given exchange of a given segment
        in n consecutive trading days

        :param mlist: report margin list only (True) or not (False)
        :param exchange: allow values in fa.exchanges. Do not allow 'all'
        :param segment: allow values in fa.segments or 'all'.
        For mlist=False only, if mlist=True left as default

        :return: pd.DataFrame (columns: 'Ticker', 'Exchange', 'Consecutive Days'
        """

        start_time = time.time()
        now = dt.datetime.now().strftime('%Y-%m-%d')
        # xét 3 tháng gần nhất, nếu sửa ở đây thì phải sửa dòng ngay dưới
        since = bdate(now,-22*3)
        three_month_ago = since

        mrate_series = internal.margin['mrate']
        drate_series = internal.margin['drate']
        maxprice_series = internal.margin['max_price']
        general_room_seires = internal.margin['general_room']
        special_room_series = internal.margin['special_room']
        total_room_series = internal.margin['total_room']

        if mlist is True:
            full_tickers = internal.mlist(exchanges=[exchange])
        else:
            full_tickers = fa.tickers(segment,exchange)

        records = []
        for ticker in full_tickers:
            df = ta.hist(ticker,fromdate=since,todate=now)[['trading_date','ref','close','total_volume']]
            df.set_index('trading_date',drop=True,inplace=True)
            df[['ref','close']] = (df[['ref','close']]*1000).round(0).astype(np.int64)
            df['total_volume'] = df['total_volume'].astype(np.int64)
            df['floor'] = df['ref'].apply(
                fc_price,
                price_type='floor',
                exchange=exchange
            )
            # giam san lien tiep
            n_floor = 0
            for i in range(df.shape[0]):
                condition1 = (df.loc[df.index[-i-1:],'floor']==df.loc[df.index[-i-1:],'close']).all()
                condition2 = (df.loc[df.index[-i-1:],'floor']!=df.loc[df.index[-i-1:],'ref']).all()
                # the second condition is to ignore trash tickers whose price
                # less than 1000 VND (a single price step equivalent to
                # more than 7%(HOSE), 10%(HNX), 15%(UPCOM))
                if condition1 and condition2:
                    n_floor += 1
                else:
                    break
            # mat thanh khoan trong 1 thang
            avg_vol_1m = df.loc[df.index[-22]:,'total_volume'].mean()
            n_illiquidity = (df.loc[df.index[-22]:,'total_volume']<avg_vol_1m).sum()
            # thanh khoan ngay gan nhat so voi thanh khoan trung binh 1 thang
            volume = df.loc[df.index[-1],'total_volume']
            volume_change_1m = volume/avg_vol_1m - 1

            n_illiquidity_bmk = 1
            n_floor_bmk = 1

            condition1 = n_floor >= n_floor_bmk
            condition2 =  n_illiquidity >= n_illiquidity_bmk

            if condition1 or condition2:
                print(ticker, '::: Failed')
                mrate = mrate_series.loc[ticker]
                drate = drate_series.loc[ticker]
                avg_vol_1w = df.loc[df.index[-5]:,'total_volume'].mean()
                avg_vol_1m = df.loc[df.index[-22]:,'total_volume'].mean()
                avg_vol_3m = df.loc[three_month_ago:,'total_volume'].mean() # để tránh out-of-bound error
                max_price = maxprice_series.loc[ticker]
                general_room = general_room_seires.loc[ticker]
                special_room = special_room_series.loc[ticker]
                total_room = total_room_series.loc[ticker]
                room_on_avg_vol_3m = total_room/avg_vol_3m
                record = pd.DataFrame({
                    'Stock': [ticker],
                    'Exchange': [exchange],
                    'Tỷ lệ vay KQ (%)': [mrate],
                    'Tỷ lệ vay TC (%)': [drate],
                    'Giá vay / Giá TSĐB tối đa (VND)': [max_price],
                    'General Room': [general_room],
                    'Special Room': [special_room],
                    'Total Room': [total_room],
                    'Consecutive Floor Days': [n_floor],
                    'Last day Volume': [volume],
                    '% Last day volume / 1M Avg.': [volume_change_1m],
                    '1W Avg. Volume': [avg_vol_1w],
                    '1M Avg. Volume': [avg_vol_1m],
                    '3M Avg. Volume': [avg_vol_3m],
                    'Approved Room / Avg. Liquidity 3 months': [room_on_avg_vol_3m],
                    '1M Illiquidity Days': [n_illiquidity],
                })
                records.append(record)
            else:
                print(ticker, '::: Passed')

            print('-------------------------')

        result_table = pd.concat(records,ignore_index=True)
        result_table.sort_values('Consecutive Floor Days',ascending=False,inplace=True)

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

        now = dt.date.today()
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
    def news_rmd():

        """
        This method runs all functions in module text_mining.newsrmd (try till success)
        and exports all resulted DataFrames to a single excel file in the specified folder
        for daily usage of RMD. This function is called in a higher-level module and
        automatically run on a daily basis

        :return: None
        """
        now = dt.datetime.now()
        time_string = now.strftime('%Y%m%d_@%H%M')
        if dt.time(hour=0,minute=0,second=0) <= now.time() <= dt.time(hour=11,minute=59,second=59):
            previous_bdate = bdate(now.strftime('%Y-%m-%d'),-1)
            time_point = dt.datetime.strptime(previous_bdate,'%Y-%m-%d')
            time_point = time_point.replace(hour=18,minute=0,second=0,microsecond=0)
        if dt.time(hour=12,minute=0,second=0) <= now.time() <= dt.time(hour=23,minute=59,second=59):
            time_point = now.replace(hour=10,minute=0,second=0,microsecond=0)

        path = r'\\192.168.10.101\phs-storge-2018' \
               r'\RiskManagementDept\RMD_Data' \
               r'\Luu tru van ban\RMC Meeting 2018' \
               r'\00. Meeting minutes\Data\News Update'
        file_name = f'{time_string}.xlsx'
        file_path = fr'{path}\{file_name}'

        while True:
            try:
                try:
                    vsd_TCPH = newsrmd.vsd.tinTCPH()
                    break
                except newsrmd.ignored_exceptions:
                    vsd_TCPH = pd.DataFrame()
                    continue
            except newsrmd.NoNewsFound:
                vsd_TCPH = pd.DataFrame()
                break

        while True:
            try:
                try:
                    vsd_TVBT = newsrmd.vsd.tinTVBT()
                    break
                except newsrmd.ignored_exceptions:
                    vsd_TVBT = pd.DataFrame()
                    continue
            except newsrmd.NoNewsFound:
                vsd_TVBT = pd.DataFrame()
                break

        while True:
            try:
                try:
                    hnx_TCPH = newsrmd.hnx.tinTCPH()
                    break
                except newsrmd.ignored_exceptions:
                    hnx_TCPH = pd.DataFrame()
                    continue
            except newsrmd.NoNewsFound:
                hnx_TCPH = pd.DataFrame()
                break

        while True:
            try:
                try:
                    hnx_tintuso = newsrmd.hnx.tintuso()
                    break
                except newsrmd.ignored_exceptions:
                    hnx_tintuso = pd.DataFrame()
                    continue
            except newsrmd.NoNewsFound:
                hnx_tintuso = pd.DataFrame()
                break

        while True:
            try:
                try:
                    hose_tintonghop = newsrmd.hose.tintonghop()
                    break
                except newsrmd.ignored_exceptions:
                    hose_tintonghop = pd.DataFrame()
                    continue
            except newsrmd.NoNewsFound:
                hose_tintonghop = pd.DataFrame()
            break

        check_dt = lambda dt_time: True if dt_time > time_point else False
        mask_vsd_TCPH = vsd_TCPH['Thời gian'].map(check_dt)
        mask_vsd_TVBT = vsd_TVBT['Thời gian'].map(check_dt)
        mask_hnx_TCPH = hnx_TCPH['Thời gian'].map(check_dt)
        mask_hnx_tintuso = hnx_tintuso['Thời gian'].map(check_dt)
        mask_hose_tintonghop = hose_tintonghop['Thời gian'].map(check_dt)

        writer = pd.ExcelWriter(file_path,engine='xlsxwriter')
        workbook = writer.book
        vsd_TCPH_sheet = workbook.add_worksheet('vsd_TCPH')
        vsd_TVBT_sheet = workbook.add_worksheet('vsd_TVBT')
        hnx_TCPH_sheet = workbook.add_worksheet('hnx_TCPH')
        hnx_tintuso_sheet = workbook.add_worksheet('hnx_tintuso')
        hose_tintonghop_sheet = workbook.add_worksheet('hose_tintonghop')

        header_fmt = workbook.add_format(
            {
                'align': 'center',
                'valign': 'vcenter',
                'bold': True,
                'border': 1,
                'text_wrap': True,
            }
        )
        regular_fmt = workbook.add_format(
            {
                'valign': 'vcenter',
                'border': 1,
                'text_wrap': True,
            }
        )
        highlight_regular_fmt = workbook.add_format(
            {
                'valign': 'vcenter',
                'border': 1,
                'bg_color': '#FFF024',
                'text_wrap': True,
            }
        )
        time_fmt = workbook.add_format(
            {
                'num_format': 'dd/mm/yyyy hh:mm:ss',
                'valign': 'vcenter',
                'border': 1,
                'text_wrap': True,
            }
        )
        highlight_time_fmt = workbook.add_format(
            {
                'num_format': 'dd/mm/yyyy hh:mm:ss',
                'valign': 'vcenter',
                'border': 1,
                'bg_color': '#FFF024',
                'text_wrap': True,
            }
        )
        vsd_TCPH_sheet.hide_gridlines(option=2)
        vsd_TVBT_sheet.hide_gridlines(option=2)
        hnx_TCPH_sheet.hide_gridlines(option=2)
        hnx_tintuso_sheet.hide_gridlines(option=2)
        hose_tintonghop_sheet.hide_gridlines(option=2)

        vsd_TCPH_sheet.set_column('A:A',18)
        vsd_TCPH_sheet.set_column('B:D',7)
        vsd_TCPH_sheet.set_column('E:E',50)
        vsd_TCPH_sheet.set_column('F:F',7)
        vsd_TCPH_sheet.write_row('A1',vsd_TCPH.columns,header_fmt)
        for col in range(vsd_TCPH.shape[1]):
            for row in range(vsd_TCPH.shape[0]):
                if col == 0 and mask_vsd_TCPH.loc[row] == True:
                    vsd_TCPH_sheet.write(row+1,col,vsd_TCPH.iloc[row,col],highlight_time_fmt)
                elif col != 0 and mask_vsd_TCPH.loc[row] == True:
                    vsd_TCPH_sheet.write(row+1,col,vsd_TCPH.iloc[row,col],highlight_regular_fmt)
                elif col == 0 and mask_vsd_TCPH.loc[row] == False:
                    vsd_TCPH_sheet.write(row+1,col,vsd_TCPH.iloc[row,col],time_fmt)
                else:
                    vsd_TCPH_sheet.write(row+1,col,vsd_TCPH.iloc[row,col],regular_fmt)

        vsd_TVBT_sheet.set_column('A:A',18)
        vsd_TVBT_sheet.set_column('B:B',90)
        vsd_TVBT_sheet.set_column('C:C',7)
        vsd_TVBT_sheet.write_row('A1',vsd_TVBT.columns,header_fmt)
        for col in range(vsd_TVBT.shape[1]):
            for row in range(vsd_TVBT.shape[0]):
                if col == 0 and mask_vsd_TVBT.loc[row] == True:
                    vsd_TVBT_sheet.write(row+1,col,vsd_TVBT.iloc[row,col],highlight_time_fmt)
                elif col != 0 and mask_vsd_TVBT.loc[row] == True:
                    vsd_TVBT_sheet.write(row+1,col,vsd_TVBT.iloc[row,col],highlight_regular_fmt)
                elif col == 0 and mask_vsd_TVBT.loc[row] == False:
                    vsd_TVBT_sheet.write(row+1,col,vsd_TVBT.iloc[row,col],time_fmt)
                else:
                    vsd_TVBT_sheet.write(row+1,col,vsd_TVBT.iloc[row,col],regular_fmt)

        hnx_TCPH_sheet.set_column('A:A',18)
        hnx_TCPH_sheet.set_column('B:D',7)
        hnx_TCPH_sheet.set_column('E:E',30)
        hnx_TCPH_sheet.set_column('F:F',120)
        hnx_TCPH_sheet.set_column('G:G',7)
        hnx_TCPH_sheet.write_row('A1',hnx_TCPH.columns,header_fmt)
        for col in range(hnx_TCPH.shape[1]):
            for row in range(hnx_TCPH.shape[0]):
                if col == 0 and mask_hnx_TCPH.loc[row] == True:
                    hnx_TCPH_sheet.write(row+1,col,hnx_TCPH.iloc[row,col],highlight_time_fmt)
                elif col != 0 and mask_hnx_TCPH.loc[row] == True:
                    hnx_TCPH_sheet.write(row+1,col,hnx_TCPH.iloc[row,col],highlight_regular_fmt)
                elif col == 0 and mask_hnx_TCPH.loc[row] == False:
                    hnx_TCPH_sheet.write(row+1,col,hnx_TCPH.iloc[row,col],time_fmt)
                else:
                    hnx_TCPH_sheet.write(row+1,col,hnx_TCPH.iloc[row,col],regular_fmt)

        hnx_tintuso_sheet.set_column('A:A',18)
        hnx_tintuso_sheet.set_column('B:D',7)
        hnx_tintuso_sheet.set_column('E:E',70)
        hnx_tintuso_sheet.set_column('F:F',7)
        hnx_tintuso_sheet.write_row('A1',hnx_tintuso,header_fmt)
        for col in range(hnx_tintuso.shape[1]):
            for row in range(hnx_tintuso.shape[0]):
                if col == 0 and mask_hnx_tintuso.loc[row] == True:
                    hnx_tintuso_sheet.write(row+1,col,hnx_tintuso.iloc[row,col],highlight_time_fmt)
                elif col != 0 and mask_hnx_tintuso.loc[row] == True:
                    hnx_tintuso_sheet.write(row+1,col,hnx_tintuso.iloc[row,col],highlight_regular_fmt)
                elif col == 0 and mask_hnx_tintuso.loc[row] == False:
                    hnx_tintuso_sheet.write(row+1,col,hnx_tintuso.iloc[row,col],time_fmt)
                else:
                    hnx_tintuso_sheet.write(row+1,col,hnx_tintuso.iloc[row,col],regular_fmt)

        hose_tintonghop_sheet.set_column('A:A',18)
        hose_tintonghop_sheet.set_column('B:D',7)
        hose_tintonghop_sheet.set_column('E:E',70)
        hose_tintonghop_sheet.set_column('F:F',80)
        hose_tintonghop_sheet.write_row('A1',hose_tintonghop,header_fmt)
        for col in range(hose_tintonghop.shape[1]):
            for row in range(hose_tintonghop.shape[0]):
                if col == 0 and mask_hose_tintonghop.iloc[row] == True:
                    hose_tintonghop_sheet.write(row+1,col,hose_tintonghop.iloc[row,col],highlight_time_fmt)
                elif col != 0 and mask_hose_tintonghop.iloc[row] == True:
                    hose_tintonghop_sheet.write(row+1,col,hose_tintonghop.iloc[row,col],highlight_regular_fmt)
                elif col == 0 and mask_hose_tintonghop.iloc[row] == False:
                    hose_tintonghop_sheet.write(row+1,col,hose_tintonghop.iloc[row,col],time_fmt)
                else:
                    hose_tintonghop_sheet.write(row+1,col,hose_tintonghop.iloc[row,col],regular_fmt)

        writer.close()


post = post()

