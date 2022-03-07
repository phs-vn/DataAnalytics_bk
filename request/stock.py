from function import *


###############################################################################

class internal:
    rmd_shared_disk = r'\\192.168.10.101\phs-storge-2018' \
                      r'\RiskManagementDept\RMD_Data\Luu tru van ban' \
                      r'\RMC Meeting 2018\00. Meeting minutes'

    # Constructor:
    def __init__(self):
        all_files = pd.Series(listdir(join(self.rmd_shared_disk,'Margin List')))
        file_times = all_files.str.split('_').str.get(1).str[:10]
        file_times = pd.to_datetime(file_times,format='%d.%m.%Y')
        self.latest = file_times.max().strftime(format='%d.%m.%Y')
        mask = all_files.str.find(self.latest).apply(lambda x:True if x!=-1 else False)
        self.file = all_files[mask].iloc[0]
        self.margin = pd.read_excel(
            join(self.rmd_shared_disk,'Margin List',self.file),
            index_col=[1],
            engine='openpyxl'
        )
        special_characters = [' ','(',')','%','/','VND']
        for c in special_characters:
            self.margin.columns = self.margin.columns.map(lambda x:x.replace(c,''))
        self.margin.columns = self.margin.columns.map(lambda x:x.lower())
        self.margin.columns = self.margin.columns.map(lambda x:unidecode.unidecode(x))
        self.margin = self.margin[[
            'tylevaykq',
            'tylevaytc',
            'giavaygiatsdbtoida',
            'roomchung',
            'specialroom',
            'totalroom',
            'sangiaodich',
        ]]
        self.margin.dropna(axis=0,how='all',inplace=True)
        self.margin.dropna(axis=1,how='all',inplace=True)
        col_names = ['mrate','drate','max_price','general_room','special_room','total_room','exchange']
        self.margin.columns = col_names

        self.fixedmp_series = pd.read_excel(join(self.rmd_shared_disk,'Data','Fixed Max Price.xlsx'),usecols=['Stock'],
                                            squeeze=True)
        self.fixedmp_list = self.fixedmp_series.tolist()

    def mlist(self,exchanges: list = 'all') -> list:

        """
        This method returns margin list of a given exchange

        :param exchanges: allow ['HOSE', 'HNX', 'UPCOM'] or 'all'
        :type exchanges: list

        :return: list
        """

        if exchanges=='all':
            exchanges = ['HOSE','HNX','UPCOM']
        table = self.margin.loc[self.margin['exchange'].isin(exchanges)]
        tickers = table.index.to_list()

        return tickers

    def mrate(self,ticker: str) -> float:

        """
        This method returns margin rate of a given ticker

        :param ticker: allow any ticker in mlist(exchanges='all')
        :type ticker: str

        :return: deposit rate
        """

        rate = self.margin.loc[ticker,'mrate']

        return rate

    def drate(self,ticker: str) -> float:

        """
        This method returns deposit rate of a given ticker

        :param ticker: allow any ticker in mlist(exchanges='all')
        :type ticker: str

        :return: deposit rate
        """

        rate = self.margin.loc[ticker,'drate']

        return rate

    def mprice(self,ticker: str) -> int:

        """
        This method returns max price of a given ticker

        :param ticker: allow any ticker in mlist(exchanges='all')
        :type ticker: str

        :return: max price
        """

        price = self.margin.loc[ticker,'max_price']

        return int(price)

    def groom(self,ticker: str) -> int:

        """
        This method returns general room of a given ticker

        :param ticker: allow any ticker in mlist(exchanges='all')
        :type ticker: str

        :return: general room
        """

        room = self.margin.loc[ticker,'general_room']

        return int(room)


internal = internal()


###############################################################################


class fa:
    database_path = join(dirname(dirname(realpath(__file__))),'database')
    financials = ['bank','sec','ins']

    # Constructor:
    def __init__(
        self
    ):
        self.folders = [f for f in listdir(self.database_path) if f.startswith('fs_')]
        self.segments = [x.split('_')[1] for x in self.folders]
        self.segments.sort()  # Returns all valid segments
        self.fs_types = [
            name.split('_')[0]
            for folder in self.folders for name in listdir(join(self.database_path,folder))
            if isfile(join(self.database_path,folder,name))
               and not name.startswith('~$')
               and not name.split('_')[0][-1].isdigit()
        ]
        self.fs_types = list(set(self.fs_types))
        self.fs_types.sort()  # Returns all valid financial statements

        self.periods = [
            name[-11:-5]
            for folder in self.folders
            for name in listdir(join(self.database_path,folder))
            if isfile(join(self.database_path,folder,name))
        ]
        self.periods = list(set(self.periods))
        self.periods.sort()  # Returns all periods
        self.latest_period = self.periods[-1]
        self.standards \
            = [name[:-5] for name in listdir(join(self.database_path,'industry'))
               if isfile(join(self.database_path,'industry',name))]  # Returns all industry classification standards

    def check_columns(
        self,
        segments: list = 'all',
    ) \
            -> None:

        """
        This function check consistency of columns in FiinPro data files and map.xlsx.
        Consistency is required for numbering collumns
        Limitation: Cannot check for fn1, fn2, fn3,...

        :return: None
        """

        # Check the template consistency
        if segments=='all':
            segments = self.segments
        map_excel = pd.ExcelFile(join(self.database_path,'map.xlsx'))
        for segment in segments:
            # map table
            map_table = map_excel.parse(
                sheet_name=segment,
                header=None,
                names=['fs_type','item','name'],
                dtype={'item':str}
            )
            map_table['fs_type'].fillna(method='ffill',inplace=True)
            map_table.set_index(['fs_type','item'],inplace=True)
            map_table = map_table.squeeze()
            map_table = map_table.str.replace(' ','')
            # data table
            folder = f'fs_{segment}'
            files = [
                f for f in listdir(join(self.database_path,folder))
                if not f.startswith('~') and f.endswith(f'.xlsm')
            ]
            files.sort()
            for file in files:
                raw_fiinpro = openpyxl.load_workbook(os.path.join(self.database_path,folder,file)).active
                raw_fiinpro.delete_rows(idx=raw_fiinpro.max_row-21,amount=1000)
                raw_fiinpro.delete_rows(idx=0,amount=7)
                raw_fiinpro.delete_rows(idx=2,amount=1)
                frame = pd.DataFrame(raw_fiinpro.values)
                headers = frame.iloc[0,4:]
                headers = headers.str.split('Year').str.get(0)
                headers = headers.str.replace(' ','')
                headers.name = 'name'
                fs_type = file.split('_')[0]
                if fs_type in ['fn1','fn2','fn3']:
                    continue
                assign_num = lambda x:[f'{i}.' for i in np.arange(x)+1]
                headers.index = pd.MultiIndex.from_product([[fs_type],assign_num(headers.shape[0])])
                result = map_table.reindex(headers.index).compare(headers)
                if not result.empty:
                    print(
                        fr'Not matching columns, please check {folder}\\{file}.\n'
                        'Check this error otherwise it might cause wrong data extraction.\n'
                    )
                    print(result)
                    conclude = input("Is everything OK? (y/n)")
                    if conclude.lower()=='y':
                        continue
                    elif conclude.lower()=='n':
                        return
                else:
                    print(fr'Checking ::: {folder}\\{file} ::: Result: Passed')

    def reload(
        self
    ) \
            -> None:

        """
        This method handles cached data in newly-added files

        :return: None
        """

        for folder in self.folders:
            file_names = [
                file for file in listdir(join(self.database_path,folder))
                if isfile(join(self.database_path,folder,file))
                   and not file.startswith('~')
            ]
            for file in file_names:
                excel = Dispatch("Excel.Application")
                excel.Visible = True
                excel.Workbooks.Open(
                    os.path.join(self.database_path,folder,file)
                )
                time.sleep(3)  # suspend 3 secs for excel to catch up python
                excel.Range("A1:XFD1048576").Select()
                excel.Selection.Copy()
                excel.Selection.PasteSpecial(Paste=-4163)
                excel.ActiveWorkbook.Save()
                excel.ActiveWorkbook.Close()

    def fin_tickers(
        self,
        sector_break: bool = False
    ) \
            -> Union[list,dict]:

        """
        This method returns all tickers of financial segments

        :param sector_break: False: ignore sectors, True: show sectors
        :type sector_break: bool

        :return: list (sector_break=False), dictionary (sector_break=True)
        """

        tickers = []
        tickers_ = dict()
        for segment in self.financials:
            folder = f'fs_{segment}'
            file = f'is_{self.latest_period}.xlsm'
            raw_fiinpro = openpyxl.load_workbook(join(self.database_path,folder,file)).active
            # delete StoxPlux Sign
            raw_fiinpro.delete_rows(idx=raw_fiinpro.max_row-21,amount=1000)
            # delete headers
            raw_fiinpro.delete_rows(idx=0,amount=7)
            raw_fiinpro.delete_rows(idx=2,amount=1)
            # import
            clean_data = pd.DataFrame(raw_fiinpro.values)
            clean_data.drop(index=[0],inplace=True)
            # remove OTC
            case = clean_data[3]!='OTC'
            if sector_break is False:
                tickers += clean_data.loc[case,1].tolist()
            else:
                tickers_[segment] = clean_data.loc[case,1].tolist()
                tickers = tickers_

        return tickers

    def core(
        self,
        year: int,
        quarter: int,
        fs_type: str,
        segment: str,
        exchange: str = 'all'
    ) \
            -> pd.DataFrame:

        """
        This method extracts data from Github server, clean up
        and make it ready for use

        :param year: reported year
        :param quarter: reported quarter
        :param fs_type: allow values in self.fs_types
        :param segment: allow values in self.segments
        :param exchange: allow values in ['HOSE', 'HNX', 'UPCOM'] or 'all'

        :return: pandas.DataFrame
        :raise ValueError: this function yet supported cashflow for securities companies
        """

        if segment not in self.segments:
            raise ValueError(f'sector must be in {self.segments}')
        if fs_type not in self.fs_types:
            raise ValueError(f'fs_type must be in {self.fs_types}')

        folder = f'fs_{segment}'
        files = [
            f for f in listdir(join(self.database_path,folder))
            if f.startswith(f'{fs_type}') and f.endswith(f'_{year}q{quarter}.xlsm')
        ]
        files.sort()  # to ensure correct order of concatination
        frames = []
        for file in files:
            # create Workbook object, select active Worksheet
            raw_fiinpro = openpyxl.load_workbook(os.path.join(self.database_path,folder,file)).active
            # delete StoxPlux Sign
            raw_fiinpro.delete_rows(idx=raw_fiinpro.max_row-21,amount=1000)
            # delete header rows
            raw_fiinpro.delete_rows(idx=0,amount=7)
            raw_fiinpro.delete_rows(idx=2,amount=1)
            # import to DataFrame, no column labels, no index
            frame = pd.DataFrame(raw_fiinpro.values)
            # assign column labels and index
            frame.columns = frame.iloc[0]
            frame.drop(index=0,inplace=True)
            frame.index = pd.MultiIndex.from_arrays(
                [[year]*frame.shape[0],[quarter]*frame.shape[0],frame['Ticker']]
            )
            frame.index.set_names(['year','quarter','ticker'],inplace=True)
            frame.drop(columns=['No','Ticker'],inplace=True)
            frame.fillna(0,inplace=True)
            # remove OTC
            frame = frame.loc[frame['Exchange']!='OTC']
            frames += [frame]
        clean_data = pd.concat(frames,axis=1)
        # drop finnacial tickers if segment = 'gen'
        if segment=='gen':
            fin = set(self.fin_tickers())&set(clean_data.index.get_level_values(2))
            clean_data.drop(index=list(fin),level=2,inplace=True)
        name_exchange = clean_data[['Name','Exchange']].copy()
        name_exchange = name_exchange.loc[:,~name_exchange.columns.duplicated()]
        clean_data.drop(['Name','Exchange'],axis=1,inplace=True)
        give_index = lambda x:[f'{num}.' for num in np.arange(x)+1]
        clean_data.columns = give_index(clean_data.shape[1])
        clean_data.insert(0,'full_name',name_exchange['Name'])
        clean_data.insert(0,'exchange',name_exchange['Exchange'])
        clean_data.columns = pd.MultiIndex.from_product(
            [[fs_type],clean_data.columns],names=['fs_type','item']
        )
        if exchange!='all':
            mask = clean_data[(fs_type,'exchange')]==exchange
            clean_data = clean_data.loc[mask]

        print('Data ::: Extracting...')
        return clean_data

    def segment(
        self,
        ticker: str
    ) \
            -> str:

        """
        This method returns the segment of a given ticker

        :param ticker: stock's ticker
        :type ticker: str

        :return: str
        """

        segment = ''
        if ticker not in self.fin_tickers():
            segment = 'gen'
        else:
            fin_dict = self.fin_tickers(True)
            for key in fin_dict:
                if ticker in fin_dict[key]:
                    segment = key
                    break

        return segment

    def fs(
        self,
        ticker: str
    ) \
            -> pd.DataFrame:

        """
        This method returns all financial statements
        of given ticker in all periods

        :param ticker: stock's ticker
        :type ticker: str

        :return: pandas.DataFrame
        """

        segment = self.segment(ticker)
        folder = f'fs_{segment}'
        file_names = listdir(join(self.database_path,folder))
        file_names = [f for f in file_names if not f.startswith('~$')]
        file_names.sort()
        refs = [
            (int(name[-11:-7]),int(name[-6]),name[:2] if name[2] in '_1234' else name[:3]) for name in file_names
        ]
        dict_idx = dict()
        for fs_type in self.fs_types:
            year = refs[-1][0]
            quarter = refs[-1][1]
            table = self.core(
                year,
                quarter,
                fs_type,
                self.segment(ticker)
            ).drop(['full_name','exchange'],level=1,axis=1)
            dict_idx[fs_type] = table.columns.get_level_values(1).tolist()
        fs = pd.concat([
            self.core(
                ref[0],ref[1],ref[2],segment
            ).loc[pd.IndexSlice[:,:,ticker]].drop(
                ['full_name','exchange'],level=1,axis=1
            ).T.set_index(pd.MultiIndex.from_product(
                [[segment],[ref[2]],dict_idx[ref[2]]]
            )) for ref in refs
        ],axis=1)
        fs = fs.groupby(level=[0,1,2],axis=1,sort=False).sum(min_count=1)

        print('Finished!')
        return fs

    def all(
        self,
        segment: str,
        exchange: str = 'all'
    ) \
            -> pd.DataFrame:

        """
        This method returns all financial statements of all companies

        :param segment: allow values in fa.segments
        :param exchange: allow values in ['HOSE', 'HNX', 'UPCOM']
        :type segment: str
        :type exchange: str

        :return: pandas.DataFrame
        """

        frames = []
        for period in self.periods:
            for fs_type in self.fs_types:
                year = int(period[:4])
                quarter = int(period[-1])
                frame = self.core(
                    year,quarter,fs_type,segment,exchange
                ).drop(labels=['full_name','exchange'],level=1,axis=1)
                frames += [frame]
        df = pd.concat(frames,axis=1,join='outer')
        df = df.groupby(level=[0,1],axis=1,sort=False).sum(min_count=1)
        df.columns.names = ['fs_type','item']

        return df

    def exchanges(
        self
    ) \
            -> pd.DataFrame:

        """
        This method returns stock exchanges of all tickers

        :param: None

        :return: pandas.DataFrame
        """

        frames = []
        for segment in self.segments:
            year = int(self.latest_period[:4])
            quarter = int(self.latest_period[-1])
            frame = self.core(year,quarter,'is',segment)
            frame = frame.loc[:,pd.IndexSlice[:,'exchange']]
            frame = frame.droplevel(level=['year','quarter'])
            frame.columns = ['exchange']
            frames += [frame]
        table = pd.concat(frames)

        return table

    def exchange(
        self,
        ticker: str
    ) \
            -> str:

        """
        This method returns stock exchange of given stock

        :param ticker: stock's ticker
        :type ticker: str

        :return: str
        """

        exchange = self.exchanges().loc[ticker,'exchange']

        return exchange

    def tickers(
        self,
        segment: str = 'all',
        exchange: str = 'all'
    ) \
            -> list:

        """
        This method returns all tickers of given segment or exchange

        :param segment: allow values in fa.segments or 'all'
        :param exchange: allow values in ['HOSE', 'HNX', 'UPCOM'] or 'all'
        :type segment: str
        :type exchange: str

        :return: list
        """

        ticker_list = []
        year = int(self.latest_period[:4])
        quarter = int(self.latest_period[-1])
        if segment in self.segments:
            ticker_list = self.core(
                year,
                quarter,
                'is',
                segment,
                exchange
            ).index.get_level_values(level=2).tolist()
        elif segment=='all':
            ticker_list = []
            for segment in self.segments:
                ticker_list += self.core(
                    year,
                    quarter,
                    'is',
                    segment,
                    exchange
                ).index.get_level_values(level=2).tolist()

        return ticker_list

    def items(
        self,
        segment: str,
        fs_type: str
    ) \
            -> list:

        """
        This method returns all financial items of given financial statement of given segment

        :param segment: allow values in self.segments()
        :param fs_type: allow valies in self.fs_types()
        :type segment: str
        :type fs_type: str

        :return: list
        """

        year = int(self.latest_period[:4])
        quarter = int(self.latest_period[-1])
        items = self.core(
            year,
            quarter,
            fs_type,
            segment
        ).columns.get_level_values(level=1).tolist()
        items.remove('full_name')
        items.remove('exchange')

        return items

    def classification(
        self,
        standard: str
    ) \
            -> pd.DataFrame:

        """
        This method returns industry classification instructed by a given standard of all stocks

        :param standard: allow values in fa.standards
        :type standard: str

        :return: pandas.DataFrame
        """

        st_dict = dict()
        folder = 'industry'

        for st in self.standards:
            # create Workbook object, select active Worksheet
            raw_bloomberg = openpyxl.load_workbook(
                os.path.join(self.database_path,folder,st+'.xlsx')
            ).active
            # delete Bloomberg Sign
            raw_bloomberg.delete_rows(idx=raw_bloomberg.max_row)
            # delete headers
            raw_bloomberg.delete_rows(idx=0,amount=2)
            clean_data = pd.DataFrame(raw_bloomberg.values)
            clean_data.iloc[1] = clean_data.iloc[0]
            clean_data.drop(index=0,inplace=True)
            # set index and columns
            clean_data.index = clean_data.iloc[:,0]
            clean_data.index.rename('ticker')
            clean_data.columns = clean_data.iloc[0,:]
            clean_data.columns.rename('level')
            # remore unwanted columns, rows
            clean_data.drop(columns=['Ticker','Name'],index=['Ticker'],inplace=True)
            # rename columns
            clean_data.columns = pd.Index(
                data=[clean_data.columns[i].split()[0].lower().split(' ',maxsplit=1)[0]+'_l'+str(i+1)
                      for i in range(clean_data.shape[1])]
            )
            # rename index
            clean_data.index = pd.Index(
                data=[clean_data.index[i].split()[0] for i in range(clean_data.shape[0])]
            )
            st_dict[st] = clean_data

        return st_dict[standard]

    def levels(
        self,
        standard: str
    ) \
            -> list:

        """
        This method returns all levels of given industry classification standard

        :param standard: allow values in fa.standards

        :return: list
        """

        levels = self.classification(standard).columns.tolist()
        return levels

    def industries(
        self,
        standard: str,
        level: int
    ) \
            -> list:

        """
        This method returns all industry names of given level of given classification standard

        :param standard: allow values in request_industry_standard()
        :param level: allow values in request_industry_level() (number only)

        :return: list
        """

        industries = self.classification(standard)
        industries = industries[f'{standard}_l{str(level)}'].drop_duplicates().tolist()

        return industries

    @staticmethod
    def peers(
        ticker: str,
        standard: str,
        level: int
    ) \
            -> list:

        df = fa.classification(standard)
        col = f'{standard}_l{level}'
        industry = df.loc[ticker,col]
        peers = df.loc[df[col]==industry].index.to_list()

        return peers

    def ownerships(self) -> pd.DataFrame:

        """
        This method returns ownership structure of all tickers

        :param: None

        :return: pandas.DataFrame
        """

        folder = 'ownership'
        file = [f for f in listdir(join(self.database_path,folder))
                if isfile(join(self.database_path,folder,f))][-1]

        excel = Dispatch("Excel.Application")
        excel.Visible = False
        for wb in [wb for wb in excel.Workbooks]:
            wb.Close(True)

        # create Workbook object, select active Worksheet
        raw_fiinpro = openpyxl.load_workbook(
            os.path.join(self.database_path,folder,file)
        ).active
        # delete StoxPlux Sign
        raw_fiinpro.delete_rows(idx=raw_fiinpro.max_row-21,amount=1000)
        # delete header rows
        raw_fiinpro.delete_rows(idx=0,amount=7)
        raw_fiinpro.delete_rows(idx=2,amount=1)
        # import to DataFrame
        clean_data = pd.DataFrame(raw_fiinpro.values)
        # drop unwanted index and columns
        clean_data.drop(columns=[0],inplace=True)
        clean_data.drop(index=[0],inplace=True)
        # set ticker as index
        clean_data.index = clean_data.iloc[:,0]
        clean_data.drop(clean_data.columns[0],axis=1,inplace=True)

        # rename columns
        columns = [
            'full_name',
            'exchange',
            'state_share',
            'state_percent',
            'frgn_share',
            'frgn_percent',
            'other_share',
            'other_percent',
            'frgn_maxpercent',
            'frgn_maxshare',
            'frgn_remainshare'
        ]
        clean_data.columns = columns

        return clean_data

    def marketcaps(
        self
    ) \
            -> pd.DataFrame:

        """
        This method returns market capitalization of all stocks

        :param: None

        :return: pandas.DataFrame
        """

        folder = 'marketcap'
        file = [f for f in listdir(join(self.database_path,folder))
                if isfile(join(self.database_path,folder,f))][-1]

        excel = Dispatch("Excel.Application")
        excel.Visible = False
        for wb in [wb for wb in excel.Workbooks]:
            wb.Close(True)
        # create Workbook object, select active Worksheet
        raw_fiinpro = openpyxl.load_workbook(
            os.path.join(self.database_path,folder,file)).active
        # delete StoxPlux Sign
        raw_fiinpro.delete_rows(idx=raw_fiinpro.max_row-21,amount=1000)
        # delete header rows
        raw_fiinpro.delete_rows(idx=0,amount=7)
        raw_fiinpro.delete_rows(idx=2,amount=1)
        raw_fiinpro.delete_cols(idx=6,amount=1000)
        # import to DataFrame
        clean_data = pd.DataFrame(raw_fiinpro.values)
        # drop unwanted index and columns
        clean_data.drop(columns=[0,2,3],inplace=True)
        clean_data.drop(index=[0],inplace=True)
        # set ticker as index
        clean_data.index = clean_data.iloc[:,0]
        clean_data.drop(clean_data.columns[0],axis=1,inplace=True)
        # rename columns
        clean_data.columns = ['marketcap']

        # change unit from VND billion
        clean_data['marketcap'] *= 1e9

        return clean_data

    def marketcap(
        self,
        ticker: str
    ) \
            -> str:

        """
        This method returns market capitalization of given stock

        :param ticker: stock's ticker
        :type ticker: str

        :return: str
        """

        table = self.marketcaps()
        marketcap = table.loc[ticker,'marketcap']

        return marketcap


fa = fa()


###############################################################################


class ta:
    address_hist = 'https://api.phs.vn/market/utilities.svc/GetShareIntraday'
    address_intra = 'https://api.phs.vn/market/Utilities.svc/GetRealShareIntraday'

    def __init__(self):
        pass

    def hist(
        self,
        ticker: str,
        fromdate=None,
        todate=None,
    ) \
            -> pd.DataFrame:

        """
        This method returns historical trading data of given ticker

        :param ticker: allow values in fa.tickers()
        :param fromdate: [optional] allow any date with format: 'yyyy-mm-dd'
        :param todate: [optional] allow any date with format: 'yyyy-mm-dd'

        :return: pandas.DataFrame
        """

        if fromdate is None:
            start = '2018-01-01'
        else:
            if dt.datetime.strptime(fromdate,'%Y-%m-%d')<dt.datetime(year=2018,month=1,day=1):
                raise Exception('Only data since 2018-01-01 is reliable')
            else:
                start = fromdate

        if todate is None or todate=='now':
            end = dt.datetime.now().strftime('%Y-%m-%d')
        else:
            end = todate

        r = requests.post(
            self.address_hist,
            data=json.dumps({
                'symbol':ticker,
                'fromdate':start,
                'todate':end,
            }),
            headers={'content-type':'application/json'}
        )
        history = pd.DataFrame(json.loads(r.json()['d']))
        history.rename(
            columns={
                'symbol':'ticker',
                'open_price':'open',
                'prior_price':'ref',
                'close_price':'close'},
            inplace=True)

        def addzero(int_: str):
            if len(int_)==1:
                int_ = '0'+int_
            return int_

        def converter(date_str:str):
            month,day,year = date_str.split('/')
            day = addzero(day)
            month = addzero(month)
            return f'{year}-{month}-{day}'

        history['trading_date'] = history['trading_date'].str.split().str.get(0).map(converter)
        history = history.sort_values('trading_date').reset_index(drop=True)

        history.loc[history['open']==0,'open'] = history['ref']
        history.loc[history['close']==0,'close'] = history['ref']
        history.loc[history['high']==0,'high'] = history['ref']
        history.loc[history['low']==0,'low'] = history['ref']

        return history

    def intra(
        self,
        ticker=str,
        fromdate=None,
        todate=None
    ) \
            -> pd.DataFrame:

        """
        This method returns intraday trading data of given ticker

        :param ticker: allow values in fa.tickers()
        :param fromdate: [optional] allow any date with format: 'yyyy-mm-dd' or 'yyyy/mm/dd'
        :param todate: [optional] allow any date with format: 'yyyy-mm-dd' or 'yyyy/mm/dd'

        :return: pandas.DataFrame
        :raise Exception: Can't extract more than 60 days
        """

        pd.options.mode.chained_assignment = None
        if fromdate is not None and todate is not None:
            if dt.datetime.strptime(todate,'%Y-%m-%d') \
                    -dt.datetime.strptime(fromdate,'%Y-%m-%d')>dt.timedelta(days=60):
                raise Exception('Can\'t extract more than 60 days')
            else:
                try:
                    r = requests.post(self.address_intra,
                                      data=json.dumps(
                                          {'symbol':ticker,
                                           'fromdate':fromdate,
                                           'todate':todate}
                                      ),
                                      headers={
                                          'content-type':'application/json'})
                    # history = pd.DataFrame(json.loads(r.json()['d'])['histories'])
                    intraday = pd.DataFrame(
                        json.loads(r.json()['d'])['intradays'])
                except KeyError:
                    raise Exception(
                        'Date Format Required: yyyy-mm-dd, yyyy/mm/dd')
        else:
            try:
                fromdate = (dt.datetime.now()-dt.timedelta(days=60)).strftime("%Y-%m-%d")
                todate = (dt.datetime.now()-dt.timedelta(days=0)).strftime("%Y-%m-%d")
                r = requests.post(self.address_intra,
                                  data=json.dumps(
                                      {'symbol':ticker,
                                       'fromdate':fromdate,
                                       'todate':todate}
                                  ),
                                  headers={'content-type':'application/json'})
                # history = pd.DataFrame(json.loads(r.json()['d'])['histories'])
                intraday = pd.DataFrame(json.loads(r.json()['d'])['intradays'])
            except KeyError:
                raise Exception('Date Format Required: yyyy-mm-dd, yyyy/mm/dd')

        def datemod(date=str):
            def addzero(int_=str):
                if len(int_)==1:
                    int_ = '0'+int_
                else:
                    pass
                return int_

            day = addzero(date.split('/')[1])
            month = addzero(date.split('/')[0])
            date = date.split('/')[2][:4]+'-' \
                   +month+'-' \
                   +day+' ' \
                   +date.split(" ",maxsplit=1)[1]
            return date

        for i in range(intraday.shape[0]):
            intraday['trading_time'].iloc[i] \
                = datemod(intraday['trading_time'].iloc[i])

        return intraday

    def prices(self,segment: str = 'all',exchange: str = 'all') \
            -> pd.DataFrame:

        """
        This method returns stock price of all tickers in all periods

        :param segment: allow values in fa.segments
        :param exchange: allow values in ['HOSE', 'HNX', 'UPCOM'] or 'all'
        :type segment: str
        :type exchange: str

        :return: pandas.DataFrame
        """

        periods = fa.periods
        tickers = fa.tickers(segment,exchange)

        prices = pd.DataFrame(data=np.zeros((len(tickers),len(periods))),
                              columns=[periods[i]
                                       for i in range(len(periods))],
                              index=[tickers[j]
                                     for j in range(len(tickers))])

        date_input = {'q1':'03-31','q2':'06-30',
                      'q3':'09-30','q4':'12-28'}
        # Note: because of the 'Bday' function structure,
        #       December must be adjusted to 28 instead of 31

        price_data = dict()
        for ticker in tickers:
            try:
                price_data[ticker] = self.hist(ticker)
            except (KeyError,IndexError):
                continue

        def Bday(date=str):
            date_ = dt.datetime(year=int(date.split('-')[0]),
                                month=int(date.split('-')[1]),
                                day=int(date.split('-')[2]))
            one_day = dt.timedelta(days=1)
            while date_.weekday() in holidays.WEEKEND or date_ in holidays.VN():
                date_ = date_+one_day

            return date_.strftime(format='%Y-%m-%d')

        for period in periods:
            for ticker in tickers:
                try:
                    date = Bday(period[:4]+'-'+date_input[period[-2:]])
                except (KeyError,IndexError):
                    continue
                try:
                    t = price_data[ticker]['trading_date']==date
                    try:
                        prices.loc[ticker,period] \
                            = price_data[ticker].loc[t]['close'].iloc[0]
                    except (IndexError,KeyError,RuntimeWarning):
                        prices.loc[ticker,period] = np.nan
                except (IndexError,KeyError,RuntimeWarning):
                    prices.loc[ticker,period] = np.nan

        prices.replace([np.inf,-np.inf,-1,1],np.nan,inplace=True)
        prices.to_csv(join(database_path,'price','prices.csv'))

        return prices

    def returns(self,segment: str = 'all',exchange: str = 'all') \
            -> pd.DataFrame:

        """
        This method returns stock returns of all tickers of given segment
        in all periods

        :param segment: allow values in fa.segments
        :param exchange: allow values in ['HOSE', 'HNX', 'UPCOM'] or 'all'
        :type segment: str
        :type exchange: str

        :return: pandas.DataFrame
        """

        periods = fa.periods
        tickers = fa.tickers(segment,exchange)

        returns = pd.DataFrame(data=np.zeros((len(tickers),len(periods))),
                               columns=[periods[i]
                                        for i in range(len(periods))],
                               index=[tickers[j]
                                      for j in range(len(tickers))])

        date_input = {'q1':['01-01','03-31'],'q2':['04-01','06-30'],
                      'q3':['07-01','09-30'],'q4':['10-01','12-28']}
        # Note: because of the 'Bday' function structure,
        #       December must be adjusted to 28 instead of 31

        price_data = dict()
        for ticker in tickers:
            try:
                price_data[ticker] = self.hist(ticker)
            except (KeyError,IndexError,ValueError):
                continue

        def Bday(date=str):
            date_ = dt.datetime(year=int(date.split('-')[0]),
                                month=int(date.split('-')[1]),
                                day=int(date.split('-')[2]))
            one_day = dt.timedelta(days=1)
            while date_.weekday() in holidays.WEEKEND or date_ in holidays.VN():
                date_ = date_+one_day

            return date_.strftime(format='%Y-%m-%d')

        for period in periods:
            for ticker in tickers:
                try:
                    fromdate \
                        = Bday(period[:4]+'-'+date_input[period[-2:]][0])
                    todate \
                        = Bday(period[:4]+'-'+date_input[period[-2:]][1])
                except (KeyError,IndexError):
                    continue
                try:
                    f = price_data[ticker]['trading_date']==fromdate
                    t = price_data[ticker]['trading_date']==todate
                    try:
                        open_price \
                            = price_data[ticker].loc[f]['close'].iloc[0]
                        close_price \
                            = price_data[ticker].loc[t]['close'].iloc[0]
                        returns.loc[ticker,period] \
                            = close_price/open_price-1
                    except (IndexError,KeyError,RuntimeWarning):
                        returns.loc[ticker,period] = np.nan
                except (IndexError,KeyError,RuntimeWarning):
                    returns.loc[ticker,period] = np.nan

        returns.replace([np.inf,-np.inf,-1,1],np.nan,inplace=True)
        returns.to_csv(join(database_path,'returns','returns.csv'))

        return returns

    @staticmethod
    def crash(benchmark=-0.5,segment: str = 'all',exchange: str = 'all') \
            -> dict:

        """
        This method returns all tickers whose stock return
        lower than 'benchmark' in a given period

        :param benchmark: negative number in [-1,0]
        :param segment: allow values in fa.segments
        :param exchange: allow values in ['HOSE', 'HNX', 'UPCOM'] or 'all'

        :return: dict[period]
        """

        folder = 'price'
        file = 'prhighlow.csv'
        table = pd.read_csv(join(database_path,folder,file),
                            index_col=[0])
        tickers = fa.tickers(segment,exchange)

        crash_dict = dict()
        for period in fa.periods:
            crash_dict[period] = []
            for ticker in tickers:
                t = table.loc[ticker,period]
                if isinstance(t,str):
                    string = t.replace('(','').replace(')','')
                    string = string.split(',')
                    low_return = float(string[3])
                    if low_return<=benchmark:
                        crash_dict[period] += [ticker]

        return crash_dict

    def prhighlow(self,ticker: str,fquarters: int = 1) \
            -> Union[dict,pd.DataFrame]:

        """
        This function returns the lowest price/return in all periods

        :param ticker: stock's ticker or 'all'
        :param fquarters: number of quarters needed after the latest reporting period

        :return: a dictionary of 'high' and 'low' (if ticker != 'all')
         or a DataFrame of (start_price, low_price, high_price, low_return, high_return)
        """

        if ticker!='all':
            periods = fa.periods+[period_cal(fa.periods[-1],0,q+1) for q in range(fquarters)]

            full_price = self.hist(ticker)[['trading_date','close']]
            full_price.set_index('trading_date',inplace=True)
            full_price = full_price['close']

            d = dict()
            for type_ in ['high','low']:
                d[type_] = pd.DataFrame(index=periods,
                                        columns=['start_price',
                                                 f'{type_}_price',
                                                 f'{type_}_return'])
                for period in periods:
                    try:
                        start_date,end_date = seopdate(period)
                        price_section = full_price.loc[start_date:end_date]
                        d[type_].loc[period,'start_price'] \
                            = price_section.iloc[0]
                        if type_=='high':
                            d[type_].loc[period,f'{type_}_price'] \
                                = price_section.max()
                        else:
                            d[type_].loc[period,f'{type_}_price'] \
                                = price_section.min()
                    except IndexError:
                        pass
                d[type_][f'{type_}_return'] \
                    = d[type_][f'{type_}_price']/d[type_]['start_price']-1

        elif ticker=='all':
            periods \
                = fa.periods+[period_cal(fa.periods[-1],0,q+1) for q in
                              range(fquarters)]
            d = pd.DataFrame(index=fa.tickers(),columns=periods)
            for ticker_ in d.index:
                try:
                    full_price = self.hist(ticker_)[['trading_date','close']]
                    full_price.set_index('trading_date',inplace=True)
                    full_price = full_price['close']
                    for period_ in d.columns:
                        try:
                            start_date,end_date = seopdate(period_)
                            price_section = full_price.loc[start_date:end_date]
                            s = price_section.iloc[0]
                            lp = price_section.min()
                            hp = price_section.max()
                            lr = lp/s-1
                            hr = hp/s-1
                            d.loc[ticker_,period_] = (s,lp,hp,lr,hr)
                        except IndexError:
                            # handle not applicable date (IndexError)
                            pass
                except ValueError:
                    # handle ticker that not available (ValueError)
                    pass

            export_name = 'prhighlow.csv'
            d.to_csv(join(dirname(dirname((realpath(__file__)))),
                          'database','price',export_name))
        else:
            d = 'Error!'

        return d

    def now(self,ticker: str) -> float:

        """
        This method returns the latest matched price of given ticker

        :param ticker: allow values in fa.tickers()
        :return: float
        """

        today = dt.datetime.now().strftime("%Y-%m-%d")
        before = bdate(today,-1)
        now = self.intra(ticker,before,today)['price'].iloc[-1]

        return now

    def last(self,ticker: str) -> float:

        """
        This method returns the latest close price of given ticker

        :param ticker: allow values in fa.tickers()
        :type ticker: str
        :return: float
        """

        close = self.hist(ticker)['close'].iloc[-1]
        return close


ta = ta()

###############################################################################
