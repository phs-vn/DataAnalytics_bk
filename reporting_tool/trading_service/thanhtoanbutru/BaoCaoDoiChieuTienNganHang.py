from reporting_tool.trading_service.thanhtoanbutru import *
import reporting_tool
# DONE
def _get_outlook_files(
        run_time=None,
):
    info = get_info('daily',run_time)
    end_date = info['end_date']
    period = info['period']

    outlook = Dispatch('outlook.application').GetNamespace("MAPI")
    for account in outlook.Accounts:
        print(f"Account {account.DeliveryStore.DisplayName} is being logged")  # print which account is being logged

    inbox = outlook.Folders.Item(1).Folders['Inbox']
    messages = inbox.Items
    # Lọc ngày
    run_date = dt.datetime.strptime(end_date,'%Y/%m/%d').date()
    def catch(m):
        try:
            result = m.ReceivedTime.date() == run_date - dt.timedelta(days=1)
        except (Exception,):
            result = True
        return result
    messages = [mes for mes in messages if catch(mes)]

    # Tạo folder
    if not os.path.isdir(join(dept_folder,'FileFromBanks',period)):
        os.mkdir(join(dept_folder,'FileFromBanks',period))

    def f(x,bank): # Hàm lọc mail
        mapper = {
            'EIB':'@eximbank',
            'OCB':'@ocb',
        }
        try:
            sender = x.SenderEmailAddress
        except (Exception,):
            try:
                sender = x.Sender.GetExchangeUser().PrimarySmtpAddress
            except (Exception,):
                sender = ''
        return re.search(f'.*{mapper[bank]}.*',sender)

    for bank in ['EIB','OCB']:
        # Tạo folder
        if not os.path.isdir(join(dept_folder,'FileFromBanks',period,bank)):
            os.mkdir(join(dept_folder,'FileFromBanks',period,bank))

        emails = [mes for mes in messages if f(mes,bank) is not None]
        received_times = [email.ReceivedTime.time() for email in emails]
        if len(received_times)==0:
            continue
        else:
            # Delete existing files (if any)
            files = os.listdir(join(dept_folder,'FileFromBanks',period,bank))
            for file in files:
                os.remove(join(dept_folder,'FileFromBanks',period,bank,file))
            # Get file
            max_time = max(received_times)
            idx = received_times.index(max_time)
            email = emails[idx]
            for attachment in email.Attachments:
                check_file_extension = attachment.FileName.endswith('.xlsx') or attachment.FileName.endswith('.xls')
                if check_file_extension:
                    attachment.SaveASFile(os.path.join(dept_folder,'FileFromBanks',period,bank,attachment.FileName))

    # Return đường dẫn
    result = {}
    for bank in ['EIB','OCB']:
        files = os.listdir(join(dept_folder,'FileFromBanks',period,bank))
        if len(files)==0: # chỉ xảy ra khi chạy file này trước khi có email
            raise RuntimeError(f'No email from {bank} today')
        else:
            file = files[0] # should only have 1 file
            result[bank] = join(dept_folder,'FileFromBanks',period,bank,file)

    return result

# run daily
def run(
        run_time=None,
):
    start = time.time()
    info = get_info('daily',run_time)
    start_date = info['start_date']
    end_date = info['end_date']
    period = info['period']
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):  # dept_folder from import
        os.mkdir(join(dept_folder,folder_name,period))

    file_path = _get_outlook_files(run_time)

    eib_sent_data = pd.read_excel(
        file_path['EIB'],
        skiprows=5,
        usecols=['FORACID','ACCOUNT BALANCE'],
        dtype={'FORACID': object},
    )
    eib_sent_data.rename(
        {
            'FORACID': 'bank_account',
            'ACCOUNT BALANCE': 'balance',
        },
        axis=1,
        inplace=True
    )
    eib_sent_data['bank_code'] = 'EIB'
    eib_sent_data.set_index(['bank_code','bank_account'],inplace=True)
    eib_sent_data.index.rename(['bank','bank_account'],inplace=True)

    ocb_sent_data = pd.read_excel(
        file_path['OCB'],
        usecols=['TÀI KHOẢN','SỐ DƯ'],
        dtype={'TÀI KHOẢN': object},
    )
    ocb_sent_data.rename(
        {
            'TÀI KHOẢN': 'bank_account',
            'SỐ DƯ': 'balance',
        },
        axis=1,
        inplace=True
    )
    ocb_sent_data['bank_code'] = 'OCB'
    ocb_sent_data.set_index(['bank_code','bank_account'],inplace=True)
    ocb_sent_data.index.rename(['bank','bank_account'],inplace=True)

    bank_sent_balance = pd.concat([eib_sent_data,ocb_sent_data])

    increase_money = pd.read_sql(
        f"""
        SELECT
        [cashflow_bank].[bank],
        [cashflow_bank].[bank_account],
        [cashflow_bank].[outflow_amount]
        FROM [cashflow_bank]
        WHERE [cashflow_bank].[bank] IN ('EIB','OCB')
        AND [cashflow_bank].[date] BETWEEN '{start_date}' AND '{end_date}'
        AND [cashflow_bank].[outflow_amount] > 0
        """,
        connect_DWH_CoSo,
        index_col=['bank','bank_account'],
    )
    increase_money_eib = increase_money.loc[increase_money.index.get_level_values(0)=='EIB'].copy()
    increase_money_ocb = increase_money.loc[increase_money.index.get_level_values(0)=='OCB'].copy()

    decrease_money = pd.read_sql(
        f"""
        SELECT 
        [cashflow_bank].[bank],
        [cashflow_bank].[bank_account], 
        [cashflow_bank].[inflow_amount]
        FROM [cashflow_bank]
        WHERE [cashflow_bank].[bank] IN ('EIB','OCB')
        AND [cashflow_bank].[date] BETWEEN '{start_date}' AND '{end_date}'
        AND [cashflow_bank].[inflow_amount] > 0
        """,
        connect_DWH_CoSo,
        index_col=['bank','bank_account'],
    )
    decrease_money_eib = decrease_money.loc[decrease_money.index.get_level_values(0)=='EIB'].copy()
    decrease_money_ocb = decrease_money.loc[decrease_money.index.get_level_values(0)=='OCB'].copy()

    in_money = pd.read_sql(
        f"""
        SELECT
        [money_in_out_transfer].[bank],
        [money_in_out_transfer].[bank_account],
        [money_in_out_transfer].[amount]
        FROM [money_in_out_transfer]
        WHERE [money_in_out_transfer].[transaction_id] = '6692'
        AND [money_in_out_transfer].[date] BETWEEN '{start_date}' AND '{end_date}'
        AND [money_in_out_transfer].[bank] IN ('EIB','OCB')
        """,
        connect_DWH_CoSo,
        index_col=['bank','bank_account'],
    )
    out_money = pd.read_sql(
        f"""
        SELECT
        [money_in_out_transfer].[bank],
        [money_in_out_transfer].[bank_account],
        [money_in_out_transfer].[amount]
        FROM [money_in_out_transfer]
        WHERE [money_in_out_transfer].[transaction_id] = '6693'
        AND [money_in_out_transfer].[date] BETWEEN '{start_date}' AND '{end_date}'
        AND [money_in_out_transfer].[bank] IN ('EIB','OCB')
        """,
        connect_DWH_CoSo,
        index_col=['bank','bank_account'],
    )
    imported_balance = pd.read_sql(
        f"""
        SELECT 
        [imported_bank_balance].[bank_code],
        [imported_bank_balance].[bank_account],
        [imported_bank_balance].[account_name],
        [imported_bank_balance].[balance]
        FROM [imported_bank_balance]
        WHERE [imported_bank_balance].[date] = '{bdate(start_date,-1)}'
        """,
        connect_DWH_CoSo,
        index_col=['bank_code','bank_account'],
    )
    imported_balance.index.rename(['bank','bank_account'],inplace=True)
    bank_account_full = set(eib_sent_data.index) | \
                        set(ocb_sent_data.index) | \
                        set(increase_money_eib.index) | \
                        set(increase_money_ocb.index) | \
                        set(decrease_money_eib.index) | \
                        set(decrease_money_ocb.index) | \
                        set(in_money.index) | \
                        set(out_money.index) | \
                        set(imported_balance.index)
    bank_sent_balance = bank_sent_balance.reindex(bank_account_full).fillna(0)
    imported_balance = imported_balance.reindex(bank_account_full)
    imported_balance['account_name'].fillna('',inplace=True)
    imported_balance['balance'].fillna(0,inplace=True)
    increase_money = increase_money.reindex(bank_account_full).fillna(0)
    decrease_money = decrease_money.reindex(bank_account_full).fillna(0)
    in_money = in_money.reindex(bank_account_full).fillna(0)
    out_money = out_money.reindex(bank_account_full).fillna(0)

    output = imported_balance.copy()
    output.sort_index(inplace=True)
    output.rename({'balance':'opening_balance'},axis=1,inplace=True)
    output['increase_money'] = increase_money['outflow_amount']
    output['decrease_money'] = decrease_money['inflow_amount']
    output['in_money'] = in_money['amount']
    output['out_money'] = out_money['amount']
    output['expected_balance'] = output['opening_balance'] + output['increase_money'] - output['decrease_money'] + output['in_money'] - output['out_money']
    output['bank_sent_balance'] = bank_sent_balance['balance']
    output['diff'] = output['expected_balance'] - output['bank_sent_balance']
    output = output.loc[output['diff']!=0]
    def name_converter(x):
        try:
            result = unidecode.unidecode(x)
        except (Exception,):
            result = ''
        return result
    output['account_name'] = output['account_name'].map(name_converter)
    output.reset_index(inplace=True)

    file_name = f'Báo cáo đối chiếu số dư Ngân Hàng trước khi Import {period}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    ## Write sheet ChenhLech
    header_format = workbook.add_format(
        {
            'bold': True,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 10,
            'text_wrap': True,
        }
    )
    text_cell_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 10,
            'text_wrap': True,
        }
    )
    number_cell_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 10,
            'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'text_wrap': True,
        }
    )
    diff_cell_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 10,
            'font_color': 'red',
            'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'text_wrap': True,
        }
    )
    worksheet = workbook.add_worksheet('Chênh Lệch')
    worksheet.hide_gridlines(option=2)
    worksheet.set_column('A:A',11)
    worksheet.set_column('B:B',18)
    worksheet.set_column('C:C',35)
    worksheet.set_column('D:K',15)
    worksheet.write_row(
        'A1',
        [
            'NGÂN HÀNG',
            'TÀI KHOẢN',
            'TÊN KHÁCH HÀNG',
            'SỐ DƯ ĐẦU KỲ',
            'TĂNG TIỀN',
            'GIẢM TIỀN',
            'NHẬP TIỀN',
            'RÚT TIỀN',
            'DỰ KIẾN CUỐI NGÀY',
            'SỐ NGÂN HÀNG GỬI',
            'LỆCH'
        ],
        header_format
    )
    worksheet.write_column('A2',output['bank'],text_cell_format)
    worksheet.write_column('B2',output['bank_account'],text_cell_format)
    worksheet.write_column('C2',output['account_name'],text_cell_format)
    worksheet.write_column('D2',output['opening_balance'],number_cell_format)
    worksheet.write_column('E2',output['increase_money'],number_cell_format)
    worksheet.write_column('F2',output['decrease_money'],number_cell_format)
    worksheet.write_column('G2',output['in_money'],number_cell_format)
    worksheet.write_column('H2',output['out_money'],number_cell_format)
    worksheet.write_column('I2',output['expected_balance'],number_cell_format)
    worksheet.write_column('J2',output['bank_sent_balance'],number_cell_format)
    worksheet.write_column('K2',output['diff'],diff_cell_format)

    writer.close()

    ###########################################################################
    ###########################################################################
    ###########################################################################

    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')