from reporting_tool.trading_service.giaodichluuky import *

def run(
        periodicity:str,
        run_time=None,
):
    start = time.time()
    info = get_info(periodicity,run_time)
    start_date = info['start_date']
    end_date = info['end_date']
    period = info['period']
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder, folder_name, period)): # dept_folder from import
        os.mkdir(join(dept_folder, folder_name, period))

    # Phi chuyen khoan sang cong ty chung khoan khac xuat theo ngay mac dinh
    transfer_fee_CTCK = pd.read_sql(
        "SELECT account_code, date, [transaction], volume FROM deposit_withdraw_stock "
        f"WHERE date BETWEEN '{start_date}' AND '{end_date}' "
        f"AND creator NOT IN ('User Online trading') "
        f"AND [transaction] LIKE 'Chuy_n CK (____)'",
        connect_DWH_CoSo,
        index_col='account_code',
    )
    # Phi chuyen khoan thanh toan bu tru xuat theo ngay duoc dieu chinh
    adj_start_date = bdate(start_date,-1)
    adj_end_date = bdate(end_date,-2)
    transfer_fee_TTBT = pd.read_sql(
        "SELECT sub_account, date, ticker, volume FROM trading_record "
        f"WHERE date BETWEEN '{adj_start_date}' AND '{adj_end_date}' "
        f"AND type_of_order = 'S' "
        f"AND depository_place LIKE 'T_i PHS'",
        connect_DWH_CoSo
    )
    account = pd.read_sql(
        "SELECT sub_account, account_code FROM sub_account",
        connect_DWH_CoSo,
        index_col='sub_account',
    ).squeeze()
    broker = pd.read_sql(
        "SELECT account_code, broker_id FROM account",
        connect_DWH_CoSo,
        index_col='account_code',
    ).squeeze()
    branch_id = pd.read_sql(
        "SELECT broker_id, branch_id FROM broker",
        connect_DWH_CoSo,
        index_col='broker_id',
    ).squeeze()
    branch_name = pd.read_sql(
        "SELECT branch_id, branch_name FROM branch;",
        connect_DWH_CoSo,
        index_col='branch_id',
    ).squeeze()

    # calculate transfer fee on TTBT
    case = transfer_fee_CTCK['date'] > dt.datetime(2020,3,19)
    transfer_fee_CTCK.loc[case,'fee'] = transfer_fee_CTCK['volume'] *  0.3
    transfer_fee_CTCK.loc[~case,'fee'] = transfer_fee_CTCK['volume'] * 0.5
    transfer_fee_CTCK['fee'] = transfer_fee_CTCK['fee'].apply(min,args=(300000,))
    transfer_fee_CTCK['branch_id'] = transfer_fee_CTCK.index.map(broker).map(branch_id)
    transfer_fee_CTCK.reset_index(drop=True,inplace=True)
    transfer_fee_CTCK = transfer_fee_CTCK.groupby('branch_id',as_index=True).sum()
    # calculate transfer fee on TTBT
    sum_volume = transfer_fee_TTBT.groupby(['date','ticker']).sum().squeeze().sort_index()
    transfer_fee_TTBT.set_index(['date','ticker'],inplace=True)
    transfer_fee_TTBT['sum_volume'] = sum_volume
    transfer_fee_TTBT.reset_index(inplace=True)
    transfer_fee_TTBT['vol_percent'] = transfer_fee_TTBT['volume']/transfer_fee_TTBT['sum_volume']
    case = transfer_fee_TTBT['volume'] > 1000000
    transfer_fee_TTBT.loc[case,'charged_volume'] = transfer_fee_TTBT['vol_percent'] * 1000000
    transfer_fee_TTBT.loc[~case,'charged_volume'] = transfer_fee_TTBT['volume']
    case = transfer_fee_TTBT['date'] > dt.datetime(2020,3,16)
    transfer_fee_TTBT.loc[case,'fee'] = transfer_fee_TTBT['charged_volume']*0.3
    transfer_fee_TTBT.loc[~case,'fee'] = transfer_fee_TTBT['charged_volume']*0.5
    transfer_fee_TTBT = transfer_fee_TTBT[['sub_account','volume','fee']]
    transfer_fee_TTBT.set_index('sub_account',inplace=True)
    transfer_fee_TTBT['branch_id'] = account.map(broker).map(branch_id)
    transfer_fee_TTBT.reset_index(drop=True,inplace=True)
    transfer_fee_TTBT = transfer_fee_TTBT.groupby('branch_id',as_index=True).sum()

    transfer_fee = pd.DataFrame(
        columns=[
            'Tên Chi Nhánh',
            'Phí Chuyển Khoản',
        ],
        index=branch_name.index
    )
    transfer_fee['Tên Chi Nhánh'] = branch_name
    transfer_fee['Phí Chuyển Khoản'] = transfer_fee_CTCK['fee'] + transfer_fee_TTBT['fee']
    transfer_fee['Phí Chuyển Khoản'].fillna(0, inplace=True)
    transfer_fee.index.name = 'Mã Chi Nhánh'
    transfer_fee.reset_index(inplace=True)
    transfer_fee.insert(0,'STT',transfer_fee.index+1)

    sum_fee = transfer_fee['Phí Chuyển Khoản'].sum()
    table_title = f'PHÍ CHUYỂN KHOẢN {period}'

    # Write to Báo cáo phí chuyển khoản
    file_name = f'Báo cáo phí chuyển khoản {period}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book
    worksheet = workbook.add_worksheet(period)
    worksheet.hide_gridlines(option=2)
    # set column width
    worksheet.set_column('A:A',8)
    worksheet.set_column('B:B',27)
    worksheet.set_column('C:C',15)
    worksheet.set_column('D:D',18)

    title_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
        }
    )
    header_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'border': 1,
        }
    )
    stt_format = workbook.add_format(
        {
            'align': 'center',
            'border': 1,
        }
    )
    tenchinhanh_format = workbook.add_format(
        {
            'border': 1
        }
    )
    machinhanh_format = workbook.add_format(
        {
            'align': 'center',
            'border': 1
        }
    )
    philuuky_format = workbook.add_format(
        {
            'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'border': 1
        }
    )
    tong_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'border': 1,
        }
    )
    tongphiluuky_format = workbook.add_format(
        {
            'bold': True,
            'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'border': 1,
        }
    )
    worksheet.merge_range('A1:D1',table_title,title_format)
    worksheet.write_row(2,0,transfer_fee.columns,header_format)
    worksheet.write_column(3,0,transfer_fee['STT'],stt_format)
    worksheet.write_column(3,1,transfer_fee['Tên Chi Nhánh'],tenchinhanh_format)
    worksheet.write_column(3,2,transfer_fee['Mã Chi Nhánh'],machinhanh_format)
    worksheet.write_column(3,3,transfer_fee['Phí Chuyển Khoản'],philuuky_format)
    tong_row = transfer_fee.shape[0]+3
    worksheet.merge_range(tong_row,0,tong_row,2,'Tổng',tong_format)
    worksheet.write(tong_row,3,sum_fee,tongphiluuky_format)
    writer.close()

    branch_index = branch_name.reset_index(name='Tên Chi Nhánh')
    lv1_col = ['Chuyển khoản thanh toán bù trừ']*3 + ['Chuyển khoản chứng khoán sang CTCK khác']*3 + ['Tổng cộng']*2
    lv2_col = ['Số lượng','Phí','Phí làm tròn']*2 + ['Số lượng','Phí']
    summary = pd.DataFrame(
        index=pd.MultiIndex.from_frame(branch_index),
        columns=pd.MultiIndex.from_arrays([lv1_col,lv2_col])
    )
    summary[('Chuyển khoản thanh toán bù trừ','Số lượng')] = summary.index.get_level_values(0).map(transfer_fee_TTBT['volume'])
    summary[('Chuyển khoản thanh toán bù trừ','Phí')] = summary.index.get_level_values(0).map(transfer_fee_TTBT['fee'])
    summary[('Chuyển khoản thanh toán bù trừ','Phí làm tròn')] = np.round(summary.loc[:,('Chuyển khoản thanh toán bù trừ','Phí')],0)
    summary[('Chuyển khoản chứng khoán sang CTCK khác','Số lượng')] = summary.index.get_level_values(0).map(transfer_fee_CTCK['volume'])
    summary[('Chuyển khoản chứng khoán sang CTCK khác','Phí')] = summary.index.get_level_values(0).map(transfer_fee_CTCK['fee'])
    summary[('Chuyển khoản chứng khoán sang CTCK khác','Phí làm tròn')] = np.round(summary[('Chuyển khoản chứng khoán sang CTCK khác','Phí')],0)
    summary[('Tổng cộng','Số lượng')] = summary[('Chuyển khoản thanh toán bù trừ','Số lượng')] + summary[('Chuyển khoản chứng khoán sang CTCK khác','Số lượng')]
    summary[('Tổng cộng','Phí')] = summary[('Chuyển khoản thanh toán bù trừ','Phí làm tròn')] + summary[('Chuyển khoản chứng khoán sang CTCK khác','Phí làm tròn')]
    summary.fillna(0,inplace=True)

    # Write to Báo cáo phí chuyển khoản tổng hợp
    file_name = f'Báo cáo phí chuyển khoản tổng hợp {period}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options': {'nan_inf_to_errors':True}}
    )
    workbook = writer.book
    worksheet = workbook.add_worksheet(period)
    worksheet.hide_gridlines(option=2)
    # set column width
    worksheet.set_column('A:A',8)
    worksheet.set_column('B:B',27)
    worksheet.set_column('C:J',15)

    header_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'border': 1,
        }
    )
    tongcong_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'border': 1,
        }
    )
    column_sum_format = workbook.add_format(
        {
            'bold': True,
            'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'align': 'center',
            'border': 1,
        }
    )
    branch_id_format = workbook.add_format(
        {
            'num_format': '@',
            'align': 'center',
            'border': 1,
        }
    )
    branch_name_format = workbook.add_format(
        {
            'valign': 'vcenter',
            'border': 1,
        }
    )
    number_format = workbook.add_format(
        {
            'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'border': 1
        }
    )
    last_col_format = workbook.add_format(
        {
            'bg_color': 'FFF300',
            'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'border': 1
        }
    )
    column_0 = summary.index.names[0]
    column_1 = summary.index.names[1]
    column_0_24 = summary.columns.get_level_values(0).drop_duplicates()[0]
    column_0_57 = summary.columns.get_level_values(0).drop_duplicates()[1]
    column_0_89 = summary.columns.get_level_values(0).drop_duplicates()[2]
    columns_1 =  summary.columns.get_level_values(1)
    branch_id_col = summary.index.get_level_values(0)
    branch_name_col = summary.index.get_level_values(1)
    worksheet.merge_range('A1:A2',column_0,header_format)
    worksheet.merge_range('B1:B2',column_1,header_format)
    worksheet.merge_range('C1:E1',column_0_24,header_format)
    worksheet.merge_range('F1:H1',column_0_57,header_format)
    worksheet.merge_range('I1:J1',column_0_89,header_format)
    worksheet.write_row('C2',columns_1,header_format)
    worksheet.write_column('A3',branch_id_col,branch_id_format)
    worksheet.write_column('B3',branch_name_col,branch_name_format)
    for col in range(summary.shape[1]):
        data = summary.iloc[:,col]
        if col < summary.shape[1]-1:
            worksheet.write_column(2,col+2,data,number_format)
        else:
            worksheet.write_column(2,col+2,data,last_col_format)
    sum_row = summary.shape[0]+2
    worksheet.merge_range(sum_row,0,sum_row,1,'Tổng cộng',tongcong_format)
    sum_values = summary.values.sum(axis=0)
    worksheet.write_row(sum_row,2,sum_values,column_sum_format)
    writer.close()

    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')