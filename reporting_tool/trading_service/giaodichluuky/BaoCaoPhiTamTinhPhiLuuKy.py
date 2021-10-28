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
    
    depository_fee = pd.read_sql(
        "SELECT sub_account, fee_amount FROM depository_fee "
        f"WHERE date BETWEEN '{start_date}' AND '{end_date}'",
        connect_DWH_CoSo,
        index_col='sub_account',
    )
    account = pd.read_sql(
        "SELECT sub_account, account_code FROM sub_account;",
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
    result = pd.DataFrame(
        columns=[
            'STT',
            'Tên Chi Nhánh',
            'Phí Lưu Ký'
        ],
        index=branch_name.index
    )
    depository_fee['account'] = account
    depository_fee['branch'] = depository_fee['account'].map(broker).map(branch_id)
    depository_fee = depository_fee.groupby('branch')['fee_amount'].sum()
    result['STT'] = np.arange(1,result.shape[0]+1)
    result['Tên Chi Nhánh'] = branch_name
    result['Phí Lưu Ký'] = depository_fee
    result['Phí Lưu Ký'].fillna(0,inplace=True)
    result.index.name = 'Mã Chi Nhánh'
    result.reset_index(inplace=True)
    result = result[['STT','Tên Chi Nhánh','Mã Chi Nhánh','Phí Lưu Ký']]

    sum_fee = result['Phí Lưu Ký'].sum()
    table_title = f'PHÍ LƯU KÝ {period}'

    # Write to Excel
    file_name = f'Báo cáo phí tạm tính phí lưu ký {period}.xlsx'
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
    for col_num, col_name in enumerate(result.columns):
        worksheet.write(2,col_num,col_name,header_format)
    for row in range(result.shape[0]):
        worksheet.write(row+3,0,result.iloc[row,0],stt_format)
        worksheet.write(row+3,1,result.iloc[row,1],tenchinhanh_format)
        worksheet.write(row+3,2,result.iloc[row,2],machinhanh_format)
        worksheet.write(row+3,3,result.iloc[row,3],philuuky_format)
    tong_row = result.shape[0]+3
    worksheet.merge_range(tong_row,0,tong_row,2,'Tổng',tong_format)
    worksheet.write(tong_row,3,sum_fee,tongphiluuky_format)
    writer.close()

    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')