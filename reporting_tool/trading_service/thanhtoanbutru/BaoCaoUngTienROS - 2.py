"""
    1. cột Doanh thu UTTB lấy từ cột fee_PHS trong table payment_in_advance
    2. có 3/6 kết quả bị lệch so với file mẫu
        a. Trên flex:
            022C017482	CÔNG TY TNHH ĐẦU TƯ XÂY DỰNG VÀ THƯƠNG MẠI ĐẠI DƯƠNG XANH	72,895,066
        Trên Database SQL:
            022C017482	CÔNG TY TNHH ĐẦU TƯ XÂY DỰNG VÀ THƯƠNG MẠI ĐẠI DƯƠNG XANH	76,204,995
        b. Trên flex:
            022C018358	CÔNG TY TNHH ĐẦU TƯ THƯƠNG MẠI VÀ XUẤT NHẬP KHẨU TÂM AN	    27,733,192
        Trên Database SQL:
            022C018358	CÔNG TY TNHH ĐẦU TƯ THƯƠNG MẠI VÀ XUẤT NHẬP KHẨU TÂM AN	    28,462,132
        c. Trên flex:
            022C019357	CÔNG TY CỔ PHẦN ĐẦU TƯ THƯƠNG MẠI VÀ PHÁT TRIỂN DỊCH VỤ PHÚC THỊNH	 69,056,567
        Trên Database SQL:
            022C019357	CÔNG TY CỔ PHẦN ĐẦU TƯ THƯƠNG MẠI VÀ PHÁT TRIỂN DỊCH VỤ PHÚC THỊNH	 69,050,864
"""

from reporting_tool.trading_service.thanhtoanbutru import *


def run(
        start_date: str,  # 2021-11-26
        end_date: str,  # 2021-12-25
        run_time=None,
):
    start = time.time()
    info = get_info('monthly', run_time)
    period = info['period']
    folder_name = info['folder_name']
    date_character = ['/', '-', '.']

    # create folder
    if not os.path.isdir(join(dept_folder, folder_name)):  # dept_folder from import
        os.mkdir(join(dept_folder, folder_name))
    if not os.path.isdir(join(dept_folder, folder_name, period)):
        os.mkdir((join(dept_folder, folder_name, period)))

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viết Query và xử lý dataframe ---------------------
    so_TK_ROS = ('022C016567',
                 '022C016608',
                 '022C015559',
                 '022C015598',
                 '022C015959',
                 '022C017264',
                 '022C017940',
                 '022C017289',
                 '022C017482',
                 '022C017535',
                 '022C018358',
                 '022C019357',
                 '022C696969',
                 '022C040921',
                 '022C041906',
                 '022C042028',
                 '022C040945'
                 )
    doanh_thu_UTTB_nhom_ROS = pd.read_sql(
        f"""
            SELECT 
                MAX(relationship.account_code) AS account_code,
                SUM(payment_in_advance.fee_at_phs) AS fee_PHS,
                SUM(payment_in_advance.total_fee) AS total_fee,
                SUM(payment_in_advance.receivable) AS receivable,
                SUM(payment_in_advance.received_amount) AS received_amount
            FROM 
                payment_in_advance
            LEFT JOIN 
                relationship ON relationship.sub_account = payment_in_advance.sub_account
            WHERE 
                payment_in_advance.date between '{start_date}' AND '{end_date}'
                AND relationship.date = '{end_date}'
                AND relationship.account_code in {so_TK_ROS}
            GROUP BY payment_in_advance.sub_account
        """,
        connect_DWH_CoSo,
        index_col='account_code'
    )
    query_ROS_account = pd.read_sql(
        f"""
            SELECT account_code, customer_name 
            FROM account
            WHERE account_code IN {so_TK_ROS}
        """,
        connect_DWH_CoSo,
        index_col='account_code'
    )
    # match 2 dataframe trên thành 1
    query_ROS_account['doanh_thu_UTTB'] = doanh_thu_UTTB_nhom_ROS['fee_PHS']
    query_ROS_account.fillna(0, inplace=True)

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File Excel ---------------------
    # Write file BÁO CÁO DOANH THU UTTB NHÓM ROS TẠO RA
    for date_char in date_character:
        if date_char in start_date and date_char in end_date:
            start_date = start_date.replace(date_char, '/')
            end_date = end_date.replace(date_char, '/')
    footer_date = bdate(end_date, 1).split('-')
    som = dt.datetime.strptime(start_date, "%Y/%m/%d").strftime("%m-%y")
    eom = dt.datetime.strptime(end_date, "%Y/%m/%d").strftime("%m-%y")
    f_name = f'Doanh thu UTTB nhóm ROS tạo ra từ {som} đến {eom}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder, folder_name, period, f_name),
        engine='xlsxwriter',
        engine_kwargs={'options': {'nan_inf_to_errors': True}}
    )
    workbook = writer.book

    ###################################################
    ###################################################
    ###################################################

    # ------------- Viết sheet -------------
    # Format
    sheet_title_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 20,
            'font_name': 'Times New Roman',
        }
    )
    STK_ten_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 12,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    STT_header_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman',
        }
    )
    doanh_Thu_UTTB_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'bg_color': '#FFF2CB'
        }
    )
    stt_col_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman'
        }
    )
    text_left_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Calibri'
        }
    )
    money_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vbottom',
            'font_size': 12,
            'font_name': 'Times New Roman',
            'num_format': '#,##0'
        }
    )
    total_format = workbook.add_format(
        {
            'bold': True,
            'border': 1,
            'align': 'center',
            'valign': 'vbottom',
            'font_size': 12,
            'font_name': 'Times New Roman'
        }
    )
    total_money_format = workbook.add_format(
        {
            'bold': True,
            'border': 1,
            'align': 'right',
            'valign': 'vbottom',
            'font_size': 12,
            'font_name': 'Times New Roman',
            'num_format': '#,##0'
        }
    )
    sod = dt.datetime.strptime(start_date, "%Y/%m/%d").strftime("%d/%m/%Y")
    eod = dt.datetime.strptime(end_date, "%Y/%m/%d").strftime("%d/%m/%Y")
    sheet_title_name = 'DOANH THU UTTB NHÓM ROS TẠO RA'
    from_to_day = f'Từ {sod} đến {eod}'

    # --------- sheet BAO CAO CAN LAM ---------
    sheet1 = workbook.add_worksheet('Sheet1')

    sum_start_row = query_ROS_account.shape[0] + 5
    total_doanh_thu = query_ROS_account['doanh_thu_UTTB'].sum()

    # Set Column Width and Row Height
    sheet1.set_column('A:A', 8.43)
    sheet1.set_column('B:B', 14)
    sheet1.set_column('C:C', 71.71)
    sheet1.set_column('D:D', 23.71)
    sheet1.merge_range('B2:D2', sheet_title_name, sheet_title_format)
    sheet1.merge_range('B3:D3', from_to_day, sheet_title_format)

    sheet1.write(
        'A5',
        'STT',
        STT_header_format
    )
    sheet1.write(
        'B5',
        'SỐ TKCK',
        STK_ten_format
    )
    sheet1.write(
        'C5',
        'TÊN',
        STK_ten_format
    )
    sheet1.write(
        'D5',
        'DOANH THU UTTB',
        doanh_Thu_UTTB_format
    )
    sheet1.write_column(
        'A6',
        [int(i) for i in np.arange(query_ROS_account.shape[0]) + 1],
        stt_col_format
    )
    sheet1.write_column(
        'B6',
        query_ROS_account.index,
        text_left_format
    )
    sheet1.write_column(
        'C6',
        query_ROS_account['customer_name'],
        text_left_format
    )
    query_ROS_account['doanh_thu_UTTB'] = query_ROS_account['doanh_thu_UTTB'].replace(0, '-')
    sheet1.write_column(
        'D6',
        query_ROS_account['doanh_thu_UTTB'],
        money_format
    )
    sheet1.merge_range(f'A{sum_start_row}:C{sum_start_row}', 'Total', total_format)
    sheet1.write(
        f'D{sum_start_row}',
        total_doanh_thu,
        total_money_format
    )

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py', '')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time() - start, 1)}s')
