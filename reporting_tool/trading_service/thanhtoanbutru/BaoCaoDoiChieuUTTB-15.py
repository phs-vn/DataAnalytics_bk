"""
    1. daily
    2. table: trading_record, cash_balance
    3. Tiền Hoàn trả UTTB T0: điều kiện mã 8851
    4. Tổng giá trị tiền bán có thể ứng: (phí bán (fee), thuế bán (tax_of_selling), phí ứng tiền (cột decrease - 1153))
    5. Tiền đã ứng: cột increase - 1153
"""
from reporting_tool.trading_service.thanhtoanbutru import *


def run(
        periodicity: str,
        start_date: str,  # 2021-11-18
        end_date: str,  # 2021-11-22
        run_time=None,
):
    start = time.time()
    info = get_info(periodicity, run_time)
    # start_date = info['start_date']
    # end_date = info['end_date']
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
    # query gets value T0,T-1. T-2
    trading_record_query = pd.read_sql(
        f"""
            SELECT
            trading_record.date,
            relationship.sub_account, 
            trading_record.value,
            trading_record.fee,
            trading_record.tax_of_selling
            FROM trading_record
            LEFT JOIN relationship ON relationship.sub_account = trading_record.sub_account
            WHERE trading_record.date BETWEEN '{start_date}' AND '{end_date}'
            AND relationship.date = '{end_date}'
            AND trading_record.type_of_order = 'S'
            ORDER BY trading_record.date ASC
        """,
        connect_DWH_CoSo,
    )
    # Query get account_code and customer_name
    account = pd.read_sql(
        f"""
            SELECT
            sub_account, 
            relationship.account_code,
            account.customer_name
            FROM relationship 
            LEFT JOIN account ON relationship.account_code = account.account_code 
            WHERE date = '{end_date}'
        """,
        connect_DWH_CoSo,
        index_col='sub_account'
    )
    # data of T-2 day
    value_T2 = trading_record_query.loc[
        trading_record_query['date'] == start_date,
        [
            'sub_account',
            'value',
        ]
    ].copy()
    value_T2 = value_T2.groupby(['sub_account']).sum()
    value_T2 = value_T2.add_suffix('_T2')

    # data of T-1 day
    value_T1 = trading_record_query.loc[
        trading_record_query['date'] == bdate(start_date, 1),
        [
            'sub_account',
            'value',
        ]
    ].copy()
    value_T1 = value_T1.groupby(['sub_account']).sum()
    value_T1 = value_T1.add_suffix('_T1')

    # data of T0 day
    value_T0 = trading_record_query.loc[
        trading_record_query['date'] == end_date,
        [
            'sub_account',
            'value',
            'fee',
            'tax_of_selling'
        ]
    ].copy()
    value_T0 = value_T0.groupby(['sub_account']).sum()
    value_T0 = value_T0.add_suffix('_T0')

    # Query to get transaction_id, increase, decrease
    cash_balance_query = pd.read_sql(
        f"""
            SELECT 
            date,
            sub_account, 
            transaction_id, 
            remark, 
            increase, 
            decrease
            FROM cash_balance 
            WHERE date = '{end_date}'
            AND (transaction_id = '8851' or transaction_id = '1153') 
        """,
        connect_DWH_CoSo,
    )
    hoan_tra_uttb_T0_table = cash_balance_query.loc[
        cash_balance_query['transaction_id'] == '8851', ['sub_account', 'decrease']].copy()
    hoan_tra_uttb_T0_table = hoan_tra_uttb_T0_table.groupby(['sub_account']).sum()
    hoan_tra_uttb_T0_table.columns = ['hoan_tra_UTTB_T0']

    tien_da_ung_table = cash_balance_query.loc[
        cash_balance_query['transaction_id'] == '1153', ['sub_account', 'increase']].copy()
    tien_da_ung_table = tien_da_ung_table.groupby(['sub_account']).sum()
    tien_da_ung_table.columns = ['tien_da_ung']

    phi_ung_tien_table = cash_balance_query.loc[
        cash_balance_query['transaction_id'] == '1153', ['sub_account', 'decrease']].copy()
    phi_ung_tien_table = phi_ung_tien_table.groupby(['sub_account']).sum()
    phi_ung_tien_table.columns = ['phi_ung_tien']

    final_table = pd.concat([
        value_T2,
        value_T1,
        value_T0,
        hoan_tra_uttb_T0_table,
        phi_ung_tien_table,
        tien_da_ung_table
    ], axis=1)
    final_table.fillna(0, inplace=True)
    final_table.insert(0, 'account_code', account['account_code'])
    final_table.insert(1, 'customer_name', account['customer_name'])

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File Excel ---------------------
    # Write file BÁO CÁO ĐỐI CHIẾU LÃI TIỀN GỬI PHÁT SINH TRÊN TÀI KHOẢN KHÁCH HÀNG
    for date_char in date_character:
        if date_char in start_date and date_char in end_date:
            start_date = start_date.replace(date_char, '/')
            end_date = end_date.replace(date_char, '/')
    footer_date = bdate(end_date, 1).split('-')
    end_date = dt.datetime.strptime(end_date, "%Y/%m/%d").strftime("%d-%m")
    f_name = f'Báo cáo đối chiếu UTTB {end_date}.xlsx'
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
    company_name_format = workbook.add_format(
        {
            'bold': True,
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    company_format = workbook.add_format(
        {
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    empty_row_format = workbook.add_format(
        {
            'bottom': 1,
            'valign': 'top',
            'font_size': 10,
            'font_name': 'Arial',
        }
    )
    sheet_title_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 16,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    from_to_format = workbook.add_format(
        {
            'italic': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    headers_format = workbook.add_format(
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
    stt_row_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    stt_col_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'top',
            'font_size': 10,
            'font_name': 'Times New Roman'
        }
    )
    text_left_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'top',
            'font_size': 10,
            'font_name': 'Times New Roman'
        }
    )
    text_left_wrap_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'text_wrap': True,
            'valign': 'top',
            'font_size': 10,
            'font_name': 'Times New Roman'
        }
    )
    money_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'top',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'num_format': '#,##0'
        }
    )
    footer_dmy_format = workbook.add_format(
        {
            'italic': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman',
        }
    )
    footer_text_format = workbook.add_format(
        {
            'bold': True,
            'italic': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 12,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    headers = [
        'STT',
        'Số tài khoản',
        'Số tiểu khoản',
        'Tên khách hàng',
        'Giá trị tiền bán T-2',
        'Giá trị tiền bán T-1',
        'Giá trị tiền bán T0',
        'Tiền Hoàn trả UTTB T0',
        'Tổng giá trị tiền bán có thể ứng',
        'Tiền đã ứng',
        'Tiền còn có thể ứng',
        'Bất Thường'
    ]

    companyAddress = 'Tầng 3, CR3-03A, 109 Tôn Dật Tiên, phường Tân Phú, Quận 7, Thành phố Hồ Chí Minh'
    sheet_title_name = 'BÁO CÁO ĐỐI CHIẾU UTTB'
    sub_title_name = f'Ngày {end_date}'

    # --------- sheet BAO CAO CAN LAM ---------
    sheet_bao_cao_can_lam = workbook.add_worksheet('BAO CAO CAN LAM')

    # Insert phu hung picture
    sheet_bao_cao_can_lam.insert_image('A1', './img/phu_hung.png', {'x_scale': 0.65, 'y_scale': 0.71})

    # Set Column Width and Row Height
    sheet_bao_cao_can_lam.set_column('A:A', 6.29)
    sheet_bao_cao_can_lam.set_column('B:B', 12.14)
    sheet_bao_cao_can_lam.set_column('C:C', 13)
    sheet_bao_cao_can_lam.set_column('D:D', 31.43)
    sheet_bao_cao_can_lam.set_column('E:F', 11.86)
    sheet_bao_cao_can_lam.set_column('G:G', 14.71)
    sheet_bao_cao_can_lam.set_column('H:H', 14.29)
    sheet_bao_cao_can_lam.set_column('I:I', 15.71)
    sheet_bao_cao_can_lam.set_column('J:J', 11.71)
    sheet_bao_cao_can_lam.set_column('K:K', 14.86)
    sheet_bao_cao_can_lam.set_column('L:L', 15.86)

    # merge row
    sheet_bao_cao_can_lam.merge_range('C1:L1', CompanyName, company_name_format)
    sheet_bao_cao_can_lam.merge_range('C2:L2', companyAddress, company_format)
    sheet_bao_cao_can_lam.merge_range('C3:L3', CompanyPhoneNumber, company_format)
    sheet_bao_cao_can_lam.merge_range('A7:L7', sheet_title_name, sheet_title_format)
    sheet_bao_cao_can_lam.merge_range('A8:L8', sub_title_name, from_to_format)
    sum_start_row = final_table.shape[0] + 13
    sheet_bao_cao_can_lam.merge_range(
        f'A{sum_start_row}:D{sum_start_row}',
        'Tổng',
        headers_format
    )
    footer_start_row = sum_start_row + 2
    sheet_bao_cao_can_lam.merge_range(
        f'J{footer_start_row}:L{footer_start_row}',
        f'Ngày {footer_date[2]} tháng {footer_date[1]} năm {footer_date[0]}',
        footer_dmy_format
    )
    sheet_bao_cao_can_lam.merge_range(
        f'J{footer_start_row + 1}:L{footer_start_row + 1}',
        'Người duyệt',
        footer_text_format
    )
    sheet_bao_cao_can_lam.merge_range(
        f'A{footer_start_row + 1}:C{footer_start_row + 1}',
        'Người lập',
        footer_text_format
    )

    # write row, column
    sheet_bao_cao_can_lam.write_row(
        'A4',
        [''] * len(headers),
        empty_row_format
    )
    sheet_bao_cao_can_lam.write_row('A11', headers, headers_format)
    sheet_bao_cao_can_lam.write_row(
        'A12',
        [f'({i})' for i in np.arange(len(headers)) + 1],
        stt_row_format,
    )
    sheet_bao_cao_can_lam.write_row(
        f'E{sum_start_row}',
        [''] * 8,
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'A13',
        [int(i) for i in np.arange(final_table.shape[0]) + 1],
        stt_col_format
    )
    sheet_bao_cao_can_lam.write_column(
        'B13',
        final_table['account_code'],
        text_left_format
    )
    sheet_bao_cao_can_lam.write_column(
        'C13',
        final_table.index,
        text_left_format
    )
    sheet_bao_cao_can_lam.write_column(
        'D13',
        final_table['customer_name'],
        text_left_wrap_format
    )
    sheet_bao_cao_can_lam.write_column(
        'E13',
        final_table['value_T2'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'F13',
        final_table['value_T1'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'G13',
        final_table['value_T0'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'H13',
        final_table['hoan_tra_UTTB_T0'],
        money_format
    )
    calc_1 = final_table['value_T1'] + final_table['value_T0']
    calc_2 = final_table['fee_T0'] + final_table['tax_of_selling_T0'] + final_table['phi_ung_tien']
    sum_gia_tri_tien_ban = calc_1 - calc_2
    final_table['tong_gia_tri_tien_ban'] = sum_gia_tri_tien_ban
    sheet_bao_cao_can_lam.write_column(
        'I13',
        final_table['tong_gia_tri_tien_ban'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'J13',
        final_table['tien_da_ung'],
        money_format
    )
    final_table['tien_con_co_the_ung'] = final_table['tong_gia_tri_tien_ban'] - final_table['tien_da_ung']
    sheet_bao_cao_can_lam.write_column(
        'K13',
        final_table['tien_con_co_the_ung'],
        money_format
    )
    # Bất thường = True - Bình thường = False
    final_table['check_bat_thuong'] = None
    for idx in final_table.index:
        value_T2 = final_table.loc[idx, 'value_T2']
        hoan_tra_uttb_T0 = final_table.loc[idx, 'hoan_tra_UTTB_T0']
        tien_da_ung = final_table.loc[idx, 'tien_da_ung']
        tong_gia_tri_tien_ban = final_table.loc[idx, 'tong_gia_tri_tien_ban']
        tien_con_co_the_ung = final_table.loc[idx, 'tien_con_co_the_ung']
        if (hoan_tra_uttb_T0 != value_T2) or (tien_da_ung > tong_gia_tri_tien_ban) or (tien_con_co_the_ung < 0):
            final_table.loc[idx, 'check_bat_thuong'] = True
        else:
            final_table.loc[idx, 'check_bat_thuong'] = False
    sheet_bao_cao_can_lam.write_column(
        'L13',
        final_table['check_bat_thuong'],
        text_left_wrap_format
    )
    sheet_bao_cao_can_lam.write(
        f'E{sum_start_row}',
        final_table['value_T2'].sum(),
        money_format
    )
    sheet_bao_cao_can_lam.write(
        f'F{sum_start_row}',
        final_table['value_T1'].sum(),
        money_format
    )
    sheet_bao_cao_can_lam.write(
        f'G{sum_start_row}',
        final_table['value_T0'].sum(),
        money_format
    )
    sheet_bao_cao_can_lam.write(
        f'H{sum_start_row}',
        final_table['hoan_tra_UTTB_T0'].sum(),
        money_format
    )
    sheet_bao_cao_can_lam.write(
        f'I{sum_start_row}',
        final_table['tong_gia_tri_tien_ban'].sum(),
        money_format
    )
    sheet_bao_cao_can_lam.write(
        f'J{sum_start_row}',
        final_table['tien_da_ung'].sum(),
        money_format
    )
    sheet_bao_cao_can_lam.write(
        f'K{sum_start_row}',
        final_table['tien_con_co_the_ung'].sum(),
        money_format
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
