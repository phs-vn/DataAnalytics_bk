"""
1. RCF1002 -> cash_balance
   ROD0040 -> trading_record
   sum các giá trị giống nhau theo price (đa số chứ ko phải tất cả, sum ko theo qui tắc nhất định)
"""
from reporting_tool.trading_service.thanhtoanbutru import *


def run(
        periodicity: str,
        start_date: str,  # 2021-11-23
        run_time=None,
):
    start = time.time()
    info = get_info(periodicity, run_time)
    end_date = bdate(start_date, 2)
    period = info['period']
    folder_name = info['folder_name']
    date_character = ['/', '-', '.']

    # create folder
    if not os.path.isdir(join(dept_folder, folder_name)):
        os.mkdir(join(dept_folder, folder_name))
    if not os.path.isdir(join(dept_folder, folder_name, period)):
        os.mkdir((join(dept_folder, folder_name, period)))

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viết Query và xử lý dataframe ---------------------
    # Viết Query
    # query ROD0040
    ket_qua_khop_lenh_query = pd.read_sql(
        f"""
            SELECT
            date,
            sub_account,
            value,
            fee,
            tax_of_selling,
            type_of_order,
            price
            FROM trading_record
            WHERE trading_record.date = '{start_date}'
            order by sub_account ASC
        """,
        connect_DWH_CoSo
    )
    # query RCF1002
    giao_dich_tien_query = pd.read_sql(
        f"""
            SELECT
            date,
            sub_account,
            transaction_id, 
            remark, 
            increase, 
            decrease
            FROM cash_balance 
            WHERE (transaction_id in ('8855', '8865') and date = '{start_date}')
            OR (transaction_id in ('8856', '8866') and date = '{end_date}')
            ORDER BY date, sub_account ASC
        """,
        connect_DWH_CoSo
    )
    get_info_customer_query = pd.read_sql(
        f"""
            SELECT
            date,
            sub_account, 
            relationship.account_code,
            account.customer_name
            FROM relationship 
            LEFT JOIN account ON relationship.account_code = account.account_code 
            WHERE date = '{start_date}'
        """,
        connect_DWH_CoSo,
        index_col='sub_account'
    )
    # Xử lý dataframe
    # Xử lý ROD0040 - ket_qua_khop_lenh_query
    ket_qua_khop_lenh_groupby = ket_qua_khop_lenh_query.groupby(['sub_account', 'type_of_order']).sum()
    ket_qua_khop_lenh_groupby = ket_qua_khop_lenh_groupby.reset_index(['type_of_order'])
    ket_qua_khop_lenh_groupby = ket_qua_khop_lenh_groupby.rename(
        columns={"value": "value_KQKL", "fee": "fee_KQKL", "tax_of_selling": "tax_of_selling_KQKL"}
    )

    # Xử lý RCF1002 - giao_dich_tien_query
    # data of start day
    res_rcf_T0 = giao_dich_tien_query.loc[
        giao_dich_tien_query['date'] == start_date,
        [
            'sub_account',
            'date',
            'transaction_id',
            'increase',
            'decrease'
        ]
    ].copy()
    res_rcf_T0 = res_rcf_T0.groupby(['sub_account', 'transaction_id', 'date'])[['increase', 'decrease']].sum()
    res_rcf_T0 = res_rcf_T0.reset_index(['transaction_id', 'date'])
    res_rcf_T0 = res_rcf_T0.add_suffix('_T0')

    # data of T+2 day
    res_rcf_T2 = giao_dich_tien_query.loc[
        giao_dich_tien_query['date'] == end_date,
        [
            'sub_account',
            'date',
            'transaction_id',
            'increase',
            'decrease'
        ]
    ].copy()
    res_rcf_T2 = res_rcf_T2.groupby(['sub_account', 'transaction_id', 'date'])[['increase', 'decrease']].sum()
    res_rcf_T2 = res_rcf_T2.reset_index(['transaction_id', 'date'])
    res_rcf_T2 = res_rcf_T2.add_suffix('_T+2')

    ket_qua_khop_lenh_groupby.loc[
        ket_qua_khop_lenh_groupby['type_of_order'] == 'B', 'value_GDT'
    ] = res_rcf_T0.loc[res_rcf_T0['transaction_id_T0'] == '8865', 'decrease_T0']
    ket_qua_khop_lenh_groupby.loc[
        ket_qua_khop_lenh_groupby['type_of_order'] == 'B', 'fee_GDT'
    ] = res_rcf_T0.loc[res_rcf_T0['transaction_id_T0'] == '8855', 'decrease_T0']
    ket_qua_khop_lenh_groupby.loc[
        ket_qua_khop_lenh_groupby['type_of_order'] == 'S', 'value_GDT'
    ] = res_rcf_T2.loc[res_rcf_T2['transaction_id_T+2'] == '8866', 'increase_T+2']
    ket_qua_khop_lenh_groupby.loc[
        ket_qua_khop_lenh_groupby['type_of_order'] == 'S', 'fee_GDT'
    ] = res_rcf_T2.loc[res_rcf_T2['transaction_id_T+2'] == '8856', 'decrease_T+2']
    ket_qua_khop_lenh_groupby.loc[
        ket_qua_khop_lenh_groupby['type_of_order'] == 'B', 'date_thanh_toan'
    ] = res_rcf_T0.loc[res_rcf_T0['transaction_id_T0'] == '8865', 'date_T0']
    ket_qua_khop_lenh_groupby.loc[
        ket_qua_khop_lenh_groupby['type_of_order'] == 'S', 'date_thanh_toan'
    ] = res_rcf_T2.loc[res_rcf_T2['transaction_id_T+2'] == '8856', 'date_T+2']
    ket_qua_khop_lenh_groupby.fillna(0, inplace=True)
    # thêm 3 cột mã tài khoản, tên KH, ngày GD vào df
    ket_qua_khop_lenh_groupby.insert(0, 'account_code', get_info_customer_query['account_code'])
    ket_qua_khop_lenh_groupby.insert(1, 'customer_name', get_info_customer_query['customer_name'])
    ket_qua_khop_lenh_groupby.insert(2, 'date', get_info_customer_query['date'])
    # thêm cột thuế giao dịch tiền vào df
    tax_GDT_S = ket_qua_khop_lenh_groupby.loc[ket_qua_khop_lenh_groupby['type_of_order'] == 'S', 'tax_of_selling_KQKL']
    tax_GDT_B = ket_qua_khop_lenh_groupby.loc[ket_qua_khop_lenh_groupby['type_of_order'] == 'B', 'tax_of_selling_KQKL']
    tax_GDT = pd.DataFrame({'tax_GDT_S': tax_GDT_S, 'tax_GDT_B': tax_GDT_B})
    tax_GDT.fillna(0, inplace=True)
    ket_qua_khop_lenh_groupby.loc[ket_qua_khop_lenh_groupby['type_of_order'] == 'B', 'tax_GDT'] = tax_GDT['tax_GDT_B']
    ket_qua_khop_lenh_groupby.loc[ket_qua_khop_lenh_groupby['type_of_order'] == 'S', 'tax_GDT'] = tax_GDT['tax_GDT_S']

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File Excel ---------------------
    # Write file BÁO CÁO ĐỐI CHIẾU THANH TOÁN BÙ TRỪ TIỀN MUA BÁN CHỨNG KHOÁN
    for date_char in date_character:
        if date_char in start_date and date_char in end_date:
            start_date = start_date.replace(date_char, '/')
            end_date = end_date.replace(date_char, '/')
    s_date_file_name = dt.datetime.strptime(start_date, "%Y/%m/%d").strftime("%d-%m")
    f_name = f'Đối chiếu TTBT tiền mua bán chứng khoán {s_date_file_name}.xlsx'
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
            'font_size': 14,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    from_to_format = workbook.add_format(
        {
            'italic': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 12,
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
    stt_col_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vbottom',
            'font_size': 12,
            'font_name': 'Calibri'
        }
    )
    text_left_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vbottom',
            'text_wrap': True,
            'font_size': 12,
            'font_name': 'Calibri'
        }
    )
    money_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vbottom',
            'font_size': 12,
            'font_name': 'Calibri',
            'num_format': '#,##0'
        }
    )
    sum_money_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'num_format': '#,##0'
        }
    )
    zero_to_dashes_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'num_format': '0;-0;-;@'
        }
    )
    date_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'num_format': 'dd/mm/yyyy'
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
        'Ngày giao dịch',
        'Ngày thanh toán tiền',
        'Số tài khoản',
        'Số tiểu khoản',
        'Tên khách hàng',
        'Loại lệnh',
        'Theo kết quả khớp lệnh',
        'Theo giao dịch tiền',
        'Lệch',
    ]
    sub_headers = [
        'Giá trị khớp',
        'Phí',
        'Thuế',
    ]

    companyAddress = 'Tầng 3, CR3-03A, 109 Tôn Dật Tiên, phường Tân Phú, Quận 7, Thành phố Hồ Chí Minh'
    sheet_title_name = 'BÁO CÁO ĐỐI CHIẾU THANH TOÁN BÙ TRỪ TIỀN MUA BÁN CHỨNG KHOÁN'
    sub_title_start_date = dt.datetime.strptime(start_date, "%Y/%m/%d").strftime("%d/%m/%Y")
    sub_title_name = f'Từ ngày {sub_title_start_date} đến {sub_title_start_date}'

    # --------- sheet BAO CAO CAN LAM ---------
    sheet_bao_cao_can_lam = workbook.add_worksheet('BAO CAO CAN LAM')

    # Insert phu hung picture
    sheet_bao_cao_can_lam.insert_image('A1', './img/phu_hung.png', {'x_scale': 0.66, 'y_scale': 0.71})

    # Set Column Width and Row Height
    sheet_bao_cao_can_lam.set_column('A:A', 4.43)
    sheet_bao_cao_can_lam.set_column('B:C', 14.14)
    sheet_bao_cao_can_lam.set_column('D:E', 12.14)
    sheet_bao_cao_can_lam.set_column('F:F', 31.5)
    sheet_bao_cao_can_lam.set_column('G:G', 9.14)
    sheet_bao_cao_can_lam.set_column('H:H', 13.71)
    sheet_bao_cao_can_lam.set_column('I:I', 11.57)
    sheet_bao_cao_can_lam.set_column('J:J', 10.14)
    sheet_bao_cao_can_lam.set_column('K:K', 13.71)
    sheet_bao_cao_can_lam.set_column('L:L', 11.57)
    sheet_bao_cao_can_lam.set_column('M:M', 10.14)
    sheet_bao_cao_can_lam.set_column('N:P', 10.5)

    # merge row
    sheet_bao_cao_can_lam.merge_range('C1:I1', CompanyName, company_name_format)
    sheet_bao_cao_can_lam.merge_range('C2:J2', companyAddress, company_format)
    sheet_bao_cao_can_lam.merge_range('C3:I3', CompanyPhoneNumber, company_format)
    sheet_bao_cao_can_lam.merge_range('B7:P7', sheet_title_name, sheet_title_format)
    sheet_bao_cao_can_lam.merge_range('B8:P8', sub_title_name, from_to_format)
    sheet_bao_cao_can_lam.merge_range('A10:A11', headers[0], headers_format)
    sheet_bao_cao_can_lam.merge_range('B10:B11', headers[1], headers_format)
    sheet_bao_cao_can_lam.merge_range('C10:C11', headers[2], headers_format)
    sheet_bao_cao_can_lam.merge_range('D10:D11', headers[3], headers_format)
    sheet_bao_cao_can_lam.merge_range('E10:E11', headers[4], headers_format)
    sheet_bao_cao_can_lam.merge_range('F10:F11', headers[5], headers_format)
    sheet_bao_cao_can_lam.merge_range('G10:G11', headers[6], headers_format)
    sheet_bao_cao_can_lam.merge_range('H10:J10', headers[7], headers_format)
    sheet_bao_cao_can_lam.merge_range('K10:M10', headers[8], headers_format)
    sheet_bao_cao_can_lam.merge_range('N10:P10', headers[9], headers_format)
    sum_start_row = ket_qua_khop_lenh_groupby.shape[0] + 12
    sheet_bao_cao_can_lam.merge_range(
        f'A{sum_start_row}:G{sum_start_row}',
        'Tổng',
        headers_format
    )
    footer_date = bdate(end_date, 1).split('-')
    footer_start_row = sum_start_row + 2
    sheet_bao_cao_can_lam.merge_range(
        f'N{footer_start_row}:P{footer_start_row}',
        f'Ngày {footer_date[2]} tháng {footer_date[1]} năm {footer_date[0]}',
        footer_dmy_format
    )
    sheet_bao_cao_can_lam.merge_range(
        f'N{footer_start_row + 1}:P{footer_start_row + 1}',
        'Người duyệt',
        footer_text_format
    )
    sheet_bao_cao_can_lam.merge_range(
        f'C{footer_start_row + 1}:E{footer_start_row + 1}',
        'Người lập',
        footer_text_format
    )

    # write row & column
    sheet_bao_cao_can_lam.write_row('H11', sub_headers * 3, headers_format)
    sheet_bao_cao_can_lam.write_row(
        'A4',
        [''] * (len(headers) + len(sub_headers) + 3),
        empty_row_format
    )
    sheet_bao_cao_can_lam.write_column(
        'A12',
        [int(i) for i in np.arange(ket_qua_khop_lenh_groupby.shape[0]) + 1],
        stt_col_format
    )
    sheet_bao_cao_can_lam.write_column(
        'B12',
        ket_qua_khop_lenh_groupby['date'],
        date_format
    )
    sheet_bao_cao_can_lam.write_column(
        'C12',
        ket_qua_khop_lenh_groupby['date_thanh_toan'],
        date_format
    )
    sheet_bao_cao_can_lam.write_column(
        'D12',
        ket_qua_khop_lenh_groupby['account_code'],
        text_left_format
    )
    sheet_bao_cao_can_lam.write_column(
        'E12',
        ket_qua_khop_lenh_groupby.index,
        text_left_format
    )
    sheet_bao_cao_can_lam.write_column(
        'F12',
        ket_qua_khop_lenh_groupby['customer_name'].str.upper(),
        text_left_format
    )
    sheet_bao_cao_can_lam.write_column(
        'G12',
        ket_qua_khop_lenh_groupby['type_of_order'],
        text_left_format
    )
    sheet_bao_cao_can_lam.write_column(
        'H12',
        ket_qua_khop_lenh_groupby['value_KQKL'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'I12',
        ket_qua_khop_lenh_groupby['fee_KQKL'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'J12',
        ket_qua_khop_lenh_groupby['tax_of_selling_KQKL'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'K12',
        ket_qua_khop_lenh_groupby['value_GDT'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'L12',
        ket_qua_khop_lenh_groupby['fee_GDT'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'M12',
        ket_qua_khop_lenh_groupby['tax_GDT'],
        money_format
    )
    ket_qua_khop_lenh_groupby['value_lech'] = ket_qua_khop_lenh_groupby['value_GDT'] - ket_qua_khop_lenh_groupby['value_KQKL']
    ket_qua_khop_lenh_groupby['fee_lech'] = ket_qua_khop_lenh_groupby['fee_GDT'] - ket_qua_khop_lenh_groupby['fee_KQKL']
    ket_qua_khop_lenh_groupby['tax_lech'] = ket_qua_khop_lenh_groupby['tax_GDT'] - ket_qua_khop_lenh_groupby['tax_of_selling_KQKL']
    sheet_bao_cao_can_lam.write_column(
        'N12',
        ket_qua_khop_lenh_groupby['value_lech'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'O12',
        ket_qua_khop_lenh_groupby['fee_lech'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'P12',
        ket_qua_khop_lenh_groupby['tax_lech'],
        money_format
    )
    sheet_bao_cao_can_lam.write(
        f'D{footer_start_row + 1}',
        'Người lập',
        footer_text_format
    )
    sheet_bao_cao_can_lam.write(
        f'H{sum_start_row}',
        ket_qua_khop_lenh_groupby['value_KQKL'].sum(),
        money_format
    )
    sheet_bao_cao_can_lam.write(
        f'I{sum_start_row}',
        ket_qua_khop_lenh_groupby['fee_KQKL'].sum(),
        money_format
    )
    sheet_bao_cao_can_lam.write(
        f'J{sum_start_row}',
        ket_qua_khop_lenh_groupby['tax_of_selling_KQKL'].sum(),
        money_format
    )
    sheet_bao_cao_can_lam.write(
        f'K{sum_start_row}',
        ket_qua_khop_lenh_groupby['value_GDT'].sum(),
        money_format
    )
    sheet_bao_cao_can_lam.write(
        f'L{sum_start_row}',
        ket_qua_khop_lenh_groupby['fee_GDT'].sum(),
        money_format
    )
    sheet_bao_cao_can_lam.write(
        f'M{sum_start_row}',
        ket_qua_khop_lenh_groupby['tax_GDT'].sum(),
        money_format
    )
    sheet_bao_cao_can_lam.write(
        f'N{sum_start_row}',
        ket_qua_khop_lenh_groupby['value_lech'].sum(),
        zero_to_dashes_format
    )
    sheet_bao_cao_can_lam.write(
        f'O{sum_start_row}',
        ket_qua_khop_lenh_groupby['fee_lech'].sum(),
        zero_to_dashes_format
    )
    sheet_bao_cao_can_lam.write(
        f'P{sum_start_row}',
        ket_qua_khop_lenh_groupby['tax_lech'].sum(),
        zero_to_dashes_format
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