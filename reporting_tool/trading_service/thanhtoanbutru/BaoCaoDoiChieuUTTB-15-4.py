"""
    1. daily
    2. table: trading_record, cash_balance
    3. Tiền Hoàn trả UTTB T0: mã giao dịch 8851
    4. Tiền đã ứng: mã giao dịch - 1153
    5. có thể dùng bảng RCI0015 thay cho RCF1002 (chị Tuyết)
    6. Số tiền UTTB nhận và phí UTTB truyền vào cột ngày nào (T-2, T-1 hay T0) dựa vào
    phần ngày đầu tiên nằm trong diễn giải
    7. Phần tử Ngày đầu tiên luôn luôn bé hơn phần tử ngày thứ 2, ko có trường hợp ngược lại (Chị Thu Anh)
    8. Giá trị của cột Số tiền UTTB nhận trong RCF1002 chưa thực hiện trừ phí UTTB
    9. TK bị lệch:
                                                Số tiền UTTB KH     Phí UTTB ngày T0
                                                đã nhận ngày T0
        022C336888	0117002702	NGUYỄN THỊ TÀI	14,008,058,561	    7,786,581
                                Giá trị tiền bán T-2
                                                    Thuế Cổ Tức
        022C019179	0104001345	PHẠM THANH TÙNG     20,000

"""
from reporting_tool.trading_service.thanhtoanbutru import *
import string


def run(
        periodicity: str,
        end_date: str,
        run_time=None,
):
    start = time.time()
    info = get_info(periodicity, run_time)
    period = info['period']
    folder_name = info['folder_name']
    start_date = bdate(end_date, -2)  # T-2 date
    date_character = ['/', '-', '.']
    for date_char in date_character:
        if date_char in start_date and date_char in end_date:
            start_date = start_date.replace(date_char, '/')
            end_date = end_date.replace(date_char, '/')

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
            trading_record.tax_of_selling,
            trading_record.tax_of_share_dividend
            FROM trading_record
            LEFT JOIN relationship ON relationship.sub_account = trading_record.sub_account
            WHERE trading_record.date BETWEEN '{start_date}' AND '{end_date}'
            AND relationship.date = '{end_date}'
            AND trading_record.type_of_order = 'S'
            ORDER BY trading_record.date, sub_account ASC
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
            ORDER BY account_code
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
            'fee',
            'tax_of_selling',
            'tax_of_share_dividend'
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
            'fee',
            'tax_of_selling',
            'tax_of_share_dividend'
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
            'tax_of_selling',
            'tax_of_share_dividend'
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
            WHERE date BETWEEN '{start_date}' AND '{end_date}'
            AND transaction_id IN ('1153', '8851')
            ORDER BY date, sub_account, transaction_id
        """,
        connect_DWH_CoSo,
    )
    cash_balance_T0 = cash_balance_query.loc[cash_balance_query['date'] == end_date]

    # column Tiền Hoàn trả UTTB T0
    hoan_tra_uttb_T0_table = cash_balance_T0.loc[
        cash_balance_T0['transaction_id'] == '8851', ['sub_account', 'decrease']].copy()
    hoan_tra_uttb_T0_table = hoan_tra_uttb_T0_table.groupby(['sub_account']).sum()
    hoan_tra_uttb_T0_table.columns = ['hoan_tra_UTTB_T0']

    # column Tiền đã ứng
    tien_da_ung_table = cash_balance_query.loc[
        cash_balance_query['transaction_id'] == '1153', ['date', 'sub_account', 'increase', 'decrease',
                                                         'remark']].copy()
    remark = tien_da_ung_table['remark']
    tmp_lst = []
    date_lst = []
    for value in remark.items():
        res = value[1].split(',')
        tmp_lst.append(res[0])
    for i in tmp_lst:
        i = i.split(" ")
        date_lst.append(i[-1])
    dates_list = [dt.datetime.strptime(date, "%d.%m.%Y").strftime("%Y/%m/%d") for date in date_lst]
    tien_da_ung_table['date_loc'] = dates_list

    # Số tiền UTTB KH đã nhận và Phí UTTB ngày T-2
    tien_da_ung_T2 = tien_da_ung_table.loc[
        tien_da_ung_table['date_loc'] == start_date,
        [
            'sub_account',
            'increase',
            'decrease',
            'remark'
        ]
    ].copy()
    tien_da_ung_T2 = tien_da_ung_T2.groupby(['sub_account']).sum()
    tien_da_ung_T2 = tien_da_ung_T2.add_suffix('_T2')
    # Số tiền UTTB KH đã nhận ngày T-2 - Phí UTTB ngày T-2
    tien_da_ung_T2['increase_T2'] = tien_da_ung_T2['increase_T2'] - tien_da_ung_T2['decrease_T2']

    # Số tiền UTTB KH đã nhận và Phí UTTB ngày T-1
    t1_date = bdate(end_date, -1)
    t1_date = dt.datetime.strptime(t1_date, "%Y-%m-%d").strftime("%Y/%m/%d")
    tien_da_ung_T1 = tien_da_ung_table.loc[
        tien_da_ung_table['date_loc'] == t1_date,
        [
            'sub_account',
            'increase',
            'decrease',
            'remark'
        ]
    ].copy()
    tien_da_ung_T1 = tien_da_ung_T1.groupby(['sub_account']).sum()
    tien_da_ung_T1 = tien_da_ung_T1.add_suffix('_T1')
    # Số tiền UTTB KH đã nhận ngày T-1 - Phí UTTB ngày T-1
    tien_da_ung_T1['increase_T1'] = tien_da_ung_T1['increase_T1'] - tien_da_ung_T1['decrease_T1']

    # Số tiền UTTB KH đã nhận và Phí UTTB ngày T0
    tien_da_ung_T0 = tien_da_ung_table.loc[
        tien_da_ung_table['date_loc'] == end_date,
        [
            'sub_account',
            'increase',
            'decrease',
            'remark'
        ]
    ].copy()
    tien_da_ung_T0 = tien_da_ung_T0.groupby(['sub_account']).sum()
    tien_da_ung_T0 = tien_da_ung_T0.add_suffix('_T0')
    # Số tiền UTTB KH đã nhận ngày T0 - Phí UTTB ngày T0
    tien_da_ung_T0['increase_T0'] = tien_da_ung_T0['increase_T0'] - tien_da_ung_T0['decrease_T0']

    final_table = pd.concat([
        value_T2,
        value_T1,
        value_T0,
        hoan_tra_uttb_T0_table,
        tien_da_ung_T2,
        tien_da_ung_T1,
        tien_da_ung_T0
    ], axis=1)
    final_table.fillna(0, inplace=True)
    final_table.insert(0, 'account_code', account['account_code'])
    final_table.insert(1, 'customer_name', account['customer_name'])

    six = final_table['value_T1'] - final_table['fee_T1'] - final_table['tax_of_selling_T1'] - final_table[
        'tax_of_share_dividend_T1']
    seven = final_table['value_T0'] - final_table['fee_T0'] - final_table['tax_of_selling_T0'] - final_table[
        'tax_of_share_dividend_T0']
    ten = final_table['increase_T1'] + final_table['decrease_T1'] + final_table['increase_T0'] + final_table[
        'decrease_T0']
    # Tính cột 'Tổng giá trị tiền bán có thể ứng'
    final_table['sum_value_selling_co_the_ung'] = six + seven
    # Tính cột 'Tổng tiền còn có thể ứng'
    final_table['sum_tien_con_co_the_ung'] = final_table['sum_value_selling_co_the_ung'] - ten
    # Xử lý cột bất thường
    d = final_table['value_T2'] - final_table['fee_T2'] - final_table['tax_of_selling_T2'] - final_table[
        'tax_of_share_dividend_T2']
    final_table.loc[final_table['hoan_tra_UTTB_T0'] > d, 'bat_thuong'] = 'A'
    final_table.loc[ten > final_table['sum_value_selling_co_the_ung'], 'bat_thuong'] = 'B'
    final_table.loc[final_table['sum_tien_con_co_the_ung'] < 0, 'bat_thuong'] = 'C'
    count_bat_thuong = final_table['bat_thuong'].value_counts()
    # fillna giá trị của cột bất thường
    final_table.fillna('', inplace=True)
    final_table = final_table.sort_values(by=['account_code'])

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File Excel ---------------------
    # Write file BÁO CÁO ĐỐI CHIẾU UTTB
    footer_date = bdate(end_date, 1).split('-')
    eod_f_name = dt.datetime.strptime(end_date, "%Y/%m/%d").strftime("%d-%m-%Y")
    f_name = f'Đối chiếu UTTB {eod_f_name}.xlsx'
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
            'bold': True,
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
    sum_name_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 12,
            'font_name': 'Times New Roman',
        }
    )
    sum_money_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'right',
            'valign': 'vcenter',
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
        'Tổng tiền còn có thể ứng',
        'Bất Thường'
    ]
    sub_headers_1 = [
        'Giá trị tiền bán',
        'Phí bán',
        'Thuế bán',
        'Thuế Cổ Tức'
    ]
    sub_headers_2 = [
        'Số tiền UTTB KH đã nhận ngày T-2',
        'Phí UTTB ngày T-2',
        'Số tiền UTTB KH đã nhận ngày T-1',
        'Phí UTTB ngày T-1',
        'Số tiền UTTB KH đã nhận ngày T0',
        'Phí UTTB ngày T0'
    ]
    companyAddress = 'Tầng 3, CR3-03A, 109 Tôn Dật Tiên, phường Tân Phú, Quận 7, Thành phố Hồ Chí Minh'
    sheet_title_name = 'BÁO CÁO ĐỐI CHIẾU UTTB'
    eod_sub = dt.datetime.strptime(end_date, "%Y/%m/%d").strftime("%d/%m/%Y")
    sub_title_name = f'Ngày {eod_sub}'

    # --------- sheet BAO CAO CAN LAM ---------
    sheet_bao_cao_can_lam = workbook.add_worksheet('BAO CAO CAN LAM')

    # Insert phu hung picture
    sheet_bao_cao_can_lam.insert_image('A1', './img/phu_hung.png', {'x_scale': 0.65, 'y_scale': 0.71})

    # Set Column Width and Row Height
    sheet_bao_cao_can_lam.set_column('A:A', 8.43)
    sheet_bao_cao_can_lam.set_column('B:B', 10.14)
    sheet_bao_cao_can_lam.set_column('C:C', 11.14)
    sheet_bao_cao_can_lam.set_column('D:D', 18.43)
    sheet_bao_cao_can_lam.set_column('E:E', 24.14)
    sheet_bao_cao_can_lam.set_column('F:F', 13.71)
    sheet_bao_cao_can_lam.set_column('G:G', 12.43)
    sheet_bao_cao_can_lam.set_column('H:H', 12.29)
    sheet_bao_cao_can_lam.set_column('I:I', 13.14)
    sheet_bao_cao_can_lam.set_column('J:J', 15.14)
    sheet_bao_cao_can_lam.set_column('K:K', 11.43)
    sheet_bao_cao_can_lam.set_column('L:L', 13.14)
    sheet_bao_cao_can_lam.set_column('M:M', 11.71)
    sheet_bao_cao_can_lam.set_column('N:Q', 8.43)
    sheet_bao_cao_can_lam.set_column('R:R', 16.29)
    sheet_bao_cao_can_lam.set_column('S:S', 14.86)
    sheet_bao_cao_can_lam.set_column('T:T', 8.43)
    sheet_bao_cao_can_lam.set_column('U:U', 13.86)
    sheet_bao_cao_can_lam.set_column('V:V', 8.43)
    sheet_bao_cao_can_lam.set_column('W:W', 8.43)
    sheet_bao_cao_can_lam.set_column('X:X', 8.43)
    sheet_bao_cao_can_lam.set_column('Y:Y', 16.29)
    sheet_bao_cao_can_lam.set_column('Z:Z', 8.43)
    # merge row
    sheet_bao_cao_can_lam.merge_range('D1:M1', CompanyName, company_name_format)
    sheet_bao_cao_can_lam.merge_range('D2:M2', companyAddress, company_format)
    sheet_bao_cao_can_lam.merge_range('D3:M3', CompanyPhoneNumber, company_format)
    sheet_bao_cao_can_lam.merge_range('A7:Z7', sheet_title_name, sheet_title_format)
    sheet_bao_cao_can_lam.merge_range('A8:Z8', sub_title_name, from_to_format)
    sheet_bao_cao_can_lam.merge_range('A11:A12', headers[0], headers_format)
    sheet_bao_cao_can_lam.merge_range('B11:B12', headers[1], headers_format)
    sheet_bao_cao_can_lam.merge_range('C11:C12', headers[2], headers_format)
    sheet_bao_cao_can_lam.merge_range('D11:D12', headers[3], headers_format)
    sheet_bao_cao_can_lam.merge_range('E11:H11', headers[4], headers_format)
    sheet_bao_cao_can_lam.merge_range('I11:L11', headers[5], headers_format)
    sheet_bao_cao_can_lam.merge_range('M11:P11', headers[6], headers_format)
    sheet_bao_cao_can_lam.merge_range('Q11:Q12', headers[7], headers_format)
    sheet_bao_cao_can_lam.merge_range('R11:R12', headers[8], headers_format)
    sheet_bao_cao_can_lam.merge_range('S11:X11', headers[9], headers_format)
    sheet_bao_cao_can_lam.merge_range('Y11:Y12', headers[10], headers_format)
    sheet_bao_cao_can_lam.merge_range('Z11:Z12', headers[11], headers_format)
    sum_start_row = final_table.shape[0] + 14
    sheet_bao_cao_can_lam.merge_range(
        f'A{sum_start_row}:D{sum_start_row}',
        'Tổng',
        sum_name_format
    )
    footer_start_row = sum_start_row + 2
    sheet_bao_cao_can_lam.merge_range(
        f'U{footer_start_row}:Z{footer_start_row}',
        f'Ngày {footer_date[2]} tháng {footer_date[1]} năm {footer_date[0]}',
        footer_dmy_format
    )
    sheet_bao_cao_can_lam.merge_range(
        f'U{footer_start_row + 1}:Z{footer_start_row + 1}',
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
        [''] * (len(headers) + len(sub_headers_1) * 3 + 2),
        empty_row_format
    )
    sheet_bao_cao_can_lam.write_row(
        'E12',
        sub_headers_1 * 3,
        headers_format,
    )
    sheet_bao_cao_can_lam.write_row(
        'S12',
        sub_headers_2,
        headers_format,
    )
    sheet_bao_cao_can_lam.write_row(
        'A13',
        [f'({i})' for i in range(1, 5)],
        stt_row_format
    )
    sheet_bao_cao_can_lam.write_row(
        'E13',
        [f'(5{alpha})' for alpha in string.ascii_lowercase[:4]],
        stt_row_format
    )
    sheet_bao_cao_can_lam.write_row(
        'I13',
        [f'(6{alpha})' for alpha in string.ascii_lowercase[:4]],
        stt_row_format
    )
    sheet_bao_cao_can_lam.write_row(
        'M13',
        [f'(7{alpha})' for alpha in string.ascii_lowercase[:4]],
        stt_row_format
    )
    sheet_bao_cao_can_lam.write_row(
        'Q13',
        [f'({i})' for i in range(8, 10)],
        stt_row_format
    )
    sheet_bao_cao_can_lam.write_row(
        'S13',
        [f'(10{alpha})' for alpha in string.ascii_lowercase[:6]],
        stt_row_format
    )
    sheet_bao_cao_can_lam.write_row(
        'Y13',
        [f'({i})' for i in range(11, 13)],
        stt_row_format
    )
    sheet_bao_cao_can_lam.write_column(
        'A14',
        [f'{i}' for i in np.arange(final_table.shape[0]) + 1],
        stt_col_format
    )
    sheet_bao_cao_can_lam.write_column(
        'B14',
        final_table['account_code'],
        text_left_format
    )
    sheet_bao_cao_can_lam.write_column(
        'C14',
        final_table.index,
        text_left_format
    )
    sheet_bao_cao_can_lam.write_column(
        'D14',
        final_table['customer_name'],
        text_left_format
    )
    sheet_bao_cao_can_lam.write_column(
        'E14',
        final_table['value_T2'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'F14',
        final_table['fee_T2'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'G14',
        final_table['tax_of_selling_T2'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'H14',
        final_table['tax_of_share_dividend_T2'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'I14',
        final_table['value_T1'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'J14',
        final_table['fee_T1'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'K14',
        final_table['tax_of_selling_T1'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'L14',
        final_table['tax_of_share_dividend_T1'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'M14',
        final_table['value_T0'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'N14',
        final_table['fee_T0'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'O14',
        final_table['tax_of_selling_T0'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'P14',
        final_table['tax_of_share_dividend_T0'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'Q14',
        final_table['hoan_tra_UTTB_T0'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'R14',
        final_table['sum_value_selling_co_the_ung'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'S14',
        final_table['increase_T2'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'T14',
        final_table['decrease_T2'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'U14',
        final_table['increase_T1'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'V14',
        final_table['decrease_T1'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'W14',
        final_table['increase_T0'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'X14',
        final_table['decrease_T0'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'Y14',
        final_table['sum_tien_con_co_the_ung'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'Z14',
        final_table['bat_thuong'],
        money_format
    )
    sheet_bao_cao_can_lam.write(
        f'E{sum_start_row}',
        final_table['value_T2'].sum(),
        sum_money_format
    )
    sheet_bao_cao_can_lam.write(
        f'F{sum_start_row}',
        final_table['fee_T2'].sum(),
        sum_money_format
    )
    sheet_bao_cao_can_lam.write(
        f'G{sum_start_row}',
        final_table['tax_of_selling_T2'].sum(),
        sum_money_format
    )
    sheet_bao_cao_can_lam.write(
        f'H{sum_start_row}',
        final_table['tax_of_share_dividend_T2'].sum(),
        sum_money_format
    )
    sheet_bao_cao_can_lam.write(
        f'I{sum_start_row}',
        final_table['value_T1'].sum(),
        sum_money_format
    )
    sheet_bao_cao_can_lam.write(
        f'J{sum_start_row}',
        final_table['fee_T1'].sum(),
        sum_money_format
    )
    sheet_bao_cao_can_lam.write(
        f'K{sum_start_row}',
        final_table['tax_of_selling_T1'].sum(),
        sum_money_format
    )
    sheet_bao_cao_can_lam.write(
        f'L{sum_start_row}',
        final_table['tax_of_share_dividend_T1'].sum(),
        sum_money_format
    )
    sheet_bao_cao_can_lam.write(
        f'M{sum_start_row}',
        final_table['value_T0'].sum(),
        sum_money_format
    )
    sheet_bao_cao_can_lam.write(
        f'N{sum_start_row}',
        final_table['fee_T0'].sum(),
        sum_money_format
    )
    sheet_bao_cao_can_lam.write(
        f'O{sum_start_row}',
        final_table['tax_of_selling_T0'].sum(),
        sum_money_format
    )
    sheet_bao_cao_can_lam.write(
        f'P{sum_start_row}',
        final_table['tax_of_share_dividend_T0'].sum(),
        sum_money_format
    )
    sheet_bao_cao_can_lam.write(
        f'Q{sum_start_row}',
        final_table['hoan_tra_UTTB_T0'].sum(),
        sum_money_format
    )
    sheet_bao_cao_can_lam.write(
        f'R{sum_start_row}',
        final_table['sum_value_selling_co_the_ung'].sum(),
        sum_money_format
    )
    sheet_bao_cao_can_lam.write(
        f'S{sum_start_row}',
        final_table['increase_T2'].sum(),
        sum_money_format
    )
    sheet_bao_cao_can_lam.write(
        f'T{sum_start_row}',
        final_table['decrease_T2'].sum(),
        sum_money_format
    )
    sheet_bao_cao_can_lam.write(
        f'U{sum_start_row}',
        final_table['increase_T1'].sum(),
        sum_money_format
    )
    sheet_bao_cao_can_lam.write(
        f'V{sum_start_row}',
        final_table['decrease_T1'].sum(),
        sum_money_format
    )
    sheet_bao_cao_can_lam.write(
        f'W{sum_start_row}',
        final_table['increase_T0'].sum(),
        sum_money_format
    )
    sheet_bao_cao_can_lam.write(
        f'X{sum_start_row}',
        final_table['decrease_T0'].sum(),
        sum_money_format
    )
    sheet_bao_cao_can_lam.write(
        f'Y{sum_start_row}',
        final_table['sum_tien_con_co_the_ung'].sum(),
        sum_money_format
    )
    sheet_bao_cao_can_lam.write(
        f'Z{sum_start_row}',
        count_bat_thuong.sum(),
        sum_money_format
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
