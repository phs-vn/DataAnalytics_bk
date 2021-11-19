from reporting_tool.trading_service.thanhtoanbutru import *


def run(
        periodicity: str,
        run_time=None,
):
    start = time.time()
    info = get_info(periodicity, run_time)
    # start_date = info['start_date']
    # end_date = info['end_date']
    start_date = '2021/11/10'
    end_date = start_date
    period = info['period']
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder, folder_name)):  # dept_folder from import
        os.mkdir(join(dept_folder, folder_name))
    if not os.path.isdir(join(dept_folder, folder_name, period)):
        os.mkdir((join(dept_folder, folder_name, period)))

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viết Query ---------------------
    # query file Tang Tien
    increase_money = pd.read_sql(
        f"""
            SELECT 
            cashflow_bank.bank, 
            account.customer_name, 
            cashflow_bank.sub_account, 
            account.account_code, 
            cashflow_bank.bank_account, 
            cashflow_bank.outflow_amount, 
            cashflow_bank.remark
            FROM cashflow_bank 
            LEFT JOIN sub_account ON cashflow_bank.sub_account = sub_account.sub_account
            LEFT JOIN account ON sub_account.account_code = account.account_code
            WHERE (cashflow_bank.bank IN ('EIB','OCB')) 
            AND (date BETWEEN '{start_date}' AND '{end_date}') 
            AND (cashflow_bank.remark LIKE '%CT sang TKCN%')
            ORDER BY sub_account, date
            """,
        connect_DWH_CoSo,
    )
    increase_money_eib = increase_money.loc[increase_money['bank'] == 'EIB']
    increase_money_ocb = increase_money.loc[increase_money['bank'] == 'OCB']

    # query file Giam Tien
    decrease_money = pd.read_sql(
        f"""
            SELECT 
            cashflow_bank.bank, 
            account.customer_name, 
            cashflow_bank.sub_account, 
            account.account_code, 
            cashflow_bank.bank_account, 
            cashflow_bank.inflow_amount, 
            cashflow_bank.remark
            FROM cashflow_bank 
            LEFT JOIN sub_account ON cashflow_bank.sub_account = sub_account.sub_account
            LEFT JOIN account ON sub_account.account_code = account.account_code
            WHERE (cashflow_bank.bank IN ('EIB','OCB')) 
            AND (date BETWEEN '{start_date}' AND '{end_date}') 
            AND (cashflow_bank.remark LIKE '%Trich tien TKCN%')
            ORDER BY sub_account, date
        """,
        connect_DWH_CoSo,
    )
    decrease_money_eib = decrease_money.loc[decrease_money['bank'] == 'EIB']
    decrease_money_ocb = decrease_money.loc[decrease_money['bank'] == 'OCB']

    # query Nhap Rut Tien
    nhap_rut = pd.read_sql(
        f"""
            SELECT 
            money_in_out_transfer.date, 
            money_in_out_transfer.time, 
            sub_account.account_code, 
            money_in_out_transfer.sub_account, 
            bank_account, 
            bank, 
            account.customer_name, 
            transaction_id, 
            amount, 
            remark, 
            branch.branch_name, 
            money_in_out_transfer.status
            FROM money_in_out_transfer 
            LEFT JOIN sub_account ON money_in_out_transfer.sub_account = sub_account.sub_account
            LEFT JOIN account ON sub_account.account_code = account.account_code
            LEFT JOIN relationship ON  money_in_out_transfer.sub_account = relationship.sub_account
            LEFT JOIN branch ON relationship.branch_id = branch.branch_id
            WHERE (money_in_out_transfer.bank IN ('EIB','OCB')) 
            AND (money_in_out_transfer.transaction_id IN ('6692', '6693')) 
            AND (money_in_out_transfer.date BETWEEN '{start_date}' AND '{end_date}') 
            AND (relationship.date='2021') 
        """,
        connect_DWH_CoSo,
    )

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File ---------------------
    # Write file excel Bao cao doi chieu file ngan hang
    # date_file_name = dt.strptime(period, "%Y.%m.%d").date() - timedelta(days=1)
    file_name = f'Báo cáo đối chiếu file ngân hàng {period}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder, folder_name, period, file_name),
        engine='xlsxwriter',
        engine_kwargs={'options': {'nan_inf_to_errors': True}}
    )

    workbook = writer.book

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet sheet ---------------------
    # Format
    header_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 12
        }
    )
    stt_below_header_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 10
        }
    )
    available_content_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 12,
        }
    )
    ocb_name_format = workbook.add_format(
        {
            'align': 'left',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 10,
            'bg_color': '#00B050'
        }
    )
    eib_name_format = workbook.add_format(
        {
            'align': 'left',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 10,
            'bg_color': '#FFFF00'
        }
    )
    sum_format = workbook.add_format(
        {
            'align': 'right',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 10,
            'num_format': '#,##0',
            'bg_color': '#FFFF00'
        }
    )
    sum_format_hom_truoc_hom_nay = workbook.add_format(
        {
            'bold': True,
            'text_wrap': True,
            'align': 'right',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 10,
            'num_format': '#,##0',
        }
    )
    text_left_format = workbook.add_format(
        {
            'border': 1,
            'text_wrap': True,
            'align': 'left',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 10
        }
    )
    text_right_format = workbook.add_format(
        {
            'border': 1,
            'text_wrap': True,
            'align': 'right',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 10
        }
    )
    text_right_not_border_format = workbook.add_format(
        {
            'text_wrap': True,
            'align': 'right',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 10
        }
    )
    text_center_format = workbook.add_format(
        {
            'border': 1,
            'text_wrap': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 10
        }
    )
    # Format lấy từ file ngân hàng
    account_format = workbook.add_format(
        {
            'border': 1,
            'num_format': '0',
            'text_wrap': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 10
        }
    )

    num_format = workbook.add_format(
        {
            'border': 1,
            'num_format': '#,##0',
            'text_wrap': True,
            'align': 'right',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 10
        }
    )
    date_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'text_wrap': True,
            'valign': 'vcenter',
            'num_format': 'dd/mm/yyyy',
            'font_name': 'Times New Roman',
            'font_size': 10
        }
    )
    time_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'num_format': 'h:m:s',
            'font_name': 'Times New Roman',
            'font_size': 10
        }
    )
    headers = [
        'STT',
        'Khách hàng',
        'Mã tiểu khoản',
        'Số lưu ký',
        'Số TK ngân hàng',
        'Số tiền',
        'Diễn giải',
        'Ghi chú',
    ]
    headers_nhap_rut = [
        'STT',
        'Ngày thực hiện',
        'Thời gian',
        'Tài khoản',
        'Tiểu khoản',
        'Mã TK ngân hàng',
        'Tên ngân hàng',
        'Họ tên khách hàng',
        'Mã giao dịch',
        'Tổng tiền',
        'Diễn giải',
        'Người nhập',
        'Người duyệt',
        'Chi nhánh',
        'Trạng thái',
    ]
    headers_hom_truoc = [
        'STT',
        'TXDATE',
        'EFFECTIVE DATE',
        'BANKCODE',
        'ACCOUNT',
        'ACCOUNT NAME',
        'BALANCE',
        'NHẬP TIỀN',
        'RÚT TIỀN',
        'File tăng tiền',
        'File giảm tiền',
        'Số dư tiền dự kiến cuối ngày',
        'File NH gửi',
        'Lệch'
    ]
    headers_hom_nay = [
        'STT',
        'TXDATE',
        'BANKCODE',
        'ACCOUNT',
        'ACCOUNT NAME',
        'BALANCE',
        'DỰ KIẾN CUỐI NGÀY',
        'LỆCH'
    ]

    ################################################

    # Sheet File Tang Tien
    sheet_tang_tien = workbook.add_worksheet('File tăng tiền')

    sheet_tang_tien.set_column('A:B', 8)
    sheet_tang_tien.set_column('C:C', 26)
    sheet_tang_tien.set_column('D:E', 10)
    sheet_tang_tien.set_column('F:F', 17)
    sheet_tang_tien.set_column('G:G', 12)
    sheet_tang_tien.set_column('H:H', 33)
    sheet_tang_tien.set_column('I:I', 8)
    sheet_tang_tien.set_row(1, 29)
    sheet_tang_tien.set_row(2, 15)

    sheet_tang_tien.write_row('B2', headers, header_format)

    sheet_tang_tien.write_row(
        'B3',
        [f'({i})' for i in np.arange(len(headers)) + 1],
        stt_below_header_format,
    )

    # -------- eib --------
    sheet_tang_tien.write_column(
        'B4',
        [f'{i}' for i in np.arange(increase_money_eib.shape[0]) + 1],
        text_center_format,
    )
    sheet_tang_tien.write_column(
        'A4',
        increase_money_eib['bank'],
        eib_name_format,
    )
    sheet_tang_tien.write_column(
        'C4',
        increase_money_eib['customer_name'],
        text_left_format,
    )
    sheet_tang_tien.write_column(
        'D4',
        increase_money_eib['sub_account'],
        text_center_format,
    )
    sheet_tang_tien.write_column(
        'E4',
        increase_money_eib['account_code'],
        text_center_format,
    )
    sheet_tang_tien.write_column(
        'F4',
        increase_money_eib['bank_account'],
        text_center_format,
    )
    sheet_tang_tien.write_column(
        'G4',
        increase_money_eib['outflow_amount'],
        num_format,
    )
    sheet_tang_tien.write_column(
        'H4',
        increase_money_eib['remark'],
        text_left_format,
    )
    sheet_tang_tien.write_column(
        'I4',
        [''] * increase_money_eib.shape[0],
        text_left_format,
    )
    eib_sum_row = increase_money_eib.shape[0] + 4
    sheet_tang_tien.write(
        f'G{eib_sum_row}',
        increase_money_eib['outflow_amount'].sum(),
        sum_format,
    )

    # -------- ocb --------
    ocb_start_row = increase_money_eib.shape[0] + 5

    sheet_tang_tien.write_column(
        f'B{ocb_start_row}',
        [f'{i}' for i in np.arange(increase_money_ocb.shape[0]) + 1],
        text_center_format,
    )
    sheet_tang_tien.write_column(
        f'A{ocb_start_row}',
        increase_money_ocb['bank'],
        ocb_name_format,
    )
    sheet_tang_tien.write_column(
        f'C{ocb_start_row}',
        increase_money_ocb['customer_name'],
        text_left_format,
    )
    sheet_tang_tien.write_column(
        f'D{ocb_start_row}',
        increase_money_ocb['sub_account'],
        text_center_format,
    )
    sheet_tang_tien.write_column(
        f'E{ocb_start_row}',
        increase_money_ocb['account_code'],
        text_center_format,
    )
    sheet_tang_tien.write_column(
        f'F{ocb_start_row}',
        increase_money_ocb['bank_account'],
        text_center_format,
    )
    sheet_tang_tien.write_column(
        f'G{ocb_start_row}',
        increase_money_ocb['outflow_amount'],
        num_format,
    )
    sheet_tang_tien.write_column(
        f'H{ocb_start_row}',
        increase_money_ocb['remark'],
        text_left_format,
    )
    sheet_tang_tien.write_column(
        f'I{ocb_start_row}',
        [''] * increase_money_ocb.shape[0],
        text_left_format,
    )
    ocb_sum_row = ocb_start_row + increase_money_ocb.shape[0]
    sheet_tang_tien.write(
        f'G{ocb_sum_row}',
        increase_money_ocb['outflow_amount'].sum(),
        sum_format,
    )
    ################################################

    # Sheet file Giam Tien
    sheet_giam_tien = workbook.add_worksheet('File giảm tiền')

    sheet_giam_tien.set_column('A:B', 8)
    sheet_giam_tien.set_column('C:C', 25)
    sheet_giam_tien.set_column('D:E', 10)
    sheet_giam_tien.set_column('F:F', 14)
    sheet_giam_tien.set_column('G:G', 12)
    sheet_giam_tien.set_column('H:H', 53)
    sheet_giam_tien.set_column('I:I', 8)
    sheet_giam_tien.set_row(1, 29)
    sheet_giam_tien.set_row(2, 15)

    sheet_giam_tien.write_row('B2', headers, header_format)

    sheet_giam_tien.write_row(
        'B3',
        [f'{i}' for i in np.arange(len(headers)) + 1],
        header_format,
    )

    # -------- eib --------
    sheet_giam_tien.write_column(
        'B4',
        [f'{i}' for i in np.arange(decrease_money_eib.shape[0]) + 1],
        text_center_format,
    )
    sheet_giam_tien.write_column(
        'A4',
        decrease_money_eib['bank'],
        eib_name_format,
    )
    sheet_giam_tien.write_column(
        'C4',
        decrease_money_eib['customer_name'],
        text_left_format,
    )
    sheet_giam_tien.write_column(
        'D4',
        decrease_money_eib['sub_account'],
        text_center_format,
    )
    sheet_giam_tien.write_column(
        'E4',
        decrease_money_eib['account_code'],
        text_center_format,
    )
    sheet_giam_tien.write_column(
        'F4',
        decrease_money_eib['bank_account'],
        text_center_format,
    )
    sheet_giam_tien.write_column(
        'G4',
        decrease_money_eib['inflow_amount'],
        num_format,
    )
    sheet_giam_tien.write_column(
        'H4',
        decrease_money_eib['remark'],
        text_left_format,
    )
    sheet_giam_tien.write_column(
        'I4',
        [''] * decrease_money_eib.shape[0],
        text_left_format,
    )
    eib_sum_row = decrease_money_eib.shape[0] + 4
    sheet_giam_tien.write(
        f'G{eib_sum_row}',
        decrease_money_eib['inflow_amount'].sum(),
        sum_format,
    )

    # -------- ocb --------
    ocb_start_row = decrease_money_eib.shape[0] + 5
    ocb_sum_row = ocb_start_row + decrease_money_ocb.shape[0]
    sheet_giam_tien.write_column(
        f'B{ocb_start_row}',
        [f'{i}' for i in np.arange(decrease_money_ocb.shape[0]) + 1],
        text_center_format,
    )
    sheet_giam_tien.write_column(
        f'A{ocb_start_row}',
        decrease_money_ocb['bank'],
        ocb_name_format,
    )
    sheet_giam_tien.write_column(
        f'C{ocb_start_row}',
        decrease_money_ocb['customer_name'],
        text_left_format,
    )
    sheet_giam_tien.write_column(
        f'D{ocb_start_row}',
        decrease_money_ocb['sub_account'],
        text_center_format,
    )
    sheet_giam_tien.write_column(
        f'E{ocb_start_row}',
        decrease_money_ocb['account_code'],
        text_center_format,
    )
    sheet_giam_tien.write_column(
        f'F{ocb_start_row}',
        decrease_money_ocb['bank_account'],
        text_center_format,
    )
    sheet_giam_tien.write_column(
        f'G{ocb_start_row}',
        decrease_money_ocb['inflow_amount'],
        num_format,
    )
    sheet_giam_tien.write_column(
        f'H{ocb_start_row}',
        decrease_money_ocb['remark'],
        text_left_format,
    )
    sheet_giam_tien.write_column(
        f'I{ocb_start_row}',
        [''] * decrease_money_ocb.shape[0],
        text_left_format,
    )
    sheet_giam_tien.write(
        f'G{ocb_sum_row}',
        decrease_money_ocb['inflow_amount'].sum(),
        sum_format,
    )
    ################################################

    # Sheet file Nhap Rut Tien
    sheet_nhap_rut = workbook.add_worksheet('Nhập Rút Tiền')

    sheet_nhap_rut.set_column('A:A', 42)
    sheet_nhap_rut.set_column('B:B', 37)
    sheet_nhap_rut.set_column('C:C', 30)
    sheet_nhap_rut.set_column('D:E', 10)
    sheet_nhap_rut.set_column('F:F', 16)
    sheet_nhap_rut.set_column('G:G', 10)
    sheet_nhap_rut.set_column('H:H', 23)
    sheet_nhap_rut.set_column('I:J', 11)
    sheet_nhap_rut.set_column('K:K', 30)
    sheet_nhap_rut.set_column('L:N', 11)
    sheet_nhap_rut.set_column('O:O', 9)
    sheet_nhap_rut.set_row(1, 47)
    sheet_nhap_rut.set_row(2, 32)

    sheet_nhap_rut.write('A1', 'CÔNG TY CỔ PHẦN CHỨNG KHOÁN PHÚ HƯNG', available_content_format)
    sheet_nhap_rut.write('B1',
                         'Tầng 3, CR3-03A, 109 Tôn Dật Tiên, phường Tân Phú, Quận 7, Thành phố Hồ Chí Minh',
                         available_content_format)
    sheet_nhap_rut.write('C1',
                         'Điện thoại: (84-28) 5 413 5479 Fax: (84-28) 5 413 5472',
                         available_content_format)
    sheet_nhap_rut.write('A2', 'DANH SÁCH NHẬP RÚT CHUYỂN TIỀN CỦA KHÁCH HÀNG', available_content_format)

    start_date_nhap_rut = dt.strptime(start_date, '%Y/%m/%d').strftime("%d/%m/%y")
    end_date_nhap_rut = dt.strptime(end_date, '%Y/%m/%d').strftime("%d/%m/%y")

    sheet_nhap_rut.write('B2',
                         f'Từ ngày {start_date_nhap_rut} đến ngày {end_date_nhap_rut}',
                         available_content_format)

    sheet_nhap_rut.write_row('A3', headers_nhap_rut, header_format)

    sheet_nhap_rut.write_column(
        'A4',
        [f'{i}' for i in np.arange(nhap_rut.shape[0]) + 1],
        text_center_format,
    )

    sheet_nhap_rut.write_column(
        'B4',
        nhap_rut['date'],
        date_format,
    )
    sheet_nhap_rut.write_column(
        'C4',
        nhap_rut['time'],
        time_format,
    )
    sheet_nhap_rut.write_column(
        'D4',
        nhap_rut['account_code'],
        text_center_format,
    )
    sheet_nhap_rut.write_column(
        'E4',
        nhap_rut['sub_account'],
        text_center_format,
    )
    sheet_nhap_rut.write_column(
        'F4',
        nhap_rut['bank_account'],
        text_center_format,
    )
    sheet_nhap_rut.write_column(
        'G4',
        nhap_rut['bank'],
        text_center_format,
    )
    sheet_nhap_rut.write_column(
        'H4',
        nhap_rut['customer_name'],
        text_center_format,
    )
    sheet_nhap_rut.write_column(
        'I4',
        nhap_rut['transaction_id'],
        text_right_format,
    )
    sheet_nhap_rut.write_column(
        'J4',
        nhap_rut['amount'],
        num_format,
    )
    sheet_nhap_rut.write_column(
        'K4',
        nhap_rut['remark'],
        text_left_format,
    )
    sheet_nhap_rut.write_column(
        'L4',
        [''] * nhap_rut.shape[0],
        text_left_format,
    )
    sheet_nhap_rut.write_column(
        'M4',
        [''] * nhap_rut.shape[0],
        text_left_format,
    )
    sheet_nhap_rut.write_column(
        'N4',
        nhap_rut['branch_name'],
        text_left_format,
    )
    sheet_nhap_rut.write_column(
        'O4',
        nhap_rut['status'],
        text_center_format,
    )
    ################################################

    # Sheet file Ngay hom truoc
    sheet_hom_truoc = workbook.add_worksheet('Ngày hôm trước')

    # read File Ngan Hang
    bank_file_eib = pd.read_excel(join(dept_folder, 'Attachment', period, 'EIB', 'SD CUOI N.10.11 - EIB.xls'),
                                  sheet_name=0, header=5,
                                  dtype={
                                      'SECURITY CODE': object,
                                      'SOL ID': object,
                                      'CIF ID': object,
                                      'FORACID': object
                                  })

    bank_file_ocb = pd.read_excel(join(dept_folder, 'Attachment', period, 'OCB', 'SỐ DƯ 10.11 - OCB.xlsx'),
                                  sheet_name=0, header=0,
                                  dtype={
                                      'STT': object,
                                      'TÀI KHOẢN': object,
                                      'STK': object
                                  })

    sheet_hom_truoc.set_column('A:A', 5)
    sheet_hom_truoc.set_column('B:B', 10)
    sheet_hom_truoc.set_column('C:C', 13)
    sheet_hom_truoc.set_column('D:D', 18)
    sheet_hom_truoc.set_column('E:E', 17)
    sheet_hom_truoc.set_column('F:F', 30)
    sheet_hom_truoc.set_column('G:G', 11)
    sheet_hom_truoc.set_column('H:N', 15)

    sheet_hom_truoc.write('H1', '6692', text_right_not_border_format)
    sheet_hom_truoc.write('I1', '6693', text_right_not_border_format)

    sheet_hom_truoc.write_row('A2', headers_hom_truoc, header_format)

    sheet_hom_truoc.write_column(
        'A3',
        [f'{i}' for i in np.arange(bank_file_eib.shape[0] + bank_file_ocb.shape[0])],
        text_center_format,
    )

    today = dt.now().date()
    sheet_hom_truoc.write_column(
        'B3',
        [today] * bank_file_eib.shape[0],
        date_format
    )

    sheet_hom_truoc.write_column(
        'C3',
        [today + timedelta(days=1)] * bank_file_eib.shape[0],
        date_format,
    )

    sheet_hom_truoc.write_column(
        'D3',
        ['EIB'] * bank_file_eib.shape[0],
        text_left_format
    )

    sheet_hom_truoc.write_column(
        'E3',
        bank_file_eib['FORACID'],
        account_format,
    )

    sheet_hom_truoc.write_column(
        'F3',
        bank_file_eib['ACCOUNT NAME'],
        text_center_format
    )

    sheet_hom_truoc.write_column(
        'G3',
        bank_file_eib['ACCOUNT BALANCE'],
        num_format
    )

    bank_file_ocb_start_row = bank_file_eib.shape[0] + 3
    sheet_hom_truoc.write_column(
        f'B{bank_file_ocb_start_row}',
        [today] * bank_file_ocb.shape[0],
        date_format
    )

    sheet_hom_truoc.write_column(
        f'C{bank_file_ocb_start_row}',
        [today + timedelta(days=1)] * bank_file_ocb.shape[0],
        date_format,
    )

    sheet_hom_truoc.write_column(
        f'D{bank_file_ocb_start_row}',
        ['OCB'] * bank_file_ocb.shape[0],
        text_left_format
    )

    sheet_hom_truoc.write_column(
        f'E{bank_file_ocb_start_row}',
        bank_file_ocb['TÀI KHOẢN'],
        account_format,
    )

    sheet_hom_truoc.write_column(
        f'F{bank_file_ocb_start_row}',
        bank_file_ocb['TÊN KHÁCH HÀNG'],
        text_center_format
    )

    bank_file_ocb['SỐ DƯ'].fillna(0, inplace=True)
    sheet_hom_truoc.write_column(
        f'G{bank_file_ocb_start_row}',
        bank_file_ocb['SỐ DƯ'],
        num_format
    )

    # Groupby trong sheet Nhap_Rut_Tien
    nhap_rut_groupby = nhap_rut.groupby(['bank_account', 'transaction_id'])['amount'].sum().unstack()
    # thêm column NHAPTIEN trong df
    bank_file_eib['NHAPTIEN'] = bank_file_eib['FORACID'].map(nhap_rut_groupby['6692'])
    bank_file_eib['NHAPTIEN'].fillna(0, inplace=True)  # thay thế NaN bằng giá trị 0
    bank_file_ocb['NHAPTIEN'] = bank_file_ocb['TÀI KHOẢN'].map(nhap_rut_groupby['6692'])
    bank_file_ocb['NHAPTIEN'].fillna(0, inplace=True)

    sheet_hom_truoc.write_column(
        'H3',
        bank_file_eib['NHAPTIEN'],
        num_format
    )
    sheet_hom_truoc.write_column(
        f'H{bank_file_ocb_start_row}',
        bank_file_ocb['NHAPTIEN'],
        num_format
    )

    bank_file_eib['RUTTIEN'] = bank_file_eib['FORACID'].map(nhap_rut_groupby['6693'])
    bank_file_eib['RUTTIEN'].fillna(0, inplace=True)
    bank_file_ocb['RUTTIEN'] = bank_file_ocb['TÀI KHOẢN'].map(nhap_rut_groupby['6693'])
    bank_file_ocb['RUTTIEN'].fillna(0, inplace=True)

    sheet_hom_truoc.write_column(
        'I3',
        bank_file_eib['RUTTIEN'],
        num_format
    )
    sheet_hom_truoc.write_column(
        f'I{bank_file_ocb_start_row}',
        bank_file_ocb['RUTTIEN'],
        num_format
    )

    # Groupby trong sheet File Tang Tien
    tang_tien_eib_groupby = increase_money_eib.groupby(['bank', 'bank_account'])['outflow_amount'].sum().unstack().T
    tang_tien_ocb_groupby = increase_money_ocb.groupby(['bank', 'bank_account'])['outflow_amount'].sum().unstack().T

    # them column TĂNG TIỀN trong df
    bank_file_eib['TANGTIEN'] = bank_file_eib['FORACID'].map(tang_tien_eib_groupby['EIB'])
    bank_file_eib['TANGTIEN'].fillna(0, inplace=True)
    bank_file_ocb['TANGTIEN'] = bank_file_ocb['TÀI KHOẢN'].map(tang_tien_ocb_groupby['OCB'])
    bank_file_ocb['TANGTIEN'].fillna(0, inplace=True)

    sheet_hom_truoc.write_column(
        'J3',
        bank_file_eib['TANGTIEN'],
        num_format
    )

    sheet_hom_truoc.write_column(
        f'J{bank_file_ocb_start_row}',
        bank_file_ocb['TANGTIEN'],
        num_format
    )

    # Groupby trong sheet File Giam Tien
    giam_tien_eib_groupby = decrease_money_eib.groupby(['bank', 'bank_account'])['inflow_amount'].sum().unstack().T
    giam_tien_ocb_groupby = decrease_money_ocb.groupby(['bank', 'bank_account'])['inflow_amount'].sum().unstack().T

    # them column GIẢM TIỀN trong df
    bank_file_eib['GIAMTIEN'] = bank_file_eib['FORACID'].map(giam_tien_eib_groupby['EIB'])
    bank_file_eib['GIAMTIEN'].fillna(0, inplace=True)
    bank_file_ocb['GIAMTIEN'] = bank_file_ocb['TÀI KHOẢN'].map(giam_tien_ocb_groupby['OCB'])
    bank_file_ocb['GIAMTIEN'].fillna(0, inplace=True)

    sheet_hom_truoc.write_column(
        'K3',
        bank_file_eib['GIAMTIEN'],
        num_format
    )

    sheet_hom_truoc.write_column(
        f'K{bank_file_ocb_start_row}',
        bank_file_ocb['GIAMTIEN'],
        num_format
    )

    # Tính Số tiền dự kiến cuối ngày
    bank_file_eib['DUKIENCUOINGAY'] = bank_file_eib['ACCOUNT BALANCE'] + bank_file_eib['NHAPTIEN'] - bank_file_eib['RUTTIEN'] + bank_file_eib['TANGTIEN'] - bank_file_eib['GIAMTIEN']
    bank_file_ocb['DUKIENCUOINGAY'] = bank_file_ocb['SỐ DƯ'] + bank_file_ocb['NHAPTIEN'] - bank_file_ocb['RUTTIEN'] + bank_file_ocb['TANGTIEN'] - bank_file_ocb['GIAMTIEN']

    sheet_hom_truoc.write_column(
        'L3',
        bank_file_eib['DUKIENCUOINGAY'],
        num_format
    )

    sheet_hom_truoc.write_column(
        f'L{bank_file_ocb_start_row}',
        bank_file_ocb['DUKIENCUOINGAY'],
        num_format
    )

    sheet_hom_truoc.write_column(
        'M3',
        bank_file_eib['ACCOUNT BALANCE'],
        num_format,
    )

    sheet_hom_truoc.write_column(
        f'M{bank_file_ocb_start_row}',
        bank_file_ocb['SỐ DƯ'],
        num_format
    )

    bank_file_eib['Lệch'] = bank_file_eib['DUKIENCUOINGAY'] - bank_file_eib['ACCOUNT BALANCE']
    bank_file_ocb['Lệch'] = bank_file_ocb['DUKIENCUOINGAY'] - bank_file_ocb['SỐ DƯ']

    sheet_hom_truoc.write_column(
        'N3',
        bank_file_eib['Lệch'],
        num_format
    )

    sheet_hom_truoc.write_column(
        f'N{bank_file_ocb_start_row}',
        bank_file_ocb['Lệch'],
        num_format
    )

    # Tính tổng các cột (Nhập tiền, rút tiền, file tăng tiền, file giảm tiền, số dư tiền dự kiến cuối ngày,File NH gửi, lệch
    hom_truoc_nay_sum_row = bank_file_eib.shape[0] + bank_file_ocb.shape[0] + 3

    sheet_hom_truoc.write(
        f'H{hom_truoc_nay_sum_row}',
        bank_file_eib['NHAPTIEN'].sum() + bank_file_ocb['NHAPTIEN'].sum(),
        sum_format_hom_truoc_hom_nay,
    )

    sheet_hom_truoc.write(
        f'I{hom_truoc_nay_sum_row}',
        bank_file_eib['RUTTIEN'].sum() + bank_file_ocb['RUTTIEN'].sum(),
        sum_format_hom_truoc_hom_nay,
    )

    sheet_hom_truoc.write(
        f'J{hom_truoc_nay_sum_row}',
        bank_file_eib['TANGTIEN'].sum() + bank_file_ocb['TANGTIEN'].sum(),
        sum_format_hom_truoc_hom_nay,
    )

    sheet_hom_truoc.write(
        f'K{hom_truoc_nay_sum_row}',
        bank_file_eib['GIAMTIEN'].sum() + bank_file_ocb['GIAMTIEN'].sum(),
        sum_format_hom_truoc_hom_nay,
    )

    sheet_hom_truoc.write(
        f'L{hom_truoc_nay_sum_row}',
        bank_file_eib['DUKIENCUOINGAY'].sum() + bank_file_ocb['DUKIENCUOINGAY'].sum(),
        sum_format_hom_truoc_hom_nay,
    )

    sheet_hom_truoc.write(
        f'M{hom_truoc_nay_sum_row}',
        bank_file_eib['ACCOUNT BALANCE'].sum() + bank_file_ocb['SỐ DƯ'].sum(),
        sum_format_hom_truoc_hom_nay,
    )

    sheet_hom_truoc.write(
        f'N{hom_truoc_nay_sum_row}',
        bank_file_eib['Lệch'].sum() + bank_file_ocb['Lệch'].sum(),
        sum_format_hom_truoc_hom_nay,
    )

    ################################################

    # Sheet file Ngay hom nay
    sheet_hom_nay = workbook.add_worksheet('Ngày hôm nay')

    sheet_hom_nay.set_column('A:A', 5)
    sheet_hom_nay.set_column('B:B', 10)
    sheet_hom_nay.set_column('C:C', 18)
    sheet_hom_nay.set_column('D:D', 17)
    sheet_hom_nay.set_column('E:E', 30)
    sheet_hom_nay.set_column('F:F', 11)
    sheet_hom_nay.set_column('G:G', 24)
    sheet_hom_nay.set_column('H:H', 15)

    sheet_hom_nay.write_row('A1', headers_hom_nay, header_format)

    sheet_hom_nay.write_column(
        'A2',
        [f'{i}' for i in np.arange(bank_file_eib.shape[0] + bank_file_ocb.shape[0])],
        text_center_format,
    )

    today = dt.now().date()
    sheet_hom_nay.write_column(
        'B2',
        [today] * bank_file_eib.shape[0],
        date_format
    )

    sheet_hom_nay.write_column(
        'C2',
        ['EIB'] * bank_file_eib.shape[0],
        text_left_format
    )

    sheet_hom_nay.write_column(
        'D2',
        bank_file_eib['FORACID'],
        account_format,
    )

    sheet_hom_nay.write_column(
        'E2',
        bank_file_eib['ACCOUNT NAME'],
        text_center_format
    )

    sheet_hom_nay.write_column(
        'F2',
        bank_file_eib['ACCOUNT BALANCE'],
        num_format
    )

    sheet_hom_nay.write_column(
        'G2',
        bank_file_eib['DUKIENCUOINGAY'],
        num_format
    )

    sheet_hom_nay.write_column(
        'H2',
        bank_file_eib['DUKIENCUOINGAY'] - bank_file_eib['ACCOUNT BALANCE'],
        num_format
    )

    bank_file_ocb_start_row = bank_file_eib.shape[0] + 2
    sheet_hom_nay.write_column(
        f'B{bank_file_ocb_start_row}',
        [today] * bank_file_ocb.shape[0],
        date_format
    )

    sheet_hom_nay.write_column(
        f'C{bank_file_ocb_start_row}',
        ['OCB'] * bank_file_ocb.shape[0],
        text_left_format
    )

    sheet_hom_nay.write_column(
        f'D{bank_file_ocb_start_row}',
        bank_file_ocb['TÀI KHOẢN'],
        account_format,
    )

    sheet_hom_nay.write_column(
        f'E{bank_file_ocb_start_row}',
        bank_file_ocb['TÊN KHÁCH HÀNG'],
        text_center_format
    )

    sheet_hom_nay.write_column(
        f'F{bank_file_ocb_start_row}',
        bank_file_ocb['SỐ DƯ'],
        num_format
    )

    sheet_hom_nay.write_column(
        f'G{bank_file_ocb_start_row}',
        bank_file_ocb['DUKIENCUOINGAY'],
        num_format
    )

    sheet_hom_nay.write_column(
        f'H{bank_file_ocb_start_row}',
        bank_file_ocb['DUKIENCUOINGAY'] - bank_file_ocb['SỐ DƯ'],
        num_format
    )

    sheet_hom_nay.write(
        f'G{hom_truoc_nay_sum_row}',
        bank_file_eib['DUKIENCUOINGAY'].sum() + bank_file_ocb['DUKIENCUOINGAY'].sum(),
        sum_format_hom_truoc_hom_nay,
    )

    sheet_hom_nay.write(
        f'H{hom_truoc_nay_sum_row}',
        bank_file_eib['Lệch'].sum() + bank_file_ocb['Lệch'].sum(),
        sum_format_hom_truoc_hom_nay,
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
