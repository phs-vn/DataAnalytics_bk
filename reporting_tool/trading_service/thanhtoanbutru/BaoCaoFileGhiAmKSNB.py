from reporting_tool.trading_service.thanhtoanbutru import *


def run(
        periodicity: str,
        start_date: str,  # 2021-07-01
        end_date: str,  # 2021-07-31
        run_time=None,
):
    start = time.time()
    info = get_info(periodicity, run_time)
    period = info['period']
    folder_name = info['folder_name']
    date_character = ['/', '-', '.']

    # create folder
    for date_char in date_character:
        if date_char in end_date:
            end_date.replace(date_char, '.')
            if not os.path.isdir(join(dept_folder, folder_name, period)):  # dept_folder from import
                os.mkdir(join(dept_folder, folder_name, period))

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viết Query ---------------------
    # query KSNB hàng tháng
    ksnb_6692_6693_query = pd.read_sql(
        f"""
            SELECT 
            money_in_out_transfer.date, 
            money_in_out_transfer.time, 
            account.account_code, 
            money_in_out_transfer.bank, 
            account.customer_name 
            FROM money_in_out_transfer
            LEFT JOIN sub_account ON money_in_out_transfer.sub_account = sub_account.sub_account
            LEFT JOIN account ON sub_account.account_code = account.account_code
            WHERE (money_in_out_transfer.bank IN ('EIB','OCB')) 
            AND (money_in_out_transfer.transaction_id IN ('6692', '6693')) 
            AND (money_in_out_transfer.date BETWEEN '{start_date}' AND '{end_date}')
            ORDER BY money_in_out_transfer.date, money_in_out_transfer.time ASC
        """,
        connect_DWH_CoSo
    )
    ksnb_1187_query = pd.read_sql(
        f"""
            SELECT 
            transactional_record.effective_date,
            money_in_out_transfer.time,
            transactional_record.transaction_id, 
            account.account_code,
            money_in_out_transfer.bank,
            account.customer_name
            FROM transactional_record
            LEFT JOIN employee ON employee.employee_id = transactional_record.approver
            LEFT JOIN department ON department.dept_id = employee.dept_id
            FULL OUTER JOIN money_in_out_transfer ON money_in_out_transfer.sub_account = transactional_record.sub_account
            LEFT JOIN sub_account ON sub_account.sub_account = transactional_record.sub_account
            LEFT JOIN account ON sub_account.account_code = account.account_code
            WHERE (transactional_record.effective_date BETWEEN '{start_date}' AND '{end_date}')
            AND (transactional_record.transaction_id = '1187')
            AND (transactional_record.approver = '0832' OR department.dept_name = 'Margin settlement')
            ORDER BY transactional_record.effective_date
        """,
        connect_DWH_CoSo
    )
    ksnb_query = pd.concat([ksnb_6692_6693_query, ksnb_1187_query])

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File Excel ---------------------
    # Write file excel Bao cao doi chieu file ngan hang
    file_name = f'BaoCao14.xlsx'
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
    sheet_title_vn_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'valign': 'vbottom',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 14
        }
    )
    sheet_title_eng_format = workbook.add_format(
        {
            'bold': True,
            'italic': True,
            'align': 'center',
            'valign': 'vbottom',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 14,
            'color': '#00B050'
        }
    )
    dia_diem_vn_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'valign': 'vbottom',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 12,
        }
    )
    dia_diem_eng_format = workbook.add_format(
        {
            'bold': True,
            'italic': True,
            'align': 'center',
            'valign': 'vbottom',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 12,
            'color': '#00B050'
        }
    )
    header_vn_format = workbook.add_format(
        {
            'bold': True,
            'border': 1,
            'align': 'center',
            'valign': 'vbottom',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 12,
        }
    )
    header_vn_format.set_bottom(0)
    header_eng_format = workbook.add_format(
        {
            'bold': True,
            'italic': True,
            'border': 1,
            'align': 'center',
            'valign': 'vbottom',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 12,
            'color': '#00B050'
        }
    )
    header_eng_format.set_top(0)
    text_center_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vbottom',
            'font_name': 'Times New Roman',
            'font_size': 12
        }
    )
    call_type_vn_left_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'top',
            'font_name': 'Times New Roman',
            'font_size': 11
        }
    )
    call_type_vn_left_format.set_bottom(0)
    call_type_eng_left_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'text_wrap': True,
            'valign': 'top',
            'font_name': 'Times New Roman',
            'font_size': 11
        }
    )
    call_type_eng_left_format.set_top(0)
    text_right_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vbottom',
            'font_name': 'Times New Roman',
            'font_size': 11
        }
    )
    sub_cont_tel_format = workbook.add_format(
        {
            'align': 'left',
            'valign': 'vbottom',
            'font_name': 'Times New Roman',
            'font_size': 11
        }
    )
    customer_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vbottom',
            'font_name': 'Times New Roman',
            'font_size': 12,
            'text_wrap': True
        }
    )
    time_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vbottom',
            'font_name': 'Times New Roman',
            'font_size': 12,
            'num_format': 'h:mm:ss'
        }
    )
    date_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vbottom',
            'font_name': 'Times New Roman',
            'font_size': 12,
            'num_format': 'm/d/yyyy'
        }
    )
    sub_cont_1_format = workbook.add_format(
        {
            'align': 'center',
            'valign': 'vbottom',
            'font_name': 'Times New Roman',
            'font_size': 12
        }
    )
    sub_cont_2_format = workbook.add_format(
        {
            'align': 'left',
            'valign': 'vbottom',
            'font_name': 'Times New Roman',
            'font_size': 12
        }
    )
    headers_vn = [
        'Loại cuộc gọi',
        'Stt',
        'Ngày',
        'Giờ',
        'Số tài khoản',
        '',
        'Tên Khách hàng',
        'Số điện thoại',
        ''
    ]
    headers_eng = [
        'Type of call',
        'No.',
        'Date',
        'Time',
        'Account number',
        '',
        'Name of customer',
        'Phone number',
        ''
    ]

    #  ----------------- Viết Sheet VIP PHS -----------------
    sheet1 = workbook.add_worksheet('Sheet1')
    # content in sheet
    sheet_title_vn = 'DANH SÁCH CUỘC GỌI ĐẾN KHÁCH HÀNG_THÁNG 07/2021 (01/07/2021- 31/07/2021)'
    sheet_title_eng = 'LIST OF PHONE CALLS TO CUSTOMERS__ July of 2021  (01/07/2021- 31/07/2021)'
    location_vn = 'Địa điểm: phòng TTBT- HỘI SỞ'
    location_eng = 'Location: TTBT department - HỘI SỞ'
    sub_cont_tel_eib = '39143152/38216082/39144080/939142155'
    sub_cont_tel_ocb = '36220104/ (028) 38302110'
    sub_cont_1 = 'HSBFCC4934'
    sub_cont_2 = 'CTBC VIETNAM EQUITY FUND'
    sub_cont_tel_CTBC = '02839155144/02462707601'
    content_type_of_call_vn = 'Xác nhận rút nộp ngân hàng liên kết'
    content_type_of_call_eng = 'Confirmation withdrawal-payment of affiliate bank'

    # Set Column Width and Row Height
    sheet1.set_column('A:A', 34.43)
    sheet1.set_column('B:B', 4.14)
    sheet1.set_column('C:C', 14.86)
    sheet1.set_column('D:D', 15.29)
    sheet1.set_column('E:E', 18.57)
    sheet1.set_column('F:F', 11.71)
    sheet1.set_column('G:G', 37.29)
    sheet1.set_column('H:H', 44.57)
    sheet1.set_column('I:I', 12.14)
    sheet1.set_column('J:J', 15.86)
    sheet1.set_column('K:K', 28)
    sheet1.set_column('L:L', 35.86)
    sheet1.set_column('M:M', 23.57)

    # Write value in each column, row
    # merge row
    sheet1.merge_range('A1:H1', sheet_title_vn, sheet_title_vn_format)
    sheet1.merge_range('A2:H2', sheet_title_eng, sheet_title_eng_format)
    sheet1.merge_range('A3:H3', location_vn, dia_diem_vn_format)
    sheet1.merge_range('A4:H4', location_eng, dia_diem_eng_format)
    sheet1.merge_range('K3:K4', 'EIB', sub_cont_tel_format)
    sheet1.merge_range('L3:L4', sub_cont_tel_eib, sub_cont_tel_format)
    sheet1.merge_range('K6:K7', sub_cont_1, sub_cont_1_format)
    sheet1.merge_range('L6:L7', sub_cont_2, sub_cont_2_format)
    sheet1.merge_range('M6:M7', sub_cont_tel_CTBC, sub_cont_tel_format)
    sheet1.write('K5', 'OCB', sub_cont_tel_format)
    sheet1.write('L5', sub_cont_tel_ocb, sub_cont_tel_format)

    sheet1.write('A8', content_type_of_call_vn, call_type_vn_left_format)
    sheet1.merge_range(f'A9:A{ksnb_query.shape[0] + 7}', content_type_of_call_eng, call_type_eng_left_format)
    # write row and write column
    sheet1.write_row('A6', headers_vn, header_vn_format)
    sheet1.write_row('A7', headers_eng, header_eng_format)

    sheet1.write_column(
        'B8',
        [int(i) for i in np.arange(ksnb_query.shape[0]) + 1],
        text_right_format
    )
    sheet1.write_column(
        'C8',
        ksnb_query['date'],
        date_format
    )
    sheet1.write_column(
        'D8',
        ksnb_query['time'],
        time_format
    )
    sheet1.write_column(
        'E8',
        ksnb_query['account_code'],
        text_center_format
    )
    sheet1.write_column(
        'F8',
        ksnb_query['bank'],
        text_center_format
    )
    sheet1.write_column(
        'G8',
        ksnb_query['customer_name'],
        customer_format
    )
    sheet1.write_column(
        'H8',
        ['(+84 28) 54135479 - Ext: 8168, 8169, 8145'] * ksnb_query.shape[0],
        text_center_format
    )
    sheet1.write_column(
        'I8',
        [''] * ksnb_query.shape[0],
        text_right_format
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
