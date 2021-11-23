"""
    1. daily
    2. table:
"""
from reporting_tool.trading_service.thanhtoanbutru import *


def run(
        periodicity: str,
        start_date: str,  # 2021-11-22
        end_date: str,    # 2021-11-23
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

    # --------------------- Viết Query ---------------------
    doi_chieu_so_du_query = pd.read_sql(
        f"""
            SELECT
            sub_account_deposit.date,
            relationship.account_code, 
            relationship.sub_account, 
            account.customer_name, 
            sub_account_deposit.opening_balance,
            sub_account_deposit.closing_balance
            FROM sub_account_deposit
            LEFT JOIN relationship ON relationship.sub_account = sub_account_deposit.sub_account
            LEFT JOIN account ON account.account_code = relationship.account_code
            WHERE sub_account_deposit.date BETWEEN '{start_date}' AND '{end_date}'
            AND relationship.date = '{end_date}'
            ORDER BY sub_account_deposit.date ASC
        """,
        connect_DWH_CoSo
    )

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File Excel ---------------------
    # Write file BÁO CÁO ĐỐI CHIẾU SỐ DƯ TIỀN TÀI KHOẢN KHÁCH HÀNG
    for date_char in date_character:
        if date_char in start_date and date_char in end_date:
            start_date = start_date.replace(date_char, '/')
            end_date = end_date.replace(date_char, '/')

    start_date = dt.datetime.strptime(start_date, "%Y/%m/%d").strftime("%d-%m")
    end_date = dt.datetime.strptime(end_date, "%Y/%m/%d").strftime("%d-%m")
    f_name = f'Báo cáo đối chiếu số dư tiền tài khoản khách hàng {end_date}.xlsx'
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
    text_right_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
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
    headers = [
        'STT',
        'Số tài khoản',
        'Số tiểu khoản',
        'Tên khách hàng',
        'Dữ liệu tại Flex',
        'Dữ liệu lưu ngoài Flex',
        'Bất thường',
        'Chi nhánh quản lý',
        'Nhân viên quản lý tài khoản'
    ]
    sub_headers = [
        'Số dư tiền đầu ngày T-1',
        'Số dư tiền cuối ngày T-1',
        'Số dư tiền đầu ngày T0',
        'Số dư tiền đầu ngày T-1',
        'Số dư tiền cuối ngày T-1'
    ]
    companyAddress = 'Tầng 3, CR3-03A, 109 Tôn Dật Tiên, phường Tân Phú, Quận 7, Thành phố Hồ Chí Minh'
    sheet_title_name = 'BÁO CÁO ĐỐI CHIẾU SỐ DƯ TIỀN TÀI KHOẢN KHÁCH HÀNG'
    sub_title_name = f'Ngày {end_date}'

    # --------- sheet BAO CAO CAN LAM ---------
    sheet_bao_cao_can_lam = workbook.add_worksheet('BAO CAO CAN LAM')

    # Insert phu hung picture
    sheet_bao_cao_can_lam.insert_image('A1', './img/phu_hung.png', {'x_scale': 0.66, 'y_scale': 0.71})

    # Set Column Width and Row Height
    sheet_bao_cao_can_lam.set_column('A:A', 6.29)
    sheet_bao_cao_can_lam.set_column('B:B', 13.29)
    sheet_bao_cao_can_lam.set_column('C:C', 15.71)
    sheet_bao_cao_can_lam.set_column('D:D', 15.86)
    sheet_bao_cao_can_lam.set_column('E:E', 16.71)
    sheet_bao_cao_can_lam.set_column('F:F', 17.43)
    sheet_bao_cao_can_lam.set_column('G:G', 13.86)
    sheet_bao_cao_can_lam.set_column('H:H', 17.14)
    sheet_bao_cao_can_lam.set_column('I:I', 17.57)
    sheet_bao_cao_can_lam.set_column('J:J', 13.86)
    sheet_bao_cao_can_lam.set_column('K:K', 17.71)
    sheet_bao_cao_can_lam.set_column('L:L', 15.86)
    sheet_bao_cao_can_lam.set_row(6, 18)
    sheet_bao_cao_can_lam.set_row(7, 15.75)
    sheet_bao_cao_can_lam.set_row(10, 47.25)

    # merge row
    sheet_bao_cao_can_lam.merge_range('C1:L1', CompanyName, company_name_format)
    sheet_bao_cao_can_lam.merge_range('C2:L2', companyAddress, company_format)
    sheet_bao_cao_can_lam.merge_range('C3:L3', CompanyPhoneNumber, company_format)
    sheet_bao_cao_can_lam.merge_range('A7:L7', sheet_title_name, sheet_title_format)
    sheet_bao_cao_can_lam.merge_range('A8:L8', sub_title_name, from_to_format)
    sheet_bao_cao_can_lam.merge_range('A10:A11', headers[0], headers_format)
    sheet_bao_cao_can_lam.merge_range('B10:B11', headers[1], headers_format)
    sheet_bao_cao_can_lam.merge_range('C10:C11', headers[2], headers_format)
    sheet_bao_cao_can_lam.merge_range('D10:D11', headers[3], headers_format)
    sheet_bao_cao_can_lam.merge_range('E10:G10', headers[4], headers_format)
    sheet_bao_cao_can_lam.merge_range('H10:I10', headers[5], headers_format)
    sheet_bao_cao_can_lam.merge_range('J10:J11', headers[6], headers_format)
    sheet_bao_cao_can_lam.merge_range('K10:K11', headers[7], headers_format)
    sheet_bao_cao_can_lam.merge_range('L10:L11', headers[8], headers_format)

    # write row, column
    sheet_bao_cao_can_lam.write_row(
        'A4',
        [''] * len(headers),
        empty_row_format
    )
    sheet_bao_cao_can_lam.write_row(
        'E11',
        sub_headers,
        headers_format
    )
    sheet_bao_cao_can_lam.write_row(
        'A12',
        [f'({i})' for i in np.arange(len(headers) + len(sub_headers) - 2) + 1],
        stt_row_format,
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

