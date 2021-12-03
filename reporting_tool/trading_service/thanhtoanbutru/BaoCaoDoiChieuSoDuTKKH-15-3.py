"""
    1. daily
    2. table:
    3. start_date : T-1
       end_date : T
"""
from reporting_tool.trading_service.thanhtoanbutru import *


def run(
        periodicity: str,
        end_date: str,  # 2021-11-29
        run_time=None,
):
    start = time.time()
    info = get_info(periodicity, run_time)
    period = info['period']
    folder_name = info['folder_name']
    start_date = bdate(end_date, -1)

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
            sub_account_deposit.closing_balance,
            branch.branch_name,
            broker.broker_name
            FROM sub_account_deposit
            LEFT JOIN relationship ON relationship.sub_account = sub_account_deposit.sub_account
            LEFT JOIN account ON account.account_code = relationship.account_code
            LEFT JOIN branch ON branch.branch_id = relationship.branch_id
            LEFT JOIN broker ON broker.broker_id = relationship.broker_id
            WHERE sub_account_deposit.date BETWEEN '{start_date}' AND '{end_date}'
            AND relationship.date = '{end_date}'
            ORDER BY sub_account_deposit.date ASC
        """,
        connect_DWH_CoSo,
        index_col='sub_account'
    )
    # lọc bảng lớn thành theo 2 ngày start_date và end_date
    info_T_tru_1 = doi_chieu_so_du_query.loc[doi_chieu_so_du_query['date'] == start_date]  # start_date
    info_T0 = doi_chieu_so_du_query.loc[doi_chieu_so_du_query['date'] == end_date]  # end_date
    # Xử lý format của date
    date_character = ['/', '-', '.']
    for date_char in date_character:
        if date_char in start_date and date_char in end_date:
            start_date = start_date.replace(date_char, '/')
            end_date = end_date.replace(date_char, '/')

    # read File Đã lưu hôm qua
    read_file_start_date = dt.datetime.strptime(start_date, "%Y/%m/%d").strftime("%d-%m")
    path = 'D:\\DataAnalytics\\reporting_tool\\trading_service\\output'
    file_name = f'Dữ liệu lưu ngoài flex {read_file_start_date}.xlsx'
    save_yesterday = pd.read_excel(join(path, file_name),
                                   sheet_name='Balance',
                                   dtype={'Sub account': object},
                                   na_filter=True)
    save_yesterday = save_yesterday.set_index('Sub account')
    save_yesterday.rename(
        {
            'Opening balance T0': 'opening_balance',
            'Closing balance T0': 'closing_balance'
        },
        axis=1,
        inplace=True
    )

    idx = info_T_tru_1.index.union(info_T0.index)  # lấy phần hợp của 2 dataframe
    # set lại index
    info_T_tru_1 = info_T_tru_1.reindex(idx)
    info_T0 = info_T0.reindex(idx)
    save_yesterday = save_yesterday.reindex(idx)
    # tạo dataframe mới chỉ lấy 2 cột opening và closing balance
    balance_tru_1 = info_T_tru_1[['opening_balance', 'closing_balance']].copy()
    # rename column
    balance_tru_1.rename(
        {
            'opening_balance': 'opening_t_tru_1',
            'closing_balance': 'closing_t_tru_1'
        },
        axis=1,
        inplace=True
    )
    balance_tru_1.fillna(0, inplace=True)

    info_T0['opening_balance_tru_1'] = balance_tru_1['opening_t_tru_1']
    info_T0['closing_balance_tru_1'] = balance_tru_1['closing_t_tru_1']
    info_T0 = info_T0.dropna()
    info_T0['yesterday_opening'] = save_yesterday['opening_balance']
    info_T0['yesterday_closing'] = save_yesterday['closing_balance']
    info_T0.fillna(0, inplace=True)

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File Excel ---------------------
    # Write file BÁO CÁO ĐỐI CHIẾU SỐ DƯ TIỀN TÀI KHOẢN KHÁCH HÀNG
    footer_date = bdate(end_date, 1).split('-')
    eod = dt.datetime.strptime(end_date, "%Y/%m/%d").strftime("%d-%m")

    f_name = f'Báo cáo đối chiếu số dư tiền tài khoản khách hàng {eod}.xlsx'
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
    money_sum_format = workbook.add_format(
        {
            'bold': True,
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
    eod_sub = dt.datetime.strptime(end_date, "%Y/%m/%d").strftime("%d/%m/%Y")
    sub_title_name = f'Ngày {eod_sub}'

    # --------- sheet BAO CAO CAN LAM ---------
    sheet_bao_cao_can_lam = workbook.add_worksheet('BAO CAO CAN LAM')

    # Insert phu hung picture
    sheet_bao_cao_can_lam.insert_image('A1', './img/phu_hung.png', {'x_scale': 0.66, 'y_scale': 0.71})

    # Set Column Width and Row Height
    sheet_bao_cao_can_lam.set_column('A:A', 6.29)
    sheet_bao_cao_can_lam.set_column('B:B', 12.14)
    sheet_bao_cao_can_lam.set_column('C:C', 13.43)
    sheet_bao_cao_can_lam.set_column('D:D', 25.43)
    sheet_bao_cao_can_lam.set_column('E:E', 14.14)
    sheet_bao_cao_can_lam.set_column('F:F', 14.57)
    sheet_bao_cao_can_lam.set_column('G:G', 13.71)
    sheet_bao_cao_can_lam.set_column('H:H', 17.14)
    sheet_bao_cao_can_lam.set_column('I:I', 17.57)
    sheet_bao_cao_can_lam.set_column('J:J', 13.86)
    sheet_bao_cao_can_lam.set_column('K:K', 19.43)
    sheet_bao_cao_can_lam.set_column('L:L', 29)
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
    sum_start_row = info_T0.shape[0] + 12
    sheet_bao_cao_can_lam.merge_range(
        f'A{sum_start_row}:D{sum_start_row}',
        'Tổng',
        headers_format
    )
    sheet_bao_cao_can_lam.merge_range(
        f'K{sum_start_row}:L{sum_start_row}',
        '',
        text_left_wrap_format
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
    sheet_bao_cao_can_lam.set_row(footer_start_row, 21.75)

    # write row, column
    sheet_bao_cao_can_lam.write_row(
        'A4',
        [''] * (len(headers) + len(sub_headers) - 2),
        empty_row_format
    )
    sheet_bao_cao_can_lam.write_row(
        f'E{sum_start_row}',
        [''] * 6,
        money_format
    )
    sheet_bao_cao_can_lam.write_row(
        'E11',
        sub_headers,
        headers_format
    )
    sheet_bao_cao_can_lam.write_column(
        'A12',
        [int(i) for i in np.arange(info_T0.shape[0]) + 1],
        stt_col_format
    )
    sheet_bao_cao_can_lam.write_column(
        'B12',
        info_T0['account_code'],
        text_left_format
    )
    sheet_bao_cao_can_lam.write_column(
        'C12',
        info_T0.index,
        text_left_format
    )
    sheet_bao_cao_can_lam.write_column(
        'D12',
        info_T0['customer_name'],
        text_left_wrap_format
    )
    sheet_bao_cao_can_lam.write_column(
        'E12',
        info_T0['opening_balance_tru_1'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'F12',
        info_T0['closing_balance_tru_1'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'G12',
        info_T0['opening_balance'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'H12',
        info_T0['yesterday_opening'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'I12',
        info_T0['yesterday_closing'],
        money_format
    )

    for idx in info_T0.index:
        five = info_T0.loc[idx, 'opening_balance_tru_1']
        six = info_T0.loc[idx, 'closing_balance_tru_1']
        seven = info_T0.loc[idx, 'opening_balance']
        eight = info_T0.loc[idx, 'yesterday_opening']
        nine = info_T0.loc[idx, 'yesterday_closing']
        if ((seven - six) != 0) or ((eight - five) != 0) or ((nine - six) != 0):
            info_T0.loc[idx, 'check_bat_thuong'] = True  # bất thường
        else:
            info_T0.loc[idx, 'check_bat_thuong'] = False  # bình thường
    sheet_bao_cao_can_lam.write_column(
        'J12',
        info_T0['check_bat_thuong'],
        text_left_format
    )
    sheet_bao_cao_can_lam.write_column(
        'K12',
        info_T0['branch_name'],
        text_left_format
    )
    sheet_bao_cao_can_lam.write_column(
        'L12',
        info_T0['broker_name'],
        text_left_wrap_format
    )
    sheet_bao_cao_can_lam.write(
        f'E{sum_start_row}',
        info_T0['opening_balance_tru_1'].sum(),
        money_sum_format
    )
    sheet_bao_cao_can_lam.write(
        f'F{sum_start_row}',
        info_T0['closing_balance_tru_1'].sum(),
        money_sum_format
    )
    sheet_bao_cao_can_lam.write(
        f'G{sum_start_row}',
        info_T0['opening_balance'].sum(),
        money_sum_format
    )

    ###########################################################################
    ###########################################################################
    ###########################################################################

    # ------------- write 2 column 'opening_t_0' and 'closing_t_0' in 'df t_0' to file -------------
    path = join(realpath(dirname(dirname(__file__))), 'output')
    eod_save_file = dt.datetime.strptime(end_date, "%Y/%m/%d").strftime("%d-%m")
    file_name = f'Dữ liệu lưu ngoài flex {eod_save_file}.xlsx'
    writer2 = pd.ExcelWriter(
        join(path, file_name),
        engine='xlsxwriter',
        engine_kwargs={'options': {'nan_inf_to_errors': True}}
    )
    workbook2 = writer2.book

    ###################################################
    ###################################################
    ###################################################

    # ------------- Viết sheet -------------
    # Format
    header_format = workbook2.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_size': 12,
            'font_name': 'Times New Roman'
        }
    )
    text_center_format = workbook2.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'top',
            'font_size': 10,
            'font_name': 'Times New Roman',
        }
    )
    money_format = workbook2.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'top',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'num_format': '#,##0'
        }
    )
    header = [
        'Sub account',
        'Opening balance T0',
        'Closing balance T0',
    ]

    # --------- sheet Balance ---------
    sheet_balance = workbook2.add_worksheet('Balance')

    sheet_balance.set_column('A:A', 12)
    sheet_balance.set_column('B:C', 13)

    sheet_balance.write_row('A1', header, header_format)
    sheet_balance.write_column(
        'A2',
        info_T0.index,
        text_center_format
    )
    sheet_balance.write_column(
        'B2',
        info_T0['opening_balance'],
        money_format
    )
    sheet_balance.write_column(
        'C2',
        info_T0['closing_balance'],
        money_format
    )

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    writer2.close()

    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py', '')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time() - start, 1)}s')
