"""
    1. daily
    2. table: cash_balance, sub_account_deposit, transactional_record, transaction_in_system
    3. tăng tiền, giảm tiền lấy từ bảng cashflow_balance
    4. mã nghiệp vụ: transaction_id
    5. số dư đầu kỳ query riêng từ bảng sub_account_deposit
"""
from reporting_tool.trading_service.thanhtoanbutru import *


def run(
        periodicity: str,
        start_date: str,  # 2021-11-22
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

    # --------------------- Viết Query ---------------------
    phat_sinh_giao_dich_query = pd.read_sql(
        f"""
            SELECT
            DISTINCT
            cash_balance.date, 
            relationship.account_code, 
            relationship.sub_account, 
            account.customer_name, 
            cash_balance.transaction_id,
            cash_balance.remark,
            cash_balance.increase, 
            cash_balance.decrease
            FROM cash_balance
            LEFT JOIN relationship ON relationship.sub_account = cash_balance.sub_account
            LEFT JOIN account ON account.account_code = relationship.account_code
            WHERE relationship.date = '{end_date}'
            AND (cash_balance.date BETWEEN '{start_date}' AND '{end_date}')
            ORDER BY decrease, transaction_id
        """,
        connect_DWH_CoSo
    )
    # lấy tổng số dư đầu kỳ
    sum_so_du_dau_ky_query = pd.read_sql(
        f"""
            SELECT sum(opening_balance) as sum_opening_balance 
            FROM sub_account_deposit 
            WHERE date = '{start_date}'
        """,
        connect_DWH_CoSo
    )

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File Excel ---------------------
    # Write file BÁO CÁO PHÁT SINH GIAO DỊCH TIỀN
    for date_char in date_character:
        if date_char in start_date and date_char in end_date:
            start_date = start_date.replace(date_char, '/')
            end_date = end_date.replace(date_char, '/')
    footer_date = bdate(end_date, 1).split('-')
    start_date = dt.datetime.strptime(start_date, "%Y/%m/%d").strftime("%d-%m")
    end_date = dt.datetime.strptime(end_date, "%Y/%m/%d").strftime("%d-%m")
    f_name = ''
    if start_date == end_date:
        f_name += f'Báo cáo phát sinh giao dịch tiền {end_date}.xlsx'
    else:
        f_name += f'Báo cáo phát sinh giao dịch tiền từ {start_date} đến {end_date}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder, folder_name, period, f_name),
        engine='xlsxwriter',
        engine_kwargs={'options': {'nan_inf_to_errors': True}}
    )
    workbook = writer.book

    ###################################################
    ###################################################
    ###################################################

    # ------------- Viet sheet -------------
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
    sub_header_empty_format = workbook.add_format(
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
    so_du_dau_cuoi_cong_ps_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
        }
    )
    num_right_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'top',
            'font_size': 10,
            'font_name': 'Times New Roman'
        }
    )
    date_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'top',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'num_format': 'dd/mm/yyyy'
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
    money_inc_dec_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'top',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'num_format': '#,##0'
        }
    )
    money_dau_cuoi_ky_cong_ps = workbook.add_format(
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
            'font_size': 10,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    empty_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
        }
    )
    empty_1_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
        }
    )
    headers = [
        'STT',
        'Ngày',
        'Số tài khoản',
        'Số tiểu khoản',
        'Tên khách hàng',
        'Mã nghiệp vụ',
        'Tên nghiệp vụ',
        'Tăng tiền',
        'Giảm tiền',
        'Người lập',
        'Người duyệt',
    ]
    companyAddress = 'Tầng 3, CR3-03A, 109 Tôn Dật Tiên, phường Tân Phú, Quận 7, Thành phố Hồ Chí Minh'
    sheet_title_name = 'BÁO CÁO PHÁT SINH GIAO DỊCH TIỀN'
    sub_title_name = f'Từ ngày {start_date} đến ngày {end_date}'
    so_du_dau_ky_name = 'Số dư tiền đầu kỳ'
    cong_phat_sinh = 'Cộng phát sinh'
    so_du_cuoi_ky_name = 'Số dư tiền cuối kỳ'

    # --------- sheet BAO CAO CAN LAM ---------
    sheet_bao_cao_can_lam = workbook.add_worksheet('BAO CAO CAN LAM')

    # Insert phu hung picture
    sheet_bao_cao_can_lam.insert_image('A1', './img/phu_hung.png', {'x_scale': 0.62, 'y_scale': 0.71})

    # Set Column Width and Row Height
    sheet_bao_cao_can_lam.set_column('A:A', 6.43)
    sheet_bao_cao_can_lam.set_column('B:B', 10.86)
    sheet_bao_cao_can_lam.set_column('C:C', 13.71)
    sheet_bao_cao_can_lam.set_column('D:D', 15.14)
    sheet_bao_cao_can_lam.set_column('E:E', 41.14)
    sheet_bao_cao_can_lam.set_column('F:F', 10)
    sheet_bao_cao_can_lam.set_column('G:G', 67.14)
    sheet_bao_cao_can_lam.set_column('H:I', 18.86)
    sheet_bao_cao_can_lam.set_column('J:K', 12.43)
    sheet_bao_cao_can_lam.set_row(5, 20.25)
    sheet_bao_cao_can_lam.set_row(9, 31.5)

    # row of Cộng phát sinh and Số dư tiền cuối kỳ
    cong_phat_sinh_start_row = phat_sinh_giao_dich_query.shape[0] + 12
    so_du_tien_cuoi_ky_start_row = phat_sinh_giao_dich_query.shape[0] + 13

    # Merge row
    sheet_bao_cao_can_lam.merge_range('C1:I1', CompanyName, company_name_format)
    sheet_bao_cao_can_lam.merge_range('C2:I2', companyAddress, company_format)
    sheet_bao_cao_can_lam.merge_range('C3:I3', CompanyPhoneNumber, company_format)
    sheet_bao_cao_can_lam.merge_range('A6:K6', sheet_title_name, sheet_title_format)
    sheet_bao_cao_can_lam.merge_range('A7:K7', sub_title_name, from_to_format)
    sheet_bao_cao_can_lam.merge_range('A11:G11', so_du_dau_ky_name, so_du_dau_cuoi_cong_ps_format)
    sheet_bao_cao_can_lam.merge_range('H11:I11', sum_so_du_dau_ky_query['sum_opening_balance'],
                                      money_dau_cuoi_ky_cong_ps)
    sheet_bao_cao_can_lam.merge_range('J11:K11', '', sub_header_empty_format)
    sheet_bao_cao_can_lam.merge_range(
        f'A{cong_phat_sinh_start_row}:G{cong_phat_sinh_start_row}',
        cong_phat_sinh,
        so_du_dau_cuoi_cong_ps_format
    )
    sheet_bao_cao_can_lam.merge_range(
        f'A{so_du_tien_cuoi_ky_start_row}:G{so_du_tien_cuoi_ky_start_row}',
        so_du_cuoi_ky_name,
        so_du_dau_cuoi_cong_ps_format
    )
    calc_increase_decrease = phat_sinh_giao_dich_query['increase'].sum() - phat_sinh_giao_dich_query['decrease'].sum()
    sheet_bao_cao_can_lam.merge_range(
        f'H{so_du_tien_cuoi_ky_start_row}:I{so_du_tien_cuoi_ky_start_row}',
        sum_so_du_dau_ky_query['sum_opening_balance'] + calc_increase_decrease,
        money_dau_cuoi_ky_cong_ps
    )
    sheet_bao_cao_can_lam.merge_range(
        f'J{so_du_tien_cuoi_ky_start_row}:K{so_du_tien_cuoi_ky_start_row}',
        '',
        empty_format
    )
    sheet_bao_cao_can_lam.merge_range(
        f'A{so_du_tien_cuoi_ky_start_row + 3}:D{so_du_tien_cuoi_ky_start_row + 3}',
        'Người lập',
        footer_text_format
    )
    sheet_bao_cao_can_lam.merge_range(
        f'I{so_du_tien_cuoi_ky_start_row + 2}:K{so_du_tien_cuoi_ky_start_row + 2}',
        f'Ngày{footer_date[2]} tháng {footer_date[1]} năm {footer_date[0]}',
        footer_dmy_format
    )
    sheet_bao_cao_can_lam.merge_range(
        f'I{so_du_tien_cuoi_ky_start_row + 3}:K{so_du_tien_cuoi_ky_start_row + 3}',
        'Người duyệt',
        footer_text_format
    )
    # write row, column
    sheet_bao_cao_can_lam.write_row(
        'A4',
        [''] * len(headers),
        empty_row_format
    )
    sheet_bao_cao_can_lam.write_row('A10', headers, headers_format)
    sheet_bao_cao_can_lam.write_column(
        'A12',
        [int(i) for i in np.arange(phat_sinh_giao_dich_query.shape[0]) + 1],
        num_right_format
    )
    sheet_bao_cao_can_lam.write_column(
        'B12',
        phat_sinh_giao_dich_query['date'],
        date_format
    )
    sheet_bao_cao_can_lam.write_column(
        'C12',
        phat_sinh_giao_dich_query['account_code'],
        text_left_format
    )
    sheet_bao_cao_can_lam.write_column(
        'D12',
        phat_sinh_giao_dich_query['sub_account'],
        text_left_format
    )
    sheet_bao_cao_can_lam.write_column(
        'E12',
        phat_sinh_giao_dich_query['customer_name'],
        text_left_format
    )
    sheet_bao_cao_can_lam.write_column(
        'F12',
        phat_sinh_giao_dich_query['transaction_id'],
        text_left_format
    )
    sheet_bao_cao_can_lam.write_column(
        'G12',
        phat_sinh_giao_dich_query['remark'],
        text_left_format
    )
    sheet_bao_cao_can_lam.write_column(
        'H12',
        phat_sinh_giao_dich_query['increase'],
        money_inc_dec_format
    )
    sheet_bao_cao_can_lam.write_column(
        'I12',
        phat_sinh_giao_dich_query['decrease'],
        money_inc_dec_format
    )
    sheet_bao_cao_can_lam.write_column(
        'J12',
        [''] * phat_sinh_giao_dich_query.shape[0],
        text_left_format
    )
    sheet_bao_cao_can_lam.write_column(
        'K12',
        [''] * phat_sinh_giao_dich_query.shape[0],
        text_left_format
    )
    sheet_bao_cao_can_lam.write(
        f'H{cong_phat_sinh_start_row}',
        phat_sinh_giao_dich_query['increase'].sum(),
        money_dau_cuoi_ky_cong_ps
    )
    sheet_bao_cao_can_lam.write(
        f'I{cong_phat_sinh_start_row}',
        phat_sinh_giao_dich_query['decrease'].sum(),
        money_dau_cuoi_ky_cong_ps
    )
    sheet_bao_cao_can_lam.write(
        f'J{cong_phat_sinh_start_row}',
        '',
        empty_1_format
    )
    sheet_bao_cao_can_lam.write(
        f'K{cong_phat_sinh_start_row}',
        '',
        empty_1_format
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
