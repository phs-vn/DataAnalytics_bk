"""
1. FDS
    - Màn hình tổng hợp tài khoản của FDS
    - DT0141
    - DT0121
    - DT0132
    - DT0136
2. daily
3. Rule
    - 2 cột E và F ('Số tiền tại công ty' và 'Số tiền ký quỹ tại VSD') lấy từ
    2 cột 'cash_balance_at_phs' và 'cash_balance_at_vsd' của rdt0121
    - cột G - 'nợ chậm trả' lấy từ cột 'deferred_payment_amount_closing' của rdt0141
    - cột I - 'giá trị tài sản ròng' lấy từ cột 'deposited_asset_value' của rdt0136
"""

from reporting.trading_service.thanhtoanbutru import *


def run(
        periodicity: str,
        run_time=None,
):
    start = time.time()
    info = get_info(periodicity, run_time)
    start_date = info['start_date']
    end_date = info['end_date']
    period = info['period']
    folder_name = info['folder_name']
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

    # --------------------- Viết Query ---------------------
    # query
    query_bao_cao_16 = pd.read_sql(
        f"""
        """,
        connect_DWH_PhaiSinh
    )

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File ---------------------
    # Write file excel Báo cáo FDS theo dõi TK cần báo KH chuyển tiền từ VSD về TKGD
    f_name = f'Báo cáo FDS theo dõi TK cần báo KH chuyển tiền từ VSD về TKGD {period}.xlsx'
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
        'STT\nNo',
        'Tên chi nhánh\nBranch Name',
        'Tài khoản ký quỹ\nMargin account',
        'Tên khách hàng\nCustomer name',
        'Số tiền tại công ty\nCash balance at Company stock',
        'Số tiền ký quỹ tại VSD\nCash balance manage at VSD',
        'Nợ chậm trả',
        'Tổng giá trị phí thuế',
        'Giá trị tài sản ròng'
    ]
    companyAddress = 'Tầng 3, CR3-03A, 109 Tôn Dật Tiên, phường Tân Phú, Quận 7, Thành phố Hồ Chí Minh'
    sheet_title_name = 'BÁO CÁO THEO DÕI CÁC TK CẦN XỬ LÝ TRÁNH ÂM TIỀN TRÊN FDS'
    eod_sub = dt.datetime.strptime(end_date, "%Y/%m/%d").strftime("%d/%m/%Y")
    sub_title_name = f'Ngày {eod_sub}'

    # --------- sheet BAO CAO CAN LAM ---------
    sheet_bao_cao_can_lam = workbook.add_worksheet('BAO CAO CAN LAM')

    # Insert phu hung picture
    sheet_bao_cao_can_lam.insert_image('A1', './img/phu_hung.png', {'x_scale': 0.65, 'y_scale': 0.71})
    sheet_bao_cao_can_lam.set_column('A:A', 6.43)
    sheet_bao_cao_can_lam.set_column('B:B', 12.43)
    sheet_bao_cao_can_lam.set_column('C:C', 21.14)
    sheet_bao_cao_can_lam.set_column('D:D', 25.71)
    sheet_bao_cao_can_lam.set_column('E:E', 18)
    sheet_bao_cao_can_lam.set_column('F:F', 28)
    sheet_bao_cao_can_lam.set_column('G:G', 15.29)
    sheet_bao_cao_can_lam.set_column('H:H', 13.57)
    sheet_bao_cao_can_lam.set_column('I:I', 12)
    sheet_bao_cao_can_lam.set_row(7, 23.25)
    sheet_bao_cao_can_lam.set_row(11, 51)

    # merge row
    sheet_bao_cao_can_lam.merge_range('C1:F1', CompanyName, company_name_format)
    sheet_bao_cao_can_lam.merge_range('C2:F2', companyAddress, company_format)
    sheet_bao_cao_can_lam.merge_range('C3:F3', CompanyPhoneNumber, company_format)
    sheet_bao_cao_can_lam.merge_range('A7:I7', sheet_title_name, sheet_title_format)
    sheet_bao_cao_can_lam.merge_range('A8:I8', sub_title_name, from_to_format)
    sum_start_row = query_bao_cao_16.shape[0] + 13
    sheet_bao_cao_can_lam.merge_range(
        f'A{sum_start_row}:D{sum_start_row}',
        'Tổng',
        headers_format
    )
    footer_date = bdate(end_date, 1).split('-')
    footer_start_row = sum_start_row + 2
    sheet_bao_cao_can_lam.merge_range(
        f'G{footer_start_row}:I{footer_start_row}',
        f'Ngày {footer_date[2]} tháng {footer_date[1]} năm {footer_date[0]}',
        footer_dmy_format
    )
    sheet_bao_cao_can_lam.merge_range(
        f'G{footer_start_row + 1}:I{footer_start_row + 1}',
        'Người duyệt',
        footer_text_format
    )
    sheet_bao_cao_can_lam.merge_range(
        f'A{footer_start_row + 1}:D{footer_start_row + 1}',
        'Người lập',
        footer_text_format
    )

    sheet_bao_cao_can_lam.write_row(
        'A4',
        [''] * len(headers),
        empty_row_format
    )
    sheet_bao_cao_can_lam.write_row('A11', headers, headers_format)
    sheet_bao_cao_can_lam.write_column(
        'A12',
        [int(i) for i in np.arange(query_bao_cao_16.shape[0]) + 1],
        stt_col_format
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