"""
1. RCF1002 -> cash_balance
   ROD0040 -> trading_record
"""
from reporting.trading_service.thanhtoanbutru import *


def run(
        periodicity: str,
        start_date: str,
        end_date: str,
        run_time=None,
):
    start = time.time()
    info = get_info(periodicity, run_time)
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
    thu_phi_LK_query = pd.read_sql(
        f"""
        
        """,
        connect_DWH_CoSo
    )

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File Excel ---------------------
    # Write file BÁO CÁO ĐỐI CHIẾU PHẢI THU PHÍ LƯU KÝ
    for date_char in date_character:
        if date_char in start_date and date_char in end_date:
            start_date = start_date.replace(date_char, '/')
            end_date = end_date.replace(date_char, '/')
    s_date_f_name = dt.datetime.strptime(start_date, "%Y/%m/%d").strftime("%d-%m")
    e_date_f_name = dt.datetime.strptime(end_date, "%Y/%m/%d").strftime("%d-%m")
    f_name = ''
    if start_date == end_date:
        f_name += f'Báo cáo đối chiếu phải thu phí lưu ký {e_date_f_name}.xlsx'
    else:
        f_name += f'Báo cáo đối chiếu phải thu phí lưu ký từ {s_date_f_name} đến {e_date_f_name}.xlsx'
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
    headers_sub_headers_format = workbook.add_format(
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
            'font_size': 11,
            'font_name': 'Times New Roman'
        }
    )
    text_left_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vbottom',
            'text_wrap': True,
            'font_size': 11,
            'font_name': 'Times New Roman'
        }
    )
    money_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'num_format': '#,##0'
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
    # create heaaders and sub-headers
    headers = [
        'STT',
        'Số tài khoản',
        'Số tiểu khoản',
        'Tên khách hàng',
        'Số tiểu khoản',
        'Tên khách hàng',
        'Tháng',
        'Dư nợ đầu kỳ',
        'Phát sinh',
        'Dư nợ cuối kỳ'
    ]
    sub_headers = [
        'Tăng',
        'Giảm',
        'Hệ thống',
        'Tính toán lại',
        'Lệch'
    ]

    companyAddress = 'Tầng 3, CR3-03A, 109 Tôn Dật Tiên, phường Tân Phú, Quận 7, Thành phố Hồ Chí Minh'
    sheet_title_name = 'BÁO CÁO ĐỐI CHIẾU PHẢI THU PHÍ LƯU KÝ'
    start_month = dt.datetime.strptime(start_date, "%Y/%m/%d").strftime("%m-%Y")
    end_month = dt.datetime.strptime(end_date, "%Y/%m/%d").strftime("%m-%Y")
    sub_title_name = f'Từ ngày {start_month} đến {end_month}'

    # --------- sheet BAO CAO CAN LAM ---------
    sheet_bao_cao_can_lam = workbook.add_worksheet('BAO CAO CAN LAM')

    # Insert phu hung picture
    sheet_bao_cao_can_lam.insert_image('A1', './img/phu_hung.png', {'x_scale': 0.66, 'y_scale': 0.71})

    # Set Column Width and Row Height
    sheet_bao_cao_can_lam.set_column('A:A', 6.5)
    sheet_bao_cao_can_lam.set_column('B:C', 14.14)
    sheet_bao_cao_can_lam.set_column('D:D', 22)
    sheet_bao_cao_can_lam.set_column('E:E', 13.5)
    sheet_bao_cao_can_lam.set_column('F:F', 15)
    sheet_bao_cao_can_lam.set_column('G:H', 15)
    sheet_bao_cao_can_lam.set_column('I:K', 15)

    # merge row
    sheet_bao_cao_can_lam.merge_range('C1:H1', CompanyName, company_name_format)
    sheet_bao_cao_can_lam.merge_range('C2:H2', companyAddress, company_format)
    sheet_bao_cao_can_lam.merge_range('C3:H3', CompanyPhoneNumber, company_format)
    sheet_bao_cao_can_lam.merge_range('B7:K7', sheet_title_name, sheet_title_format)
    sheet_bao_cao_can_lam.merge_range('B8:K8', sub_title_name, from_to_format)
    sheet_bao_cao_can_lam.merge_range('A10:A11', headers[0], headers_sub_headers_format)
    sheet_bao_cao_can_lam.merge_range('B10:B11', headers[1], headers_sub_headers_format)
    sheet_bao_cao_can_lam.merge_range('C10:C11', headers[2], headers_sub_headers_format)
    sheet_bao_cao_can_lam.merge_range('D10:D11', headers[3], headers_sub_headers_format)
    sheet_bao_cao_can_lam.merge_range('E10:E11', headers[4], headers_sub_headers_format)
    sheet_bao_cao_can_lam.merge_range('F10:F11', headers[5], headers_sub_headers_format)
    sheet_bao_cao_can_lam.merge_range('G10:H10', headers[6], headers_sub_headers_format)
    sheet_bao_cao_can_lam.merge_range('I10:K10', headers[7], headers_sub_headers_format)
    sum_start_row = thu_phi_LK_query.shape[0] + 13
    sheet_bao_cao_can_lam.merge_range(
        f'A{sum_start_row}:E{sum_start_row}',
        'Tổng',
        headers_sub_headers_format
    )
    footer_start_row = sum_start_row + 2
    footer_date = bdate(end_date, 1).split('-')
    sheet_bao_cao_can_lam.merge_range(
        f'H{footer_start_row}:K{footer_start_row}',
        f'Ngày {footer_date[2]} tháng {footer_date[1]} năm {footer_date[0]}',
        footer_dmy_format
    )
    sheet_bao_cao_can_lam.merge_range(
        f'H{footer_start_row + 1}:K{footer_start_row + 1}',
        'Người duyệt',
        footer_text_format
    )
    sheet_bao_cao_can_lam.merge_range(
        f'A{footer_start_row + 1}:D{footer_start_row + 1}',
        'Người lập',
        footer_text_format
    )

    # write row & column
    sheet_bao_cao_can_lam.write_row('G11', sub_headers, headers_sub_headers_format)
    sheet_bao_cao_can_lam.write_row(
        'A12',
        [f'({i})' for i in np.arange(len(headers) + len(sub_headers) + 3) + 1],
        stt_row_format,
    )
    sheet_bao_cao_can_lam.write_row(
        'A4',
        [''] * (len(headers) + len(sub_headers) + 3),
        empty_row_format
    )
    sheet_bao_cao_can_lam.write_column(
        'A13',
        [int(i) for i in np.arange(thu_phi_LK_query.shape[0]) + 1],
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


