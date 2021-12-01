"""
1. RCF1002 -> cash_balance
   ROD0040 -> trading_record
"""
from reporting_tool.trading_service.thanhtoanbutru import *


def run(
        periodicity: str,
        start_date: str,
        end_date: str,
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
    if not os.path.isdir(join(dept_folder, folder_name)):
        os.mkdir(join(dept_folder, folder_name))
    if not os.path.isdir(join(dept_folder, folder_name, period)):
        os.mkdir((join(dept_folder, folder_name, period)))

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viết Query và xử lý dataframe ---------------------
    phi_LK_NDT_query = pd.read_sql(
        f"""

        """,
        connect_DWH_CoSo
    )

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File Excel ---------------------
    # Write file BÁO CÁO ĐỐI CHIẾU CHI TIẾT PHÍ LƯU KÝ NHÀ ĐẦU TƯ
    for date_char in date_character:
        if date_char in start_date and date_char in end_date:
            start_date = start_date.replace(date_char, '/')
            end_date = end_date.replace(date_char, '/')
    s_date_f_name = dt.datetime.strptime(start_date, "%Y/%m/%d").strftime("%d-%m")
    e_date_f_name = dt.datetime.strptime(end_date, "%Y/%m/%d").strftime("%d-%m")
    f_name = ''
    if start_date == end_date:
        f_name += f'Báo cáo đối chiếu chi tiết phí lưu ký nhà đầu tư {e_date_f_name}.xlsx'
    else:
        f_name += f'Báo cáo đối chiếu chi tiết phí lưu ký nhà đầu tư {s_date_f_name} đến {e_date_f_name}.xlsx'
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
        'Ngày',
        'Số tài khoản',
        'Số tiểu khoản',
        'Tên khách hàng',
        'Loại chứng khoán',
        'Mã chứng khoán',
        'Số lượng',
        'Theo tham số hệ thống',
        'Theo tính toán',
        'Lệch',
    ]
    sub_headers = [
        'Mức phí',
        'Số phí',
        'Số phí cộng dồn',
    ]

    companyAddress = 'Tầng 3, CR3-03A, 109 Tôn Dật Tiên, phường Tân Phú, Quận 7, Thành phố Hồ Chí Minh'
    sheet_title_name = 'BÁO CÁO ĐỐI CHIẾU CHI TIẾT PHÍ LƯU KÝ NHÀ ĐẦU TƯ'
    sub_title_start_date = dt.datetime.strptime(start_date, "%Y/%m/%d").strftime("%d/%m/%Y")
    sub_title_end_date = dt.datetime.strptime(end_date, "%Y/%m/%d").strftime("%d/%m/%Y")
    sub_title_name = f'Từ ngày {sub_title_start_date} Đến ngày {sub_title_end_date}'

    # --------- sheet BAO CAO CAN LAM ---------
    sheet_bao_cao_can_lam = workbook.add_worksheet('BAO CAO CAN LAM')

    # Insert phu hung picture
    sheet_bao_cao_can_lam.insert_image('A1', './img/phu_hung.png', {'x_scale': 0.66, 'y_scale': 0.71})

    # Set Column Width and Row Height
    sheet_bao_cao_can_lam.set_column('A:A', 7)
    sheet_bao_cao_can_lam.set_column('B:B', 13)
    sheet_bao_cao_can_lam.set_column('C:C', 15)
    sheet_bao_cao_can_lam.set_column('D:D', 15)
    sheet_bao_cao_can_lam.set_column('E:E', 22)
    sheet_bao_cao_can_lam.set_column('F:F', 15)
    sheet_bao_cao_can_lam.set_column('G:G', 15)
    sheet_bao_cao_can_lam.set_column('H:H', 7)
    sheet_bao_cao_can_lam.set_column('I:I', 18)
    sheet_bao_cao_can_lam.set_column('J:J', 18)
    sheet_bao_cao_can_lam.set_column('K:K', 18)
    sheet_bao_cao_can_lam.set_column('L:L', 18)
    sheet_bao_cao_can_lam.set_column('M:M', 18)
    sheet_bao_cao_can_lam.set_column('N:N', 18)
    sheet_bao_cao_can_lam.set_column('O:O', 18)
    sheet_bao_cao_can_lam.set_column('P:P', 18)
    sheet_bao_cao_can_lam.set_column('Q:Q', 18)

    # merge row
    sheet_bao_cao_can_lam.merge_range('D1:K1', CompanyName, company_name_format)
    sheet_bao_cao_can_lam.merge_range('D2:K2', companyAddress, company_format)
    sheet_bao_cao_can_lam.merge_range('D3:K3', CompanyPhoneNumber, company_format)
    sheet_bao_cao_can_lam.merge_range('B7:Q7', sheet_title_name, sheet_title_format)
    sheet_bao_cao_can_lam.merge_range('B8:Q8', sub_title_name, from_to_format)
    sheet_bao_cao_can_lam.merge_range('A10:A11', headers[0], headers_sub_headers_format)
    sheet_bao_cao_can_lam.merge_range('B10:B11', headers[1], headers_sub_headers_format)
    sheet_bao_cao_can_lam.merge_range('C10:C11', headers[2], headers_sub_headers_format)
    sheet_bao_cao_can_lam.merge_range('D10:D11', headers[3], headers_sub_headers_format)
    sheet_bao_cao_can_lam.merge_range('E10:E11', headers[4], headers_sub_headers_format)
    sheet_bao_cao_can_lam.merge_range('F10:F11', headers[5], headers_sub_headers_format)
    sheet_bao_cao_can_lam.merge_range('G10:G11', headers[6], headers_sub_headers_format)
    sheet_bao_cao_can_lam.merge_range('H10:H11', headers[7], headers_sub_headers_format)
    sheet_bao_cao_can_lam.merge_range('I10:K10', headers[8], headers_sub_headers_format)
    sheet_bao_cao_can_lam.merge_range('L10:N10', headers[9], headers_sub_headers_format)
    sheet_bao_cao_can_lam.merge_range('O10:Q10', headers[10], headers_sub_headers_format)
    sum_start_row = phi_LK_NDT_query.shape[0] + 13
    sheet_bao_cao_can_lam.merge_range(
        f'A{sum_start_row}:G{sum_start_row}',
        'Tổng',
        headers_sub_headers_format
    )
    footer_start_row = sum_start_row + 2
    footer_date = bdate(end_date, 1).split('-')
    sheet_bao_cao_can_lam.merge_range(
        f'O{footer_start_row}:Q{footer_start_row}',
        f'Ngày {footer_date[2]} tháng {footer_date[1]} năm {footer_date[0]}',
        footer_dmy_format
    )
    sheet_bao_cao_can_lam.merge_range(
        f'O{footer_start_row + 1}:Q{footer_start_row + 1}',
        'Người duyệt',
        footer_text_format
    )
    sheet_bao_cao_can_lam.merge_range(
        f'C{footer_start_row + 1}:F{footer_start_row + 1}',
        'Người lập',
        footer_text_format
    )

    # write row & column
    sheet_bao_cao_can_lam.write_row('I11', sub_headers * 3, headers_sub_headers_format)
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
        [int(i) for i in np.arange(phi_LK_NDT_query.shape[0]) + 1],
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


