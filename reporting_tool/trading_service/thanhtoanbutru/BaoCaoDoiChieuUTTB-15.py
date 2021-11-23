"""
    1. daily
    2. table:
"""
from reporting_tool.trading_service.thanhtoanbutru import *


def run(
        periodicity: str,
        start_date: str,  # 2021-11-22
        end_date: str,    # 2021-11-22
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



    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File Excel ---------------------
    # Write file BÁO CÁO ĐỐI CHIẾU LÃI TIỀN GỬI PHÁT SINH TRÊN TÀI KHOẢN KHÁCH HÀNG
    for date_char in date_character:
        if date_char in start_date and date_char in end_date:
            start_date = start_date.replace(date_char, '/')
            end_date = end_date.replace(date_char, '/')

    start_date = dt.strptime(start_date, "%Y/%m/%d").strftime("%d-%m")
    end_date = dt.strptime(end_date, "%Y/%m/%d").strftime("%d-%m")
    f_name = ''
    if start_date == end_date:
        f_name += f'Báo cáo đối chiếu UTTB {end_date}.xlsx'
    else:
        f_name += f'Báo cáo đối chiếu UTTB {start_date} đến {end_date}.xlsx'
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
        'Giá trị tiền bán T-2',
        'Giá trị tiền bán T-1',
        'Giá trị tiền bán T0',
        'Tiền Hoàn trả UTTB T0',
        'Tổng giá trị tiền bán có thể ứng',
        'Tiền đã ứng',
        'Tiền còn có thể ứng',
        'Bất Thường'
    ]

    companyAddress = 'Tầng 3, CR3-03A, 109 Tôn Dật Tiên, phường Tân Phú, Quận 7, Thành phố Hồ Chí Minh'
    sheet_title_name = 'BÁO CÁO ĐỐI CHIẾU UTTB'
    sub_title_name = f'Ngày {end_date}'

    # --------- sheet BAO CAO CAN LAM ---------
    sheet_bao_cao_can_lam = workbook.add_worksheet('BAO CAO CAN LAM')

    # Insert phu hung picture
    sheet_bao_cao_can_lam.insert_image('A1', './img/phu_hung.png', {'x_scale': 0.66, 'y_scale': 0.71})

    # Set Column Width and Row Height
    sheet_bao_cao_can_lam.set_column('A:A', 6.29)
    sheet_bao_cao_can_lam.set_column('B:B', 13.29)
    sheet_bao_cao_can_lam.set_column('C:C', 13)
    sheet_bao_cao_can_lam.set_column('D:D', 10.29)
    sheet_bao_cao_can_lam.set_column('E:E', 14.43)
    sheet_bao_cao_can_lam.set_column('F:F', 14.57)
    sheet_bao_cao_can_lam.set_column('G:G', 15.29)
    sheet_bao_cao_can_lam.set_column('H:H', 17.43)
    sheet_bao_cao_can_lam.set_column('I:I', 15.71)
    sheet_bao_cao_can_lam.set_column('J:J', 11.71)
    sheet_bao_cao_can_lam.set_column('K:K', 14.86)
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

    # write row, column
    sheet_bao_cao_can_lam.write_row(
        'A4',
        [''] * len(headers),
        empty_row_format
    )
    sheet_bao_cao_can_lam.write_row('A11', headers, headers_format)
    sheet_bao_cao_can_lam.write_row(
        'A12',
        [f'({i})' for i in np.arange(len(headers)) + 1],
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