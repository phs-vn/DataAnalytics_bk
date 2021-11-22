"""
    1. daily
    2. table: cash_balance, sub_account_deposit, transactional_record, transaction_in_system
    3. tăng tiền, giảm tiền lấy từ bảng cashflow_balance
    4. mã nghiệp vụ: transaction_id
"""
from reporting_tool.trading_service.thanhtoanbutru import *


def run(
        periodicity: str,
        start_date: str,  # 2021-11-18
        end_date: str,  # 2021-11-19
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
    # Write file Bao cao phat sinh giao dich Tien
    for date_char in date_character:
        if date_char in start_date and date_char in end_date:
            start_date = start_date.replace(date_char, '/')
            end_date = end_date.replace(date_char, '/')

    start_date = dt.strptime(start_date, "%Y/%m/%d").strftime("%d-%m")
    end_date = dt.strptime(end_date, "%Y/%m/%d").strftime("%d-%m")
    f_name = 'Báo cáo phát sinh giao dịch tiền.xlsx'
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
    workbook.add_format(
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
    sub_title_empty_format = workbook.add_format(
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
        'Chi nhánh quản lý',
        'Nhân viên quản lý tài khoản',
    ]
    companyAddress = 'Tầng 3, CR3-03A, 109 Tôn Dật Tiên, phường Tân Phú, Quận 7, Thành phố Hồ Chí Minh'
    sheet_title_name = 'BÁO CÁO PHÁT SINH GIAO DỊCH TIỀN'
    sub_title_name = f'Từ ngày {start_date} đến ngày {end_date}'
    # --------- sheet BAO CAO CAN LAM ---------
    sheet_bao_cao_can_lam = workbook.add_worksheet('BAO CAO CAN LAM')

    # Set Column Width and Row Height
    sheet_bao_cao_can_lam.set_column('A:A', 4.43)
    sheet_bao_cao_can_lam.set_column('B:B', 14.14)
    sheet_bao_cao_can_lam.set_column('C:C', 12.14)
    sheet_bao_cao_can_lam.set_column('D:D', 13.43)
    sheet_bao_cao_can_lam.set_column('E:E', 15.57)
    sheet_bao_cao_can_lam.set_column('F:F', 13.57)
    sheet_bao_cao_can_lam.set_column('G:G', 13.86)
    sheet_bao_cao_can_lam.set_column('H:I', 9.29)
    sheet_bao_cao_can_lam.set_column('J:K', 12.43)
    sheet_bao_cao_can_lam.set_column('L:L', 13.71)
    sheet_bao_cao_can_lam.set_column('M:M', 16.43)
    sheet_bao_cao_can_lam.set_row(4, 27)
    sheet_bao_cao_can_lam.set_row(5, 24)

    # Write values to row and column
    # merge row
    sheet_bao_cao_can_lam.merge_range('C1:I1', CompanyName, company_name_format)
    sheet_bao_cao_can_lam.merge_range('C2:I2', companyAddress, company_format)
    sheet_bao_cao_can_lam.merge_range('C3:I3', CompanyPhoneNumber, company_format)

    sheet_bao_cao_can_lam.merge_range('A6:M6', sheet_title_name, sheet_title_format)
    sheet_bao_cao_can_lam.merge_range('A7:M7', sub_title_name, from_to_format)
    sheet_bao_cao_can_lam.write_row('A10', headers, headers_format)
    sheet_bao_cao_can_lam.merge_range('A11:G11', 'Số dư đầu kỳ', headers_format)
    sheet_bao_cao_can_lam.merge_range('H11:I11', '', sub_title_empty_format)
    sheet_bao_cao_can_lam.merge_range('J11:M11', '', sub_title_empty_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py', '')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time() - start, 1)}s')
