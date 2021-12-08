"""
1. Loại báo cáo: monthly
2. Chỉ xuất ra 2 số đầu kỳ và cuối kỳ
3. Table: rdt0121
"""

from reporting_tool.trading_service.thanhtoanbutru import *


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
    date_dau_ky = bdate(start_date, -1)
    date_cuoi_ky = end_date
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
    # query tổng số dư ký quỹ (tiền mặt)
    phai_sinh_thang_query = pd.read_sql(
        f"""
            SELECT date, cash_balance_at_vsd
            from rdt0121
            where date = '{date_dau_ky}' or date = '{date_cuoi_ky}'
        """,
        connect_DWH_PhaiSinh
    )
    dau_thang = phai_sinh_thang_query.loc[phai_sinh_thang_query['date'] == date_dau_ky]
    cuoi_thang = phai_sinh_thang_query.loc[phai_sinh_thang_query['date'] == date_cuoi_ky]

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File ---------------------
    # Write file excel Bao cao phai sinh hang thang
    f_name = f'Báo cáo phái sinh tháng {period}.xlsx'
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
    header_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 14,
            'font_name': 'Times New Roman',
        }
    )
    text_left_format = workbook.add_format(
        {
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 12,
            'font_name': 'Times New Roman',
        }
    )
    money_format = workbook.add_format(
        {
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 14,
            'font_name': 'Times New Roman',
            'num_format': '#,##0'
        }
    )
    sheet_title_name = 'Tổng số dư ký quỹ'
    headers = [
        'Đầu tháng',
        'Cuối tháng'
    ]
    idx_name = 'Tiền mặt'

    # --------- sheet BAO CAO CAN LAM ---------
    sheet = workbook.add_worksheet('sheet1')
    sheet.set_column('A:A', 9.86)
    sheet.set_column('B:C', 17.14)
    sheet.merge_range('A1:C1', sheet_title_name, sheet_title_format)
    sheet.write_row('B3', headers, header_format)
    sheet.write('A4', idx_name, text_left_format)
    sheet.write('B4', dau_thang['cash_balance_at_vsd'].sum(), money_format)
    sheet.write('C4', cuoi_thang['cash_balance_at_vsd'].sum(), money_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py', '')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time() - start, 1)}s')