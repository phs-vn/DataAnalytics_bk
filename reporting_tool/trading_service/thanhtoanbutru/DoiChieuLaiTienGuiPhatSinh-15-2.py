"""
    1. daily
    2. Chị Tuyết muốn xuất thành 2 file:
        - file 1: xuất tại thời điểm ngày đang chạy
        - file 2: xuất tổng tiền từ ngày làm việc cuối cùng của tháng trước tới ngày đang chạy
        - VD: báo cáo chạy ngày 05/12/2021
            --> file 1: xuất của ngày 05/12/2021
            --> file 2: xuất file tổng tiền từ ngày 29/10/2021 tới ngày 05/12/2021
"""
from reporting_tool.trading_service.thanhtoanbutru import *


def run(
        start_date: str,    # 2021-12-01
        end_date: str,      # 2021-12-01
        run_time=None,
):
    start = time.time()
    info = get_info('daily', run_time)
    period = info['period']
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder, folder_name)):  # dept_folder from import
        os.mkdir(join(dept_folder, folder_name))
    if not os.path.isdir(join(dept_folder, folder_name, period)):
        os.mkdir((join(dept_folder, folder_name, period)))

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viết Query ---------------------
    interest_query = pd.read_sql(
        f"""
            SELECT
                [account].[account_code],
                [relationship].[sub_account],
                [account].[customer_name],
                SUM([rcf11].[interest]) AS [lai_tren_TKKH],
                SUM([rcf02].[value]) AS [tien_lai_da_tra],
                SUM((([rci01].[closing_balance] * (0.1 / 100)) / 360)) AS [lai_tinh_lai]
            FROM [relationship]
            LEFT JOIN
                [account] 
            ON 
                [account].[account_code] = [relationship].[account_code]
            LEFT JOIN (
                SELECT 
                    [rcf0011].[sub_account], 
                    [rcf0011].[interest] 
                FROM [rcf0011]
                WHERE [rcf0011].[date] BETWEEN '{start_date}' AND '{end_date}'
            ) [rcf11]
            ON
                [rcf11].[sub_account] = [relationship].[sub_account]
            LEFT JOIN (
                SELECT 
                    [cash_balance].[date],
                    [cash_balance].[sub_account],
                    [cash_balance].[decrease] AS [value]
                FROM 
                    [cash_balance]
                WHERE
                    [cash_balance].[transaction_id] = '1190' 
                    AND [cash_balance].[date] BETWEEN '{start_date}' AND '{end_date}'
                UNION ALL
                SELECT
                    [cash_balance].[date],
                    [cash_balance].[sub_account],
                    [cash_balance].[increase] AS [value]
                FROM 
                    [cash_balance]
                WHERE
                    [cash_balance].[transaction_id] = '1191' 
                    AND [cash_balance].[date] BETWEEN '{start_date}' AND '{end_date}'
            ) [rcf02]
            ON [rcf02].[sub_account] = [relationship].[sub_account]
            LEFT JOIN (
                SELECT 
                    [sub_account_deposit].[sub_account],
                    [sub_account_deposit].[closing_balance]
                FROM
                    [sub_account_deposit]
                WHERE
                    [sub_account_deposit].[date] BETWEEN '{start_date}' AND '{end_date}'
            ) [rci01]
            ON [rci01].[sub_account] = [relationship].[sub_account]
            WHERE
                [relationship].[date] BETWEEN '{start_date}' AND '{end_date}'
                AND (
                    ([rcf11].[interest] IS NOT NULL OR [rcf11].[interest] <> 0)
                )
            GROUP BY [account].[account_code], [relationship].[sub_account], [account].[customer_name]
            ORDER BY [sub_account]
        """,
        connect_DWH_CoSo
    )
    interest_query = interest_query.dropna(subset=['tien_lai_da_tra', 'lai_tinh_lai'], how='all')
    interest_query.fillna(0, inplace=True)
    # lọc những tài khoản có giá trị bằng 0 ở cả 3 cột lãi trên TKKH, tiền lãi đã trả, lãi tính lại
    # interest_query = interest_query.loc[
    #     ~(interest_query[
    #           ['lai_tren_TKKH', 'tien_lai_da_tra', 'lai_tinh_lai']
    #       ] == 0).all(axis=1)]
    interest_query['bat_thuong'] = interest_query['lai_tinh_lai'] - interest_query['lai_tren_TKKH']

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File Excel ---------------------
    # Write file BÁO CÁO ĐỐI CHIẾU LÃI TIỀN GỬI PHÁT SINH TRÊN TÀI KHOẢN KHÁCH HÀNG
    f_date = dt.datetime.strptime(start_date, "%Y-%m-%d").strftime("%d-%m-%Y")
    f_name = f'Đối chiếu lãi tiền gửi phát sinh trên tài khoản khách hàng {f_date}.xlsx'
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
            'valign': 'vbottom',
            'font_size': 10,
            'font_name': 'Times New Roman'
        }
    )
    text_left_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vbottom',
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
            'num_format': '#,##0.00_);(#,##0.00)'
        }
    )
    tien_lai_da_tra_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'top',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'num_format': '#,##0'
        }
    )
    footer_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 12,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    footer_dmy_format = workbook.add_format(
        {
            'italic': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman'
        }
    )
    sum_name_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 12,
            'font_name': 'Times New Roman'
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
    headers = [
        'STT',
        'Số tài khoản',
        'Số tiểu khoản',
        'Tên khách hàng',
        'Lãi tiền gửi cộng dồn trên TKKH',
        'Số tiền lãi đã trả',
        'Lãi tiền gửi cộng dồn tính lại',
        'Bất thường',
        'Chi nhánh quản lý',
        'Môi giới quản lý'
    ]

    companyAddress = 'Tầng 3, CR3-03A, 109 Tôn Dật Tiên, phường Tân Phú, Quận 7, Thành phố Hồ Chí Minh'
    sheet_title_name = 'BÁO CÁO ĐỐI CHIẾU LÃI TIỀN GỬI PHÁT SINH TRÊN TÀI KHOẢN KHÁCH HÀNG'
    report_s_date = dt.datetime.strptime(start_date, "%Y-%m-%d").strftime("%d/%m/%Y")
    report_e_date = dt.datetime.strptime(end_date, "%Y-%m-%d").strftime("%d/%m/%Y")
    sub_title_name = f'Từ ngày {report_s_date} đến ngày {report_e_date}'
    sum_name = 'Cộng phát sinh'

    # --------- sheet BAO CAO CAN LAM ---------
    sheet_bao_cao_can_lam = workbook.add_worksheet('BAO CAO CAN LAM')

    # Insert phu hung picture
    sheet_bao_cao_can_lam.insert_image('A1', './img/phu_hung.png', {'x_scale': 0.66, 'y_scale': 0.71})

    # Set Column Width and Row Height
    sheet_bao_cao_can_lam.set_column('A:A', 6.22)
    sheet_bao_cao_can_lam.set_column('B:B', 14.11)
    sheet_bao_cao_can_lam.set_column('C:C', 13.78)
    sheet_bao_cao_can_lam.set_column('D:D', 29.89)
    sheet_bao_cao_can_lam.set_column('E:E', 21.89)
    sheet_bao_cao_can_lam.set_column('F:F', 18.89)
    sheet_bao_cao_can_lam.set_column('G:G', 21.89)
    sheet_bao_cao_can_lam.set_column('H:H', 11.89)
    sheet_bao_cao_can_lam.set_column('I:I', 28.89)
    sheet_bao_cao_can_lam.set_column('J:J', 43.89)
    sheet_bao_cao_can_lam.set_row(6, 18)
    sheet_bao_cao_can_lam.set_row(10, 47.3)

    sum_start_row = interest_query.shape[0] + 12
    # merge row
    sheet_bao_cao_can_lam.merge_range('C1:J1', CompanyName, company_name_format)
    sheet_bao_cao_can_lam.merge_range('C2:J2', companyAddress, company_format)
    sheet_bao_cao_can_lam.merge_range('C3:J3', CompanyPhoneNumber, company_format)
    sheet_bao_cao_can_lam.merge_range('A7:J7', sheet_title_name, sheet_title_format)
    sheet_bao_cao_can_lam.merge_range('A8:J8', sub_title_name, from_to_format)
    sheet_bao_cao_can_lam.merge_range(
        f'H{sum_start_row + 2}:J{sum_start_row + 2}',
        f'Ngày  tháng  năm ',
        footer_dmy_format
    )
    sheet_bao_cao_can_lam.merge_range(
        f'A{sum_start_row + 3}:C{sum_start_row + 3}',
        'Người lập',
        footer_format
    )
    sheet_bao_cao_can_lam.merge_range(
        f'H{sum_start_row + 3}:J{sum_start_row + 3}',
        'Người duyệt',
        footer_format
    )

    # write row, column
    sheet_bao_cao_can_lam.write_row(
        'A4',
        [''] * len(headers),
        empty_row_format
    )
    sheet_bao_cao_can_lam.write_row('A11', headers, headers_format)
    sheet_bao_cao_can_lam.write_column(
        'A12',
        [int(i) for i in np.arange(interest_query.shape[0]) + 1],
        stt_col_format
    )
    sheet_bao_cao_can_lam.write_column(
        'B12',
        interest_query['account_code'],
        text_left_format
    )
    sheet_bao_cao_can_lam.write_column(
        'C12',
        interest_query['sub_account'],
        text_left_format
    )
    sheet_bao_cao_can_lam.write_column(
        'D12',
        interest_query['customer_name'],
        text_left_format
    )
    sheet_bao_cao_can_lam.write_column(
        'E12',
        interest_query['lai_tren_TKKH'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'F12',
        interest_query['tien_lai_da_tra'],
        tien_lai_da_tra_format
    )
    sheet_bao_cao_can_lam.write_column(
        'G12',
        interest_query['lai_tinh_lai'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'H12',
        interest_query['bat_thuong'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'I12',
        [''] * interest_query.shape[0],
        text_left_format
    )
    sheet_bao_cao_can_lam.write_column(
        'J12',
        [''] * interest_query.shape[0],
        text_left_format
    )

    sheet_bao_cao_can_lam.merge_range(
        f'A{sum_start_row}:D{sum_start_row}',
        sum_name,
        sum_name_format
    )
    sheet_bao_cao_can_lam.write(
        f'E{sum_start_row}',
        interest_query['lai_tren_TKKH'].sum(),
        sum_money_format
    )
    sheet_bao_cao_can_lam.write(
        f'F{sum_start_row}',
        interest_query['tien_lai_da_tra'].sum(),
        sum_money_format
    )
    sheet_bao_cao_can_lam.write(
        f'G{sum_start_row}',
        interest_query['lai_tinh_lai'].sum(),
        sum_money_format
    )
    sheet_bao_cao_can_lam.write(
        f'H{sum_start_row}',
        interest_query['bat_thuong'].sum(),
        sum_money_format
    )
    sheet_bao_cao_can_lam.merge_range(
        f'I{sum_start_row}:J{sum_start_row}',
        '',
        sum_name_format
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