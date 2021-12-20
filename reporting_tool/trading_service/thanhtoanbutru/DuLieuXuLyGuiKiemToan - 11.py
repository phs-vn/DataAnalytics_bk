"""
1. quarterly, ngày cuối của kỳ trước = 2021-09-30
2. bản chất của báo cáo này là mình xuất gốc từ RMR1062, bớt 1 số dòng va thêm 1 số dòng (Mrs. Tuyết)
3. Báo cáo RLN0006: lấy loại vay của trả trậm, margin, và bảo lãnh (bảo lãnh trừ tk tự doanh 022p... ) ra
4. Dựa theo file cách làm
    - cột H và cột O: bắt buộc phải khớp với báo cáo RLN0006
    - cột J: bắt buộc phải  khớp với báo cáo RCI0001
    - cột K: là số tiền còn có thể UTTB: cách lấy: Lấy ROD0040 (2 ngày làm việc cuối cùng
của tháng) cột giá trị bán trừ ( phí bán và thuế bán) trừ tiếp số tiền đã ứng (xuất
RCI0015, 2 ngày làm việc cuối cùng, chọn ngày bán là 2 ngày làm việc cuối cùng) là ra.
    - có thể LOẠI các tài khoản mà giá trị của cả 5 cột (H, J, K, L, O) đều bằng 0
    (chú ý: miễn sao các cột này có số và khớp với mấy báo cáo kia là dc,
    còn nếu nếu họ có số ở các cột này thì không dc phép loại)
    - kết quả xuất ra có thể ít hơn hoặc bằng so với file kết quả của báo cáo
"""
from reporting_tool.trading_service.thanhtoanbutru import *


def run(
        run_time=None
):
    start = time.time()
    info = get_info('quarterly', run_time)
    end_date = info['end_date'].replace('/', '-')
    period = info['period']
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder, folder_name)):
        os.mkdir(join(dept_folder, folder_name))
    if not os.path.isdir(join(dept_folder, folder_name, period)):
        os.mkdir((join(dept_folder, folder_name, period)))

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viết Query và xử lý dataframe ---------------------
    # query RMR1062 và VCF0051
    query_rmr = pd.read_sql(
        f"""
            SELECT
                [relationship].[date],
                [branch].[branch_name],
                [relationship].[account_code],
                [relationship].[sub_account],
                [account].[customer_name],
                [rmr62].[credit_line],
                ISNULL([rmr63].[rtt], 0) [rtt],
                ISNULL([rci01].[closing_balance], 0) [total_cash_RCI01],
                (ISNULL([rod40].[calc_rod], 0) - ISNULL([rci15].[receivable], 0)) [uttb_con_lai],
                ISNULL([rmr63].[buying_power], 0) [buying_power],
                [rmr62].[total_outstanding],
                [rmr62].[total_cash],
                [rmr62].[total_margin_value],
                [rmr62].[total_asset_value],
                [rmr62].[total_outstanding_plust_interest],
                50 AS [rai_mortgage_ratio]
            FROM [relationship]
            LEFT JOIN ( 
                SELECT
                    [sub_account],
                    [status],
                    [contract_type]
                FROM
                    [customer_information]
            ) [vcf51]
            ON [vcf51].[sub_account] = [relationship].[sub_account]
            LEFT JOIN (
                SELECT
                    [branch].[branch_id],
                    [branch].[branch_name]
                FROM [branch]
            ) [branch]
            ON [branch].[branch_id] = [relationship].[branch_id]
            LEFT JOIN (
                SELECT
                    [account].[account_code],
                    [account].[customer_name]
                FROM
                    [account]
            ) [account]
            ON [account].[account_code] = [relationship].[account_code]
            RIGHT JOIN (
                SELECT 
                    [rmr1062].[account_code], 
                    [rmr1062].[credit_line], 
                    [rmr1062].[total_outstanding], 
                    [rmr1062].[total_cash], 
                    [rmr1062].[total_margin_value], 
                    [rmr1062].[total_asset_value], 
                    [rmr1062].[total_outstanding_plust_interest]
                FROM
                    [rmr1062]
                WHERE
                    [rmr1062].[date] = '{end_date}'
            ) [rmr62]
            ON [rmr62].[account_code] = [relationship].[account_code]
            LEFT JOIN(
                SELECT
                    [sub_account_report].[sub_account],
                    [sub_account_report].[buying_power],
                    [sub_account_report].[rtt]
                FROM
                    [sub_account_report]
                WHERE
                    [sub_account_report].[date] = '{end_date}'
            ) [rmr63]
            ON [rmr63].[sub_account] = [relationship].[sub_account]
            LEFT JOIN (
                SELECT
                    [sub_account],
                    [closing_balance]
                FROM 
                    [sub_account_deposit]
                WHERE 
                    [sub_account_deposit].[date] = '{end_date}'
            )[rci01]
            ON [rci01].[sub_account] = [relationship].[sub_account]
            LEFT JOIN (
                SELECT 
                    [sub_account], 
                    (SUM([value]) - SUM([fee]) - SUM([tax_of_selling])) AS [calc_rod]
                FROM [trading_record]
                WHERE [date] BETWEEN '{bdate(end_date, -1)}' AND '{end_date}'
                AND [type_of_order] = 'S'
                GROUP BY [sub_account]
            ) [rod40]
            ON [rod40].[sub_account] = [relationship].[sub_account]
            LEFT JOIN(
                SELECT
                    [payment_in_advance].[sub_account],
                    SUM([payment_in_advance].[receivable]) AS [receivable]
                FROM
                    [payment_in_advance]
                WHERE
                    [payment_in_advance].[date] BETWEEN '{bdate(end_date, -1)}' AND '{end_date}'
                AND [trading_date] BETWEEN '{bdate(end_date, -1)}' AND '{end_date}'
                GROUP BY [sub_account]
            ) [rci15]
            ON [rci15].[sub_account] = [relationship].[sub_account]
            WHERE
                [relationship].[date] = '{end_date}'
                AND [vcf51].[contract_type] like '%MR%'
                AND [vcf51].[status] in ('A', 'B')
            ORDER BY [account_code]
        """,
        connect_DWH_CoSo,
        index_col='sub_account'
    )
    # lọc những tài khoản có value=0 ở cả 5 cột
    # 'UTTB còn lại, buying power, total outstanding, total cash, total outstanding plus total interest'
    query_rmr = query_rmr.loc[
        ~(query_rmr[
              [
                  'credit_line',
                  'rtt',
                  'buying_power'
              ]
          ] == 0).all(axis=1)]
    query_rmr = query_rmr.loc[
        ~(query_rmr[
              [
                  'uttb_con_lai',
                  'buying_power',
                  'total_outstanding',
                  'total_cash',
                  'total_outstanding_plust_interest'
              ]
          ] == 0).all(axis=1)]

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File Excel ---------------------
    # Write file BÁO CÁO DỮ LIỆU XỬ LÝ GỬI KIỂM TOÁN
    f_date = dt.datetime.strptime(end_date, "%Y-%m-%d").strftime("%d-%m-%Y")
    f_name = f'RMR1062 {f_date}.xlsx'
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
    info_company_format = workbook.add_format(
        {
            'align': 'left',
            'valign': 'top',
            'font_size': 10,
            'font_name': 'Arial',
        }
    )
    headers_format = workbook.add_format(
        {
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Arial',
            'text_wrap': True
        }
    )
    stt_col_format = workbook.add_format(
        {
            'align': 'right',
            'valign': 'top',
            'font_size': 10,
            'font_name': 'Arial'
        }
    )
    text_left_format = workbook.add_format(
        {
            'align': 'left',
            'valign': 'top',
            'font_size': 10,
            'font_name': 'Arial'
        }
    )
    money_format = workbook.add_format(
        {
            'align': 'right',
            'valign': 'top',
            'font_size': 10,
            'font_name': 'Arial',
            'num_format': '_(* #,##0_);_(* (#,##0)'
        }
    )
    sum_name_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'valign': 'top',
            'font_size': 10,
            'font_name': 'Arial',
        }
    )
    sum_money_format = workbook.add_format(
        {
            'bold': True,
            'align': 'right',
            'valign': 'top',
            'font_size': 10,
            'font_name': 'Arial',
            'num_format': '_(* #,##0_);_(* (#,##0)'
        }
    )
    headers = [
        'No.',
        'Branch',
        'Account',
        'Tiểu khoản MR',
        'Name',
        'Creditline',
        'Margin ratio',
        'Total Cash',
        'UTTB còn lại',
        'Buying power',
        'Total Outstanding',
        'Total Cash',
        'Total Margin Value',
        'Total Asset Value',
        'Total Outstanding plus total interest',
        'Rai - Mortgage ratio of Account'
    ]
    sheet_title_name = 'ACCOUNT DETAILS REPORT'
    report_date = dt.datetime.strptime(end_date, "%Y-%m-%d").strftime("%d/%m/%Y")
    date_col = f'Date: {report_date}'

    # ---------------------------- sheet BAO CAO CAN LAM ----------------------------
    sheet1 = workbook.add_worksheet('Sheet1')
    # Set Column Width and Row Height
    sheet1.set_column('A:A', 6.11)
    sheet1.set_column('B:B', 13.22)
    sheet1.set_column('C:D', 11.56)
    sheet1.set_column('E:E', 24.33)
    sheet1.set_column('F:F', 25.89)
    sheet1.set_column('G:G', 11.56)
    sheet1.set_column('H:P', 18.56)
    sheet1.set_row(1, 39.6)

    sum_start_row = query_rmr.shape[0] + 3
    sheet1.merge_range(f'A{sum_start_row}:E{sum_start_row}', 'TỔNG', sum_name_format)

    sheet1.write(
        'A1',
        CompanyName,
        info_company_format
    )
    sheet1.write(
        'B1',
        CompanyAddress,
        info_company_format
    )
    sheet1.write(
        'C1',
        CompanyPhoneNumber,
        info_company_format
    )
    sheet1.write(
        'D1',
        sheet_title_name,
        info_company_format
    )
    sheet1.write(
        'F1',
        date_col,
        info_company_format
    )
    sheet1.write_row(
        'A2',
        headers,
        headers_format
    )
    sheet1.write_column(
        'A3',
        [int(i) for i in np.arange(query_rmr.shape[0]) + 1],
        stt_col_format
    )
    sheet1.write_column(
        'B3',
        query_rmr['branch_name'],
        text_left_format
    )
    sheet1.write_column(
        'C3',
        query_rmr['account_code'],
        text_left_format
    )
    sheet1.write_column(
        'D3',
        query_rmr.index,
        text_left_format
    )
    sheet1.write_column(
        'E3',
        query_rmr['customer_name'],
        text_left_format
    )
    sheet1.write_column(
        'F3',
        query_rmr['credit_line'],
        money_format
    )
    sheet1.write_column(
        'G3',
        query_rmr['rtt'],
        money_format
    )
    sheet1.write_column(
        'H3',
        query_rmr['total_cash_RCI01'],
        money_format
    )
    sheet1.write_column(
        'I3',
        query_rmr['uttb_con_lai'],
        money_format
    )
    sheet1.write_column(
        'J3',
        query_rmr['buying_power'],
        money_format
    )
    sheet1.write_column(
        'K3',
        query_rmr['total_outstanding'],
        money_format
    )
    sheet1.write_column(
        'L3',
        query_rmr['total_cash'],
        money_format
    )
    sheet1.write_column(
        'M3',
        query_rmr['total_margin_value'],
        money_format
    )
    sheet1.write_column(
        'N3',
        query_rmr['total_asset_value'],
        money_format
    )
    sheet1.write_column(
        'O3',
        query_rmr['total_outstanding_plust_interest'],
        money_format
    )
    sheet1.write_column(
        'P3',
        query_rmr['rai_mortgage_ratio'],
        money_format
    )
    # Row sum money
    sheet1.write(
        f'F{sum_start_row}',
        query_rmr['credit_line'].sum(),
        sum_money_format
    )
    sheet1.write(
        f'G{sum_start_row}',
        query_rmr['rtt'].sum(),
        sum_money_format
    )
    sheet1.write(
        f'H{sum_start_row}',
        query_rmr['total_cash_RCI01'].sum(),
        sum_money_format
    )
    sheet1.write(
        f'I{sum_start_row}',
        query_rmr['uttb_con_lai'].sum(),
        sum_money_format
    )
    sheet1.write(
        f'J{sum_start_row}',
        query_rmr['buying_power'].sum(),
        sum_money_format
    )
    sheet1.write(
        f'K{sum_start_row}',
        query_rmr['total_outstanding'].sum(),
        sum_money_format
    )
    sheet1.write(
        f'L{sum_start_row}',
        query_rmr['total_cash'].sum(),
        sum_money_format
    )
    sheet1.write(
        f'M{sum_start_row}',
        query_rmr['total_margin_value'].sum(),
        sum_money_format
    )
    sheet1.write(
        f'N{sum_start_row}',
        query_rmr['total_asset_value'].sum(),
        sum_money_format
    )
    sheet1.write(
        f'O{sum_start_row}',
        query_rmr['total_outstanding_plust_interest'].sum(),
        sum_money_format
    )
    sheet1.write(
        f'P{sum_start_row}',
        query_rmr['rai_mortgage_ratio'].sum(),
        sum_money_format
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

