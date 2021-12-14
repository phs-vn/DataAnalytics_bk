"""
1. quarterly, ngày cuối của kỳ trước = 2021-09-30
2. bản  chất của báo cáo này là mình xuất gốc từ RMR1062, bớt 1 số dòng va thêm 1 số dòng (Mrs. Tuyết)
3. Báo cáo RLN0006: lấy loại vay của trả trậm, margin, và bảo lãnh (bảo lãnh trừ tk tự doanh 022p... ) ra
4. Dựa theo file cách làm
    - cột H và cột O: bắt buộc phải khớp với báo cáo RLN0006
    - cột J: bắt buộc phải  khớp với báo cáo RCI0001
    - cột K: là số tiền còn có thể UTTB: cách lấy: Lấy ROD0040 (2 ngày làm việc cuối cùng
của tháng) cột giá trị bán trừ ( phí bán và thuế bán) trừ tiếp số tiền đã ứng (xuất
RCI0015, 2 ngày làm việc cuối cùng, chọn ngày bán là 2 ngày lam việc cuối cùng) là ra.
    - có thể LOẠI các tài khoản mà giá trị của cả 5 cột (H, J, K, L, O) đều bằng 0
    (chú ý: miễn sao các cột này có số và khớp với mấy báo cáo kia là dc,
    còn nếu nếu họ có số ở các cột này thì không dc phép loại)
"""
from reporting_tool.trading_service.thanhtoanbutru import *


def run(
        run_time=None,
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
    # query VCF0051
    query_vcf = pd.read_sql(
        f"""
            SELECT 
            DISTINCT
                [branch_name],
                [relationship].[account_code],
                [relationship].[sub_account],
                [account].[customer_name]
            FROM [relationship]
            LEFT JOIN [customer_information] 
            ON [customer_information].[sub_account] = [relationship].[sub_account]
            LEFT JOIN [branch] 
            ON [branch].[branch_id] = [relationship].[branch_id]
            LEFT JOIN [account] 
            ON [account].[account_code] = [relationship].[account_code]
            WHERE 
                [relationship].[date] = '{end_date}'
                AND [customer_information].[contract_type] LIKE '%MR%'
                AND [customer_information].[status] IN ('A', 'B')
        """,
        connect_DWH_CoSo,
        index_col='sub_account'
    )
    # query RMR1063
    query_rmr = pd.read_sql(
        f"""
            SELECT
                [sub_account], 
                [credit_line], 
                [rtt], 
                [buying_power], 
                [total_outstanding], 
                [total_cash], 
                [margin_value], 
                [asset_value], 
                [total_outstanding_and_interest]
            FROM [sub_account_report]
            WHERE [date] = '{end_date}'
        """,
        connect_DWH_CoSo,
        index_col='sub_account'
    )
    # query RCI0001
    query_rci01 = pd.read_sql(
        f"""
            SELECT 
                [sub_account], 
                [closing_balance]
            FROM [sub_account_deposit]
            WHERE [date] = '{end_date}'
            ORDER BY [sub_account]
        """,
        connect_DWH_CoSo,
        index_col='sub_account'
    )
    # query RLN0006
    query_rln = pd.read_sql(
        f"""
            SELECT 
                [account_code], 
                SUM([principal_outstanding]) AS [principal_outstanding], 
                (SUM([principal_outstanding]) + SUM([interest_outstanding])) AS [total_outstanding_plus_interest]
            FROM [margin_outstanding]
            WHERE 
                [date] = '{end_date}'
                AND [type] NOT IN (N'Ứng trước cổ tức', N'Cầm cố')
                AND [account_code] <> '022P002222'
            GROUP BY [account_code]
        """,
        connect_DWH_CoSo,
        index_col='account_code'
    )
    # query ROD0040
    query_rod = pd.read_sql(
        f"""
            SELECT 
                [sub_account], 
                sum([value]) AS [value], 
                sum([fee]) AS [fee], 
                sum([tax_of_selling]) AS [tax_of_selling]
            FROM [trading_record]
            WHERE [date] BETWEEN '{bdate(end_date, -1)}' AND '{end_date}'
            AND [type_of_order] = 'S'
            GROUP BY [date], [sub_account]
            ORDER BY [sub_account]
        """,
        connect_DWH_CoSo,
        index_col='sub_account'
    )
    query_rod = query_rod.groupby(query_rod.index).sum()
    query_rod['calc_rod'] = query_rod['value'] - (query_rod['fee'] + query_rod['tax_of_selling'])
    # query RCI0015
    query_rci15 = pd.read_sql(
        f"""
            SELECT  
                [sub_account],
                SUM([receivable]) AS [receivable]
            FROM [payment_in_advance]
            WHERE [date] BETWEEN '{bdate(end_date, -1)}' AND '{end_date}'
            AND [trading_date] BETWEEN '{bdate(end_date, -1)}' AND '{end_date}'
            GROUP BY [date], [sub_account], [trading_date]
        """,
        connect_DWH_CoSo,
        index_col='sub_account'
    )
    query_rci15 = query_rci15.groupby(query_rci15.index).sum()

    # Xử lý dataframe
    query_rmr['total_cash_rci01'] = query_rci01['closing_balance']
    final_table = pd.concat([query_vcf, query_rmr], axis=1)
    final_table = final_table.merge(query_rln, how='outer', left_on='account_code', right_index=True)
    final_table['rod'] = query_rod['calc_rod']
    final_table['rci15'] = query_rci15['receivable']
    final_table['uttb_con_lai'] = final_table['rod'] - final_table['rci15']
    final_table = final_table.dropna(subset=['branch_name', 'account_code', 'customer_name'])
    final_table['rai_mortgage_ratio'] = 50
    final_table.fillna(0, inplace=True)

    a = final_table.loc[
        (final_table['buying_power'] == 0) &
        (final_table['total_outstanding'] == 0) &
        (final_table['total_cash'] == 0) &
        (final_table['total_outstanding_and_interest'] == 0) &
        (final_table['uttb_con_lai'] == 0)
    ].index
    for i in a:
        final_table.drop(i, inplace=True)

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

    sum_start_row = final_table.shape[0] + 3
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
        [int(i) for i in np.arange(final_table.shape[0]) + 1],
        stt_col_format
    )
    sheet1.write_column(
        'B3',
        final_table['branch_name'],
        text_left_format
    )
    sheet1.write_column(
        'C3',
        final_table['account_code'],
        text_left_format
    )
    sheet1.write_column(
        'D3',
        final_table.index,
        text_left_format
    )
    sheet1.write_column(
        'E3',
        final_table['customer_name'],
        text_left_format
    )
    sheet1.write_column(
        'F3',
        final_table['credit_line'],
        money_format
    )
    sheet1.write_column(
        'G3',
        final_table['rtt'],
        money_format
    )
    sheet1.write_column(
        'H3',
        final_table['total_cash_rci01'],
        money_format
    )
    sheet1.write_column(
        'I3',
        final_table['uttb_con_lai'],
        money_format
    )
    sheet1.write_column(
        'J3',
        final_table['buying_power'],
        money_format
    )
    sheet1.write_column(
        'K3',
        final_table['total_outstanding'],
        money_format
    )
    sheet1.write_column(
        'L3',
        final_table['total_cash'],
        money_format
    )
    sheet1.write_column(
        'M3',
        final_table['margin_value'],
        money_format
    )
    sheet1.write_column(
        'N3',
        final_table['asset_value'],
        money_format
    )
    sheet1.write_column(
        'O3',
        final_table['total_outstanding_and_interest'],
        money_format
    )
    sheet1.write_column(
        'P3',
        final_table['rai_mortgage_ratio'],
        money_format
    )
    # Row sum money
    sheet1.write(
        f'F{sum_start_row}',
        final_table['credit_line'].sum(),
        sum_money_format
    )
    sheet1.write(
        f'G{sum_start_row}',
        final_table['rtt'].sum(),
        sum_money_format
    )
    sheet1.write(
        f'H{sum_start_row}',
        final_table['total_cash_rci01'].sum(),
        sum_money_format
    )
    sheet1.write(
        f'I{sum_start_row}',
        final_table['uttb_con_lai'].sum(),
        sum_money_format
    )
    sheet1.write(
        f'J{sum_start_row}',
        final_table['buying_power'].sum(),
        sum_money_format
    )
    sheet1.write(
        f'K{sum_start_row}',
        final_table['total_outstanding'].sum(),
        sum_money_format
    )
    sheet1.write(
        f'L{sum_start_row}',
        final_table['total_cash'].sum(),
        sum_money_format
    )
    sheet1.write(
        f'M{sum_start_row}',
        final_table['margin_value'].sum(),
        sum_money_format
    )
    sheet1.write(
        f'N{sum_start_row}',
        final_table['asset_value'].sum(),
        sum_money_format
    )
    sheet1.write(
        f'O{sum_start_row}',
        final_table['total_outstanding_and_interest'].sum(),
        sum_money_format
    )
    sheet1.write(
        f'P{sum_start_row}',
        final_table['rai_mortgage_ratio'].sum(),
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

