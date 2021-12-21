"""
    1. Các cột G, H, I phải tính mới ra được kết quả. Tất cả đều có công thức trong file mẫu
    2. Nếu là GOLD PHS và SILV PHS thì cứ 6 tháng set 1 lần (VD: 01/01/2021 tới 30/06/2021)
       Nếu là VIP Branch thì 3 tháng set 1 lần (từ đầu quý trước tới cuối quý trước)
       - Ví dụ báo cáo suất vào quý 3 của năm (30/9/2021) --> chỉ xét 3 tháng
       hoặc dễ hiểu hơn là chỉ review VIP CN (từ 01/07 tới 30/09)
       - Nếu báo cáo suất vào quý 2 hoặc quý 4 --> xét cả 3 tháng và 6 tháng
       hoặc dễ hiểu hơn là review cả VIP CN, GOLD, SILV (
            ví dụ quý 2: 30/06
            --> xét 3 tháng: từ 01/04 tới 30/06, xét 6 tháng: từ 01/01/2021 tới 30/06/2021)
            quý 4: 31/12
            --> xét 3 tháng: từ 01/10 tới 31/12, xét 6 tháng: từ 01/07/2021 tới 31/12/2021)
"""
from reporting_tool.trading_service.thanhtoanbutru import *


# DONE
def run(
        run_time=None,
):
    start = time.time()
    info = get_info('quarterly', run_time)
    start_date = info['start_date']
    end_date = info['end_date']
    period = info['period']
    folder_name = info['folder_name']

    if run_time is None:
        run_time = dt.datetime.now()

    # create folder
    if not os.path.isdir(join(dept_folder, folder_name, period)):
        os.mkdir(join(dept_folder, folder_name, period))

    ###################################################
    ###################################################
    ###################################################
    # month of end_date
    moe = dt.datetime.strptime(end_date, "%Y/%m/%d")

    review_vip = pd.read_sql(
        f"""
        SELECT
            CONCAT(N'Tháng ',MONTH([account].[date_of_birth])) [birth_month],
            [account].[account_code], 
            [account].[customer_name], 
            [branch].[branch_name], 
            [account].[date_of_birth], 
            [customer_information].[contract_type],
            [customer_information_change].[date_of_change],
            [customer_information_change].[time_of_change],
            [customer_information_change].[date_of_approval],
            [customer_information_change].[time_of_approval],
            [broker].[broker_id],
            [broker].[broker_name], 
            [customer_information_change].[change_content]
        FROM 
            [customer_information]
        LEFT JOIN 
            (SELECT * FROM [relationship] WHERE [relationship].[date] = '{end_date}') [relationship]
        ON 
            [customer_information].[sub_account] = [relationship].[sub_account]
        LEFT JOIN 
            [account]
        ON 
            [relationship].[account_code] = [account].[account_code]
        LEFT JOIN 
            [broker]
        ON 
            [relationship].[broker_id] = [broker].[broker_id]
        LEFT JOIN
            [branch]
        ON
            [relationship].[branch_id] = [branch].[branch_id]
        LEFT JOIN
            [customer_information_change]
        ON
            [customer_information_change].[account_code] = [relationship].[account_code]
        WHERE (
            [customer_information].[contract_type] LIKE N'%GOLD%' 
            OR [customer_information].[contract_type] LIKE N'%SILV%' 
            OR [customer_information].[contract_type] LIKE N'%VIPCN%'
        )
        AND -- De chay baktest
            [account].[date_of_open] <= '{end_date}' 
        AND -- De chay baktest
            ([account].[date_of_close] > '{end_date}' OR [account].[date_of_close] IS NULL OR [account].[date_of_close] = '2099-12-31')
        AND 
            [customer_information_change].[change_content] = 'Loai hinh hop dong'
        ORDER BY
            [birth_month], [date_of_birth]
        """,
        connect_DWH_CoSo,
        index_col='account_code'
    )
    review_vip.drop_duplicates(inplace=True)

    # Groupby and condition
    review_vip['effective_date'] = review_vip.groupby('account_code')['date_of_change'].min()
    review_vip['approved_date'] = review_vip.groupby('account_code')['date_of_approval'].max()
    mask = (
                review_vip['date_of_change'] == review_vip['effective_date']
           ) | (
                review_vip['date_of_approval'] == review_vip['approved_date'])
    review_vip = review_vip.loc[mask]

    review_vip = review_vip[[
        'birth_month',
        'customer_name',
        'branch_name',
        'date_of_birth',
        'contract_type',
        'broker_id',
        'broker_name',
        'effective_date',
        'approved_date',
    ]]
    # convert cột contract_type
    review_vip['contract_type'] = review_vip['contract_type'].apply(
        lambda x: 'SILV PHS' if 'SILV' in x else ('GOLD PHS' if 'GOLD' in x else 'VIP Branch')
    )
    review_vip.drop_duplicates(keep='last', inplace=True)

    # điều kiện phân biệt giữa SILV, GOLD và VIP Branch
    for idx in review_vip.index:
        contract_type = review_vip.loc[idx, 'contract_type']
        if contract_type == 'SILV PHS' or contract_type == 'GOLD PHS':
            if 1 <= run_time.month < 6:
                review_vip.loc[idx, 'review_date'] = dt.datetime(run_time.year, 6, 30).strftime("%d/%m/%Y")
            elif 6 <= run_time.month < 12:
                review_vip.loc[idx, 'review_date'] = dt.datetime(run_time.year, 12, 31).strftime("%d/%m/%Y")
            else:
                review_vip.loc[idx, 'review_date'] = dt.datetime(run_time.year+1, 6, 30).strftime("%d/%m/%Y")
        elif contract_type == 'VIP Branch':
            if 1 <= run_time.month < 3:
                review_vip.loc[idx, 'review_date'] = dt.datetime(run_time.year, 3, 31).strftime("%d/%m/%Y")
            elif 3 <= run_time.month < 6:
                review_vip.loc[idx, 'review_date'] = dt.datetime(run_time.year, 6, 30).strftime("%d/%m/%Y")
            elif 6 <= run_time.month < 9:
                review_vip.loc[idx, 'review_date'] = dt.datetime(run_time.year, 9, 30).strftime("%d/%m/%Y")
            else:
                review_vip.loc[idx, 'review_date'] = dt.datetime(run_time.year+1, 3, 31).strftime("%d/%m/%Y")

    branch_groupby_table = review_vip.groupby(['branch_name', 'contract_type'])['customer_name'].count().unstack()
    branch_groupby_table.fillna(0, inplace=True)
    review_vip['criteria_fee'] = review_vip['contract_type'].apply(
        lambda x: 40000000 if ('GOLD' in x or 'VIP Branch' in x) else 20000000
    )
    review_vip_phs = review_vip.loc[
        (review_vip['contract_type'] == 'GOLD PHS')|(review_vip['contract_type'] == 'SILV PHS')].copy()
    review_vip_branch = review_vip.loc[(review_vip['contract_type'] == 'VIP Branch')].copy()
    # Query RCF3001
    # setup 6 tháng
    query_rcf01_6m = pd.read_sql(
        f"""
            SELECT
                [account_code],
                SUM([trading_fee])[trading_fee],
                SUM([loan_fee])[loan_fee]
            FROM [revenue_from_vip]
            WHERE 
                [date] BETWEEN '2021-01-01' AND '2021-06-30'
            GROUP BY [account_code]
            ORDER BY [account_code]
        """,
        connect_DWH_CoSo,
        index_col='account_code'
    )
    query_rcf01_3m = pd.read_sql(
        f"""
            SELECT
                [account_code],
                SUM([trading_fee])[trading_fee],
                SUM([loan_fee])[loan_fee]
            FROM [revenue_from_vip]
            WHERE 
                [date] BETWEEN '2021-04-01' AND '2021-06-30'
            GROUP BY [account_code]
            ORDER BY [account_code]
        """,
        connect_DWH_CoSo,
        index_col='account_code'
    )
    review_vip_phs['trading_fee'] = query_rcf01_6m['trading_fee']
    review_vip_branch['trading_fee'] = query_rcf01_3m['trading_fee']

    # --------------------- Viet File ---------------------
    # Write file excel Báo cáo review KH vip
    file_name = f'Báo cáo review KH vip.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder, folder_name, period, file_name),
        engine='xlsxwriter',
        engine_kwargs={'options': {'nan_inf_to_errors': True}}
    )
    workbook = writer.book

    ###################################################
    ###################################################
    ###################################################

    sheet_title_format = workbook.add_format(
        {
            'bold': True,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 20
        }
    )
    sub_title_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        }
    )
    sub_title_1_format = workbook.add_format(
        {
            'bold': True,
            'font_name': 'Times New Roman',
            'font_size': 14,
        }
    )
    sub_title_2_format = workbook.add_format(
        {
            'italic': True,
            'font_name': 'Times New Roman',
            'font_size': 14,
        }
    )
    no_and_date_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 12,
        }
    )
    description_format = workbook.add_format(
        {
            'bold': True,
            'align': 'left',
            'valign': 'top',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 12,
        }
    )
    header_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 8,
            'bg_color': '#FFC000'
        }
    )
    header_2_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 8,
            'bg_color': '#FFFF00'
        }
    )
    header_3_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 8,
        }
    )
    empty_format = workbook.add_format(
        {
            'border': 1,
            'text_wrap': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 12,
        }
    )
    stt_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 8,
        }
    )
    text_left_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 8,
        }
    )
    text_left_wrap_text_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 8,
        }
    )
    date_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'num_format': 'dd/mm/yyyy',
            'font_name': 'Times New Roman',
            'font_size': 8
        }
    )
    fee_and_value_format = workbook.add_format(
        {
            'border': 1,
            'text_wrap': True,
            'align': 'right',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 8,
            'num_format': '(#,##0.0_);_( (#,##0.0);_(* "-"??_);_(@)'
        }
    )
    num_percent_format = workbook.add_format(
        {
            'border': 1,
            'text_wrap': True,
            'align': 'right',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 8,
            'num_format': '0.00"%"'
        }
    )
    footer_format = workbook.add_format(
        {
            'text_wrap': True,
            'align': 'left',
            'valign': 'top',
            'font_name': 'Times New Roman',
            'font_size': 11,
        }
    )
    table_header_format = workbook.add_format(
        {
            'bold': True,
            'border': 1,
            'text_wrap': True,
            'align': 'center',
            'valign': 'top',
            'font_name': 'Times New Roman',
            'font_size': 11,
        }
    )
    table_sub_header_format = workbook.add_format(
        {
            'bold': True,
            'border': 1,
            'text_wrap': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 9,
        }
    )
    table_content_format =  workbook.add_format(
        {
            'border': 1,
            'text_wrap': True,
            'align': 'left',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 11,
        }
    )
    date_header = dt.datetime.strptime(end_date, "%Y/%m/%d").strftime("%d.%m.%Y")
    headers = [
        'No.',
        'Account',
        'Name',
        'Branch',
        'Approve day',
        'Criteria Fee',
        'Fee for assessment',
        '% Fee for assessment / Criteria Fee',
        'Average Net Asset Value',
        'Current VIP',
        'After review',
        'Rate',
        f'Group & Deal đợt {date_header}',
        'Opinion of Location Manager',
        'Opinion of Trading Service',
        'Decision of  Deputy General Director',
        'MOI GIOI QUAN LY',
        'NOTE'
    ]

    #  Viết Sheet VIP PHS
    review_vip_sheet = workbook.add_worksheet('REVIEW VIP')
    review_vip_sheet.hide_gridlines(option=2)

    # Content of sheet
    sheet_title_name = 'SUBMISSION'
    description = 'To: Deputy General Director of Phu Hung Securities Corporation\nProposer: Nguyen Thi Tuyet'
    review_vip_sheet.merge_range('A2:C3', '', empty_format)
    review_vip_sheet.insert_image('A2', './img/phu_hung.png', {'x_scale': 1.27, 'y_scale': 0.52})

    # Set Column Width and Row Height
    review_vip_sheet.set_column('A:A', 4)
    review_vip_sheet.set_column('B:B', 10.43)
    review_vip_sheet.set_column('C:C', 21.86)
    review_vip_sheet.set_column('D:D', 9.71)
    review_vip_sheet.set_column('E:E', 10.86)
    review_vip_sheet.set_column('F:F', 12.43)
    review_vip_sheet.set_column('G:G', 15)
    review_vip_sheet.set_column('H:H', 12.57)
    review_vip_sheet.set_column('I:I', 19.71)
    review_vip_sheet.set_column('J:J', 21.14)
    review_vip_sheet.set_column('K:K', 16)
    review_vip_sheet.set_column('L:L', 12.14)
    review_vip_sheet.set_column('M:N', 19)
    review_vip_sheet.set_column('O:O', 17.14)
    review_vip_sheet.set_column('P:P', 27.86)
    review_vip_sheet.set_column('Q:Q', 23.43)
    review_vip_sheet.set_column('R:R', 14.14)

    review_vip_sheet.set_row(1, 25.5)
    review_vip_sheet.set_row(2, 18.75)
    review_vip_sheet.set_row(4, 36.5)
    review_vip_sheet.set_row(5, 46.5)

    # merge row
    review_vip_sheet.merge_range('D2:M2', sheet_title_name, sheet_title_format)
    report_date = dt.datetime.strptime(end_date, "%Y/%m/%d").strftime("%B, %Y")
    review_vip_sheet.merge_range('D3:M3', '', empty_format)
    review_vip_sheet.write_rich_string(
        'D3',
        sub_title_1_format,
        'Subject :',
        sub_title_2_format,
        f' REVIEW VIP (THE END OF {report_date.upper()})',
        sub_title_format,
    )

    review_vip_sheet.write('N2', 'No.:', no_and_date_format)
    review_vip_sheet.write('N3', 'Date:', no_and_date_format)
    review_vip_sheet.merge_range('O2:P2', '129/2021/TTr-TRS', no_and_date_format)
    review_vip_sheet.merge_range('O3:P3', f'{run_time.strftime("%d/%m/%Y")}', no_and_date_format)

    review_vip_sheet.merge_range('A5:P5', description, description_format)

    review_vip_sheet.write_row('A6', headers, header_format)
    review_vip_sheet.write('G6', headers[6], header_2_format)
    review_vip_sheet.write('H6', headers[7], header_2_format)
    review_vip_sheet.write('I6', headers[8], header_2_format)
    review_vip_sheet.write('Q6', headers[-2], header_3_format)
    review_vip_sheet.write('R6', headers[-1], header_2_format)

    review_vip_sheet.write_column(
        'A7',
        np.arange(review_vip.shape[0]) + 1,
        stt_format
    )
    review_vip_sheet.write_column(
        'B7',
        review_vip.index,
        text_left_format
    )
    review_vip_sheet.write_column(
        'C7',
        review_vip['customer_name'],
        text_left_wrap_text_format
    )
    review_vip_sheet.write_column(
        'D7',
        review_vip['branch_name'],
        text_left_format
    )
    review_vip_sheet.write_column(
        'E7',
        review_vip['approved_date'],
        date_format
    )
    review_vip_sheet.write_column(
        'F7',
        review_vip['criteria_fee'],
        fee_and_value_format
    )
    review_vip_sheet.write_column(
        'J7',
        review_vip['contract_type'],
        text_left_format
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