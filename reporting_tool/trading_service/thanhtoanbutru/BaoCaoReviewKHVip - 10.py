"""
    1. Các cột G, H, I phải tính mới ra được kết quả. Tất cả đều có công thức trong file mẫu
    2. Nguyên tắc là VIP CN thì 3 tháng review 1 lần chốt cuối mỗi quý thì từ thời gian chốt
    đó tính ngược lại 3 tháng cứ thế mà đếm ra, còn GOLD và SILV thì sáu tháng tức là 1 năm 2 lần
    cuối tháng 6 và cuối tháng 12 cứ thế mà đếm ngược lại.
    Nguyên tắc thời gian, nếu báo cáo chạy vào thời điểm quý 2, quý 4 là sẽ xét toàn bộ KH (gold, vip, silv),
    còn nếu chạy vào thời điểm quý 1, quý 3 thì chỉ xét KH VIP CN
    3. Các cột M, N, O, P lấy ý kiến từ các phòng ban khác
    4. Cột Note để trắng cho CN họ note cái gì thì note.
       Cột MOI GIOI QUAN LY thì lên dữ liệu cho chị Tuyết
    5. TK 022C036979 có contract type là 'MR - KHTN GOLD 70% - OD 0.15% - Margin. PIA 10% - MR 2 - DP 5 (Lãi suất)'
"""
from reporting_tool.trading_service.thanhtoanbutru import *


# DONE
def run(
        run_time=None,
):
    start = time.time()
    info = get_info('quarterly', run_time)
    start_date = info['start_date'].replace('/', '-')
    end_date = info['end_date'].replace('/', '-')
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

    sod = start_date

    # query danh sách khách hàng
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
            (
                [account].[date_of_close] > '{end_date}' 
                OR [account].[date_of_close] IS NULL 
                OR [account].[date_of_close] = '2099-12-31'
            )
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
    mask = (review_vip['date_of_change']==review_vip['effective_date'])|(review_vip['date_of_approval']==review_vip['approved_date'])
    review_vip = review_vip.loc[mask]

    review_vip = review_vip[[
        'customer_name',
        'branch_name',
        'contract_type',
        'broker_name',
        'approved_date',
    ]]

    review_vip['approved_date'] = pd.to_datetime(review_vip['approved_date']).dt.date
    # convert cột contract_type
    review_vip['current_vip'] = review_vip['contract_type'].apply(
        lambda x: 'SILV PHS' if 'SILV' in x else ('GOLD PHS' if 'GOLD' in x else 'VIP Branch')
    )
    review_vip.drop_duplicates(keep='last', inplace=True)

    review_vip['criteria_fee'] = review_vip['current_vip'].apply(
        lambda x: 40000000 if ('GOLD' in x or 'VIP Branch' in x) else 20000000
    )
    review_vip['rate'] = review_vip['contract_type'].str.split('Margin.PIA ').str.get(1).str.split('%').str.get(0)
    review_vip['rate'] = review_vip['rate'].astype(float)
    mask_phs = (review_vip['current_vip'] == 'GOLD PHS') | (review_vip['current_vip'] == 'SILV PHS')
    review_vip_phs = review_vip.loc[mask_phs].copy()
    mask_branch = review_vip['current_vip'] == 'VIP Branch'
    review_vip_branch = review_vip.loc[mask_branch].copy()

    # month, year of end_date
    check_date = dt.datetime.strptime(end_date, "%Y-%m-%d")
    check_month = check_date.month
    check_year = check_date.year

    if check_month == 6:
        start_date = dt.datetime(check_year, 1, 1).strftime('%Y-%m-%d')
        end_date = end_date
    elif check_month == 12:
        start_date = dt.datetime(check_year, 7, 1).strftime('%Y-%m-%d')
        end_date = end_date
    else:
        start_date = start_date
        end_date = end_date

    # query RCF3001
    query_rcf01 = pd.read_sql(
        f"""
            SELECT
                [date],
                [account_code],
                [trading_fee],
                [loan_fee]
            FROM [revenue_from_vip]
            WHERE 
                [date] BETWEEN '{start_date}' AND '{end_date}'
            ORDER BY [account_code]
        """,
        connect_DWH_CoSo,
    )
    # query RCF3002
    query_rcf02 = pd.read_sql(
        f"""
            SELECT 
                *
            FROM [vip_evaluation]
            WHERE 
                [date] BETWEEN '{start_date}' AND '{end_date}'
            ORDER BY [account_code]
        """,
        connect_DWH_CoSo,
    )

    if start_date != sod:
        start_date_3m = sod
        # loc RCF3001
        query_rcf01['date'] = pd.to_datetime(query_rcf01['date'], format='%Y-%m-%d')
        query_rcf01_3m = query_rcf01.loc[(query_rcf01['date'] >= start_date_3m) and (query_rcf01['date'] <= end_date)]
        query_rcf01_3m = query_rcf01_3m.groupby('account_code').sum()
        query_rcf01_6m = query_rcf01.groupby('account_code').sum()
        # loc RCF3002
        query_rcf02['date'] = pd.to_datetime(query_rcf02['date'], format='%Y-%m-%d')
        query_rcf02_3m = query_rcf02.loc[(query_rcf02['date'] >= start_date_3m) and (query_rcf02['date'] <= end_date)]
        query_rcf02_3m = query_rcf02_3m.groupby('account_code').sum()
        query_rcf02_6m = query_rcf02.groupby('account_code').sum()
        # review vip GOLD, SILV 6 months
        review_vip_phs['trading_fee'] = query_rcf01_6m['trading_fee']
        # review_vip_phs['tai_san_rong'] = query_rcf02_6m['tai_san_rong']
        review_vip_phs.fillna(0, inplace=True)
        # review vip VIP CN 3 months
        review_vip_branch['trading_fee'] = query_rcf01_3m['trading_fee']
        # review_vip_branch['tai_san_rong'] = query_rcf02_3m['tai_san_rong']
        review_vip_branch.fillna(0, inplace=True)
    else:
        query_rcf01 = query_rcf01.groupby('account_code').sum()
        review_vip_branch['trading_fee'] = query_rcf01['trading_fee']
        # review_vip_branch['tai_san_rong'] = query_rcf02['tai_san_rong']
        review_vip_branch.fillna(0, inplace=True)

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
    money_format = workbook.add_format(
        {
            'border': 1,
            'text_wrap': True,
            'align': 'right',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 8,
            'num_format': '#,##0'
        }
    )
    per_cri_fee_format = workbook.add_format(
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
    rate_format = workbook.add_format(
        {
            'border': 1,
            'text_wrap': True,
            'align': 'right',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 8,
            'num_format': '0.0"%"'
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
    table_empty_format = workbook.add_format(
        {
            'border': 1,
            'text_wrap': True,
            'align': 'center',
            'valign': 'top',
            'font_name': 'Times New Roman',
            'font_size': 11,
        }
    )
    table_content_format = workbook.add_format(
        {
            'border': 1,
            'text_wrap': True,
            'align': 'left',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 11,
        }
    )
    date_header = dt.datetime.strptime(end_date, "%Y-%m-%d").strftime("%d.%m.%Y")
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
    review_vip_sheet.set_column('D:D', 16.71)
    review_vip_sheet.set_column('E:E', 10.86)
    review_vip_sheet.set_column('F:F', 12.43)
    review_vip_sheet.set_column('G:G', 15)
    review_vip_sheet.set_column('H:H', 12.57)
    review_vip_sheet.set_column('I:I', 19.71)
    review_vip_sheet.set_column('J:J', 13.3)
    review_vip_sheet.set_column('K:K', 16)
    review_vip_sheet.set_column('L:L', 9.14)
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
    report_date = dt.datetime.strptime(end_date, "%Y-%m-%d").strftime("%B, %Y")
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
    review_vip_sheet.merge_range('O2:P2', f'     /{run_time.year}/TTr-TRS', no_and_date_format)
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
        money_format
    )
    review_vip_sheet.write_column(
        'J7',
        review_vip['current_vip'],
        text_left_format
    )
    review_vip_sheet.write_column(
        'L7',
        review_vip['rate'],
        rate_format
    )
    review_vip_sheet.write_column(
        'M7',
        [''] * review_vip.shape[0],
        text_left_format
    )
    review_vip_sheet.write_column(
        'N7',
        [''] * review_vip.shape[0],
        text_left_format
    )
    review_vip_sheet.write_column(
        'O7',
        [''] * review_vip.shape[0],
        text_left_format
    )
    review_vip_sheet.write_column(
        'P7',
        [''] * review_vip.shape[0],
        text_left_format
    )
    review_vip_sheet.write_column(
        'Q7',
        review_vip['branch_name'],
        text_left_format
    )
    review_vip_sheet.write_column(
        'R7',
        [''] * review_vip.shape[0],
        text_left_format
    )
    table_start_row = review_vip.shape[0] + 9
    footer = 'Effective: Promoted accounts be applied from ......./......./2021 & another accounts be applied from ........../........../2021'
    review_vip_sheet.merge_range(
        f'A{table_start_row}:P{table_start_row}',
        footer,
        footer_format
    )
    review_vip_sheet.merge_range(
        f'B{table_start_row+2}:P{table_start_row+2}',
        'TRADING SERVICE DIVISION',
        table_header_format
    )
    review_vip_sheet.merge_range(
        f'B{table_start_row+3}:H{table_start_row+3}',
        'PROPOSER',
        table_sub_header_format
    )
    review_vip_sheet.merge_range(
        f'I{table_start_row + 3}:P{table_start_row + 3}',
        'DIRECTOR OF TRADING SERVICE DIVISION',
        table_sub_header_format
    )
    review_vip_sheet.merge_range(
        f'B{table_start_row + 4}:H{table_start_row + 7}',
        '',
        table_empty_format
    )
    review_vip_sheet.merge_range(
        f'I{table_start_row + 4}:P{table_start_row + 7}',
        '',
        table_empty_format
    )
    review_vip_sheet.merge_range(
        f'B{table_start_row + 9}:H{table_start_row + 9}',
        'Decision of Deputy General Director:',
        table_sub_header_format
    )
    review_vip_sheet.merge_range(
        f'I{table_start_row + 9}:P{table_start_row + 9}',
        'DEPUTY GENERAL DIRECTOR',
        table_sub_header_format
    )
    review_vip_sheet.merge_range(
        f'B{table_start_row + 10}:H{table_start_row + 14}',
        "o Agree\no Disagree\no Others:................",
        table_content_format
    )
    review_vip_sheet.merge_range(
        f'I{table_start_row + 10}:P{table_start_row + 14}',
        '',
        table_empty_format
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
