"""
    1. VIP CN => mỗi quý review 1 lần => lấy data quý đó
    NORMAL, GOLD và SILV => 6 tháng review 1 lần => lấy data 6 tháng
    =>  Báo cáo chạy vào cuối quý 2, quý 4 => xét toàn bộ KH (normal,gold,silv,vip)
        Báo cáo chạy vào cuối quý 1, quý 3 => chỉ xét vip

    2. Nếu từ được lên VIP (approved date) cho tới ngày cuối chu kỳ > 30 ngày => giữ (có review)
    còn KH nào thời gian lên VIP tới ngày cuối chu kỳ <= 30 => bỏ (không review)

    3. Qui tắc xét ngày lấy giá trị:
    - Nếu ngày KH lên VIP vào trước chu kỳ => xét giá trị từ đầu chu kỳ tới cuối chu kỳ
    - Nếu ngày KH lên VIP nằm trong chu kỳ => xét giá trị từ ngày lên VIP tới cuối chu kỳ

    4. Fee for assessment
    công thức: Fee for assessment = 100% * Phí GD + 30% * (Phí UTTB + lãi vay)
    chú ý: tính tổng phí trung bình tháng và tổng lãi trung bình tháng
           trước khi thay vào công thức trên
    - phí GD: ROD0040
    - phí UTTB: RCI0015
    = lãi vay: RLN0019

    5. Cột I Average Net Asset Value -> lấy từ bảng nav

    6. cột H - % Fee for assessment / Criteria Fee

    7. cột Rate lấy từ số % nằm trong 'contract_type' nằm sau cụm từ Margin.PIA ...

    8. Rule cột After review
    KH GOLD:
        if: KH có Fee for Aseessment >= 40tr -> GOLD
        (nếu khác GOLD thì tăng bậc là PROMOTE GOLD, còn nếu là GOLD thì giữ nguyên GOLD)
        else:
            if: KH GOLD có cột H >= 80%, I >= 4 tỷ -> giữ là GOLD
            else: DEMOTE SILV
    KH SILV:
        if: KH có 20tr <= Fee for Aseessment < 40tr -> SILV
        (nếu khác SILV tăng bậc là PROMOTE SILV, còn nếu là SILV thì giữ nguyên SILV)
        else:
            if: KH SILV có cột H >= 80%, I >= 2 tỷ -> giữ là SILV
            else: DEMOTE DP
    KH VIP CN:
        if: fee < 20tr OR H < 80%, I < 4 tỷ -> DEMOTE DP
        else:
            if: 20tr <= fee < 40tr -> PROMOTE SILV
            else: fee >= 40tr -> PROMOTE GOLD
    KH Normal:
        if: 20tr <= fee < 40tr -> PROMOTE SILV
        else: fee >= 40tr -> PROMOTE GOLD

    9. Các cột M, N, O, P, R để trống
"""
from reporting.trading_service.thanhtoanbutru import *


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

    # lưu start_date đầu tiên
    save_sod = start_date

    # month, year of end_date
    check_date = dt.datetime.strptime(end_date, "%Y-%m-%d")
    check_month = check_date.month
    check_year = check_date.year

    # kiểm tra điều kiện để thay đổi start date và end date
    if check_month == 6:
        start_date = dt.datetime(check_year, 1, 1).strftime('%Y-%m-%d')
        end_date = end_date
    elif check_month == 12:
        start_date = dt.datetime(check_year, 7, 1).strftime('%Y-%m-%d')
        end_date = end_date
    else:
        start_date = start_date
        end_date = end_date

    # query danh sách khách hàng
    review_vip = pd.read_sql(
        f"""
            WITH [i] AS (
                SELECT 
                    [relationship].[sub_account],
                    [relationship].[account_code],
                    [account].[customer_name],
                    [branch].[branch_name],
                    [c].[approve_date],
                    CASE 
                        WHEN 
                            [c].[approve_date] < '{start_date}' 
                            AND (
                                [vcf0051].[contract_type] LIKE N'%GOLD%' 
                                OR [vcf0051].[contract_type] LIKE N'%SILV%'
                            ) THEN '{start_date}'
                        WHEN 
                            [c].[approve_date] < '{save_sod}' 
                            AND [vcf0051].[contract_type] LIKE N'%VIPCN%' THEN '{save_sod}'
                        WHEN [vcf0051].[contract_type] LIKE N'%NOR%' THEN '{start_date}'
                        ELSE [c].[approve_date]
                    END [data_fdate],
                    [vcf0051].[contract_type],
                    CASE
                        WHEN 
                            [vcf0051].[contract_type] LIKE '%GOLD%' 
                            OR [vcf0051].[contract_type] LIKE '%VIPCN%' THEN 40000000
                        ELSE 20000000
                    END [criteria_fee],
                    [broker].[broker_name]
                FROM
                    [relationship]
                LEFT JOIN
                    [vcf0051]
                    ON [vcf0051].[sub_account] = [relationship].[sub_account]
                    AND [vcf0051].[date] = [relationship].[date]
                LEFT JOIN (
                    SELECT 
                        [customer_information_change].[account_code], 
                        MAX([customer_information_change].[date_of_approval]) [approve_date]
                    FROM [customer_information_change]
                    WHERE
                        [customer_information_change].[date_of_approval] <= '{end_date}'
                    GROUP BY
                        [customer_information_change].[account_code]
                ) [c]
                    ON [relationship].[account_code] = [c].[account_code]
                LEFT JOIN
                    [account]
                    ON [relationship].[account_code] = [account].[account_code]
                LEFT JOIN
                    [branch]
                    ON [relationship].[branch_id] = [branch].[branch_id]
                LEFT JOIN
                    [broker]
                    ON [relationship].[broker_id] = [broker].[broker_id]
                WHERE (
                    [vcf0051].[contract_type] LIKE N'%MR%'
                )
                    AND [relationship].[date] = '{end_date}'
                    AND [vcf0051].[status] IN ('A','B')
                    AND DATEDIFF(day,[c].[approve_date], '{end_date}') > 30
            ),
            [rod40] AS (
                SELECT 
                    [sum].[account_code],
                    SUM([sum].[fee])/ROUND(CAST(DATEDIFF(day,[sum].[data_fdate],'{end_date}')+1 AS FLOAT)/30,0) [fee]
                FROM (
                    SELECT 
                        [i].[account_code],
                        [i].[data_fdate],
                        SUM([trading_record].[fee]) [fee]
                    FROM
                        [trading_record]
                    RIGHT JOIN [i] ON [i].[sub_account] = [trading_record].[sub_account]
                    WHERE [trading_record].[date] BETWEEN [i].[data_fdate] AND '{end_date}'
                    GROUP BY
                        [trading_record].[date],
                        [i].[account_code],
                        [i].[data_fdate]
                ) [sum]
                GROUP BY 
                    [sum].[account_code],
                    [sum].[data_fdate]
            ),
            [rci15] AS (
                SELECT
                    [sum].[account_code],
                    SUM([sum].[fee])/ROUND(CAST(DATEDIFF(day,[sum].[data_fdate],'{end_date}')+1 AS FLOAT)/30,0) [fee]
                FROM (
                    SELECT
                        [i].[account_code],
                        [i].[data_fdate],
                        SUM([payment_in_advance].[total_fee]) [fee]
                    FROM
                        [payment_in_advance]
                    RIGHT JOIN [i] ON [i].[sub_account] = [payment_in_advance].[sub_account]
                    WHERE [payment_in_advance].[date] BETWEEN [i].[data_fdate] AND '{end_date}'
                    GROUP BY
                        [payment_in_advance].[date],
                        [i].[account_code],
                        [i].[data_fdate]
                ) [sum]
                GROUP BY 
                    [sum].[account_code],
                    [sum].[data_fdate]
            ),
            [rln19] AS (
                SELECT
                    [sum].[account_code],
                    SUM([sum].[int])/ROUND(CAST(DATEDIFF(day,[sum].[data_fdate],'{end_date}')+1 AS FLOAT)/30,0) [int]
                FROM (
                    SELECT
                        [i].[account_code],
                        [i].[data_fdate],
                        SUM([rln0019].[interest]) [int]
                    FROM
                        [rln0019]
                    RIGHT JOIN [i] ON [i].[sub_account] = [rln0019].[sub_account]
                    WHERE [rln0019].[date] BETWEEN [i].[data_fdate] AND '{end_date}'
                    GROUP BY
                        [rln0019].[date], 
                        [i].[account_code],
                        [i].[data_fdate]
                ) [sum]
                GROUP BY 
                    [sum].[account_code],
                    [sum].[data_fdate]
            ),
            [vip_customer] AS (
                SELECT [full].* FROM (
                    SELECT
                        [i].[account_code],
                        [i].[sub_account],
                        [i].[customer_name],
                        [i].[branch_name],
                        [i].[approve_date],
                        [i].[data_fdate],
                        [i].[broker_name],
                        [i].[contract_type],
                        [i].[criteria_fee],
                        ISNULL([rod40].[fee],0) + 0.3*(ISNULL([rci15].[fee],0)+ISNULL([rln19].[int],0)) [fee_for_assm]
                    FROM [i]
                    FULL JOIN [rod40] ON [rod40].[account_code] = [i].[account_code]
                    FULL JOIN [rci15] ON [rci15].[account_code] = [i].[account_code]
                    FULL JOIN [rln19] ON [rln19].[account_code] = [i].[account_code]
                ) [full]
            ),
            [anav] AS (
                SELECT 
                    [sum].[account_code],
                    AVG([sum].[nav]) [nav]
                FROM (
                    SELECT 
                        [n].[date],
                        [i].[account_code],
                        SUM([n].[nav]) [nav]
                    FROM [i]
                    LEFT JOIN (
                        SELECT [nav].[date], [nav].[sub_account], [nav].[nav]
                        FROM [nav]
                        RIGHT JOIN [i] ON [i].[sub_account] = [nav].[sub_account]
                        WHERE [nav].[date] BETWEEN [i].[data_fdate] AND '{end_date}'
                    ) [n]
                    ON [n].[sub_account] = [i].[sub_account]
                    GROUP BY 
                        [n].[date], 
                        [i].[account_code]
                ) [sum]
                GROUP BY
                    [sum].[account_code]
            )
            SELECT 
                [all].[account_code],
                [all].[sub_account],
                [all].[customer_name],
                [all].[contract_type],
                [all].[branch_name],
                [all].[approve_date],
                [all].[criteria_fee],
                [all].[fee_for_assm],
                (([all].[fee_for_assm] / [all].[criteria_fee]) * 100) [percent_fee],
                [all].[nav],
                [all].[broker_name]
            FROM (
                SELECT
                    [vip_customer].*,
                    [anav].[nav]
                FROM [vip_customer]
                LEFT JOIN [anav] ON [anav].[account_code] = [vip_customer].[account_code]
            ) [all]
        """,
        connect_DWH_CoSo,
        index_col='account_code'
    )
    # thêm cột current vip (trích dữ liệu từ cột contract type)
    review_vip['current_vip'] = review_vip['contract_type'].apply(
        lambda x: 'SILV PHS' if 'SILV' in x else (
            'GOLD PHS' if 'GOLD' in x else ('Nor Margin' if 'NOR' in x else 'VIP Branch')))

    # thêm cột rate vào dataframe, trích dữ liệu từ contract type (lấy số n% sau ... Margin.PIA n% ...)
    review_vip['contract_type'] = review_vip['contract_type'].str.replace(' ', '')
    review_vip['rate'] = review_vip['contract_type'].str.split('Margin.PIA').str.get(1).str.split('%').str.get(
        0).astype(float)
    review_vip = review_vip.drop(columns=['contract_type'])

    # Chia dataframe thành 3 loại vip PHS, vip CN, Nor Margin
    mask_phs = (review_vip['current_vip'] == 'GOLD PHS') | (review_vip['current_vip'] == 'SILV PHS')
    review_vip_phs = review_vip.loc[mask_phs].copy()
    mask_branch = review_vip['current_vip'] == 'VIP Branch'
    review_vip_branch = review_vip.loc[mask_branch].copy()
    mask_nor = review_vip['current_vip'] == 'Nor Margin'
    review_nor_mr = review_vip.loc[mask_nor].copy()
    review_nor_mr = review_nor_mr.loc[review_nor_mr['fee_for_assm'] >= 20000000]

    if save_sod != start_date:
        review_vip = pd.concat([review_vip_branch, review_vip_phs, review_nor_mr], join='outer')
        review_vip = review_vip.reset_index()

        check_gold = (review_vip['current_vip'] == 'GOLD PHS')
        check_silv = (review_vip['current_vip'] == 'SILV PHS')
        check_vipcn = (review_vip['current_vip'] == 'VIP Branch')

        check_percentage = review_vip['percent_fee'] >= 80

        check_gold_fee_assm = review_vip['fee_for_assm'] >= 40000000
        check_gold_vipcn_avg = review_vip['nav'] >= 4000000000
        check_silv_fee_assm = (20000000 <= review_vip['fee_for_assm']) & (review_vip['fee_for_assm'] < 40000000)
        check_silv_avg = review_vip['nav'] >= 2000000000
        check_vipcn_fee_assm = review_vip['fee_for_assm'] < 20000000

        gold = (check_gold_fee_assm & check_gold) | (check_gold & check_percentage & check_gold_vipcn_avg)
        gold_1 = (check_gold_fee_assm & ~check_gold)
        gold_2 = check_gold & ~check_percentage & ~check_gold_vipcn_avg
        gold_3 = check_gold & ~check_percentage & check_gold_vipcn_avg
        silv = (check_silv_fee_assm & check_silv) | (check_silv & check_percentage & check_silv_avg)
        silv_1 = check_silv_fee_assm & ~check_silv
        silv_2 = check_silv & ~check_percentage & ~check_silv_avg
        vip_cn = check_vipcn & check_percentage & check_gold_vipcn_avg

        review_vip.loc[gold, 'after_review'] = 'GOLD_PHS'
        review_vip.loc[gold_1, 'after_review'] = 'Promote Gold PHS'
        review_vip.loc[gold_2 | gold_3, 'after_review'] = 'Demote Silv PHS'
        review_vip.loc[silv & review_vip['after_review'].isna(), 'after_review'] = 'SILV_PHS'
        review_vip.loc[silv_1 & review_vip['after_review'].isna(), 'after_review'] = 'Promote SILV PHS'
        review_vip.loc[silv_2 & review_vip['after_review'].isna(), 'after_review'] = 'Demote DP'
        review_vip.loc[check_vipcn_fee_assm & review_vip['after_review'].isna(), 'after_review'] = 'Demote DP'
        review_vip.loc[vip_cn & review_vip['after_review'].isna(), 'after_review'] = 'VIP Branch'
    else:
        review_vip = review_vip_branch.reset_index()

        check_percentage = review_vip['percent_fee'] >= 80

        check_gold_fee_assm = review_vip['fee_for_assm'] >= 40000000
        check_silv_fee_assm = (20000000 <= review_vip['fee_for_assm']) & (review_vip['fee_for_assm'] < 40000000)
        check_vipcn_fee_assm = review_vip['fee_for_assm'] < 20000000

        check_gold_vipcn_avg = review_vip['nav'] >= 4000000000

        gold_1 = check_gold_fee_assm
        silv_1 = check_silv_fee_assm
        vip_cn = check_percentage & check_gold_vipcn_avg

        review_vip.loc[gold_1, 'after_review'] = 'Promote GOLD PHS'
        review_vip.loc[silv_1, 'after_review'] = 'Promote SILV PHS'
        review_vip.loc[check_vipcn_fee_assm, 'after_review'] = 'Demote DP'
        review_vip.loc[vip_cn & review_vip['after_review'].isna(), 'after_review'] = 'VIP Branch'

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
    text_center_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 8,
        }
    )
    text_center_bg_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 8,
            'color': '#222B33',
            'bg_color': '#FFCCFF'
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
    criteria_format = workbook.add_format(
        {
            'border': 1,
            'text_wrap': True,
            'align': 'right',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 8,
            'num_format': '#,##0',
            'color': '#222B33',
            'bg_color': '#FFCCFF'
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
            'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
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
        'Decision of Deputy General Director',
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
    review_vip_sheet.insert_image('A2', './img/phu_hung.png', {'x_scale': 1.26, 'y_scale': 0.52})

    # Set Column Width and Row Height
    review_vip_sheet.set_column('A:A', 4)
    review_vip_sheet.set_column('B:B', 11)
    review_vip_sheet.set_column('C:C', 23)
    review_vip_sheet.set_column('D:D', 16.71)
    review_vip_sheet.set_column('E:E', 10.86)
    review_vip_sheet.set_column('F:F', 12.43)
    review_vip_sheet.set_column('G:G', 15)
    review_vip_sheet.set_column('H:H', 12.57)
    review_vip_sheet.set_column('I:I', 13)
    review_vip_sheet.set_column('J:J', 15)
    review_vip_sheet.set_column('K:K', 15)
    review_vip_sheet.set_column('L:L', 9)
    review_vip_sheet.set_column('M:N', 19)
    review_vip_sheet.set_column('O:O', 17.14)
    review_vip_sheet.set_column('P:P', 15.71)
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
        review_vip['account_code'],
        text_center_format
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
        review_vip['approve_date'],
        date_format
    )
    review_vip_sheet.write_column(
        'F7',
        review_vip['criteria_fee'],
        criteria_format
    )
    review_vip_sheet.write_column(
        'G7',
        review_vip['fee_for_assm'],
        money_format
    )
    review_vip_sheet.write_column(
        'H7',
        review_vip['percent_fee'],
        per_cri_fee_format
    )
    review_vip_sheet.write_column(
        'I7',
        review_vip['nav'],
        money_format
    )
    review_vip_sheet.write_column(
        'J7',
        review_vip['current_vip'],
        text_center_bg_format
    )
    review_vip_sheet.write_column(
        'K7',
        review_vip['after_review'],
        text_left_wrap_text_format
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
        review_vip['broker_name'],
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
        f'B{table_start_row + 2}:P{table_start_row + 2}',
        'TRADING SERVICE DIVISION',
        table_header_format
    )
    review_vip_sheet.merge_range(
        f'B{table_start_row + 3}:H{table_start_row + 3}',
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
