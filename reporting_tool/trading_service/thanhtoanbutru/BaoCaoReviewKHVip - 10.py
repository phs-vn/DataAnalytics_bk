"""
    1. Nguyên tắc là VIP CN thì 3 tháng review 1 lần chốt cuối mỗi quý thì từ thời gian chốt
    đó tính ngược lại 3 tháng cứ thế mà đếm ra, còn GOLD và SILV thì sáu tháng tức là 1 năm 2 lần
    cuối tháng 6 và cuối tháng 12 cứ thế mà đếm ngược lại.
    Nguyên tắc thời gian, nếu báo cáo chạy vào thời điểm quý 2, quý 4 là sẽ xét toàn bộ KH (gold, vip, silv),
    còn nếu chạy vào thời điểm quý 1, quý 3 thì chỉ xét KH VIP CN
    2. Các cột M, N, O, P lấy ý kiến từ các phòng ban khác
    3. Cột Note để trắng cho CN họ note cái gì thì note.
       Cột MOI GIOI QUAN LY thì lên dữ liệu cho chị Tuyết
    4. Cột I (Average Net Asset Value - Tài sản Ròng BQ) không lấy từ bảng RCF3002 nữa mà lấy từ bảng nav theo sub_account
    5. Nếu ngày trở thành vip nằm trong khoảng chu kỳ đang được setup thì lấy từ ngày họ bắt đầu trở thành vip tới cuối kì.
    Nếu ngày trở thành vip nằm ngoài khoảng chu kỳ đang được setup thì lấy full toàn bộ giá trị trong chu kỳ đang được setup
    (điều kiện này được xét cho tất cả các giá trị của Phí GD, lãi vay và phí ứng tiền, tài sản ròng BQ).
    Phần Tài sản ròng BQ ta có thể sum rồi chia bình quân cũng dc.
    6. Cột G Fee for assessment
        công thức: 100% * Phí GD + 30% * (Phí UTTB và lãi vay)
            phí GD: ROD0040
            phí UTTB: RCI0015
            lãi vay: RLN0019
    7. Rule cột After review
    GOLD:
        - toàn bộ những KH có cột G >= 40tr -> GOLD (nếu khác GOLD thì tăng bậc là PROMOTE GOLD,
        còn nếu là GOLD thì giữ nguyên GOLD)
        - tất cả nhũng KH đang là GOLD, cột H >= 80%, I >= 4 tỷ, giữ là GOLD, ngược lại sẽ bị hạ bậc hết xuống DEMOTE SILV
    SILV:
        - tất cả những KH thuộc điều kiện 20tr <= cột G < 40tr -> SILV (nếu khác SILV tăng bậc là PROMOTE SILV,
        còn nếu là SILV thì giữ nguyên SILV)
        - KH nào đang là SILV, cột H >= 80%, I >= 2 tỷ, giữ là SILV, ngược lại sẽ bị DEMOTE DP
    VIP CN:
        tất cả những KH còn lại (cột G < 20tr) --> Demote DP
        KH nào đang là VIP CN, cột H >= 80%, I >= 4 tỷ, giữ là VIP CN, ngược lại sẽ bị DEMOTE DP
    8. cột H - % Fee for assessment / Criteria Fee
        - công thức tính nằm trong file kết quả
    9. cột Rate lấy từ số % nằm trong 'contract_type' nằm sau cụm từ Margin.PIA ...
    10. Trong file mẫu hay file kết quả có 1 vài trường hợp đặc biệt, DSKH VIP trong kỳ review sẽ có 1 số kh mà tại
    tháng cuối cùng của chu kỳ họ dc là vip thì họ sẽ không bị review
    và thêm điều kiện xét 6 tháng của TK nào không thuộc DS vip (VCF0051) mà họ có phí >=20 tr thì lấy
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

    # lưu start_date đầu tiên
    save_sod = start_date

    # query danh sách khách hàng
    review_vip = pd.read_sql(
        f"""
        SELECT
        DISTINCT
            CONCAT(N'Tháng ',MONTH([account].[date_of_birth])) [birth_month],
            [account].[account_code],
            [relationship].[sub_account], 
            [account].[customer_name], 
            [branch].[branch_name], 
            [account].[date_of_birth], 
            [customer_information].[contract_type],
            [customer_information_change].[date_of_change],
            [customer_information_change].[time_of_change],
            [customer_information_change].[date_of_approval],
            [customer_information_change].[time_of_approval],
            [broker].[broker_name]
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

    # Groupby and condition
    review_vip['effective_date'] = review_vip.groupby('account_code')['date_of_change'].min()
    review_vip['approved_date'] = review_vip.groupby('account_code')['date_of_approval'].max()
    mask = (review_vip['date_of_change'] == review_vip['effective_date']) | (
            review_vip['date_of_approval'] == review_vip['approved_date'])
    review_vip = review_vip.loc[mask]

    review_vip = review_vip[[
        'sub_account',
        'customer_name',
        'branch_name',
        'contract_type',
        'broker_name',
        'approved_date',
    ]]
    review_vip['approved_date'] = pd.to_datetime(review_vip['approved_date']).dt.date
    end_day = dt.datetime.strptime(end_date, '%Y-%m-%d').date()
    review_vip['Ngay_GD'] = (end_day - review_vip['approved_date']).dt.days + 1
    review_vip = review_vip.loc[~(review_vip['Ngay_GD'] <= 30)]
    # thêm cột current vip (trích dữ liệu từ cột contract type)
    review_vip['current_vip'] = review_vip['contract_type'].apply(
        lambda x: 'SILV PHS' if 'SILV' in x else ('GOLD PHS' if 'GOLD' in x else 'VIP Branch')
    )
    review_vip.drop_duplicates(keep='last', inplace=True)
    # thêm cột criteria fee
    review_vip['criteria_fee'] = review_vip['current_vip'].apply(
        lambda x: 40000000 if ('GOLD' in x or 'VIP Branch' in x) else 20000000
    )
    # thêm cột rate vào dataframe, trích dữ liệu từ contract type (lấy số n% sau ... Margin.PIA n% ...)
    review_vip['contract_type'] = review_vip['contract_type'].str.replace(' ', '')
    review_vip['rate'] = review_vip['contract_type'].str.split('Margin.PIA').str.get(1).str.split('%').str.get(0)
    review_vip['rate'] = review_vip['rate'].astype(float)

    # Chia dataframe thành 2 loại vip PHS và vip CN
    mask_phs = (review_vip['current_vip'] == 'GOLD PHS') | (review_vip['current_vip'] == 'SILV PHS')
    review_vip_phs = review_vip.loc[mask_phs].copy()
    mask_branch = review_vip['current_vip'] == 'VIP Branch'
    review_vip_branch = review_vip.loc[mask_branch].copy()

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

    # query data 6 months
    if save_sod != start_date:
        query_nav_6m = pd.read_sql(
            f"""
                SELECT
                    [nav].[date],
                    [nav].[sub_account],
                    [nav].[nav] [avg_net_asset_val]
                FROM [nav] 
                WHERE
                    [nav].[date] BETWEEN '{start_date}' AND '{end_date}'
                ORDER BY
                    [date], [sub_account]
            """,
            connect_DWH_CoSo,
        )
        query_rln_6m = pd.read_sql(
            f"""
                SELECT
                    [date],
                    [sub_account],
                    SUM([interest]) [lai_vay]
                FROM
                    [rln0019]
                WHERE
                    [date] BETWEEN '{start_date}' AND '{end_date}'
                GROUP BY
                    [date], [sub_account]
            """,
            connect_DWH_CoSo,
        )
        query_rod_6m = pd.read_sql(
            f"""
                SELECT
                    [date],
                    [sub_account],
                    SUM([fee]) [phi_gd]
                FROM
                    [trading_record]
                WHERE
                    [date] BETWEEN '{start_date}' AND '{end_date}'
                GROUP BY
                    [date], [sub_account]
                ORDER BY
                    [date], [sub_account]
            """,
            connect_DWH_CoSo,
        )
        query_rci_6m = pd.read_sql(
            f"""
                SELECT
                    [date],
                    [sub_account],
                    SUM([total_fee]) AS [phi_uttb]
                FROM
                    [payment_in_advance]
                WHERE
                    [date] BETWEEN '{start_date}' AND '{end_date}'
                GROUP BY
                    [date], [sub_account]
            """,
            connect_DWH_CoSo,
        )
        # xử lý data thời điểm xét 6 tháng
        query_rod_6m = query_rod_6m.set_index(['date', 'sub_account'])
        query_rci_6m = query_rci_6m.set_index(['date', 'sub_account'])
        query_rln_6m = query_rln_6m.set_index(['date', 'sub_account'])
        query_nav_6m = query_nav_6m.set_index(['date', 'sub_account'])
        table_6m = query_rod_6m.merge(
            query_rln_6m, how='outer', left_index=True, right_index=True
        ).merge(
            query_rci_6m, how='outer', left_index=True, right_index=True
        ).merge(
            query_nav_6m, how='outer', left_index=True, right_index=True
        )
        table_6m.fillna(0, inplace=True)
        table_6m = table_6m.reset_index(['date', 'sub_account'])
        table_6m['fee_for_assm'] = table_6m['phi_gd'] + (table_6m['lai_vay'] * 0.3) + (table_6m['phi_uttb'] * 0.3)

        table_6m = table_6m.merge(
            review_vip_phs[['sub_account', 'approved_date']],
            how='right',
            left_on=['sub_account'],
            right_on=['sub_account'],
        )
        table_6m = table_6m.loc[~(table_6m['date'] < table_6m['approved_date'])]
        table_6m = table_6m.drop(columns=['phi_gd', 'lai_vay', 'phi_uttb'])
        table_6m.fillna(0, inplace=True)

        fee_for_assessment_6m = table_6m.groupby('sub_account')['fee_for_assm'].sum()
        avg_net_asset_val_6m = table_6m.groupby('sub_account')['avg_net_asset_val'].mean()

        review_vip_phs = review_vip_phs.merge(
            fee_for_assessment_6m,
            left_on='sub_account',
            right_index=True,
            how='left').merge(
            avg_net_asset_val_6m,
            left_on='sub_account',
            right_index=True,
            how='left'
        )
        review_vip_phs.fillna(0, inplace=True)
        review_vip_phs['%_fee_div_cri'] = (review_vip_phs['fee_for_assm'] / review_vip_phs['criteria_fee']) * 100
    # query data 3 months
    query_nav_3m = pd.read_sql(
        f"""
            SELECT
                [nav].[date],
                [nav].[sub_account],
                [nav].[nav] [avg_net_asset_val]
            FROM [nav] 
            WHERE 
                [nav].[date] BETWEEN '{save_sod}' AND '{end_date}'
            ORDER BY
                [date], [sub_account]
        """,
        connect_DWH_CoSo,
    )
    query_rln_3m = pd.read_sql(
        f"""
            SELECT
                [date],
                [sub_account],
                SUM([interest]) [lai_vay]
            FROM
                [rln0019]
            WHERE
                [date] BETWEEN '{save_sod}' AND '{end_date}'
            GROUP BY
                [date], [sub_account]
        """,
        connect_DWH_CoSo,
    )
    query_rod_3m = pd.read_sql(
        f"""
            SELECT
                [date],
                [sub_account],
                SUM([fee]) [phi_gd]
            FROM
                [trading_record]
            WHERE
                [date] BETWEEN '{save_sod}' AND '{end_date}'
            GROUP BY
                [date], [sub_account]
            ORDER BY
                [date], [sub_account]
        """,
        connect_DWH_CoSo,
    )
    query_rci_3m = pd.read_sql(
        f"""
            SELECT
                [date],
                [sub_account],
                SUM([total_fee]) AS [phi_uttb]
            FROM
                [payment_in_advance]
            WHERE
                [date] BETWEEN '{start_date}' AND '{end_date}'
            GROUP BY
                [date], [sub_account]
        """,
        connect_DWH_CoSo,
    )

    # xử lý data thời điểm xét 3 tháng
    query_rod_3m = query_rod_3m.set_index(['date', 'sub_account'])
    query_rci_3m = query_rci_3m.set_index(['date', 'sub_account'])
    query_rln_3m = query_rln_3m.set_index(['date', 'sub_account'])
    query_nav_3m = query_nav_3m.set_index(['date', 'sub_account'])
    table_3m = query_rod_3m.merge(
        query_rln_3m, how='outer', left_index=True, right_index=True
    ).merge(
        query_rci_3m, how='outer', left_index=True, right_index=True
    ).merge(
        query_nav_3m, how='outer', left_index=True, right_index=True
    )
    table_3m.fillna(0, inplace=True)
    table_3m = table_3m.reset_index(['date', 'sub_account'])
    table_3m['fee_for_assm'] = table_3m['phi_gd'] + (table_3m['lai_vay'] * 0.3) + (table_3m['phi_uttb'] * 0.3)

    table_3m = table_3m.merge(
        review_vip_branch[['sub_account', 'approved_date']],
        how='right',
        left_on=['sub_account'],
        right_on=['sub_account'],
    )
    table_3m = table_3m.loc[~(table_3m['date'] < table_3m['approved_date'])]
    table_3m = table_3m.drop(columns=['phi_gd', 'lai_vay', 'phi_uttb'])
    table_3m.fillna(0, inplace=True)

    # tính trung bình cộng và bình quân của 2 cột fee for assessment và tài sản ròng BQ
    fee_for_assessment_3m = table_3m.groupby('sub_account')['fee_for_assm'].sum()
    avg_net_asset_val_3m = table_3m.groupby('sub_account')['avg_net_asset_val'].mean()

    review_vip_branch = review_vip_branch.merge(
        fee_for_assessment_3m,
        left_on='sub_account',
        right_index=True,
        how='left').merge(
        avg_net_asset_val_3m,
        left_on='sub_account',
        right_index=True,
        how='left'
    )
    review_vip_branch.fillna(0, inplace=True)
    review_vip_branch['%_fee_div_cri'] = (review_vip_branch['fee_for_assm'] / review_vip_branch['criteria_fee']) * 100

    if save_sod != start_date:
        review_vip = pd.concat([review_vip_branch, review_vip_phs], join='outer')
        review_vip = review_vip.reset_index()[[
            'account_code',
            'sub_account',
            'customer_name',
            'branch_name',
            'approved_date',
            'criteria_fee',
            'fee_for_assm',
            '%_fee_div_cri',
            'avg_net_asset_val',
            'current_vip',
            'rate',
            'broker_name',
        ]]

        check_gold = (review_vip['current_vip'] == 'GOLD PHS')
        check_silv = (review_vip['current_vip'] == 'SILV PHS')
        check_vipcn = (review_vip['current_vip'] == 'VIP Branch')

        check_percentage = review_vip['%_fee_div_cri'] >= 80

        check_gold_fee_assm = review_vip['fee_for_assm'] >= 40000000
        check_gold_vipcn_avg = review_vip['avg_net_asset_val'] >= 4000000000
        check_silv_fee_assm = (20000000 <= review_vip['fee_for_assm']) & (review_vip['fee_for_assm'] < 40000000)
        check_silv_avg = review_vip['avg_net_asset_val'] >= 2000000000
        check_vipcn_fee_assm = review_vip['fee_for_assm'] < 20000000

        gold = (check_gold_fee_assm & check_gold) | (check_gold & check_percentage & check_gold_vipcn_avg)
        gold_1 = (check_gold_fee_assm & ~check_gold)
        gold_2 = check_gold & ~check_percentage & ~check_gold_vipcn_avg
        silv = (check_silv_fee_assm & check_silv) | (check_silv & check_percentage & check_silv_avg)
        silv_1 = check_silv_fee_assm & ~check_silv
        silv_2 = check_silv & ~check_percentage & ~check_silv_avg
        vip_cn = check_vipcn & check_percentage & check_gold_vipcn_avg

        review_vip.loc[gold, 'after_review'] = 'GOLD_PHS'
        review_vip.loc[gold_1, 'after_review'] = 'Promote Gold PHS'
        review_vip.loc[gold_2, 'after_review'] = 'Demote Silv PHS'
        review_vip.loc[silv, 'after_review'] = 'SILV_PHS'
        review_vip.loc[silv_1, 'after_review'] = 'Promote Silv PHS'
        review_vip.loc[silv_2, 'after_review'] = 'Demote DP'
        review_vip.loc[check_vipcn_fee_assm, 'after_review'] = 'Demote DP'
        review_vip.loc[vip_cn & review_vip['after_review'].isna(), 'after_review'] = 'VIP Branch'

    review_vip = review_vip_branch.reset_index()[[
        'account_code',
        'sub_account',
        'customer_name',
        'branch_name',
        'approved_date',
        'criteria_fee',
        'fee_for_assm',
        '%_fee_div_cri',
        'avg_net_asset_val',
        'current_vip',
        'rate',
        'broker_name',
    ]]

    check_gold = (review_vip['current_vip'] == 'GOLD PHS')
    check_silv = (review_vip['current_vip'] == 'SILV PHS')
    check_vipcn = (review_vip['current_vip'] == 'VIP Branch')

    check_percentage = review_vip['%_fee_div_cri'] >= 80

    check_gold_fee_assm = review_vip['fee_for_assm'] >= 40000000
    check_gold_vipcn_avg = review_vip['avg_net_asset_val'] >= 4000000000
    check_silv_fee_assm = (20000000 <= review_vip['fee_for_assm']) & (review_vip['fee_for_assm'] < 40000000)
    check_vipcn_fee_assm = review_vip['fee_for_assm'] < 20000000

    gold_1 = check_gold_fee_assm & ~check_gold
    silv_1 = check_silv_fee_assm & ~check_silv
    vip_cn = check_vipcn & check_percentage & check_gold_vipcn_avg

    review_vip.loc[gold_1, 'after_review'] = 'Promote Gold PHS'
    review_vip.loc[silv_1, 'after_review'] = 'Promote Silv PHS'
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
            'num_format': '#,###0',
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
    review_vip_sheet.set_column('A:A', 3.86)
    review_vip_sheet.set_column('B:B', 10.43)
    review_vip_sheet.set_column('C:C', 21.86)
    review_vip_sheet.set_column('D:D', 22)
    review_vip_sheet.set_column('E:E', 10.86)
    review_vip_sheet.set_column('F:F', 12.43)
    review_vip_sheet.set_column('G:G', 15)
    review_vip_sheet.set_column('H:H', 12.57)
    review_vip_sheet.set_column('I:I', 14)
    review_vip_sheet.set_column('J:J', 18.71)
    review_vip_sheet.set_column('K:K', 11.86)
    review_vip_sheet.set_column('L:L', 8)
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
        criteria_format
    )
    review_vip_sheet.write_column(
        'G7',
        review_vip['fee_for_assm'],
        money_format
    )
    review_vip_sheet.write_column(
        'H7',
        review_vip['%_fee_div_cri'],
        per_cri_fee_format
    )
    review_vip_sheet.write_column(
        'I7',
        review_vip['avg_net_asset_val'],
        money_format
    )
    review_vip_sheet.write_column(
        'J7',
        review_vip['current_vip'],
        text_center_format
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
