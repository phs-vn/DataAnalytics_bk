"""
    1. Nguyên tắc là VIP CN thì 3 tháng review 1 lần chốt cuối mỗi quý thì từ thời gian chốt
    đó tính ngược lại 3 tháng cứ thế mà đếm ra, còn GOLD và SILV thì sáu tháng tức là 1 năm 2 lần
    cuối tháng 6 và cuối tháng 12 cứ thế mà đếm ngược lại.
    Nguyên tắc thời gian, nếu báo cáo chạy vào thời điểm quý 2, quý 4 là sẽ xét toàn bộ KH (gold, vip, silv),
    còn nếu chạy vào thời điểm quý 1, quý 3 thì chỉ xét KH VIP CN
    2. Các cột M, N, O, P lấy ý kiến từ các phòng ban khác
    3. Cột Note để trắng cho CN họ note cái gì thì note.
       Cột MOI GIOI QUAN LY thì lên dữ liệu cho chị Tuyết
    4. Cột I (Average Net Asset Value - Tài sản Ròng BQ) không lấy từ bảng RCF3002 nữa mà lấy từ
    bảng nav theo sub_account
        nav BQ = tổng nav của chu kỳ / tổng số ngày làm việc của chu kỳ
    5. Cột G Fee for assessment
        công thức: Fee for assessment = 100% * Phí GD + 30% * (Phí UTTB và lãi vay)
        chú ý: ta phải tính tổng phí trung bình và tổng lãi trung bình trước khi thay vào công thức trên
        cách tinh tổng phí GD BQ tháng, lãi vay BQ tháng, phí UTTB BQ tháng -> trang 5 file pdf
        nhưng theo rule của chị Tuyết thì công thức sẽ khác 1 chút so với file pdf là:
            phí GD BQ = Tổng phí GD của chu kỳ / round(((ngay ket thuc chu ky - ngay bat dau len vip) + 1) / 30, 0)
                thêm vào 1 đơn vị ở đây (+ 1 trên công thức) là vì bản chất ngày cuối cùng vẫn dc tính, trên thực tế tại
                cái ngày cuối đó KH vẫn dc tính, cthuc trong pdf tính như thế là thiếu của KH 1 ngày.
            phí UTTB BQ và lãi vay BQ cũng tương tự
            phí GD: ROD0040
            phí UTTB: RCI0015
            lãi vay: RLN0019
    6. Rule cột After review
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
    7. cột H - % Fee for assessment / Criteria Fee
        - công thức: fee for assessment / criteria fee
    8. cột Rate lấy từ số % nằm trong 'contract_type' nằm sau cụm từ Margin.PIA ...
    9. Trong file mẫu hay file kết quả có 1 vài trường hợp đặc biệt, DSKH VIP trong kỳ review sẽ có 1 số kh mà tại
    tháng cuối cùng của chu kỳ họ dc là vip thì họ sẽ không bị review
    và thêm điều kiện xét 6 tháng của TK nào không thuộc DS vip (VCF0051) mà họ có phí >=20 tr thì lấy
    10. Nếu khoảng thời gian mà từ lúc KH được lên VIP (approved date) tới ngày cuối chu kỳ > 30 ngày
    => KH đó ở giữa chu kỳ => tính như qui tắc bình thường
    còn KH nào thời gian lên VIP tới cuối chu kỳ <= 30 (tháng cuối cùng của chu kỳ)
    => không review KH này => bỏ KH này ra khỏi danh sách
    11. Qui tắc xét ngày lấy giá trị:
    - Nếu ngày KH lên VIP trước chu kỳ, thì xét toàn bộ giá trị từ đầu chu kỳ tới cuối chu kỳ
    - Nếu ngày KH lên VIP nằm trong chu kỳ, thì xét giá trị từ ngày lên VIP tới cuối chu kỳ
    - theo qui tắc số 10, nếu ngày lên VIP nằm trong khoảng thời gian tháng cuối cùng của chu kỳ thì
    ko review KH tại thời điểm của báo cáo này mà để sang thời điểm sau.
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
                        WHEN [c].[approve_date] < '{start_date}' AND [vcf0051].[contract_type] NOT LIKE N'%VIPCN%' 
                            THEN '{start_date}'
                        WHEN [c].[approve_date] < '{save_sod}' AND [vcf0051].[contract_type] LIKE N'%VIPCN%' 
                            THEN '{save_sod}'
                        ELSE [c].[approve_date]
                    END [data_fdate],
                    [vcf0051].[contract_type],
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
                    [vcf0051].[contract_type] LIKE N'%GOLD%' 
                    OR [vcf0051].[contract_type] LIKE N'%SILV%' 
                    OR [vcf0051].[contract_type] LIKE N'%VIPCN%'
                    OR [vcf0051].[contract_type] LIKE N'%NOR%'
                )
                    AND [relationship].[date] = '{end_date}'
                    AND [vcf0051].[status] IN ('A','B')
                    AND DATEDIFF(day,[c].[approve_date], '{end_date}') > 30
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
                [all].[data_fdate],
                [all].[nav],
                [all].[broker_name]
            FROM (
                SELECT
                    [i].*,
                    [anav].[nav]
                FROM [i]
                LEFT JOIN [anav] ON [anav].[account_code] = [i].[account_code]
            ) [all]
        """,
        connect_DWH_CoSo,
        index_col='account_code'
    )
    # thêm cột current vip (trích dữ liệu từ cột contract type)
    review_vip['current_vip'] = review_vip['contract_type'].apply(
        lambda x: 'SILV PHS' if 'SILV' in x else (
            'GOLD PHS' if 'GOLD' in x else ('Nor Margin' if 'NOR' in x else 'VIP Branch')))

    # thêm cột criteria fee
    review_vip['criteria_fee'] = review_vip['current_vip'].apply(
        lambda x: 40000000 if ('GOLD' in x or 'VIP Branch' in x) else 20000000
    )
    # thêm cột rate vào dataframe, trích dữ liệu từ contract type (lấy số n% sau ... Margin.PIA n% ...)
    review_vip['contract_type'] = review_vip['contract_type'].str.replace(' ', '')
    review_vip['rate'] = review_vip['contract_type'].str.split('Margin.PIA').str.get(1).str.split('%').str.get(
        0).astype(float)
    review_vip['data_fdate'] = pd.to_datetime(review_vip['data_fdate']).dt.date
    review_vip = review_vip.drop(columns=['contract_type'])

    # Chia dataframe thành 3 loại vip PHS, vip CN, Nor Margin
    mask_phs = (review_vip['current_vip'] == 'GOLD PHS') | (review_vip['current_vip'] == 'SILV PHS')
    review_vip_phs = review_vip.loc[mask_phs].copy()
    mask_branch = review_vip['current_vip'] == 'VIP Branch'
    review_vip_branch = review_vip.loc[mask_branch].copy()
    mask_nor = review_vip['current_vip'] == 'Nor Margin'
    review_nor_mr = review_vip.loc[mask_nor].copy()

    review_nor_mr['data_fdate'] = dt.datetime.strptime(start_date, '%Y-%m-%d').date()
    review_nor_mr['data_fdate'] = pd.to_datetime(review_nor_mr['data_fdate']).dt.date
    end_day = dt.datetime.strptime(end_date, '%Y-%m-%d').date()

    # query data 6 months
    if save_sod != start_date:

        review_vip_phs['ngay_gd'] = round(((end_day - review_vip_phs['data_fdate']).dt.days + 1) / 30, 0)

        review_nor_mr['ngay_gd'] = round(((end_day - review_nor_mr['data_fdate']).dt.days + 1) / 30, 0)

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
                    SUM([total_fee]) [phi_uttb]
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

        table_6m = query_rod_6m.merge(
            query_rln_6m, how='outer', left_index=True, right_index=True
        ).merge(
            query_rci_6m, how='outer', left_index=True, right_index=True
        )
        table_6m.fillna(0, inplace=True)
        table_6m = table_6m.reset_index(['date', 'sub_account'])
        table_6m['fee_for_assm'] = table_6m['phi_gd'] + 0.3 * (table_6m['lai_vay'] + table_6m['phi_uttb'])

        table_6m_vip = table_6m.merge(
            review_vip_phs[['sub_account', 'data_fdate']],
            how='right',
            left_on='sub_account',
            right_on='sub_account',
        )
        table_6m_nor = table_6m.merge(
            review_nor_mr[['sub_account', 'data_fdate']],
            how='right',
            left_on='sub_account',
            right_on='sub_account',
        )
        table_6m_vip = table_6m_vip.dropna()
        table_6m_vip = table_6m_vip.loc[~(table_6m_vip['date'] < table_6m_vip['data_fdate'])]
        table_6m_vip = table_6m_vip.drop(columns=['phi_gd', 'lai_vay', 'phi_uttb', 'data_fdate'])

        table_6m_nor = table_6m_nor.dropna()
        table_6m_nor = table_6m_nor.drop(columns=['phi_gd', 'lai_vay', 'phi_uttb', 'data_fdate'])

        # tính trung bình cộng và bình quân của 2 cột fee for assessment và tài sản ròng BQ
        fee_for_assessment_6m_vip = table_6m_vip.groupby('sub_account')['fee_for_assm'].sum()
        fee_for_assessment_6m_nor = table_6m_nor.groupby('sub_account')['fee_for_assm'].sum()

        review_vip_phs = review_vip_phs.merge(
            fee_for_assessment_6m_vip,
            left_on='sub_account',
            right_index=True,
            how='left'
        )
        review_vip_phs.fillna(0, inplace=True)
        review_vip_phs['fee_for_assm'] = review_vip_phs['fee_for_assm'] / review_vip_phs['ngay_gd']
        review_vip_phs['%_fee_div_cri'] = (review_vip_phs['fee_for_assm'] / review_vip_phs['criteria_fee']) * 100

        review_nor_mr = review_nor_mr.merge(
            fee_for_assessment_6m_nor,
            left_on='sub_account',
            right_index=True,
            how='left'
        )
        review_nor_mr.fillna(0, inplace=True)
        review_nor_mr['fee_for_assm'] = review_nor_mr['fee_for_assm'] / review_nor_mr['ngay_gd']
        review_nor_mr['%_fee_div_cri'] = (review_nor_mr['fee_for_assm'] / review_nor_mr['criteria_fee']) * 100
        review_nor_mr = review_nor_mr.loc[review_nor_mr['fee_for_assm'] >= 20000000]
        review_nor_mr['approved_date'] = 'New'

    # query data 3 months
    review_vip_branch['ngay_gd'] = round(((end_day - review_vip_branch['data_fdate']).dt.days + 1) / 30, 0)

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
                SUM([total_fee]) [phi_uttb]
            FROM
                [payment_in_advance]
            WHERE
                [date] BETWEEN '{save_sod}' AND '{end_date}'
            GROUP BY
                [date], [sub_account]
        """,
        connect_DWH_CoSo,
    )

    # xử lý data thời điểm xét 3 tháng
    query_rod_3m = query_rod_3m.set_index(['date', 'sub_account'])
    query_rci_3m = query_rci_3m.set_index(['date', 'sub_account'])
    query_rln_3m = query_rln_3m.set_index(['date', 'sub_account'])

    table_3m = query_rod_3m.merge(
        query_rln_3m, how='outer', left_index=True, right_index=True
    ).merge(
        query_rci_3m, how='outer', left_index=True, right_index=True
    )

    table_3m.fillna(0, inplace=True)
    table_3m = table_3m.reset_index(['date', 'sub_account'])
    table_3m['fee_for_assm'] = table_3m['phi_gd'] + 0.3 * (table_3m['lai_vay'] + table_3m['phi_uttb'])

    table_3m = table_3m.merge(
        review_vip_branch[['sub_account', 'data_fdate']],
        how='right',
        left_on='sub_account',
        right_on='sub_account',
    )
    table_3m = table_3m.dropna()
    table_3m = table_3m.loc[~(table_3m['date'] < table_3m['data_fdate'])]
    table_3m = table_3m.drop(columns=['phi_gd', 'lai_vay', 'phi_uttb', 'data_fdate'])
    # tính trung bình cộng và bình quân của 2 cột fee for assessment và tài sản ròng BQ
    fee_for_assessment_3m = table_3m.groupby('sub_account')['fee_for_assm'].sum()

    review_vip_branch = review_vip_branch.merge(
        fee_for_assessment_3m,
        left_on='sub_account',
        right_index=True,
        how='left'
    )
    review_vip_branch.fillna(0, inplace=True)
    review_vip_branch['fee_for_assm'] = review_vip_branch['fee_for_assm'] / review_vip_branch['ngay_gd']
    review_vip_branch['%_fee_div_cri'] = (review_vip_branch['fee_for_assm'] / review_vip_branch['criteria_fee']) * 100

    if save_sod != start_date:
        review_vip = pd.concat([review_vip_branch, review_vip_phs, review_nor_mr], join='outer')
        review_vip = review_vip.reset_index()[[
            'account_code',
            'sub_account',
            'customer_name',
            'branch_name',
            'approve_date',
            'criteria_fee',
            'fee_for_assm',
            '%_fee_div_cri',
            'nav',
            'current_vip',
            'rate',
            'broker_name',
        ]]

        check_gold = (review_vip['current_vip'] == 'GOLD PHS')
        check_silv = (review_vip['current_vip'] == 'SILV PHS')
        check_vipcn = (review_vip['current_vip'] == 'VIP Branch')

        check_percentage = review_vip['%_fee_div_cri'] >= 80

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
        review_vip.loc[silv_1 & review_vip['after_review'].isna(), 'after_review'] = 'Promote Silv PHS'
        review_vip.loc[silv_2 & review_vip['after_review'].isna(), 'after_review'] = 'Demote DP'
        review_vip.loc[check_vipcn_fee_assm & review_vip['after_review'].isna(), 'after_review'] = 'Demote DP'
        review_vip.loc[vip_cn & review_vip['after_review'].isna(), 'after_review'] = 'VIP Branch'
    else:
        review_vip = review_vip_branch.reset_index()[[
            'account_code',
            'sub_account',
            'customer_name',
            'branch_name',
            'approve_date',
            'criteria_fee',
            'fee_for_assm',
            '%_fee_div_cri',
            'nav',
            'current_vip',
            'rate',
            'broker_name',
        ]]

        check_percentage = review_vip['%_fee_div_cri'] >= 80

        check_gold_fee_assm = review_vip['fee_for_assm'] >= 40000000
        check_silv_fee_assm = (20000000 <= review_vip['fee_for_assm']) & (review_vip['fee_for_assm'] < 40000000)
        check_vipcn_fee_assm = review_vip['fee_for_assm'] < 20000000

        check_gold_vipcn_avg = review_vip['nav'] >= 4000000000

        gold_1 = check_gold_fee_assm
        silv_1 = check_silv_fee_assm
        vip_cn = check_percentage & check_gold_vipcn_avg

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
    review_vip_sheet.set_column('K:K', 15)
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
        review_vip['%_fee_div_cri'],
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
