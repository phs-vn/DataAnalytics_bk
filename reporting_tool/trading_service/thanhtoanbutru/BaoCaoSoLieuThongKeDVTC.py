"""
1. monthly
2. qui tắc đầu kỳ, cuối kỳ
    Tháng 11:
        + số đầu kỳ: lấy RLN0006 cuối tháng 10 (là số đầu kỳ tháng 11)
        + số cuối kỳ : lấy RLN0006 cuối tháng 11
3. hỏi về cách lấy dữ liệu đầu kỳ (cuối kỳ tháng trước) trong bảng VCF0051
4. Dữ liệu bị lệch khá nhiều so với dữ liệu trên flex
5. Cách để thêm danh sách KH nợ xấu vào báo cáo (Set cứng hay sao)
"""
from reporting_tool.trading_service.thanhtoanbutru import *


def run(
        periodicity:str,
        run_time=None,
):
    start = time.time()
    info = get_info(periodicity,run_time)
    period = info['period']
    folder_name = info['folder_name']
    end_date = info['end_date']
    start_date = info['start_date']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir((join(dept_folder,folder_name,period)))

    ###################################################
    ###################################################
    ###################################################

    outstandings = pd.read_sql(
        f"""
        WITH 
        [opening] AS (
            SELECT
                [margin_outstanding].[type],
                SUM([margin_outstanding].[principal_outstanding]) [o_outs]
            FROM [margin_outstanding]
            WHERE [margin_outstanding].[date] = '{bdate(start_date,-1)}'
            GROUP BY [margin_outstanding].[type]
        ),
        [closing] AS (
            SELECT 
                [margin_outstanding].[type],
                SUM([margin_outstanding].[principal_outstanding]) [c_outs]
            FROM [margin_outstanding]
            WHERE [margin_outstanding].[date] = '{end_date}'
            GROUP BY [margin_outstanding].[type]
        )
        SELECT
            [t].[type],
            [t].[o_outs],
            [t].[c_outs] - [t].[o_outs] [change],
            [t].[c_outs]
        FROM (
            SELECT
                CASE [closing].[type]
                    WHEN N'Ứng trước cổ tức' THEN 'UTCT'
                    WHEN N'Trả chậm' THEN 'DP'
                    WHEN N'Margin' THEN 'MR'
                    WHEN N'Bảo lãnh' THEN 'BL'
                    ELSE [closing].[type]
                END [type],
                [opening].[o_outs],
                [closing].[c_outs]
            FROM [opening]
            FULL JOIN [closing]
                ON [opening].[type] = [closing].[type]) [t]
        ORDER BY 
            CASE [t].[type]
                WHEN N'UTCT' THEN 1
                WHEN N'DP' THEN 2
                WHEN N'MR' THEN 3
                WHEN N'BL' THEN 4
                WHEN N'Cầm cố' THEN 5
            END ASC;
        """,
        connect_DWH_CoSo,
    )

    """
    QUY TẮC MAP TRÊN margin_accounts:

    dim_1:
    contain: MR 1, VIPCN -> VIPCN T1
    contain: MR 1, SILV -> SILV T1
    contain: MR 0, GOLD -> GOLD T0
    contain: MR 2, GOLD -> GOLD T2
    contain: other -> NOR

    dim_2:
    contain: VIPCN, DP 3 -> VIPCN DP+3
    contain: SILV, DP 4 -> SILV DP+4
    contain: GOLD, DP 5 -> GOLD DP+5
    contain: NOR, DP 2 -> NORMAL DP+2
    contain: other -> NODP

    dim_3:
    contain: SILV -> SILVER
    contain: GOLD -> GOLDEN
    contain: VIPCN -> VIPCN
    contain: other -> NOR

    """

    margin_accounts = pd.read_sql(
        f"""
        SELECT DISTINCT
            [vcf0051].[sub_account], 
            -- [vcf0051].[contract_type],
            CASE
                WHEN
                    CHARINDEX('MR1',REPLACE([vcf0051].[contract_type],' ','')) > 0 
                    AND CHARINDEX('VIPCN',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    THEN 'VIPCN T1'
                WHEN
                    CHARINDEX('MR1',REPLACE([vcf0051].[contract_type],' ','')) > 0 
                    AND CHARINDEX('SILV',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    THEN 'SILV T1'
                WHEN
                    CHARINDEX('MR0',REPLACE([vcf0051].[contract_type],' ','')) > 0 
                    AND CHARINDEX('GOLD',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    THEN 'GOLD T0'
                WHEN
                    CHARINDEX('MR2',REPLACE([vcf0051].[contract_type],' ','')) > 0 
                    AND CHARINDEX('GOLD',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    THEN 'GOLD T2'
                ELSE 'NOR'
            END [dim_1],
            CASE
                WHEN
                    CHARINDEX('VIPCN',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    AND CHARINDEX('DP3',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    THEN 'VIPCN DP+3'
                WHEN
                    CHARINDEX('SILV',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    AND CHARINDEX('DP4',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    THEN 'SILV DP+4'
                WHEN
                    CHARINDEX('GOLD',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    AND CHARINDEX('DP5',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    THEN 'GOLD DP+5'
                WHEN
                    CHARINDEX('NOR',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    AND CHARINDEX('DP2',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    THEN 'NORMAL DP+2'
                ELSE 'NODP'
            END [dim_2],
            CASE
                WHEN
                    CHARINDEX('SILV',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    THEN 'SILVER'
                WHEN
                    CHARINDEX('GOLD',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    THEN 'GOLDEN'
                WHEN
                    CHARINDEX('VIPCN',REPLACE([vcf0051].[contract_type],' ','')) > 0
                    THEN 'VIPCN'
                ELSE 'NOR'
            END [dim_3]
            
        FROM 
            [vcf0051]
    
        RIGHT JOIN (SELECT [t].[time], [t].[date], [t].[sub_account]
            FROM (SELECT MAX(time) [time], [date], [sub_account] 
                    FROM [vcf0051] WHERE [date] <= '{end_date}'
                    -- AND [vcf0051].[action] IN ('EDIT','ADD')
                    GROUP BY [sub_account], [date]) [t]
            RIGHT JOIN (SELECT MAX(date) [date], [sub_account]
                    FROM [vcf0051] WHERE [date] <= '{end_date}' 
                    -- AND [vcf0051].[action] IN ('EDIT','ADD')
                    GROUP BY [sub_account]) [d]
            ON 
                [t].[sub_account] = [d].[sub_account]
                AND [t].[date] = [d].[date]
        ) [last_record]
    
        ON [last_record].[sub_account] = [vcf0051].[sub_account] 
            AND [last_record].[date] = [vcf0051].[date] 
            AND [last_record].[time] = [vcf0051].[time]
        
        WHERE
            [vcf0051].[contract_type] NOT LIKE N'%Thường%'
            AND [vcf0051].[contract_type] NOT LIKE N'%Tự doanh%'
            AND [vcf0051].[status] IN ('A','B')
        ORDER BY 
            [vcf0051].[sub_account]
        """,
        connect_DWH_CoSo,
    )
    
    """
    INB01 -> branch_id = '0111'
    """

    outstandings_inb01 = pd.read_sql(
        f"""
        SELECT
            CASE [margin_outstanding].[type]
                WHEN N'Ứng trước cổ tức' THEN 'UTCT'
                WHEN N'Trả chậm' THEN 'DP'
                WHEN N'Margin' THEN 'MR'
                WHEN N'Bảo lãnh' THEN 'BL'
                ELSE [margin_outstanding].[type]
            END [type],
            SUM([margin_outstanding].[principal_outstanding]) [c_outs]
        FROM [margin_outstanding]
        LEFT JOIN (
            SELECT DISTINCT
                [relationship].[account_code], 
                [relationship].[branch_id]
            FROM [relationship]
            WHERE [relationship].[date] = '{end_date}'
            ) [r]
            ON [r].[account_code] = [margin_outstanding].[account_code] 
        WHERE 
            [margin_outstanding].[date] = '{end_date}' 
            AND [r].[branch_id] = '0111'
        GROUP BY [type]
        ORDER BY 
            CASE [type]
                WHEN N'UTCT' THEN 1
                WHEN N'DP' THEN 2
                WHEN N'MR' THEN 3
                WHEN N'BL' THEN 4
                WHEN N'Cầm cố' THEN 5
            END ASC;
        """,
        connect_DWH_CoSo,
    )
    margin_accounts_inb01 = pd.read_sql(
        f"""
        SELECT 
            COUNT(DISTINCT [vcf0051].[sub_account]) [count]
        FROM 
            [vcf0051]
        RIGHT JOIN (SELECT [t].[time], [t].[date], [t].[sub_account]
            FROM (SELECT MAX(time) [time], [date], [sub_account] 
                    FROM [vcf0051] WHERE [date] <= '{end_date}'
                    -- AND [vcf0051].[action] IN ('EDIT','ADD')
                    GROUP BY [sub_account], [date]) [t]
            RIGHT JOIN (SELECT MAX(date) [date], [sub_account]
                    FROM [vcf0051] WHERE [date] <= '{end_date}' 
                    -- AND [vcf0051].[action] IN ('EDIT','ADD')
                    GROUP BY [sub_account]) [d]
            ON 
                [t].[sub_account] = [d].[sub_account]
                AND [t].[date] = [d].[date]
        ) [last_record]
        ON [last_record].[sub_account] = [vcf0051].[sub_account] 
        AND [last_record].[date] = [vcf0051].[date] 
        AND [last_record].[time] = [vcf0051].[time]
        LEFT JOIN (
            SELECT 
                [relationship].[sub_account],
                [relationship].[branch_id]
            FROM [relationship]
            WHERE [relationship].[date] = '{end_date}'
        ) [r]
        ON [r].[sub_account] = [vcf0051].[sub_account]        
        WHERE
            [vcf0051].[contract_type] NOT LIKE N'%Thường%'
            AND [vcf0051].[contract_type] NOT LIKE N'%Tự doanh%'
            AND [vcf0051].[status] IN ('A','B')
            AND [r].[branch_id] = '0111'
        """,
        connect_DWH_CoSo,
    ).squeeze()

    bad_debt_path = r"\\192.168.10.101\phs-storge-2018\RiskManagementDept" \
                    r"\RMD_Data\Luu tru van ban\Daily Report\03. Not Matched" \
                    r"\DS nợ xấu\DS nợ xấu.xlsx"
    bad_debt = pd.read_excel(
        bad_debt_path,
        names=['no.','account_code','principal','interest','total','group'],
        skipfooter=1,
    )
    bad_debt['group'].fillna(method='ffill',inplace=True)
    bad_debt.drop(['interest','total'],axis=1,inplace=True)

    name = pd.read_sql(
        """
        SELECT [account].[account_code], [account].[customer_name] 
        FROM [account]
        """,
        connect_DWH_CoSo,
        index_col='account_code',
    ).squeeze()



    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File Excel ---------------------
    # Write file BÁO CÁO SỐ LIỆU THỐNG KÊ DVTC THÁNG
    for date_char in date_character:
        if date_char in start_date and date_char in end_date:
            start_date = start_date.replace(date_char, '/')
            end_date = end_date.replace(date_char, '/')
    som = dt.datetime.strptime(start_date, "%Y/%m/%d").strftime("%m-%Y")
    f_name = f'Số liệu thống kê DVTC tháng {som}.xlsx'
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
            'font_name': 'Arial',
            'bg_color': '#FFFF00',
        }
    )
    INB_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'bg_color': '#FFFF00',
            'color': '#00B050'
        }
    )
    ty_le_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
        }
    )
    num_ty_le_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'num_format': '#,##0.00'
        }
    )
    header_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'bg_color': '#FFC000'
        }
    )
    no_xau_header_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'bg_color': '#00B050'
        }
    )
    cuoi_ky_no_xau_values_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'color': '#FF0000',
            'num_format': '#,##0'
        }
    )
    normal_money_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'num_format': '#,##0'
        }
    )
    trong_ky_money_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'num_format': '#,##0.00;(#,##0.00)'
        }
    )
    trong_ky_num_account_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'num_format': '#,##0;(#,##0)'
        }
    )
    tong_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
        }
    )
    sum_money_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'num_format': '#,##0'
        }
    )
    text_left_wrap_text_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'text_wrap': True
        }
    )
    SSC_name_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'text_wrap': True,
            'bg_color': '#E0FFC1'
        }
    )
    SSC_empty_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'bg_color': '#E0FFC1'
        }
    )
    SSC_value_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'bg_color': '#E0FFC1'
        }
    )
    TK_margin_empty_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'bg_color': '#FFCCFF'
        }
    )

    text_left_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
        }
    )
    text_center_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
        }
    )
    text_center_wrap_text_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'text_wrap': True
        }
    )

    # --------- sheet BAO CAO CAN LAM ---------
    sheet_DVTC = workbook.add_worksheet('DVTC')
    sheet_title_name = 'TOÀN CÔNG TY'
    INB_name = 'INB - 01'
    ty_le = 'TỶ LỆ %'
    TK_margin = 'Tổng số lượng TK Margin (Active & Block)'
    TK_vip = 'Tổng số lượng TK VIP miễn lãi vay MR (Active & Block)'
    TK_use_DP = 'Tổng số lượng TK sử dụng hạn mức DP (Active &Block)'
    total_account_vipcn = 'Tổng số lượng TK VIPCN'
    total_account_silv = 'Tổng số lượng TK SILVER'
    total_account_gold = 'Tổng số lượng TK GOLDEN'
    header = [
        '`',
        'Loại',
        'Đầu kỳ',
        'Trong kỳ',
        'Cuối kỳ',
        'Nợ xấu',
        'Ghi chú'
    ]
    # Set Column Width and Row Height
    sheet_DVTC.set_column('A:A', 30.43)
    sheet_DVTC.set_column('B:B', 15)
    sheet_DVTC.set_column('C:C', 21.14)
    sheet_DVTC.set_column('D:D', 18.86)
    sheet_DVTC.set_column('E:E', 19.14)
    sheet_DVTC.set_column('F:F', 18)
    sheet_DVTC.set_column('G:G', 37)
    sheet_DVTC.set_column('H:H', 16.14)
    sheet_DVTC.set_column('I:I', 10.14)

    sheet_DVTC.merge_range('A1:G1', sheet_title_name, sheet_title_format)
    sheet_DVTC.merge_range('H1:H2', INB_name, INB_format)
    sheet_DVTC.merge_range('I1:I2', ty_le, ty_le_format)
    sheet_DVTC.write_row('A2', header, header_format)
    sheet_DVTC.write('F2', 'Nợ xấu', no_xau_header_format)
    sheet_DVTC.write_column('B3', sum_du_no_goc_dau_ky['type'], text_center_format)
    sheet_DVTC.write_column(
        'C3',
        sum_du_no_goc_dau_ky['principal_outstanding'],
        normal_money_format
    )
    sheet_DVTC.write_column(
        'E3',
        sum_du_no_goc_cuoi_ky['principal_outstanding'],
        cuoi_ky_no_xau_values_format
    )
    sheet_DVTC.write('B5', 'MR', TK_margin_empty_format)

    sum_row = sum_du_no_goc_dau_ky.shape[0] + 3
    sheet_DVTC.merge_range(f'A3:A{sum_row}', 'Tổng Dư nợ gốc', text_left_wrap_text_format)
    sheet_DVTC.write(f'A{sum_row}', '', text_left_wrap_text_format)
    sheet_DVTC.write(f'B{sum_row}', 'TỔNG', tong_format)
    sheet_DVTC.write(
        f'C{sum_row}',
        sum_du_no_goc_dau_ky['principal_outstanding'].sum(),
        sum_money_format
    )
    sheet_DVTC.write(
        f'D{sum_row}',
        du_no_goc_trong_ky['values'].abs().sum(),
        sum_money_format
    )
    du_no_goc_trong_ky['values'] = du_no_goc_trong_ky['values'].replace(0, '-')
    sheet_DVTC.write_column(
        'D3',
        du_no_goc_trong_ky['values'],
        trong_ky_money_format
    )
    sheet_DVTC.write(
        f'E{sum_row}',
        sum_du_no_goc_cuoi_ky['principal_outstanding'].sum(),
        sum_money_format
    )
    sheet_DVTC.write(
        f'A{sum_row + 1}',
        'Dư nợ báo cáo SSC',
        SSC_name_format
    )
    sheet_DVTC.write(f'B{sum_row + 1}', '', SSC_empty_format)
    sheet_DVTC.write(f'C{sum_row + 1}', '', SSC_value_format)
    sheet_DVTC.write(f'D{sum_row + 1}', '', SSC_value_format)
    sheet_DVTC.write(f'E{sum_row + 1}', '', SSC_value_format)
    sheet_DVTC.write(f'F{sum_row + 1}', '', SSC_value_format)

    # Tổng số lượng TK Margin (Active & Block)
    sheet_DVTC.write(f'A{sum_row + 2}', TK_margin, text_left_wrap_text_format)
    sheet_DVTC.write(f'B{sum_row + 2}', '', TK_margin_empty_format)
    sheet_DVTC.write(f'C{sum_row + 2}', amount_TK_dau_ky.shape[0], normal_money_format)
    sheet_DVTC.write(
        f'D{sum_row + 2}',
        amount_TK_cuoi_ky.shape[0] - amount_TK_dau_ky.shape[0],
        trong_ky_num_account_format
    )
    sheet_DVTC.write(f'E{sum_row + 2}', amount_TK_cuoi_ky.shape[0], cuoi_ky_no_xau_values_format)
    sheet_DVTC.write(
        f'H{sum_row + 2}',
        query_amount_account_INB.shape[0],
        cuoi_ky_no_xau_values_format
    )
    sheet_DVTC.write(
        f'I{sum_row + 2}',
        (query_amount_account_INB.shape[0]/amount_TK_cuoi_ky.shape[0]) * 100,
        num_ty_le_format
    )

    # Tổng số lượng TK VIP miễn lãi vay MR  (Active &Block)
    sheet_DVTC.merge_range(f'A{sum_row + 3}:A{sum_row + 6}', TK_vip, text_center_wrap_text_format)
    sheet_DVTC.write_column(f'B{sum_row + 3}', df_tk_vip_dau_ky.index, text_left_format)
    sheet_DVTC.write_column(f'C{sum_row + 3}', df_tk_vip_dau_ky['num_account'], normal_money_format)
    sheet_DVTC.write_column(
        f'D{sum_row + 3}',
        (df_tk_vip_cuoi_ky['num_account'] - df_tk_vip_dau_ky['num_account']).abs(),
        trong_ky_num_account_format
    )
    sheet_DVTC.write_column(f'E{sum_row + 3}', df_tk_vip_cuoi_ky['num_account'], cuoi_ky_no_xau_values_format)

    # Tổng số lượng TK sử dụng hạn mức DP  (Active &Block)
    sheet_DVTC.merge_range(f'A{sum_row + 7}:A{sum_row + 10}', TK_use_DP, text_center_wrap_text_format)
    sheet_DVTC.write_column(f'B{sum_row + 7}', df_tk_dp_dau_ky.index, text_left_format)
    sheet_DVTC.write_column(f'C{sum_row + 7}', df_tk_dp_dau_ky['num_account'], normal_money_format)
    sheet_DVTC.write_column(
        f'D{sum_row + 7}',
        df_tk_dp_cuoi_ky['num_account'] - df_tk_dp_dau_ky['num_account'],
        trong_ky_num_account_format
    )
    sheet_DVTC.write_column(f'E{sum_row + 7}', df_tk_dp_cuoi_ky['num_account'], cuoi_ky_no_xau_values_format)

    # Tổng số lượng TK VIPCN, TK SILVER, TK GOLDEN
    sheet_DVTC.write(f'A{sum_row + 11}', total_account_vipcn, text_left_format)
    sheet_DVTC.write(f'A{sum_row + 12}', total_account_silv, text_left_format)
    sheet_DVTC.write(f'A{sum_row + 13}', total_account_gold, text_left_format)
    sheet_DVTC.write_column(f'C{sum_row + 11}', df_tk_all_dau_ky['num_account'], normal_money_format)
    sheet_DVTC.write_column(
        f'D{sum_row + 11}',
        df_tk_all_cuoi_ky['num_account'] - df_tk_all_dau_ky['num_account'],
        trong_ky_num_account_format
    )
    sheet_DVTC.write_column(
        f'E{sum_row + 11}',
        df_tk_all_cuoi_ky['num_account'],
        cuoi_ky_no_xau_values_format
    )

    sheet_DVTC.write_column(f'B{sum_row + 11}', [''] * 3, text_center_format)

    # Ghi chú column
    sheet_DVTC.write_column(
        'G3',
        [''] * (sum_row + df_tk_vip_dau_ky.shape[0] + df_tk_dp_dau_ky.shape[0] + df_tk_all_dau_ky.shape[0]),
        text_center_format
    )
    sheet_DVTC.write(f'G{sum_row + 5}', 'HPX', text_left_format)

    # INB-01 column
    sheet_DVTC.write_column(
        'H3',
        sum_du_no_goc_cuoi_ky['INB-01'],
        cuoi_ky_no_xau_values_format
    )
    # Tỷ lệ column
    sheet_DVTC.write_column(
        'I3',
        sum_du_no_goc_cuoi_ky['ty_le'],
        num_ty_le_format
    )
    # DANH SÁCH KH NỢ XẤU

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py', '')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time() - start, 1)}s')