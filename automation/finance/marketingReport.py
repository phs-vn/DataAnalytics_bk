from automation.finance import *


def run(
    run_time=None
):
    start = time.time()
    info = get_info('weekly',run_time)
    start_date = '2022/02/14'
    end_date = dt.datetime.now().strftime('%Y/%m/%d')
    period = info['period']
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir((join(dept_folder,folder_name,period)))

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viết Query ---------------------
    # Xuất ra thông tin theo chiều ngang
    mtk_info = pd.read_sql(
        f"""
        WITH [b] AS (
            SELECT
                [vcf0051].[date],
                [relationship].[account_code],
                CASE
                    WHEN [relationship].[branch_id] IN ('0001','0101','0102','0104','0105','0106','0111','0113','0114','0116','0117','0118','0119','0999') THEN 'HoChiMinhCity'
                    WHEN [relationship].[branch_id] IN ('0201','0202','0203') THEN 'HaNoi'
                    WHEN [relationship].[branch_id] IN ('0301') THEN 'HaiPhong'
                    ELSE 'Unidentified'
                END [Location]
            FROM [vcf0051]
            LEFT JOIN [relationship] ON [vcf0051].[sub_account] = [relationship].[sub_account]
            AND [relationship].[date] = [vcf0051].[date]
            WHERE [vcf0051].[date] BETWEEN '2022-02-14' AND GETDATE()
            AND [vcf0051].[contract_type] LIKE N'Thường%'
        )
        SELECT
            [customer_change].[account_code],
            FORMAT([customer_change].[open_date], 'HH:mm - dd/MM/yyyy') [open_date_time],
            CONCAT(
                FORMAT(DATEADD(DAY, 2 - DATEPART(WEEKDAY, [customer_change].[open_date]), CAST([customer_change].[open_date] AS DATE)),'dd/MM')
                , '-', 
                FORMAT(DATEADD(DAY, 8 - DATEPART(WEEKDAY, [customer_change].[open_date]), CAST([customer_change].[open_date] AS DATE)),'dd/MM')
            ) [weekday],
            CASE
                WHEN YEAR(GETDATE()) - YEAR([account].[date_of_birth]) BETWEEN 18 AND 24 THEN '18-24'
                WHEN YEAR(GETDATE()) - YEAR([account].[date_of_birth]) BETWEEN 25 AND 34 THEN '25-34'
                WHEN YEAR(GETDATE()) - YEAR([account].[date_of_birth]) BETWEEN 35 AND 44 THEN '35-44'
                WHEN YEAR(GETDATE()) - YEAR([account].[date_of_birth]) BETWEEN 45 AND 54 THEN '45-54'
                WHEN YEAR(GETDATE()) - YEAR([account].[date_of_birth]) BETWEEN 55 AND 64 THEN '55-64'
                WHEN YEAR(GETDATE()) - YEAR([account].[date_of_birth]) >= 65 THEN '65+'
                ELSE 'Unidentified'
            END [AgeGroup],
            CASE
                WHEN [account].[gender] = '001' THEN 'Male'
                WHEN [account].[gender] = '002' THEN 'Female'
                ELSE 'Unidentified'
            END [Gender],
            [b].[Location],
            CASE
                WHEN RIGHT([account].[contract_number_normal],1) = 'I' THEN N'KH tự mở tài khoản'
                WHEN RIGHT([account].[contract_number_normal],1) = 'O' THEN N'KH mở tài khoản qua môi giới'
                ELSE 'Unidentified'
            END [Type]
        FROM [customer_change]
        LEFT JOIN [account] ON [customer_change].[account_code] = [account].[account_code]
        LEFT JOIN [b] ON [b].[account_code] = [customer_change].[account_code]
            AND [b].[date] = [customer_change].[open_date]
        WHERE [customer_change].[open_or_close] = 'Mo'
            AND [customer_change].[open_date] BETWEEN '{start_date}' AND '{end_date}'
            AND [account].[account_type] LIKE N'Cá nhân%'
        ORDER BY [customer_change].[open_date]
        """
        ,connect_DWH_CoSo
    )
    # Tổng số lượng mở tài khoản mới theo tuần
    mtk_sum = pd.read_sql(
        f"""
        SELECT
            CONCAT(
                FORMAT(DATEADD(DAY, 2 - DATEPART(WEEKDAY, [t].[open_date]), CAST([t].[open_date] AS DATE)),'dd/MM')
                , '-', 
                FORMAT(DATEADD(DAY, 8 - DATEPART(WEEKDAY, [t].[open_date]), CAST([t].[open_date] AS DATE)),'dd/MM')
            ) [week],
            COUNT([t].[account_code]) [NewAccounts]
        FROM [customer_change] [t]
        WHERE 
            [t].[open_or_close] = 'Mo'
            AND [t].[open_date] BETWEEN '{start_date}' AND '{end_date}'
            AND EXISTS (
                SELECT [account].[account_code] 
                FROM [account] 
                WHERE [account].[account_code] = [t].[account_code] AND [account].[account_type] LIKE N'Cá nhân%'
            )
        GROUP BY CONCAT(
            FORMAT(DATEADD(DAY, 2 - DATEPART(WEEKDAY, [t].[open_date]), CAST([t].[open_date] AS DATE)),'dd/MM')
            , '-', 
            FORMAT(DATEADD(DAY, 8 - DATEPART(WEEKDAY, [t].[open_date]), CAST([t].[open_date] AS DATE)),'dd/MM')
        )
        ORDER BY MONTH(MAX([open_date])), DAY(MAX([open_date]))
        """,
        connect_DWH_CoSo
    )

    # Yêu cầu số 1: Breakdown số lượng mở tài khoản mới theo tuần
    # 1a: breakdown theo tuổi

    mtk_age = pd.read_sql(
        f"""
        WITH [SourceTable] AS (
        SELECT 
            [customer_change].[account_code],
            CASE
                WHEN YEAR(GETDATE()) - YEAR([account].[date_of_birth]) BETWEEN 18 AND 24 THEN '18-24'
                WHEN YEAR(GETDATE()) - YEAR([account].[date_of_birth]) BETWEEN 25 AND 34 THEN '25-34'
                WHEN YEAR(GETDATE()) - YEAR([account].[date_of_birth]) BETWEEN 35 AND 44 THEN '35-44'
                WHEN YEAR(GETDATE()) - YEAR([account].[date_of_birth]) BETWEEN 45 AND 54 THEN '45-54'
                WHEN YEAR(GETDATE()) - YEAR([account].[date_of_birth]) BETWEEN 55 AND 64 THEN '55-64'
                WHEN YEAR(GETDATE()) - YEAR([account].[date_of_birth]) >= 65 THEN '65+'
                ELSE 'Unidentified'
            END [AgeGroup],
            CONCAT(
                FORMAT(DATEADD(DAY, 2 - DATEPART(WEEKDAY, [customer_change].[open_date]), CAST([customer_change].[open_date] AS DATE)),'dd/MM')
                , '-', 
                FORMAT(DATEADD(DAY, 8 - DATEPART(WEEKDAY, [customer_change].[open_date]), CAST([customer_change].[open_date] AS DATE)),'dd/MM')
            ) [weekday],
            DATEPART(WEEK, [customer_change].[open_date]) [week]
        FROM [customer_change]
        LEFT JOIN [account] ON [customer_change].[account_code] = [account].[account_code]
        WHERE [customer_change].[open_or_close] = 'Mo'
            AND [customer_change].[open_date] BETWEEN '{start_date}' AND '{end_date}'
            AND [account].[account_type] LIKE N'Cá nhân%'
        )
        SELECT
            [weekday], [18-24], [25-34], [35-44], [45-54], [55-64], [65+], [Unidentified]
        FROM [SourceTable]
        PIVOT (
            COUNT([account_code]) FOR [AgeGroup] IN ([18-24],[25-34],[35-44],[45-54],[55-64],[65+],[Unidentified])
        ) [PivotTable]
        ORDER BY [week]
        """,
        connect_DWH_CoSo
    )

    # 1d: breakdown theo nguồn mở tk
    mtk_type = pd.read_sql(
        f"""
        WITH [SourceTable] AS (
            SELECT 
                [customer_change].[account_code],
                CASE
                    WHEN RIGHT([account].[contract_number_normal],1) = 'I' THEN N'KH tự mở tài khoản'
                    WHEN RIGHT([account].[contract_number_normal],1) = 'O' THEN N'KH mở tài khoản qua môi giới'
                    ELSE 'Unidentified'
                END [Type],
                CONCAT(
                    FORMAT(DATEADD(DAY, 2 - DATEPART(WEEKDAY, [customer_change].[open_date]), CAST([customer_change].[open_date] AS DATE)),'dd/MM')
                    , '-', 
                    FORMAT(DATEADD(DAY, 8 - DATEPART(WEEKDAY, [customer_change].[open_date]), CAST([customer_change].[open_date] AS DATE)),'dd/MM')
                ) [weekday],
                DATEPART(WEEK, [customer_change].[open_date]) [week]
            FROM [customer_change]
            LEFT JOIN [account] ON [account].[account_code] = [customer_change].[account_code]
            WHERE [customer_change].[open_or_close] = 'Mo'
                AND [customer_change].[open_date] BETWEEN '{start_date}' AND '{end_date}'
                AND [account].[account_type] LIKE N'Cá nhân%'
        )
        SELECT
            [weekday], [KH tự mở tài khoản], [KH mở tài khoản qua môi giới], [Unidentified]
        FROM [SourceTable]
        PIVOT (
            COUNT([account_code]) FOR [Type] IN ([KH tự mở tài khoản], [KH mở tài khoản qua môi giới], [Unidentified])
        ) [PivotTable]
        ORDER BY [week]
        """,
        connect_DWH_CoSo
    )

    # Yêu cầu 3: Biến động mtk giữa các tuần (tính từ tuần 14/2) so với các tuần trước đó và các tuần cùng kỳ năm trước
    year = dt.datetime.strptime(start_date,'%Y/%m/%d').date().year
    # Lùi về 4 tuần liên tục trước đó tính từ ngày 14/2/2022
    predate = bdate(start_date, -20).replace('-', '/')  # 'Tuần từ 31/1 tới 6/2 ko có data vì là tuần nghỉ Tết dương lịch'
    # ngày đầu tiên của tuần thuộc ngày lùi về 4 tuần liên tục
    sdate = getDateRangeFromWeek(year, dt.datetime.strptime(predate,'%Y/%m/%d').date().isocalendar()[1])[0]
    sdate = sdate.strftime('%Y/%m/%d')
    # ngày cuối cùng của tuần 14/02/2022
    edate = getDateRangeFromWeek(year, dt.datetime.strptime(start_date,'%Y/%m/%d').date().isocalendar()[1])[1]
    edate = edate.strftime('%Y/%m/%d')

    # Tổng số lượng mở tài khoản mới 4 tuần liên tục trước đó
    mtk_sum_prev_week = pd.read_sql(
        f"""
        SELECT
            CONCAT(
                FORMAT(DATEADD(DAY, 2 - DATEPART(WEEKDAY, [t].[open_date]), CAST([t].[open_date] AS DATE)),'dd/MM')
                , '-', 
                FORMAT(DATEADD(DAY, 8 - DATEPART(WEEKDAY, [t].[open_date]), CAST([t].[open_date] AS DATE)),'dd/MM')
            ) [weekday],
            MAX(YEAR([t].[open_date])) [year],
            COUNT([t].[account_code]) [NewAccounts]
        FROM [customer_change] [t]
        WHERE 
            [t].[open_or_close] = 'Mo'
            AND [t].[open_date] BETWEEN '{sdate}' AND '{edate}'
            AND EXISTS (SELECT [account].[account_code] FROM [account] WHERE [account].[account_code] = [t].[account_code] AND [account].[account_type] LIKE N'Cá nhân%')
        GROUP BY CONCAT(
            FORMAT(DATEADD(DAY, 2 - DATEPART(WEEKDAY, [t].[open_date]), CAST([t].[open_date] AS DATE)),'dd/MM')
            , '-', 
            FORMAT(DATEADD(DAY, 8 - DATEPART(WEEKDAY, [t].[open_date]), CAST([t].[open_date] AS DATE)),'dd/MM')
        )
        ORDER BY MAX(DATEPART(WEEK, [t].[open_date]))
        """,
        connect_DWH_CoSo
    )
    mtk_age_prev_week = pd.read_sql(
        f"""
        WITH [SourceTable] AS (
            SELECT 
                [customer_change].[account_code],
                CASE
                    WHEN YEAR(GETDATE()) - YEAR([account].[date_of_birth]) BETWEEN 18 AND 24 THEN '18-24'
                    WHEN YEAR(GETDATE()) - YEAR([account].[date_of_birth]) BETWEEN 25 AND 34 THEN '25-34'
                    WHEN YEAR(GETDATE()) - YEAR([account].[date_of_birth]) BETWEEN 35 AND 44 THEN '35-44'
                    WHEN YEAR(GETDATE()) - YEAR([account].[date_of_birth]) BETWEEN 45 AND 54 THEN '45-54'
                    WHEN YEAR(GETDATE()) - YEAR([account].[date_of_birth]) BETWEEN 55 AND 64 THEN '55-64'
                    WHEN YEAR(GETDATE()) - YEAR([account].[date_of_birth]) >= 65 THEN '65+'
                    ELSE 'Unidentified'
                END [AgeGroup],
                CONCAT(
                    FORMAT(DATEADD(DAY, 2 - DATEPART(WEEKDAY, [customer_change].[open_date]), CAST([customer_change].[open_date] AS DATE)),'dd/MM')
                    , '-', 
                    FORMAT(DATEADD(DAY, 8 - DATEPART(WEEKDAY, [customer_change].[open_date]), CAST([customer_change].[open_date] AS DATE)),'dd/MM')
                ) [weekday],
                DATEPART(WEEK, [customer_change].[open_date]) [week],
                YEAR([customer_change].[open_date]) [year]
            FROM [customer_change]
            LEFT JOIN [account] ON [customer_change].[account_code] = [account].[account_code]
            WHERE [customer_change].[open_or_close] = 'Mo'
                AND [customer_change].[open_date] BETWEEN '{sdate}' AND '{edate}'
                AND [account].[account_type] LIKE N'Cá nhân%'
        )
        SELECT
            [weekday], [year], [18-24], [25-34], [35-44], [45-54], [55-64], [65+], [Unidentified]
        FROM [SourceTable]
        PIVOT (
            COUNT([account_code]) FOR [AgeGroup] IN ([18-24],[25-34],[35-44],[45-54],[55-64],[65+],[Unidentified])
        ) [PivotTable]
        ORDER BY [week]
            """,
        connect_DWH_CoSo
    )
    mtk_type_prev_week = pd.read_sql(
        f"""
        WITH [SourceTable] AS (
            SELECT 
                [customer_change].[account_code],
                CASE
                    WHEN RIGHT([account].[contract_number_normal],1) = 'I' THEN N'KH tự mở tài khoản'
                    WHEN RIGHT([account].[contract_number_normal],1) = 'O' THEN N'KH mở tài khoản qua môi giới'
                    ELSE 'Unidentified'
                END [Type],
                CONCAT(
                    FORMAT(DATEADD(DAY, 2 - DATEPART(WEEKDAY, [customer_change].[open_date]), CAST([customer_change].[open_date] AS DATE)),'dd/MM')
                    , '-', 
                    FORMAT(DATEADD(DAY, 8 - DATEPART(WEEKDAY, [customer_change].[open_date]), CAST([customer_change].[open_date] AS DATE)),'dd/MM')
                ) [weekday],
                DATEPART(WEEK, [customer_change].[open_date]) [week],
                YEAR([customer_change].[open_date]) [year]
            FROM [customer_change]
            LEFT JOIN [account] ON [account].[account_code] = [customer_change].[account_code]
            WHERE [customer_change].[open_or_close] = 'Mo'
                AND [customer_change].[open_date] BETWEEN '{sdate}' AND '{edate}'
                AND [account].[account_type] LIKE N'Cá nhân%'
        )
        SELECT
            [weekday], [year], [KH tự mở tài khoản], [KH mở tài khoản qua môi giới], [Unidentified]
        FROM [SourceTable]
        PIVOT (
            COUNT([account_code]) FOR [Type] IN ([KH tự mở tài khoản], [KH mở tài khoản qua môi giới], [Unidentified])
        ) [PivotTable]
        ORDER BY [week]
        """,
        connect_DWH_CoSo
    )
    # Tổng số lượng mở tài khoản mới các tuần cùng kỳ năm trước
    year_prev = dt.datetime.strptime(start_date, '%Y/%m/%d').date().year - 1

    date_prev_year = dt.datetime.strptime(start_date, '%Y/%m/%d').date()
    # đổi 2022 thành 2021
    date_prev_year = date_prev_year.replace(year=date_prev_year.year - 1).strftime('%Y/%m/%d')
    # Lùi về 4 tuần liên tục trước đó tính từ ngày 14/2/2021
    predate_prev_year = bdate(date_prev_year, -20).replace('-', '/')
    # ngày đầu tiên của tuần thuộc ngày sau khi lùi về liên tục 4 tuần
    sdate_prev_year = getDateRangeFromWeek(year_prev, dt.datetime.strptime(predate_prev_year, '%Y/%m/%d').date().isocalendar()[1])[0]
    sdate_prev_year = sdate_prev_year.strftime('%Y/%m/%d')
    # ngày cuối cùng của tuần 14/02/2021
    edate_prev_year = getDateRangeFromWeek(year_prev, dt.datetime.strptime(date_prev_year, '%Y/%m/%d').date().isocalendar()[1])[1]
    edate_prev_year = edate_prev_year.strftime('%Y/%m/%d')

    mtk_sum_prev_year = pd.read_sql(
        f"""
        SELECT
           CONCAT(
               FORMAT(DATEADD(DAY, 2 - DATEPART(WEEKDAY, [t].[open_date]), CAST([t].[open_date] AS DATE)),'dd/MM')
               , '-', 
               FORMAT(DATEADD(DAY, 8 - DATEPART(WEEKDAY, [t].[open_date]), CAST([t].[open_date] AS DATE)),'dd/MM')
           ) [weekday],
           MAX(YEAR([t].[open_date])) [year],
           COUNT([t].[account_code]) [NewAccounts]
        FROM [customer_change] [t]
        WHERE 
           [t].[open_or_close] = 'Mo'
           AND [t].[open_date] BETWEEN '{sdate_prev_year}' AND '{edate_prev_year}'
           AND EXISTS (SELECT [account].[account_code] FROM [account] WHERE [account].[account_code] = [t].[account_code] AND [account].[account_type] LIKE N'Cá nhân%')
        GROUP BY CONCAT(
           FORMAT(DATEADD(DAY, 2 - DATEPART(WEEKDAY, [t].[open_date]), CAST([t].[open_date] AS DATE)),'dd/MM')
           , '-', 
           FORMAT(DATEADD(DAY, 8 - DATEPART(WEEKDAY, [t].[open_date]), CAST([t].[open_date] AS DATE)),'dd/MM')
        )
        ORDER BY MAX(DATEPART(WEEK, [t].[open_date]))
        """,
        connect_DWH_CoSo
    )
    mtk_age_prev_year = pd.read_sql(
        f"""
        WITH [SourceTable] AS (
            SELECT 
                [customer_change].[account_code],
                CASE
                    WHEN YEAR(GETDATE()) - YEAR([account].[date_of_birth]) BETWEEN 18 AND 24 THEN '18-24'
                    WHEN YEAR(GETDATE()) - YEAR([account].[date_of_birth]) BETWEEN 25 AND 34 THEN '25-34'
                    WHEN YEAR(GETDATE()) - YEAR([account].[date_of_birth]) BETWEEN 35 AND 44 THEN '35-44'
                    WHEN YEAR(GETDATE()) - YEAR([account].[date_of_birth]) BETWEEN 45 AND 54 THEN '45-54'
                    WHEN YEAR(GETDATE()) - YEAR([account].[date_of_birth]) BETWEEN 55 AND 64 THEN '55-64'
                    WHEN YEAR(GETDATE()) - YEAR([account].[date_of_birth]) >= 65 THEN '65+'
                    ELSE 'Unidentified'
                END [AgeGroup],
                CONCAT(
                    FORMAT(DATEADD(DAY, 2 - DATEPART(WEEKDAY, [customer_change].[open_date]), CAST([customer_change].[open_date] AS DATE)),'dd/MM')
                    , '-', 
                    FORMAT(DATEADD(DAY, 8 - DATEPART(WEEKDAY, [customer_change].[open_date]), CAST([customer_change].[open_date] AS DATE)),'dd/MM')
                ) [weekday],
                YEAR([customer_change].[open_date]) [year],
                DATEPART(WEEK, [customer_change].[open_date]) [week]
            FROM [customer_change]
            LEFT JOIN [account] ON [customer_change].[account_code] = [account].[account_code]
            WHERE [customer_change].[open_or_close] = 'Mo'
                AND [customer_change].[open_date] BETWEEN '{sdate_prev_year}' AND '{edate_prev_year}'
                AND [account].[account_type] LIKE N'Cá nhân%'
        )
        SELECT
            [weekday], [year], [18-24], [25-34], [35-44], [45-54], [55-64], [65+], [Unidentified]
        FROM [SourceTable]
        PIVOT (
            COUNT([account_code]) FOR [AgeGroup] IN ([18-24],[25-34],[35-44],[45-54],[55-64],[65+],[Unidentified])
        ) [PivotTable]
        ORDER BY [week]
        """,
        connect_DWH_CoSo
    )
    mtk_type_prev_year = pd.read_sql(
        f"""
        WITH [SourceTable] AS (
           SELECT 
               [customer_change].[account_code],
               CASE
                   WHEN RIGHT([account].[contract_number_normal],1) = 'I' THEN N'KH tự mở tài khoản'
                   WHEN RIGHT([account].[contract_number_normal],1) = 'O' THEN N'KH mở tài khoản qua môi giới'
                   ELSE 'Unidentified'
               END [Type],
               CONCAT(
                   FORMAT(DATEADD(DAY, 2 - DATEPART(WEEKDAY, [customer_change].[open_date]), CAST([customer_change].[open_date] AS DATE)),'dd/MM')
                   , '-', 
                   FORMAT(DATEADD(DAY, 8 - DATEPART(WEEKDAY, [customer_change].[open_date]), CAST([customer_change].[open_date] AS DATE)),'dd/MM')
               ) [weekday],
               DATEPART(WEEK, [customer_change].[open_date]) [week],
               YEAR([customer_change].[open_date]) [year]
           FROM [customer_change]
           LEFT JOIN [account] ON [account].[account_code] = [customer_change].[account_code]
           WHERE [customer_change].[open_or_close] = 'Mo'
               AND [customer_change].[open_date] BETWEEN '{sdate_prev_year}' AND '{edate_prev_year}'
               AND [account].[account_type] LIKE N'Cá nhân%'
        )
        SELECT
           [weekday], [year], [KH tự mở tài khoản], [KH mở tài khoản qua môi giới], [Unidentified]
        FROM [SourceTable]
        PIVOT (
           COUNT([account_code]) FOR [Type] IN ([KH tự mở tài khoản], [KH mở tài khoản qua môi giới], [Unidentified])
        ) [PivotTable]
        ORDER BY [week]
        """,
        connect_DWH_CoSo
    )

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File Excel ---------------------
    # Write file excel Bao cao Marketing
    file_name = 'Báo Cáo Marketing.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    ###################################################
    ###################################################
    ###################################################

    # Format
    header_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'text_wrap': True,
            'font_size':11,
            'font_name':'Times New Roman',
        }
    )
    num_right_format = workbook.add_format(
        {
            'border':1,
            'align':'right',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman'
        }
    )
    text_left_format = workbook.add_format(
        {
            'border':1,
            'align':'left',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Times New Roman'
        }
    )
    text_center_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman'
        }
    )

    # ------------------ Viết Sheet Marketing report ------------------
    # --------- sheet Xuất ra thông tin theo chiều ngang phục vụ mục đích Khảo sát ---------
    sheet_ks = workbook.add_worksheet('Khảo Sát')
    # header
    headers = [
        'STT',
        'Mã số TK',
        'Tuần \n (từ ngày tới ngày)',
        'Thời gian mở tài khoản (HH:MM - DD/MM/YYYY)',
        'Độ tuổi',
        'Giới tính',
        'Địa chỉ TK',
        'Nguồn mở TK'
    ]

    # Set Column Width and Row Height
    sheet_ks.set_column('A:A',5)
    sheet_ks.set_column('B:B',12)
    sheet_ks.set_column('C:C',17.5)
    sheet_ks.set_column('D:D',27.5)
    sheet_ks.set_column('E:E',10)
    sheet_ks.set_column('F:F',8.5)
    sheet_ks.set_column('G:G',14)
    sheet_ks.set_column('H:H',26)

    # Xử lý Dataframe
    sheet_ks.write_row('A1', headers, header_format)
    sheet_ks.write_column('A2', np.arange(mtk_info.shape[0]) + 1, num_right_format)
    sheet_ks.write_column('B2', mtk_info['account_code'], text_left_format)
    sheet_ks.write_column('C2', mtk_info['weekday'], text_center_format)
    sheet_ks.write_column('D2', mtk_info['open_date_time'], text_center_format)
    sheet_ks.write_column('E2', mtk_info['AgeGroup'], text_center_format)
    sheet_ks.write_column('F2', mtk_info['Gender'], text_left_format)
    sheet_ks.write_column('G2', mtk_info['Location'], text_left_format)
    sheet_ks.write_column('H2', mtk_info['Type'], text_left_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    # --------- sheet Tổng số lượng mở tài khoản mới ---------
    sheet_sum = workbook.add_worksheet('mtk-Tổng')
    title = 'Tổng số lượng mở tài khoản mới'
    # header
    headers = [
        'STT',
        'Tuần \n (từ ngày tới ngày)',
        'Số lượng mở tài khoản mới'
    ]

    # Set Column Width and Row Height
    sheet_sum.set_column('A:A', 5)
    sheet_sum.set_column('B:B', 17.5)
    sheet_sum.set_column('C:C', 17.5)

    # Xử lý Dataframe
    sheet_sum.merge_range('A1:C1', title, header_format)
    sheet_sum.write_row('A2', headers, header_format)
    sheet_sum.write_column('A3', np.arange(mtk_sum.shape[0]) + 1, num_right_format)
    sheet_sum.write_column('B3', mtk_sum['week'], text_center_format)
    sheet_sum.write_column('C3', mtk_sum['NewAccounts'], text_left_format)

    # draw chart
    chart = workbook.add_chart({'type': 'column'})

    chart.add_series({
        'name':'number of accounts',
        'categories': ['mtk-Tổng', 2, 1, mtk_sum.shape[0] + 2, 1],
        'values': f'=mtk-Tổng!$C$3:$C${mtk_sum.shape[0]+2}',
    })
    sheet_sum.insert_chart('F10', chart, {'x_scale': 1.75, 'y_scale': 1})

    ###########################################################################
    ###########################################################################
    ###########################################################################

    # --------- sheet số lượng mở tài khoản mới theo độ tuổi ---------
    sheet_age = workbook.add_worksheet('mtk-Tuổi')
    title = 'Tổng số lượng mở tài khoản mới theo độ tuổi'
    # header
    headers = [
        'STT',
        'Tuần \n (từ ngày tới ngày)',
        '18-24',
        '25-34',
        '35-44',
        '45-54',
        '55-64',
        '65+',
        'Không xác định'
    ]

    # Set Column Width and Row Height
    sheet_age.set_column('A:A', 5)
    sheet_age.set_column('B:B', 17.5)
    sheet_age.set_column('C:I', 10)

    # Xử lý Dataframe
    sheet_age.merge_range('A1:I1', title, header_format)
    sheet_age.write_row('A2', headers, header_format)
    sheet_age.write_column('A3', np.arange(mtk_age.shape[0]) + 1, num_right_format)
    sheet_age.write_column('B3', mtk_age['weekday'], text_center_format)
    sheet_age.write_column('C3', mtk_age['18-24'], num_right_format)
    sheet_age.write_column('D3', mtk_age['25-34'], num_right_format)
    sheet_age.write_column('E3', mtk_age['35-44'], num_right_format)
    sheet_age.write_column('F3', mtk_age['45-54'], num_right_format)
    sheet_age.write_column('G3', mtk_age['55-64'], num_right_format)
    sheet_age.write_column('H3', mtk_age['65+'], num_right_format)
    sheet_age.write_column('I3', mtk_age['Unidentified'], num_right_format)

    # draw chart
    chart = workbook.add_chart({'type': 'column'})

    chart.add_series({
        'name': '18-24',
        'categories': ['mtk-Tuổi', 2, 1, mtk_age.shape[0] + 2, 1],
        'values': F'=mtk-Tuổi!$C$3:$C${mtk_age.shape[0]+2}',
    })
    chart.add_series({
        'name': '25-34',
        'categories': ['mtk-Tuổi', 2, 1, mtk_age.shape[0] + 2, 1],
        'values': f'=mtk-Tuổi!$D$3:$D${mtk_age.shape[0]+2}',
    })
    chart.add_series({
        'name': '35-44',
        'categories': ['mtk-Tuổi', 2, 1, mtk_age.shape[0] + 2, 1],
        'values': f'=mtk-Tuổi!$E$3:$E${mtk_age.shape[0]+2}',
    })
    chart.add_series({
        'name': '45-54',
        'categories': ['mtk-Tuổi', 2, 1, mtk_age.shape[0] + 2, 1],
        'values': f'=mtk-Tuổi!$F$3:$F${mtk_age.shape[0]+2}',
    })
    chart.add_series({
        'name': '55-64',
        'categories': ['mtk-Tuổi', 2, 1, mtk_age.shape[0] + 2, 1],
        'values': f'=mtk-Tuổi!$G$3:$G${mtk_age.shape[0]+2}',
    })
    chart.add_series({
        'name': '65+',
        'categories': ['mtk-Tuổi', 2, 1, mtk_age.shape[0] + 2, 1],
        'values': f'=mtk-Tuổi!$H$3:$H${mtk_age.shape[0]+2}',
    })
    chart.add_series({
        'name': 'Không xác định',
        'categories': ['mtk-Tuổi', 2, 1, mtk_age.shape[0] + 2, 1],
        'values': f'=mtk-Tuổi!$I$3:$I${mtk_age.shape[0]+2}',
    })
    chart.set_title({'name': 'Number of accounts by age'})
    sheet_age.insert_chart('E12', chart, {'x_scale': 2, 'y_scale': 1})

    ###########################################################################
    ###########################################################################
    ###########################################################################

    # --------- sheet số lượng mở tài khoản mới theo nguồn ---------
    sheet_open_source = workbook.add_worksheet('mtk-Nguồn')
    title = 'Tổng số lượng mở tài khoản mới theo nguồn'
    # header
    headers = [
        'STT',
        'Tuần \n (từ ngày tới ngày)',
        'KH tự mở tài khoản',
        'KH mở tài khoản qua môi giới',
        'Không xác định'
    ]

    # Set Column Width and Row Height
    sheet_open_source.set_column('A:A', 5)
    sheet_open_source.set_column('B:B', 17.5)
    sheet_open_source.set_column('C:E', 12)

    # Xử lý Dataframe
    sheet_open_source.merge_range('A1:E1', title, header_format)
    sheet_open_source.write_row('A2', headers, header_format)
    sheet_open_source.write_column('A3', np.arange(mtk_type.shape[0]) + 1, num_right_format)
    sheet_open_source.write_column('B3', mtk_type['weekday'], text_center_format)
    sheet_open_source.write_column('C3', mtk_type['KH tự mở tài khoản'], num_right_format)
    sheet_open_source.write_column('D3', mtk_type['KH mở tài khoản qua môi giới'], num_right_format)
    sheet_open_source.write_column('E3', mtk_type['Unidentified'], num_right_format)

    # draw chart
    chart = workbook.add_chart({'type': 'column'})

    chart.add_series({
        'name': 'KH tự mở tài khoản',
        'categories': ['mtk-Nguồn', 2, 1, mtk_type.shape[0] + 2, 1],
        'values': F'=mtk-Nguồn!$C$3:$C${mtk_type.shape[0] + 2}',
    })
    chart.add_series({
        'name': 'KH mở tài khoản qua môi giới',
        'categories': ['mtk-Nguồn', 2, 1, mtk_type.shape[0] + 2, 1],
        'values': f'=mtk-Nguồn!$D$3:$D${mtk_type.shape[0] + 2}',
    })
    chart.add_series({
        'name': 'Không xác định',
        'categories': ['mtk-Nguồn', 2, 1, mtk_type.shape[0] + 2, 1],
        'values': f'=mtk-Nguồn!$E$3:$E${mtk_type.shape[0] + 2}',
    })
    chart.set_title({'name': 'Number of accounts by account open source'})
    sheet_open_source.insert_chart('E14', chart, {'x_scale': 2, 'y_scale': 1})

    ###########################################################################
    ###########################################################################
    ###########################################################################

    # --------- sheet Tổng số lượng mở tài khoản mới 4 tuần liên tục trước đó và 4 tuần cùng kì 2021 ---------
    sheet_sum_prev = workbook.add_worksheet('mtk-Tổng-previous')
    title = 'Tổng số lượng mở tài khoản mới (tính từ tuần 14/02 so với 4 tuần liên tục trước đó và 4 tuần cùng kì 2021)'
    # header
    headers = [
        'STT',
        'Tuần \n (từ ngày tới ngày)',
        'Năm',
        'Số lượng mở tài khoản mới'
    ]

    # Set Column Width and Row Height
    sheet_sum_prev.set_column('A:A', 5)
    sheet_sum_prev.set_column('B:B', 20)
    sheet_sum_prev.set_column('C:C', 8)
    sheet_sum_prev.set_column('D:D', 17.5)
    sheet_sum_prev.set_row(0, 47)

    # Xử lý Dataframe
    sheet_sum_prev.merge_range('A1:D1', title, header_format)
    sheet_sum_prev.write_row('A3', headers, header_format)
    sheet_sum_prev.write_column('A4', np.arange(mtk_sum_prev_week.shape[0]) + 1, num_right_format)
    sheet_sum_prev.write_column('B4', mtk_sum_prev_week['weekday'], text_center_format)
    sheet_sum_prev.write_column('C4', mtk_sum_prev_week['year'], num_right_format)
    sheet_sum_prev.write_column('D4', mtk_sum_prev_week['NewAccounts'], text_left_format)

    new_row = mtk_sum_prev_week.shape[0] + 11
    sheet_sum_prev.write_row(f'A{new_row}', headers, header_format)
    sheet_sum_prev.write_column(f'A{new_row+1}', np.arange(mtk_sum_prev_year.shape[0]) + 1, num_right_format)
    sheet_sum_prev.write_column(f'B{new_row+1}', mtk_sum_prev_year['weekday'], text_center_format)
    sheet_sum_prev.write_column(f'C{new_row+1}', mtk_sum_prev_year['year'], num_right_format)
    sheet_sum_prev.write_column(f'D{new_row+1}', mtk_sum_prev_year['NewAccounts'], text_left_format)

    # draw chart
    # chart 2022
    chart = workbook.add_chart({'type': 'column'})
    chart.add_series({
        'name': 'number of accounts 2022',
        'categories': ['mtk-Tổng-previous', 3, 1, mtk_sum_prev_week.shape[0] + 2, 1],
        'values': f'=mtk-Tổng-previous!$D$4:$D${mtk_sum_prev_week.shape[0] + 3}',
    })
    sheet_sum_prev.insert_chart('F1', chart, {'x_scale': 1.3, 'y_scale': 0.8})
    # chart 2021
    chart1 = workbook.add_chart({'type': 'column'})
    chart1.add_series({
        'name': 'number of accounts 2021',
        'categories': ['mtk-Tổng-previous',mtk_sum_prev_week.shape[0]+11,1,mtk_sum_prev_week.shape[0]+mtk_sum_prev_year.shape[0]+11,1],
        'values': f'=mtk-Tổng-previous!$D${mtk_sum_prev_week.shape[0]+12}:$D${mtk_sum_prev_week.shape[0]+mtk_sum_prev_year.shape[0]+11}',
    })
    sheet_sum_prev.insert_chart(f'F{mtk_age_prev_week.shape[0]+9}', chart1, {'x_scale': 1.3, 'y_scale': 0.8})

    ###########################################################################
    ###########################################################################
    ###########################################################################

    # --------- sheet số lượng mở tài khoản mới theo độ tuổi 4 tuần liên tục trước đó và 4 tuần cùng kì 2021 ---------
    sheet_age_prev = workbook.add_worksheet('mtk-Tuổi-previous')
    title = 'Tổng số lượng mở tài khoản mới theo độ tuổi (tính từ tuần 14/02 so với 4 tuần liên tục trước đó và 4 ' \
            'tuần cùng kì 2021)'
    # header
    headers = [
        'STT',
        'Tuần \n (từ ngày tới ngày)',
        'Năm',
        '18-24',
        '25-34',
        '35-44',
        '45-54',
        '55-64',
        '65+',
        'Không xác định'
    ]

    # Set Column Width and Row Height
    sheet_age_prev.set_column('A:A', 4.5)
    sheet_age_prev.set_column('B:B', 17.5)
    sheet_age_prev.set_column('C:C', 5)
    sheet_age_prev.set_column('D:H', 5.6)
    sheet_age_prev.set_column('I:I', 4)
    sheet_age_prev.set_column('J:J', 8)
    sheet_age_prev.set_row(0, 34.5)

    # Xử lý Dataframe
    sheet_age_prev.merge_range('A1:J1', title, header_format)
    sheet_age_prev.write_row('A3', headers, header_format)
    sheet_age_prev.write_column('A4', np.arange(mtk_age_prev_week.shape[0]) + 1, num_right_format)
    sheet_age_prev.write_column('B4', mtk_age_prev_week['weekday'], text_center_format)
    sheet_age_prev.write_column('C4', mtk_age_prev_week['year'], num_right_format)
    sheet_age_prev.write_column('D4', mtk_age_prev_week['18-24'], num_right_format)
    sheet_age_prev.write_column('E4', mtk_age_prev_week['25-34'], num_right_format)
    sheet_age_prev.write_column('F4', mtk_age_prev_week['35-44'], num_right_format)
    sheet_age_prev.write_column('G4', mtk_age_prev_week['45-54'], num_right_format)
    sheet_age_prev.write_column('H4', mtk_age_prev_week['55-64'], num_right_format)
    sheet_age_prev.write_column('I4', mtk_age_prev_week['65+'], num_right_format)
    sheet_age_prev.write_column('J4', mtk_age_prev_week['Unidentified'], num_right_format)

    new_row = mtk_age_prev_week.shape[0] + 12
    sheet_age_prev.write_row(f'A{new_row}', headers, header_format)
    sheet_age_prev.write_column(f'A{new_row+1}', np.arange(mtk_age_prev_year.shape[0]) + 1, num_right_format)
    sheet_age_prev.write_column(f'B{new_row+1}', mtk_age_prev_year['weekday'], text_center_format)
    sheet_age_prev.write_column(f'C{new_row+1}', mtk_age_prev_year['year'], num_right_format)
    sheet_age_prev.write_column(f'D{new_row+1}', mtk_age_prev_year['18-24'], num_right_format)
    sheet_age_prev.write_column(f'E{new_row+1}', mtk_age_prev_year['25-34'], num_right_format)
    sheet_age_prev.write_column(f'F{new_row+1}', mtk_age_prev_year['35-44'], num_right_format)
    sheet_age_prev.write_column(f'G{new_row+1}', mtk_age_prev_year['45-54'], num_right_format)
    sheet_age_prev.write_column(f'H{new_row+1}', mtk_age_prev_year['55-64'], num_right_format)
    sheet_age_prev.write_column(f'I{new_row+1}', mtk_age_prev_year['65+'], num_right_format)
    sheet_age_prev.write_column(f'J{new_row+1}', mtk_age_prev_year['Unidentified'], num_right_format)

    # draw chart
    # chart 2022
    chart = workbook.add_chart({'type': 'column'})
    chart.add_series({
        'name': '18-24',
        'categories': ['mtk-Tuổi-previous', 3, 1, mtk_age_prev_week.shape[0] + 2, 1],
        'values': F'=mtk-Tuổi-previous!$D$4:$D${mtk_age_prev_week.shape[0] + 3}',
    })
    chart.add_series({
        'name': '25-34',
        'categories': ['mtk-Tuổi-previous', 3, 1, mtk_age_prev_week.shape[0] + 2, 1],
        'values': f'=mtk-Tuổi-previous!$E$4:$E${mtk_age_prev_week.shape[0] + 3}',
    })
    chart.add_series({
        'name': '35-44',
        'categories': ['mtk-Tuổi-previous', 3, 1, mtk_age_prev_week.shape[0] + 2, 1],
        'values': f'=mtk-Tuổi-previous!$F$4:$F${mtk_age_prev_week.shape[0] + 3}',
    })
    chart.add_series({
        'name': '45-54',
        'categories': ['mtk-Tuổi-previous', 3, 1, mtk_age_prev_week.shape[0] + 2, 1],
        'values': f'=mtk-Tuổi-previous!$G$4:$G${mtk_age_prev_week.shape[0] + 3}',
    })
    chart.add_series({
        'name': '55-64',
        'categories': ['mtk-Tuổi-previous', 3, 1, mtk_age_prev_week.shape[0] + 2, 1],
        'values': f'=mtk-Tuổi-previous!$H$4:$H${mtk_age_prev_week.shape[0] + 3}',
    })
    chart.add_series({
        'name': '65+',
        'categories': ['mtk-Tuổi-previous', 3, 1, mtk_age_prev_week.shape[0] + 2, 1],
        'values': f'=mtk-Tuổi-previous!$I$4:$I${mtk_age_prev_week.shape[0] + 3}',
    })
    chart.add_series({
        'name': 'Không xác định',
        'categories': ['mtk-Tuổi-previous', 3, 1, mtk_age_prev_week.shape[0] + 2, 1],
        'values': f'=mtk-Tuổi-previous!$J$4:$J${mtk_age_prev_week.shape[0] + 3}',
    })
    chart.set_title({'name': 'Number of accounts by age 2022'})
    sheet_age_prev.insert_chart('L1', chart, {'x_scale': 1.4, 'y_scale': 0.8})
    # chart 2021
    chart1 = workbook.add_chart({'type': 'column'})
    chart1.add_series({
        'name': '18-24',
        'categories': ['mtk-Tuổi-previous', mtk_age_prev_week.shape[0]+12,1,mtk_age_prev_week.shape[0]+mtk_age_prev_year.shape[0]+11,1],
        'values': f'=mtk-Tuổi-previous!$D${mtk_age_prev_week.shape[0]+13}:$D${mtk_age_prev_week.shape[0]+mtk_age_prev_year.shape[0]+12}',
    })
    chart1.add_series({
        'name': '25-34',
        'categories': ['mtk-Tuổi-previous', mtk_age_prev_week.shape[0]+12,1,mtk_age_prev_week.shape[0]+mtk_age_prev_year.shape[0]+11,1],
        'values': f'=mtk-Tuổi-previous!$E${mtk_age_prev_week.shape[0]+13}:$E${mtk_age_prev_week.shape[0]+mtk_age_prev_year.shape[0]+12}',
    })
    chart1.add_series({
        'name': '35-44',
        'categories': ['mtk-Tuổi-previous', mtk_age_prev_week.shape[0]+12,1,mtk_age_prev_week.shape[0]+mtk_age_prev_year.shape[0]+11,1],
        'values': f'=mtk-Tuổi-previous!$F${mtk_age_prev_week.shape[0]+13}:$F${mtk_age_prev_week.shape[0]+mtk_age_prev_year.shape[0]+12}',
    })
    chart1.add_series({
        'name': '45-54',
        'categories': ['mtk-Tuổi-previous', mtk_age_prev_week.shape[0]+12,1,mtk_age_prev_week.shape[0]+mtk_age_prev_year.shape[0]+11,1],
        'values': f'=mtk-Tuổi-previous!$G${mtk_age_prev_week.shape[0]+13}:$G${mtk_age_prev_week.shape[0]+mtk_age_prev_year.shape[0]+12}',
    })
    chart1.add_series({
        'name': '55-64',
        'categories': ['mtk-Tuổi-previous', mtk_age_prev_week.shape[0]+12,1,mtk_age_prev_week.shape[0]+mtk_age_prev_year.shape[0]+11,1],
        'values': f'=mtk-Tuổi-previous!$H${mtk_age_prev_week.shape[0]+13}:$H${mtk_age_prev_week.shape[0]+mtk_age_prev_year.shape[0]+12}',
    })
    chart1.add_series({
        'name': '65+',
        'categories': ['mtk-Tuổi-previous', mtk_age_prev_week.shape[0]+12,1,mtk_age_prev_week.shape[0]+mtk_age_prev_year.shape[0]+11,1],
        'values': f'=mtk-Tuổi-previous!$I${mtk_age_prev_week.shape[0]+13}:$I${mtk_age_prev_week.shape[0]+mtk_age_prev_year.shape[0]+12}',
    })
    chart1.add_series({
        'name': 'Không xác định',
        'categories': ['mtk-Tuổi-previous', mtk_age_prev_week.shape[0]+12,1,mtk_age_prev_week.shape[0]+mtk_age_prev_year.shape[0]+11,1],
        'values': f'=mtk-Tuổi-previous!$J${mtk_age_prev_week.shape[0]+13}:$J${mtk_age_prev_week.shape[0]+mtk_age_prev_year.shape[0]+12}',
    })
    chart1.set_title({'name': 'Number of accounts by age 2021'})
    sheet_age_prev.insert_chart('L14', chart1, {'x_scale': 1.4, 'y_scale': 0.8})

    # --------- sheet số lượng mở tài khoản mới theo nguồn 4 tuần liên tục trước đó và 4 tuần cùng kì 2021 ---------
    sheet_open_source_prev = workbook.add_worksheet('mtk-Nguồn-previous')
    title = 'Tổng số lượng mở tài khoản mới theo nguồn (tính từ tuần 14/02 so với 4 tuần liên tục trước đó và 4 ' \
            'tuần cùng kì 2021)'
    # header
    headers = [
        'STT',
        'Tuần \n (từ ngày tới ngày)',
        'Năm',
        'KH tự mở tài khoản',
        'KH mở tài khoản qua môi giới',
        'Không xác định'
    ]

    # Set Column Width and Row Height
    sheet_open_source_prev.set_column('A:A', 5)
    sheet_open_source_prev.set_column('B:B', 17.5)
    sheet_open_source_prev.set_column('C:C', 5)
    sheet_open_source_prev.set_column('D:F', 12)
    sheet_open_source_prev.set_row(0, 33)

    # Xử lý Dataframe
    sheet_open_source_prev.merge_range('A1:F1', title, header_format)
    sheet_open_source_prev.write_row('A3', headers, header_format)
    sheet_open_source_prev.write_column('A4', np.arange(mtk_type_prev_week.shape[0]) + 1, num_right_format)
    sheet_open_source_prev.write_column('B4', mtk_type_prev_week['weekday'], text_center_format)
    sheet_open_source_prev.write_column('C4', mtk_type_prev_week['year'], num_right_format)
    sheet_open_source_prev.write_column('D4', mtk_type_prev_week['KH tự mở tài khoản'], num_right_format)
    sheet_open_source_prev.write_column('E4', mtk_type_prev_week['KH mở tài khoản qua môi giới'], num_right_format)
    sheet_open_source_prev.write_column('F4', mtk_type_prev_week['Unidentified'], num_right_format)

    new_row = mtk_type_prev_week.shape[0] + 11
    sheet_open_source_prev.write_row(f'A{new_row}', headers, header_format)
    sheet_open_source_prev.write_column(f'A{new_row+1}', np.arange(mtk_type_prev_year.shape[0]) + 1, num_right_format)
    sheet_open_source_prev.write_column(f'B{new_row+1}', mtk_type_prev_year['weekday'], text_center_format)
    sheet_open_source_prev.write_column(f'C{new_row+1}', mtk_type_prev_year['year'], num_right_format)
    sheet_open_source_prev.write_column(f'D{new_row+1}', mtk_type_prev_year['KH tự mở tài khoản'], num_right_format)
    sheet_open_source_prev.write_column(f'E{new_row+1}', mtk_type_prev_year['KH mở tài khoản qua môi giới'], num_right_format)
    sheet_open_source_prev.write_column(f'F{new_row+1}', mtk_type_prev_year['Unidentified'], num_right_format)

    # draw chart
    # chart 2022
    chart = workbook.add_chart({'type': 'column'})
    chart.add_series({
        'name': 'KH tự mở tài khoản',
        'categories': ['mtk-Nguồn-previous', 3, 1, mtk_type_prev_week.shape[0] + 2, 1],
        'values': F'=mtk-Nguồn-previous!$D$4:$D${mtk_type_prev_week.shape[0] + 3}',
    })
    chart.add_series({
        'name': 'KH mở tài khoản qua môi giới',
        'categories': ['mtk-Nguồn-previous', 3, 1, mtk_type_prev_week.shape[0] + 2, 1],
        'values': f'=mtk-Nguồn-previous!$E$4:$E${mtk_type_prev_week.shape[0] + 3}',
    })
    chart.add_series({
        'name': 'Không xác định',
        'categories': ['mtk-Nguồn-previous', 3, 1, mtk_type_prev_week.shape[0] + 2, 1],
        'values': f'=mtk-Nguồn-previous!$F$4:$F${mtk_type_prev_week.shape[0] + 3}',
    })
    chart.set_title({'name': 'Number of accounts by account open source 2022'})
    sheet_open_source_prev.insert_chart('H1', chart, {'x_scale': 1.4, 'y_scale': 0.8})

    # chart 2021
    chart1 = workbook.add_chart({'type': 'column'})
    chart1.add_series({
        'name': 'KH tự mở tài khoản',
        'categories': ['mtk-Nguồn-previous', mtk_type_prev_week.shape[0]+11,1,mtk_type_prev_week.shape[0]+mtk_type_prev_year.shape[0]+10, 1],
        'values': F'=mtk-Nguồn-previous!$D${mtk_type_prev_week.shape[0]+12}:$D${mtk_type_prev_week.shape[0]+mtk_type_prev_year.shape[0]+11}',
    })
    chart1.add_series({
        'name': 'KH mở tài khoản qua môi giới',
        'categories': ['mtk-Nguồn-previous', mtk_type_prev_week.shape[0]+11,1,mtk_type_prev_week.shape[0]+mtk_type_prev_year.shape[0]+10, 1],
        'values': f'=mtk-Nguồn-previous!$E${mtk_type_prev_week.shape[0]+12}:$E${mtk_type_prev_week.shape[0]+mtk_type_prev_year.shape[0]+11}',
    })
    chart1.add_series({
        'name': 'Không xác định',
        'categories': ['mtk-Nguồn-previous', 3, 1, mtk_type_prev_week.shape[0] + 2, 1],
        'values': f'=mtk-Nguồn-previous!$F${mtk_type_prev_week.shape[0]+12}:$F${mtk_type_prev_week.shape[0]+mtk_type_prev_year.shape[0]+11}',
    })
    chart1.set_title({'name': 'Number of accounts by account open source 2021'})
    sheet_open_source_prev.insert_chart('H14', chart1, {'x_scale': 1.4, 'y_scale': 0.8})

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
