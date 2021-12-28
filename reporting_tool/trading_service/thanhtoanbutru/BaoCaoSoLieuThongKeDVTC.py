from reporting_tool.trading_service.thanhtoanbutru import *

# DONE
def run(
        periodicity:str,
        run_time=None,
):
    start = time.time()
    info = get_info(periodicity,run_time)
    period = info['period']
    folder_name = info['folder_name']
    end_date = info['end_date']
    start_date = bdate(info['start_date'],-1)

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
            WHERE [margin_outstanding].[date] = '{start_date}'
                AND [margin_outstanding].[account_code] NOT IN (
                    SELECT [account_code] FROM [account] WHERE [account_type] = N'Tự doanh')
            GROUP BY [margin_outstanding].[type]
        ),
        [closing] AS (
            SELECT 
                [margin_outstanding].[type],
                SUM([margin_outstanding].[principal_outstanding]) [c_outs]
            FROM [margin_outstanding]
            WHERE [margin_outstanding].[date] = '{end_date}'
                AND [margin_outstanding].[account_code] NOT IN (
                    SELECT [account_code] FROM [account] WHERE [account_type] = N'Tự doanh')
            GROUP BY [margin_outstanding].[type]
        )
        SELECT
            [t].[type],
            [t].[o_outs],
            [t].[c_outs] - [t].[o_outs] [d_oust],
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

    def get_margin_accounts(t):
        return pd.read_sql(
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
                    FROM [vcf0051] WHERE [date] <= '{t}'
                    AND [vcf0051].[action] IN ('EDIT','ADD')
                    GROUP BY [sub_account], [date]) [t]
            RIGHT JOIN (SELECT MAX(date) [date], [sub_account]
                    FROM [vcf0051] WHERE [date] <= '{t}' 
                    AND [vcf0051].[action] IN ('EDIT','ADD')
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
    o_margin_table = get_margin_accounts(start_date)
    c_margin_table = get_margin_accounts(end_date)

    o_margin_accounts = o_margin_table.shape[0]
    c_margin_accounts = c_margin_table.shape[0]

    def get_dim_1(x):
        if x == 'o':
            table = o_margin_table
        elif x == 'c':
            table = c_margin_table
        else:
            raise ValueError('Invalid Input')
        count_table = table['dim_1'].value_counts()
        return count_table[['VIPCN T1','SILV T1','GOLD T0','GOLD T2']]

    o_dim_1, c_dim_1 = get_dim_1('o'), get_dim_1('c')
    d_dim_1 = c_dim_1 - o_dim_1

    def get_dim_2(x):
        if x == 'o':
            table = o_margin_table
        elif x == 'c':
            table = c_margin_table
        else:
            raise ValueError('Invalid Input')
        count_table = table['dim_2'].value_counts()
        return count_table[['NORMAL DP+2','VIPCN DP+3','SILV DP+4','GOLD DP+5']]

    o_dim_2, c_dim_2 = get_dim_2('o'), get_dim_2('c')
    d_dim_2 = c_dim_2 - o_dim_2

    def get_dim_3(x):
        if x == 'o':
            table = o_margin_table
        elif x == 'c':
            table = c_margin_table
        else:
            raise ValueError('Invalid Input')
        count_table = table['dim_3'].value_counts()
        return count_table[['VIPCN','SILVER','GOLDEN']]

    o_dim_3, c_dim_3 = get_dim_3('o'), get_dim_3('c')
    d_dim_3 = c_dim_3 - o_dim_3

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
            AND [margin_outstanding].[account_code] NOT IN (
                SELECT [account_code] FROM [account] WHERE [account_type] = N'Tự doanh')
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
        index_col='type',
    ).squeeze(axis=1).reindex(['UTTB','DP','MR','BL','Cầm cố']).fillna(0)
    margin_accounts_inb01 = pd.read_sql(
        f"""
        SELECT 
            COUNT(DISTINCT [vcf0051].[sub_account]) [count]
        FROM 
            [vcf0051]
        RIGHT JOIN (SELECT [t].[time], [t].[date], [t].[sub_account]
            FROM (SELECT MAX(time) [time], [date], [sub_account] 
                    FROM [vcf0051] WHERE [date] <= '{end_date}'
                    AND [vcf0051].[action] IN ('EDIT','ADD')
                    GROUP BY [sub_account], [date]) [t]
            RIGHT JOIN (SELECT MAX(date) [date], [sub_account]
                    FROM [vcf0051] WHERE [date] <= '{end_date}' 
                    AND [vcf0051].[action] IN ('EDIT','ADD')
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

    path = r"\\192.168.10.101\phs-storge-2018\RiskManagementDept"\
           r"\RMD_Data\Luu tru van ban\Monthly Report"\
           r"\2. Monthly Strategy Meeting Report"
    bad_loan_path = join(
        path,
        end_date[:4],
        f'RMD_Monthly Summary Risk Report for Margin Operations_{month_mapper(end_date[5:7])} {end_date[:4]}.xlsx',
    )
    bad_loan = pd.read_excel(bad_loan_path,sheet_name='Bad loan',usecols=[0,1,4])
    bad_loan.columns = ['account_code','principal','group']
    last_row = (~bad_loan['account_code'].isna()).idxmin()
    bad_loan = bad_loan.iloc[:last_row]
    bad_loan.dropna(subset=['principal'],inplace=True)
    bad_loan = bad_loan.loc[bad_loan['principal']!=0]
    bad_loan['group'] = bad_loan['group'].str.replace(' ','').str.replace("case",'')
    bad_loan['group'].fillna(method='ffill',inplace=True)

    debt_type = pd.read_sql(
        f"""
        SELECT DISTINCT
        	[margin_outstanding].[account_code],
            CASE [margin_outstanding].[type]
                WHEN N'Ứng trước cổ tức' THEN 'UTCT'
                WHEN N'Trả chậm' THEN 'DP'
                WHEN N'Margin' THEN 'MR'
                WHEN N'Bảo lãnh' THEN 'BL'
            ELSE [margin_outstanding].[type]
            END [type]
        FROM [margin_outstanding]
        WHERE [margin_outstanding].[account_code] IN {iterable_to_sqlstring(bad_loan['account_code'])}
        """,
        connect_DWH_CoSo,
        index_col='account_code'
    ).squeeze()

    bad_loan['type'] = bad_loan['account_code'].map(debt_type)
    name = pd.read_sql(
        f"""
        SELECT [account].[account_code], [account].[customer_name] 
        FROM [account]
        WHERE [account].[account_code] IN {iterable_to_sqlstring(bad_loan['account_code'])}
        """,
        connect_DWH_CoSo,
        index_col='account_code',
    ).squeeze()
    bad_loan['cutomer_name'] = bad_loan['account_code'].map(name)
    bad_loan.sort_values('type',inplace=True)

    bad_loan_g = bad_loan.groupby('type')['principal'].sum()
    outstandings = outstandings.join(bad_loan_g,'type','outer').fillna(0)
    outstandings.rename({'principal':'bad_loan'},axis=1,inplace=True)

    # Lấy Dư nợ gửi SSC
    def get_ssc_outs(t):
        folder_path \
            = r"\\192.168.10.101\phs-storge-2018\RiskManagementDept\RMD_Data"\
              r"\Luu tru van ban\SSC Report\Daily SSC before 8AM"
        if t=='c':
            d = end_date
        elif t=='o':
            d = start_date
        else:
            raise ValueError('Invalid Input')
        path = join(
            folder_path,
            f'{d[:4]}',
            f'{d[5:7]}.{d[:4]}',
            f'180426_SCMS_Bao cao ngay truoc 8AM {d[-2:]}{d[5:7]}{d[:4]}.xlsx',
        )
        ssc_outs = pd.read_excel(
            path,
            sheet_name='1.f_thgdkq_06692',
            skiprows=2,
            skipfooter=1,
            usecols=[2],
        ).sum().squeeze() * 1e6
        return ssc_outs

    ssc_o_outs, ssc_c_outs = get_ssc_outs('o'), get_ssc_outs('c')

    ###################################################
    ###################################################
    ###################################################
    
    file_name = f'Số liệu thống kê DVTC tháng {period}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder, folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    title_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 14,
            'font_name': 'Times New Roman',
            'bg_color': '#4BABC6',
        }
    )
    percent_format = workbook.add_format(
        {
            'border': 1,
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'num_format': '0.00%'
        }
    )
    header_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman',
        }
    )
    bad_loan_header_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'bg_color': '#FFBF00'
        }
    )
    num_format = workbook.add_format(
        {
            'border': 1,
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'
        }
    )
    num_bold_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'
        }
    )
    text_left_format = workbook.add_format(
        {
            'border': 1,
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    text_center_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman',
        }
    )
    text_noborder_format = workbook.add_format(
        {
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman',
        }
    )
    worksheet = workbook.add_worksheet('DVTC')
    worksheet.hide_gridlines(option=2)
    headers = [
        '',
        'Loại',
        'Đầu kỳ',
        'Trong kỳ',
        'Cuối kỳ',
        'Nợ xấu',
        'Ghi chú',
        'INB-01',
        'Tỷ lệ',
    ]
    worksheet.set_column('A:A',26)
    worksheet.set_column('B:B',16)
    worksheet.set_column('C:F',20)
    worksheet.set_column('G:G',27)
    worksheet.set_column('H:H',16)
    worksheet.set_column('I:I',10)

    worksheet.merge_range('A1:I1','TOÀN CÔNG TY',title_format)
    worksheet.write_row('A2',headers,header_format)
    worksheet.write_column('B3',outstandings['type'],text_center_format)
    worksheet.write_column('C3',outstandings['o_outs'],num_format)
    worksheet.write_column('D3',outstandings['d_oust'],num_format)
    worksheet.write_column('E3',outstandings['c_outs'],num_format)
    worksheet.write_column('F3',outstandings['bad_loan'],num_format)
    worksheet.write('F9',0,num_format)

    worksheet.merge_range('A3:A7','Tổng Dư nợ gốc',text_left_format)
    worksheet.write('A8','',text_left_format)
    worksheet.write('B8','TỔNG',header_format)
    worksheet.write_row('C8',outstandings.iloc[:,1:].sum(),num_bold_format)
    worksheet.write('A9','Dư nợ báo cáo SSC',text_left_format)
    worksheet.write_row('C9',[ssc_o_outs,ssc_c_outs-ssc_o_outs,ssc_c_outs],num_format)

    worksheet.write('A10','Tổng số lượng TK Margin\n(Active & Block)',text_left_format)
    worksheet.write('B10','',text_left_format)
    worksheet.write_row(
        'C10',
        [o_margin_accounts,c_margin_accounts-o_margin_accounts,c_margin_accounts],
        num_format,
    )
    worksheet.merge_range(
        'A11:A14',
        'Tổng số lượng TK VIP miễn lãi vay MR\n(Active &Block)',
        text_left_format
    )
    worksheet.write_column('B11',o_dim_1.index,text_center_format)
    worksheet.write_column('C11',o_dim_1,num_format)
    worksheet.write_column('D11',d_dim_1,num_format)
    worksheet.write_column('E11',c_dim_1,num_format)

    worksheet.merge_range(
        'A15:A18',
        'Tổng số lượng TK sử dụng hạn mức DP\n(Active &Block)',
        text_left_format
    )
    worksheet.write_column('B15',o_dim_2.index,text_center_format)
    worksheet.write_column('C15',o_dim_2,num_format)
    worksheet.write_column('D15',d_dim_2,num_format)
    worksheet.write_column('E15',c_dim_2,num_format)

    worksheet.write_column(
        'A19',
        ['Tổng số lượng TK VIPCN','Tổng số lượng TK SILVER','Tổng số lượng TK GOLDEN'],
        text_left_format,
    )
    worksheet.write_column('B19',['']*3,text_center_format)
    worksheet.write_column('C19',o_dim_3,num_format)
    worksheet.write_column('D19',d_dim_3,num_format)
    worksheet.write_column('E19',c_dim_3,num_format)

    worksheet.write_column('F10',['']*12,num_format)
    for col in ['G','H','I']:
        worksheet.write_column(f'{col}3',['']*19,num_format)

    worksheet.write_column('H3',outstandings_inb01,num_format)
    worksheet.write('H8',outstandings_inb01.sum(),num_bold_format)
    worksheet.write('H10',margin_accounts_inb01,num_format)
    for row in range(3,9):
        worksheet.write(f'I{row}',f'=H{row}/E{row}',percent_format)
    worksheet.write('I10','=H10/E10',percent_format)

    worksheet.merge_range('A23:E23','DANH SÁCH KH NỢ XẤU',bad_loan_header_format)
    headers = [
        'STT',
        'Số TKCK',
        'Dư nợ gốc',
        'Note',
        'Loại nợ',
    ]
    worksheet.write_row('A24',headers,header_format)
    worksheet.write_column('A25',np.arange(bad_loan.shape[0])+1,text_center_format)
    worksheet.write_column('B25',bad_loan['account_code'],text_center_format)
    worksheet.write_column('C25',bad_loan['principal'],num_format)
    worksheet.write_column('D25',bad_loan['group'],text_center_format)
    worksheet.write_column('E25',bad_loan['type'],text_center_format)
    worksheet.write_column('F25',bad_loan['cutomer_name'],text_noborder_format)

    last_row = 25 + bad_loan.shape[0]
    worksheet.merge_range(f'A{last_row}:B{last_row}','Tổng',header_format)
    worksheet.write(f'C{last_row}',bad_loan['principal'].sum(),num_bold_format)
    worksheet.write_row(f'D{last_row}',['']*2,header_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py', '')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time() - start, 1)}s')