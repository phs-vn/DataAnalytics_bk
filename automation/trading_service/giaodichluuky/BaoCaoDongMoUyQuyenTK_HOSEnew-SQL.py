from automation.trading_service.giaodichluuky import *


def run(
    run_time=None
):

    start = time.time()
    info = get_info('monthly',run_time)
    start_date = info['start_date']
    end_date = info['end_date']
    period = info['period']
    folder_name = info['folder_name']

    month = int(period.split('.')[0])
    year = int(period.split('.')[1])

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):  # dept_folder from import
        os.mkdir(join(dept_folder,folder_name,period))

    summary = pd.read_sql(
        f"""
            WITH
            [FlagTable] AS (
            SELECT 
                [ao].[account_type] [Type],
                CASE 
                    WHEN [ao].[account_type] = N'Cá nhân trong nước' THEN 2
                    WHEN [ao].[account_type] = N'Tổ chức trong nước' THEN 3
                    WHEN [ao].[account_type] = N'Cá nhân nước ngoài' THEN 5
                    WHEN [ao].[account_type] = N'Tổ chức nước ngoài' THEN 6
                END [Number],
                CASE WHEN [ao].[date_of_open] BETWEEN '{start_date}' AND '{end_date}' THEN 1 ELSE 0 END [Open],
                CASE WHEN [ac].[date_of_close] BETWEEN '{start_date}' AND '{end_date}' THEN 1 ELSE 0 END [Close]
            FROM [account] [ao] FULL JOIN [account] [ac] ON [ao].[account_code] = [ac].[account_code]
            WHERE [ao].[account_type] IN (N'Cá nhân trong nước',N'Tổ chức trong nước',N'Cá nhân nước ngoài',N'Tổ chức nước ngoài')
            ),
            [AggregateTable] AS (
            SELECT 
                [FlagTable].[Number],
                [FlagTable].[Type],
                SUM([FlagTable].[Open]) [Open],
                SUM([FlagTable].[Close]) [Close]
            FROM [FlagTable]
            GROUP BY [FlagTable].[Type], [FlagTable].[Number]
            ),
            [Change] AS (
                SELECT 
                    1 [Number], 
                    N'Trong nước' [Type], 
                    SUM([AggregateTable].[Open]) [Open], 
                    SUM([AggregateTable].[Close]) [Close] 
                FROM [AggregateTable] WHERE [AggregateTable].[Type] IN (N'Cá nhân trong nước',N'Tổ chức trong nước')
                UNION 
                SELECT *
                FROM [AggregateTable]
                UNION
                SELECT 
                    4 [Number], 
                    N'Nước ngoài' [Type], 
                    SUM([AggregateTable].[Open]) [Open], 
                    SUM([AggregateTable].[Close]) [Close]
                FROM [AggregateTable] WHERE [AggregateTable].[Type] IN (N'Cá nhân nước ngoài',N'Tổ chức nước ngoài')
                UNION 
                SELECT 
                    7 [Number], 
                    N'Tổng Cộng' [Type], 
                    SUM([AggregateTable].[Open]) [Open], 
                    SUM([AggregateTable].[Close]) [Close]
                FROM [AggregateTable]
            ),
            [tempEnd] AS (
                SELECT 
                    COUNT([account].[account_code]) [count], 
                    [account].[account_type]
                FROM [account]
                WHERE 
                    [account].[date_of_open] <= '{end_date}'
                    AND ([account].[date_of_close] IS NULL 
                        OR ([account].[date_of_close] > '{end_date}' AND [account].[date_of_close] != '2099-12-31')
                    ) -- mot so tai khoan dong rat lau roi duoc gan ngay dong la ngay nay
                    AND [account].[account_type] IN (N'Cá nhân trong nước',N'Tổ chức trong nước',N'Cá nhân nước ngoài',N'Tổ chức nước ngoài')
                GROUP BY [account].[account_type]
            ),
            [End] AS (
                SELECT * FROM [tempEnd]
                UNION SELECT SUM([tempEnd].[count]) [count], N'Trong nước' FROM [tempEnd] WHERE [tempEnd].[account_type] IN (N'Cá nhân trong nước',N'Tổ chức trong nước')
                UNION SELECT SUM([tempEnd].[count]) [count], N'Nước ngoài' FROM [tempEnd] WHERE [tempEnd].[account_type] IN (N'Cá nhân nước ngoài',N'Tổ chức nước ngoài')
                UNION SELECT SUM([tempEnd].[count]) [count], N'Tổng Cộng' FROM [tempEnd]
            )
            SELECT 
                COALESCE([Change].[Type],[End].[account_type]) [Type],
                [End].[count] + [Change].[Close] - [Change].[Open] [Start],
                [Change].[Open],
                [Change].[Close],
                [End].[count] [End]
            FROM [Change] 
            FULL JOIN [End] ON [End].[account_type] = [Change].[Type]
            ORDER BY [Change].[Number]
        """
        ,
        connect_DWH_CoSo
    )

    # country code
    country_code = pd.read_excel(join(dirname(dirname(__file__)),'country_code.xlsx'),index_col=0,squeeze=True)
    # country_code = pd.read_excel(
    #     r'D:\DataAnalytics\automation\trading_service\country_code.xlsx', index_col=0, squeeze=True)
    split_at = int(country_code.shape[0]/2)+1
    country_code_head = country_code.head(split_at)
    country_code_tail = country_code.tail(country_code.shape[0]-split_at)

    country_mapper = country_code.str.replace(' ','').str[2:].reset_index().set_index('QUỐC GIA').squeeze()

    account_open = pd.read_sql(
        f"""
            WITH
            [m] AS (
                SELECT DISTINCT
                    [sub_account].[account_code]
                FROM [vcf0051] 
                LEFT JOIN [sub_account] ON [vcf0051].[sub_account] = [sub_account].[sub_account]
                WHERE [vcf0051].[contract_type] NOT LIKE N'%Thường%' AND [vcf0051].[date] = '{end_date}'
            )
            SELECT
                ROW_NUMBER() OVER (ORDER BY [a].[account_type], [a].[account_code], [a].[date_of_open]) [no.],
                [a].[account_type],
                [a].[account_code],
                [a].[customer_name],
                REPLACE([a].[nationality], ' ', '') [nationality],
                [a].[address],
                [a].[customer_id_number],
                [a].[date_of_issue],
                [a].[place_of_issue],
                [a].[date_of_open],
                [a].[date_of_close],
                CASE
                    WHEN [m].[account_code] IS NULL THEN '' ELSE 'TKKQ' END [remark],
                CASE 
                    WHEN [a].[account_type] LIKE N'%Cá nhân%' THEN 'CN'
                    WHEN [a].[account_type] LIKE N'%Tổ chức%' THEN 'TC'
                END [entity_type]
            FROM [account] [a]
            LEFT JOIN [m] ON [m].[account_code] = [a].[account_code]
            WHERE [a].[date_of_open] BETWEEN '{start_date}' AND '{end_date}'
            ORDER BY [a].[account_type], [a].[account_code], [a].[date_of_open]
        """,
        connect_DWH_CoSo
    )
    account_close = pd.read_sql(
        f"""
            WITH
            [m] AS (
                SELECT DISTINCT
                    [sub_account].[account_code]
                FROM [vcf0051] 
                LEFT JOIN [sub_account] ON [vcf0051].[sub_account] = [sub_account].[sub_account]
                WHERE [vcf0051].[contract_type] NOT LIKE N'%Thường%' AND [vcf0051].[date] = '{end_date}'
            )
            SELECT
                ROW_NUMBER() OVER (ORDER BY [a].[account_type], [a].[account_code], [a].[date_of_close]) [no.],
                [a].[account_type],
                [a].[account_code],
                [a].[customer_name],
                REPLACE([a].[nationality], ' ', '') [nationality],
                [a].[address],
                [a].[customer_id_number],
                [a].[date_of_issue],
                [a].[place_of_issue],
                [a].[date_of_open],
                [a].[date_of_close],
                CASE
                    WHEN [m].[account_code] IS NULL THEN '' ELSE 'TKKQ' END [remark],
                CASE 
                    WHEN [a].[account_type] LIKE N'%Cá nhân%' THEN 'CN'
                    WHEN [a].[account_type] LIKE N'%Tổ chức%' THEN 'TC'
                END [entity_type]
            FROM [account] [a]
            LEFT JOIN [m] ON [m].[account_code] = [a].[account_code]
            WHERE [a].[date_of_close] BETWEEN '{start_date}' AND '{end_date}'
            ORDER BY [a].[account_type], [a].[account_code], [a].[date_of_close]
        """,
        connect_DWH_CoSo
    )
    account_open['nationality'] = account_open['nationality'].map(country_mapper).fillna('')
    account_close['nationality'] = account_close['nationality'].map(country_mapper).fillna('')

    customer_information_change = pd.read_sql(
        f"""
            SELECT 
                CONCAT('(',(ROW_NUMBER() OVER (ORDER BY [z].[account_code], [z].[date_of_change])),')') [no.],
                *
            FROM (
                SELECT
                DISTINCT
                    ISNULL([t].[account_code],'') [account_code],
                    ISNULL([a].[customer_name],'') [customer_name],
                    ISNULL([t].[date_of_change],'') [date_of_change],
                    ISNULL([t].[old_id_number],'') [old_id_number],
                    ISNULL([t].[new_id_number],'') [new_id_number],
                    ISNULL([t].[old_date_of_issue],'') [old_date_of_issue],
                    ISNULL([t].[new_date_of_issue],'') [new_date_of_issue],
                    ISNULL([t].[old_place_of_issue],'') [old_place_of_issue],
                    ISNULL([t].[new_place_of_issue],'') [new_place_of_issue],
                    ISNULL([t].[old_address],'') [old_address],
                    ISNULL([t].[new_address],'') [new_address],
                    ISNULL(REPLACE([t].[old_nationality], ' ', ''),'') [old_nationality],
                    ISNULL(REPLACE([t].[new_nationality], ' ', ''),'') [new_nationality],
                    '' [old_note],
                    '' [new_note]
                FROM [rcf0005] [t]
                LEFT JOIN [account] [a] ON [a].[account_code] = [t].[account_code]
                WHERE [t].[date_of_change] BETWEEN '{start_date}' AND '{end_date}'
            ) [z]
        """,
        connect_DWH_CoSo,
    )
    customer_information_change['old_nationality']=customer_information_change['old_nationality'].map(country_mapper).fillna('')
    customer_information_change['new_nationality']=customer_information_change['new_nationality'].map(country_mapper).fillna('')
    authorization = pd.read_sql(
        f"""
            SELECT
                ROW_NUMBER() OVER (ORDER BY [authorization].[account_code]) [no.],
                ISNULL([authorization].[account_code],'') [account_code],
                ISNULL([authorization].[authorizing_person_id],'') [authorizing_person_id],
                ISNULL([authorization].[authorizing_person_name],'') [authorizing_person_name],
                ISNULL([authorization].[authorizing_person_address],'') [authorizing_person_address],
                ISNULL([authorization].[authorized_person_id],'') [authorized_person_id],
                ISNULL([authorization].[authorized_person_name],'') [authorized_person_name],
                CASE
                    WHEN [authorization].[authorized_person_name] = N'CTY CP CHỨNG KHOÁN PHÚ HƯNG'
                        THEN N'{CompanyAddress}'
                    ELSE [authorization].[authorized_person_address]
                END [authorized_person_address],
                [authorization].[date_of_authorization],
                'I,II,IV,V,VII,IX,X' [scope_of_authorization]
            FROM [authorization]  
            WHERE [authorization].[date_of_authorization] BETWEEN '{start_date}' AND '{end_date}'
                AND [authorization].[scope_of_authorization] IS NOT NULL
                AND [authorization].[scope_of_authorization] <> 'I,IV,V'
        """,
        connect_DWH_CoSo
    )
    # Highlight cac uy quyen duoc mo moi chi de dang ky uy quyen them (rule ben DVKH)
    highlight_account = pd.read_sql(
        f"""
            SELECT
                [authorization_change].[account_code]
            FROM 
                [authorization_change]
            WHERE 
                [authorization_change].[new_end_date] BETWEEN '{start_date}' AND '{end_date}'
        """,
        connect_DWH_CoSo
    )
    authorization_change = pd.read_sql(
        f"""
            SELECT 
                ROW_NUMBER() OVER (ORDER BY [account_code],[date_of_change]) [no.],
                ISNULL([c].[account_code],'') [account_code],
                ISNULL([c].[authorizing_person_id],'') [authorizing_person_id],
                ISNULL([c].[authorizing_person_name],'') [authorizing_person_name],
                ISNULL(CONVERT(VARCHAR(20),[c].[date_of_authorization],103),'') [date_of_authorization],
                ISNULL(CONVERT(VARCHAR(20),[c].[date_of_termination],103),'') [date_of_termination],
                [c].[date_of_change],
                ISNULL([c].[authorized_person_name],'') [authorized_person_name],
                ISNULL([c].[old_authorized_person_id],'') [old_authorized_person_id],
                ISNULL([c].[new_authorized_person_id],'') [new_authorized_person_id],
                ISNULL([c].[old_authorized_person_address],'') [old_authorized_person_address],
                ISNULL([c].[new_authorized_person_address],'') [new_authorized_person_address],
                ISNULL([c].[old_scope_of_authorization],'') [old_scope_of_authorization],
                ISNULL([c].[new_scope_of_authorization],'') [new_scope_of_authorization],
                ISNULL(CONVERT(VARCHAR(20),[c].[old_end_date],103),'') [old_end_date],
                ISNULL(CONVERT(VARCHAR(20),[c].[new_end_date],103),'') [new_end_date]
            FROM [authorization_change] [c]
            WHERE [c].[date_of_change] BETWEEN '{start_date}' AND '{end_date}'
        """,
        connect_DWH_CoSo
    )

    ###########################################################################
    ###########################################################################
    ###########################################################################

    file_name = f'Danh sách KH đóng mở ủy quyền tài khoản PHS (new) HOSE {period}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    ###########################################################################
    ###########################################################################
    ###########################################################################

    ## Write sheet TONG HOP
    company_name_format = workbook.add_format(
        {
            'bold':True,
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':11,
        }
    )
    company_info_format = workbook.add_format(
        {
            'italic':True,
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':11,
        }
    )
    sup_title_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':12
        }
    )
    sup_note_format = workbook.add_format(
        {
            'italic':True,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':11
        }
    )
    kinhgui_format = workbook.add_format(
        {
            'bold':False,
            'italic':True,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':12
        }
    )
    header_format = workbook.add_format(
        {
            'bold':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':12
        }
    )
    stt_heading_format = workbook.add_format(
        {
            'bold':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':12
        }
    )
    stt_column_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':12
        }
    )
    header_cell_format = workbook.add_format(
        {
            'bold':True,
            'valign':'vcenter',
            'border':1,
            'font_name':'Times New Roman',
            'font_size':12
        }
    )
    normal_cell_format = workbook.add_format(
        {
            'border':1,
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':12
        }
    )
    header_value = workbook.add_format(
        {
            'border':1,
            'valign':'vcenter',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'bold':True,
            'font_name':'Times New Roman',
            'font_size':12
        }
    )
    normal_value = workbook.add_format(
        {
            'border':1,
            'valign':'vcenter',
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'font_name':'Times New Roman',
            'font_size':12
        }
    )
    sheet_tonghop = workbook.add_worksheet('Tổng hợp')
    sheet_tonghop.hide_gridlines(option=2)

    sheet_tonghop.set_column('A:A',8)
    sheet_tonghop.set_column('B:B',26)
    sheet_tonghop.set_column('C:F',22)
    sheet_tonghop.set_row(0,15)
    sheet_tonghop.set_row(1,15)
    sheet_tonghop.set_row(2,15)
    sheet_tonghop.set_row(3,15)

    sheet_tonghop.merge_range('A1:F1',CompanyName,company_name_format)
    sheet_tonghop.merge_range('A2:F2',CompanyAddress,company_info_format)
    sheet_tonghop.merge_range('A3:F3',CompanyPhoneNumber,company_info_format)
    sheet_tonghop.merge_range('A4:F4',CompanyEmail,company_info_format)
    sup_title = r'BÁO CÁO TÌNH HÌNH ĐÓNG/MỞ TÀI KHOẢN'
    sheet_tonghop.merge_range('A6:F6',sup_title,sup_title_format)
    note_title = f'(Tháng {convert_int(month)} năm {year})'
    sheet_tonghop.merge_range('A7:F7',note_title,sup_note_format)
    kinhgui_content = 'Kính gửi: Sở Giao dịch Chứng khoán Thành phố Hồ Chí Minh'
    sheet_tonghop.merge_range('A8:F8',kinhgui_content,kinhgui_format)
    sheet_tonghop.merge_range('A10:A11','STT',header_format)
    sheet_tonghop.merge_range('B10:B11','KHÁCH HÀNG',header_format)
    sheet_tonghop.merge_range('C10:F10','SỐ LƯỢNG TÀI KHOẢN',header_format)
    sheet_tonghop.write_row('C11',['Đầu kỳ','Mở trong kỳ','Đóng trong kỳ','Cuối kỳ'],header_format)
    sheet_tonghop.write_column('A12',['','1','2','','1','2',''],stt_column_format)
    sheet_tonghop.write('A12','I',stt_heading_format)
    sheet_tonghop.write('A15','II',stt_heading_format)
    sheet_tonghop.write('B12','Trong nước',header_cell_format)
    sheet_tonghop.write('B15','Nước ngoài',header_cell_format)
    sheet_tonghop.write_column('B13',['     Cá nhân','     Tổ chức'],normal_cell_format)
    sheet_tonghop.write_column('B16',['     Cá nhân','     Tổ chức'],normal_cell_format)
    sheet_tonghop.write('B18','Tổng cộng',header_cell_format)

    cols = ['Start', 'Open', 'Close', 'End']
    sheet_tonghop.write_row('C12', summary.loc[summary['Type'] == 'Trong nước', cols].squeeze(), header_value)
    sheet_tonghop.write_row('C13', summary.loc[summary['Type'] == 'Cá nhân trong nước', cols].squeeze(), normal_value)
    sheet_tonghop.write_row('C14', summary.loc[summary['Type'] == 'Tổ chức trong nước', cols].squeeze(), normal_value)
    sheet_tonghop.write_row('C15', summary.loc[summary['Type'] == 'Nước ngoài', cols].squeeze(), header_value)
    sheet_tonghop.write_row('C16', summary.loc[summary['Type'] == 'Cá nhân nước ngoài', cols].squeeze(), normal_value)
    sheet_tonghop.write_row('C17', summary.loc[summary['Type'] == 'Tổ chức nước ngoài', cols].squeeze(), normal_value)
    sheet_tonghop.write_row('C18', summary.loc[summary['Type'] == 'Tổng Cộng', cols].squeeze(), header_value)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    ###########################################################################
    ###########################################################################
    ###########################################################################

    ## Write sheet MO TAi KHOAN
    sup_title_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':18
        }
    )
    header_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    text_left_format = workbook.add_format(
        {
            'border':1,
            'align':'left',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    text_center_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    date_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'num_format':'dd\/mm\/yyyy',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    sheet_motaikhoan = workbook.add_worksheet('Mở TK')
    sheet_motaikhoan.hide_gridlines(option=2)
    # set column width
    sheet_motaikhoan.set_column('A:A',4)
    sheet_motaikhoan.set_column('B:B',25)
    sheet_motaikhoan.set_column('C:D',12)
    sheet_motaikhoan.set_column('E:E',70)
    sheet_motaikhoan.set_column('F:F',12)
    sheet_motaikhoan.set_column('G:G',30)
    sheet_motaikhoan.set_column('H:J',12)
    sheet_motaikhoan.set_column('K:M',10)
    sheet_motaikhoan.set_row(0,36)
    sheet_motaikhoan.merge_range('A1:M1','Danh sách khách hàng mở tài khoản',sup_title_format)
    headers = [
        'STT',
        'Tên khách hàng',
        'Mã TK',
        'Số CMND/ Hộ chiếu/Giấy ĐKKD',
        'Địa chỉ',
        'Ngày cấp',
        'Nơi cấp',
        'Loại hình',
        'Ngày mở',
        'Quốc tịch',
        'Cấp lại mới',
        'Chức vụ',
        'Ghi chú',
    ]
    sheet_motaikhoan.write_row('A2',headers,header_format)
    header_num = [f'({i})' for i in np.arange(len(headers))+1]
    sheet_motaikhoan.write_row('A3',header_num,header_format)

    sheet_motaikhoan.write_column('A4',account_open['no.'],text_center_format)
    sheet_motaikhoan.write_column('B4',account_open['customer_name'],text_left_format)
    sheet_motaikhoan.write_column('C4',account_open['account_code'],text_center_format)
    sheet_motaikhoan.write_column('D4',account_open['customer_id_number'],text_center_format)
    sheet_motaikhoan.write_column('E4',account_open['address'],text_left_format)
    sheet_motaikhoan.write_column('F4',account_open['date_of_issue'].map(convertNaTtoSpaceString),date_format)
    sheet_motaikhoan.write_column('G4',account_open['place_of_issue'].map(convertNaTtoSpaceString),text_left_format)
    sheet_motaikhoan.write_column('H4',account_open['entity_type'],text_center_format)
    sheet_motaikhoan.write_column('I4',account_open['date_of_open'],date_format)
    sheet_motaikhoan.write_column('J4',account_open['nationality'],text_center_format)
    sheet_motaikhoan.write_column('K4',['']*account_open.shape[0],text_center_format)
    sheet_motaikhoan.write_column('L4',['']*account_open.shape[0],text_center_format)
    sheet_motaikhoan.write_column('M4',account_open['remark'],text_center_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################
    # Write sheet DONG TAi KHOAN

    sup_title_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':18
        }
    )
    header_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    text_left_format = workbook.add_format(
        {
            'border':1,
            'align':'left',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    text_center_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    date_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'num_format':'dd\/mm\/yyyy',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    sheet_dongtaikhoan = workbook.add_worksheet('Đóng TK')
    sheet_dongtaikhoan.hide_gridlines(option=2)

    sheet_dongtaikhoan.set_column('A:A',4)
    sheet_dongtaikhoan.set_column('B:B',25)
    sheet_dongtaikhoan.set_column('C:D',12)
    sheet_dongtaikhoan.set_column('E:E',70)
    sheet_dongtaikhoan.set_column('F:F',12)
    sheet_dongtaikhoan.set_column('G:G',30)
    sheet_dongtaikhoan.set_column('H:L',12)
    sheet_dongtaikhoan.set_default_row(27)  # set all row height = 27
    sheet_dongtaikhoan.set_row(0,36)
    sheet_dongtaikhoan.set_row(1,42)
    sheet_dongtaikhoan.set_row(2,15)
    sheet_dongtaikhoan.merge_range('A1:L1','Danh sách khách hàng đóng tài khoản',sup_title_format)
    headers = [
        'STT',
        'Tên khách hàng',
        'Mã TK',
        'Số CMND/ Hộ chiếu/Giấy ĐKKD',
        'Địa chỉ',
        'Ngày cấp',
        'Nơi cấp',
        'Loại hình',
        'Ngày mở',
        'Ngày đóng',
        'Quốc tịch',
        'Ghi chú',
    ]
    sheet_dongtaikhoan.write_row('A2',headers,header_format)
    header_num = [f'({i})' for i in np.arange(len(headers))+1]
    sheet_dongtaikhoan.write_row('A3',header_num,header_format)

    sheet_dongtaikhoan.write_column('A4', account_close['no.'], text_center_format)
    sheet_dongtaikhoan.write_column('B4',account_close['customer_name'],text_left_format)
    sheet_dongtaikhoan.write_column('C4', account_close['account_code'], text_center_format)
    sheet_dongtaikhoan.write_column('D4',account_close['customer_id_number'],text_center_format)
    sheet_dongtaikhoan.write_column('E4',account_close['address'],text_left_format)
    sheet_dongtaikhoan.write_column('F4',account_close['date_of_issue'].map(convertNaTtoSpaceString),date_format)
    sheet_dongtaikhoan.write_column('G4',account_close['place_of_issue'],text_center_format)
    sheet_dongtaikhoan.write_column('H4',account_close['entity_type'],text_center_format)
    sheet_dongtaikhoan.write_column('I4',account_close['date_of_open'].map(convertNaTtoSpaceString),date_format)
    sheet_dongtaikhoan.write_column('J4',account_close['date_of_close'].map(convertNaTtoSpaceString),date_format)
    sheet_dongtaikhoan.write_column('K4',account_close['nationality'],text_center_format)
    sheet_dongtaikhoan.write_column('L4',account_close.shape[0]*[''],text_center_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    # Write sheet THAY DOI THONG TIN
    sup_title_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':18
        }
    )
    header_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    text_left_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'left',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    text_center_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    date_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'num_format':'dd\/mm\/yyyy',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    sheet_thaydoithongtin = workbook.add_worksheet('Thay đổi thông tin')
    sheet_thaydoithongtin.hide_gridlines(option=2)

    sheet_thaydoithongtin.set_column('A:A',4)
    sheet_thaydoithongtin.set_column('B:B',25)
    sheet_thaydoithongtin.set_column('C:F',13)
    sheet_thaydoithongtin.set_column('G:G',24)
    sheet_thaydoithongtin.set_column('H:I',13)
    sheet_thaydoithongtin.set_column('J:J',24)
    sheet_thaydoithongtin.set_column('K:L',26)
    sheet_thaydoithongtin.set_column('M:P',8)
    sheet_thaydoithongtin.set_row(0,30)
    sheet_thaydoithongtin.set_row(1,28)
    sheet_thaydoithongtin.set_row(2,39)

    sheet_thaydoithongtin.merge_range('A1:P1','Danh sách khách hàng thay đổi thông tin',sup_title_format)
    sheet_thaydoithongtin.merge_range('A2:A3','STT',header_format)
    sheet_thaydoithongtin.merge_range('B2:B3','Tên khách hàng',header_format)
    sheet_thaydoithongtin.merge_range('C2:C3','Mã TK cũ',header_format)
    sheet_thaydoithongtin.merge_range('D2:D3','Ngày thay đổi thông tin',header_format)
    sheet_thaydoithongtin.merge_range('E2:J2','Thay đổi thông tin về CMND/ Hộ chiếu/ Giấy ĐKKD',header_format)
    sheet_thaydoithongtin.merge_range('K2:L2','Thay đổi thông tin về địa chỉ',header_format)
    sheet_thaydoithongtin.merge_range('M2:N2','Thay đổi TT về Q.tịch',header_format)
    sheet_thaydoithongtin.merge_range('O2:P2','Thay đổi thông tin về Ghi chú',header_format)
    sub_header = [
        'Số CMND/ Hộ chiếu/ Giấy ĐKKD cũ',
        'Ngày cấp',
        'Nơi cấp',
        'Số CMND/ Hộ chiếu/ Giấy ĐKKD mới',
        'Ngày cấp',
        'Nơi cấp',
        'Địa chỉ cũ',
        'Địa chỉ mới',
        'Quốc tịch cũ',
        'Quốc tịch mới',
        'Ghi chú cũ',
        'Ghi chú mới',
    ]
    sheet_thaydoithongtin.write_row('E3',sub_header,header_format)
    sheet_thaydoithongtin.write_row(
        'A4',
        [f'({i})' for i in np.arange(16)+1],  # cong them 2 cot ghi chu
        header_format,
    )
    sheet_thaydoithongtin.write_column('A5', customer_information_change['no.'], text_center_format)
    sheet_thaydoithongtin.write_column(
        'B5',
        customer_information_change['customer_name'],
        text_left_format,
    )
    sheet_thaydoithongtin.write_column('C5', customer_information_change['account_code'], text_center_format)
    sheet_thaydoithongtin.write_column('D5', customer_information_change['date_of_change'], date_format)
    sheet_thaydoithongtin.write_column('E5', customer_information_change['old_id_number'], text_center_format)
    sheet_thaydoithongtin.write_column('F5', customer_information_change['old_date_of_issue'], date_format)
    sheet_thaydoithongtin.write_column('G5', customer_information_change['old_place_of_issue'], text_center_format)
    sheet_thaydoithongtin.write_column('H5', customer_information_change['new_id_number'], text_center_format)
    sheet_thaydoithongtin.write_column('I5', customer_information_change['new_date_of_issue'], date_format)
    sheet_thaydoithongtin.write_column(
        'J5',
        customer_information_change['new_place_of_issue'],
        text_center_format,
    )
    sheet_thaydoithongtin.write_column(
        'K5',
        customer_information_change['old_address'],
        text_left_format,
    )
    sheet_thaydoithongtin.write_column(
        'L5',
        customer_information_change['new_address'],
        text_left_format,
    )
    sheet_thaydoithongtin.write_column(
        'M5',
        customer_information_change['old_nationality'],
        text_center_format,
    )
    sheet_thaydoithongtin.write_column(
        'N5',
        customer_information_change['new_nationality'],
        text_center_format,
    )
    sheet_thaydoithongtin.write_column('O5', customer_information_change['old_note'], text_left_format)
    sheet_thaydoithongtin.write_column('P5', customer_information_change['new_note'], text_left_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    # Write sheet UY QUYEN
    sup_title_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':18
        }
    )
    header_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    text_left_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'left',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    text_highlight_left_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'left',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10,
            'bg_color':'#FFFF00'
        }
    )
    text_center_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    text_highlight_center_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10,
            'bg_color':'#FFFF00'
        }
    )
    date_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'num_format':'dd\/mm\/yyyy',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    date_highlight_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'num_format':'dd\/mm\/yyyy',
            'font_name':'Times New Roman',
            'font_size':10,
            'bg_color':'#FFFF00'
        }
    )
    sheet_uyquyen = workbook.add_worksheet('Ủy quyền')
    sheet_uyquyen.hide_gridlines(option=2)

    sheet_uyquyen.set_column('A:A',4)
    sheet_uyquyen.set_column('B:B',25)
    sheet_uyquyen.set_column('C:D',12)
    sheet_uyquyen.set_column('E:E',35)
    sheet_uyquyen.set_column('F:F',10)
    sheet_uyquyen.set_column('G:G',20)
    sheet_uyquyen.set_column('H:H',12)
    sheet_uyquyen.set_column('I:I',35)
    sheet_uyquyen.set_column('J:J',15)
    sheet_uyquyen.set_column('K:K',6)
    sheet_uyquyen.set_row(0,36)

    sheet_uyquyen.merge_range('A1:J1','Danh sách khách hàng ủy quyền',sup_title_format)
    headers = [
        'STT',
        'Tên khách hàng ủy quyền',
        'Mã TK',
        'Số CMND/ Hộ chiếu/ Giấy ĐKKD người UQ',
        'Địa chỉ  người UQ',
        'Ngày Uỷ quyền',
        'Tên người nhận uỷ quyền',
        'Số CMND/ Hộ chiếu/ Giấy ĐKKD người nhận UQ',
        'Địa chỉ người nhận UQ',
        'Phạm vi uỷ quyền',
    ]
    sheet_uyquyen.write_row('A2', headers, header_format)
    sheet_uyquyen.write_row('A3', [f'({i})' for i in np.arange(len(headers)) + 1], header_format)
    for row in range(authorization.shape[0]):
        ticker = authorization.iloc[row, authorization.columns.get_loc('account_code')]
        if ticker in highlight_account.values:
            fmt1 = text_highlight_center_format
            fmt2 = text_highlight_left_format
            fmt3 = date_highlight_format
        else:
            fmt1 = text_center_format
            fmt2 = text_left_format
            fmt3 = date_format
        sheet_uyquyen.write(row+3,0,row+1,fmt1)
        sheet_uyquyen.write(row+3,1,authorization.iloc[row,authorization.columns.get_loc('authorizing_person_name')],fmt2)
        sheet_uyquyen.write(row+3,2,authorization.iloc[row,authorization.columns.get_loc('account_code')],fmt1)
        sheet_uyquyen.write(row+3,3,authorization.iloc[row,authorization.columns.get_loc('authorizing_person_id')],fmt1)
        sheet_uyquyen.write(row+3,4,authorization.iloc[row,authorization.columns.get_loc('authorizing_person_address')],fmt2)
        sheet_uyquyen.write(row+3,5,authorization.iloc[row,authorization.columns.get_loc('date_of_authorization')],fmt3)
        sheet_uyquyen.write(row+3,6,authorization.iloc[row, authorization.columns.get_loc('authorized_person_name')],fmt1)
        sheet_uyquyen.write(row+3,7,authorization.iloc[row,authorization.columns.get_loc('authorized_person_id')],fmt1)
        sheet_uyquyen.write(row+3,8,authorization.iloc[row,authorization.columns.get_loc('authorized_person_address')],fmt2)
        sheet_uyquyen.write(row+3,9,authorization.iloc[row,authorization.columns.get_loc('scope_of_authorization')],fmt1)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    # Write sheet THAY DOI UY QUYEN
    sup_title_format = workbook.add_format(
        {
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':18,
        }
    )
    header_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
            'font_name':'Times New Roman',
            'font_size':10,
        }
    )
    signature_format = workbook.add_format(
        {
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':12
        }
    )
    text_left_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'left',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    text_center_format = workbook.add_format(
        {
            'border':1,
            'text_wrap':True,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    date_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'num_format':'dd\/mm\/yyyy',
            'font_name':'Times New Roman',
            'font_size':10
        }
    )
    sheet_thaydoiuyquyen = workbook.add_worksheet('Thay đổi ủy quyền')
    sheet_thaydoiuyquyen.hide_gridlines(option=2)

    sheet_thaydoiuyquyen.set_column('A:A',4)
    sheet_thaydoiuyquyen.set_column('B:B',22)
    sheet_thaydoiuyquyen.set_column('C:E',12)
    sheet_thaydoiuyquyen.set_column('F:F',22)
    sheet_thaydoiuyquyen.set_column('G:J',12)
    sheet_thaydoiuyquyen.set_column('K:L',17)
    sheet_thaydoiuyquyen.set_column('M:P',12)
    sheet_thaydoiuyquyen.set_row(0,36)
    sheet_thaydoiuyquyen.set_row(1,30)
    sheet_thaydoiuyquyen.set_row(2,36)

    sheet_thaydoiuyquyen.merge_range(
        'A1:P1',
        'Danh sách khách hàng chấm dứt, thay đổi nội dung ủy quyền',
        sup_title_format
    )
    sheet_thaydoiuyquyen.merge_range('A2:A3','STT',header_format)
    sheet_thaydoiuyquyen.merge_range('B2:B3','Tên khách hàng uỷ quyền',header_format)
    sheet_thaydoiuyquyen.merge_range('C2:C3','Mã TK',header_format)
    sheet_thaydoiuyquyen.merge_range('D2:D3','Số CMND/ Hộ chiếu/ Giấy ĐKKD của khách hàng UQ',header_format)
    sheet_thaydoiuyquyen.merge_range('E2:E3','Ngày Uỷ quyền',header_format)
    sheet_thaydoiuyquyen.merge_range('F2:F3','Tên người nhận UQ',header_format)
    sheet_thaydoiuyquyen.merge_range('G2:G3','Ngày chấm dứt Uỷ quyền',header_format)
    sheet_thaydoiuyquyen.merge_range('H2:H3','Ngày thay đổi ND uỷ quyền',header_format)
    sheet_thaydoiuyquyen.merge_range('I2:J2','Thay đổi CMND/ Hộ chiếu người nhận UQ',header_format)
    sheet_thaydoiuyquyen.merge_range('K2:L2','Thay đổi địa chỉ người nhận UQ',header_format)
    sheet_thaydoiuyquyen.merge_range('M2:N2','Thay đổi phạm vi uỷ quyền',header_format)
    sheet_thaydoiuyquyen.merge_range('O2:P2','Thay đổi thời hạn ủy quyền',header_format)
    sub_header = [
        'Số CMND/ Hộ chiếu cũ',
        'Số CMND/ Hộ chiếu mới',
        'Địa chỉ cũ',
        'Địa chỉ mới',
        'Phạm vi uỷ quyền cũ',
        'Phạm vi uỷ quyền mới',
        'Thời hạn cũ',
        'Thời hạn mới',
    ]
    sheet_thaydoiuyquyen.write_row('I3',sub_header,header_format)
    sheet_thaydoiuyquyen.write_row('A4',[f'({i})' for i in np.arange(16)+1],header_format)
    sheet_thaydoiuyquyen.write_column('A5', authorization_change['no.'], text_center_format)
    sheet_thaydoiuyquyen.write_column('B5',authorization_change['authorizing_person_name'],text_left_format)
    sheet_thaydoiuyquyen.write_column('C5', authorization_change['account_code'], text_center_format)
    sheet_thaydoiuyquyen.write_column('D5', authorization_change['authorizing_person_id'], text_left_format)
    sheet_thaydoiuyquyen.write_column('E5', authorization_change['date_of_authorization'], date_format)
    sheet_thaydoiuyquyen.write_column('F5', authorization_change['authorized_person_name'], text_center_format)
    sheet_thaydoiuyquyen.write_column('G5', authorization_change['date_of_termination'], date_format)
    sheet_thaydoiuyquyen.write_column('H5', authorization_change['date_of_change'], date_format)
    sheet_thaydoiuyquyen.write_column('I5', authorization_change['old_authorized_person_id'], text_center_format)
    sheet_thaydoiuyquyen.write_column('J5', authorization_change['new_authorized_person_id'], text_center_format)
    sheet_thaydoiuyquyen.write_column('K5', authorization_change['old_authorized_person_address'], text_center_format)
    sheet_thaydoiuyquyen.write_column('L5', authorization_change['new_authorized_person_address'], text_center_format)
    sheet_thaydoiuyquyen.write_column('M5', authorization_change['old_scope_of_authorization'], text_center_format)
    sheet_thaydoiuyquyen.write_column('N5', authorization_change['new_scope_of_authorization'], text_center_format)
    sheet_thaydoiuyquyen.write_column('O5', authorization_change['old_end_date'], date_format)
    sheet_thaydoiuyquyen.write_column('P5', authorization_change['new_end_date'], date_format)

    row_of_signature = 3+authorization_change.shape[0]+2
    sheet_thaydoiuyquyen.set_row(row_of_signature-1,55)
    sheet_thaydoiuyquyen.set_row(row_of_signature+2,80)
    sheet_thaydoiuyquyen.merge_range(row_of_signature,1,row_of_signature,4,'Người lập',signature_format)
    sheet_thaydoiuyquyen.merge_range(row_of_signature+1,1,row_of_signature+1,4,'(Ký, ghi rõ họ tên)',signature_format)
    sheet_thaydoiuyquyen.merge_range(row_of_signature+3,1,row_of_signature+3,4,'ĐIỀN HỌ TÊN VÀO Ô',signature_format)
    sheet_thaydoiuyquyen.merge_range(row_of_signature,10,row_of_signature,15,'ĐIỀN CHỨC DANH VÀO Ô',signature_format)
    sheet_thaydoiuyquyen.merge_range(row_of_signature+1,10,row_of_signature+1,15,'(Ký, ghi rõ họ tên)',signature_format)
    sheet_thaydoiuyquyen.merge_range(row_of_signature+3,10,row_of_signature+3,15,'ĐIỀN HỌ TÊN VÀO Ô',signature_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    ###########################################################################
    ###########################################################################
    ###########################################################################

    # Write sheet THAY DOI UY QUYEN
    header_format = workbook.add_format(
        {
            'bold':True,
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':12,
        }
    )
    stt_format = workbook.add_format(
        {
            'border':1,
            'align':'center',
            'valign':'vcenter',
            'num_format':'0',
            'font_name':'Times New Roman',
            'font_size':'12',
        }
    )
    country_format = workbook.add_format(
        {
            'border':1,
            'valign':'vcenter',
            'font_name':'Times New Roman',
            'font_size':12,
        }
    )
    sheet_country_code = workbook.add_worksheet('country_code')
    sheet_country_code.hide_gridlines(option=2)

    sheet_country_code.set_column('A:A',6)
    sheet_country_code.set_column('B:B',35)
    sheet_country_code.set_column('C:C',6)
    sheet_country_code.set_column('D:D',35)

    sheet_country_code.write_row('A1',['STT','QUỐC GIA'],header_format)
    sheet_country_code.write_column('A2',country_code_head.index,stt_format)
    sheet_country_code.write_column('B2',country_code_head,country_format)
    sheet_country_code.write_row('C1',['STT','QUỐC GIA'],header_format)
    sheet_country_code.write_column('C2',country_code_tail.index,stt_format)
    sheet_country_code.write_column('D2',country_code_tail,country_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
