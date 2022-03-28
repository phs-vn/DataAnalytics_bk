from automation.trading_service.giaodichluuky import *


def run(
    run_time=None
):
    start = time.time()
    info = get_info('monthly',run_time)
    start_date = info['start_date']
    end_date = info['end_date']
    period = info['period']
    begin_of_year = f'{start_date[:4]}/01/01'
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):  # dept_folder from import
        os.mkdir(join(dept_folder,folder_name,period))

    result = pd.read_sql(
        f"""
            WITH [cal_period] AS (
                SELECT
                    CASE
                        WHEN [trading_record].[exchange] = 'UPCOM' THEN 'HNX'
                        WHEN [trading_record].[exchange] = 'HOSE' THEN 'HSX' 
                        ELSE [trading_record].[exchange]
                    END [exchange],
                    CASE
                        WHEN [account].[account_type] LIKE N'%trong nước%' THEN 'domestic'
                        WHEN [account].[account_type] LIKE N'%nước ngoài%' THEN 'foreign'
                        ELSE 'dealing'
                    END [account_type],
                    CASE
                        WHEN [trading_record].[type_of_asset] = N'Cổ phiếu thường' THEN 'stock'
                        WHEN [trading_record].[type_of_asset] = N'Chứng chỉ quỹ' THEN 'fund_certificate'
                        WHEN [trading_record].[type_of_asset] = N'Chứng quyền' THEN 'cw'
                        ELSE 'bond'
                    END [type_of_asset],
                    [trading_record].[type_of_order],
                    [trading_record].[value]
                FROM [trading_record]
                LEFT JOIN [sub_account]
                ON [sub_account].[sub_account] = [trading_record].[sub_account]
                LEFT JOIN [account]
                ON [account].[account_code] = [sub_account].[account_code]
                WHERE 
                    [trading_record].[date] BETWEEN '{start_date}' AND '{end_date}'
                    AND [trading_record].[settlement_period] IN (1,2)
            ),
            [cal_ytd] AS (
                SELECT
                    [trading_record].[type_of_order],
                    CASE
                        WHEN [trading_record].[exchange] = 'UPCOM' THEN 'HNX'
                        WHEN [trading_record].[exchange] = 'HOSE' THEN 'HSX' 
                        ELSE [trading_record].[exchange]
                    END [exchange],
                    CASE
                        WHEN [trading_record].[type_of_asset] = N'Cổ phiếu thường' THEN 'stock'
                        WHEN [trading_record].[type_of_asset] = N'Chứng chỉ quỹ' THEN 'fund_certificate'
                        WHEN [trading_record].[type_of_asset] = N'Chứng quyền' THEN 'cw'
                        ELSE 'bond'
                    END [type_of_asset],
                    CASE
                        WHEN [account].[account_type] LIKE N'%trong nước%' THEN 'domestic'
                        WHEN [account].[account_type] LIKE N'%nước ngoài%' THEN 'foreign'
                        ELSE 'dealing'
                    END [account_type],
                    [trading_record].[value]
                FROM [trading_record]
                LEFT JOIN [sub_account]
                ON [sub_account].[sub_account] = [trading_record].[sub_account]
                LEFT JOIN [account]
                ON [account].[account_code] = [sub_account].[account_code]
                WHERE
                    [trading_record].[date] BETWEEN '{begin_of_year}' AND '{end_date}'
                    AND [trading_record].[settlement_period] IN (1,2)
            ),
            [period_tab] AS (
                SELECT
                    [cal_period].[type_of_order], 
                    [cal_period].[exchange], 
                    [cal_period].[type_of_asset], 
                    [cal_period].[account_type],
                    SUM([cal_period].[value]) [value_period] 
                FROM [cal_period]
                GROUP BY [cal_period].[type_of_order],[cal_period].[exchange],[cal_period].[type_of_asset],[cal_period].[account_type]
            ),
            [ytd_tab] AS (
                SELECT
                    [cal_ytd].[type_of_order], 
                    [cal_ytd].[exchange], 
                    [cal_ytd].[type_of_asset], 
                    [cal_ytd].[account_type],
                    SUM([cal_ytd].[value]) [value_ytd]
                FROM [cal_ytd]
                GROUP BY [cal_ytd].[type_of_order],[cal_ytd].[exchange],[cal_ytd].[type_of_asset],[cal_ytd].[account_type]
            ),
            [val_period] AS (
                SELECT
                    [type_of_asset],
                    [account_type],
                    ISNULL(SUM([B_HNX_val_period]), 0) [B_HNX_val_period],
                    ISNULL(SUM([B_HSX_val_period]), 0) [B_HSX_val_period],
                    ISNULL(SUM([S_HNX_val_period]), 0) [S_HNX_val_period],
                    ISNULL(SUM([S_HSX_val_period]), 0) [S_HSX_val_period]
                FROM (
                    SELECT
                        CONCAT([period_tab].[type_of_order],'_',[period_tab].[exchange],'_val_period') [type_of_order_exchange_period],
                        [period_tab].[type_of_asset],
                        [period_tab].[account_type],
                        ISNULL([period_tab].[value_period],0) [value_period]
                    FROM [period_tab]
                ) [tab1]
                PIVOT (
                    SUM([tab1].[value_period]) FOR [type_of_order_exchange_period] IN (B_HNX_val_period,S_HNX_val_period,B_HSX_val_period,S_HSX_val_period)
                ) [tab2]
                GROUP BY [type_of_asset], [account_type]
            ),
            [val_ytd] AS (
                SELECT
                    [type_of_asset],
                    [account_type],
                    ISNULL(SUM([B_HNX_val_ytd]),0) [B_HNX_val_ytd],
                    ISNULL(SUM([B_HSX_val_ytd]),0) [B_HSX_val_ytd],
                    ISNULL(SUM([S_HNX_val_ytd]),0) [S_HNX_val_ytd],
                    ISNULL(SUM([S_HSX_val_ytd]),0) [S_HSX_val_ytd]
                FROM (
                    SELECT
                        CONCAT([ytd_tab].[type_of_order],'_',[ytd_tab].[exchange],'_val_ytd') [type_of_order_exchange_ytd],
                        [ytd_tab].[type_of_asset],
                        [ytd_tab].[account_type],
                        ISNULL([ytd_tab].[value_ytd],0) [value_ytd]
                    FROM [ytd_tab]
                ) [tab1]
                PIVOT (
                    SUM([tab1].[value_ytd]) FOR [type_of_order_exchange_ytd] IN (B_HNX_val_ytd,S_HNX_val_ytd,B_HSX_val_ytd,S_HSX_val_ytd)
                ) [tab2]
                GROUP BY [type_of_asset], [account_type]
            ),
            [result] AS (
                SELECT
                    CONCAT([vp].[type_of_asset], '_', [vp].[account_type]) [stock_type],
                    [vp].[B_HNX_val_period],
                    [vp].[B_HSX_val_period],
                    [vy].[B_HNX_val_ytd],
                    [vy].[B_HSX_val_ytd],
                    [vp].[S_HNX_val_period],
                    [vp].[S_HSX_val_period],
                    [vy].[S_HNX_val_ytd],
                    [vy].[S_HSX_val_ytd],
                    ([vp].[B_HNX_val_period] + [vp].[S_HNX_val_period]) [T_HNX_val_period],
                    ([vp].[B_HSX_val_period] + [vp].[S_HSX_val_period]) [T_HSX_val_period],
                    ([vy].[B_HNX_val_ytd] + [vy].[S_HNX_val_ytd]) [T_HNX_val_ytd],
                    ([vy].[B_HSX_val_ytd] + [vy].[S_HSX_val_ytd]) [T_HSX_val_ytd]
                FROM 
                    [val_period] [vp],
                    [val_ytd] [vy]
                WHERE [vp].[type_of_asset] = [vy].[type_of_asset]
                AND [vp].[account_type] = [vy].[account_type]
            )
            SELECT 
                vals.b_s_time,
                MAX(CASE WHEN vals.stock_type = 'stock_domestic' THEN [value] END) [stock_domestic],
                MAX(CASE WHEN vals.stock_type = 'stock_foreign' THEN [value] END) [stock_foreign],
                MAX(CASE WHEN vals.stock_type = 'fund_certificate_domestic' THEN [value] END) [fund_certificate_domestic],
                MAX(CASE WHEN vals.stock_type = 'fund_certificate_foreign' THEN [value] END) [fund_certificate_foreign],
                MAX(CASE WHEN vals.stock_type = 'stock_dealing' THEN [value] END) [stock_dealing],
                MAX(CASE WHEN vals.stock_type = 'bond_dealing' THEN [value] END) [bond_dealing],
                ISNULL(MAX(CASE WHEN vals.stock_type='fund_certificate_dealing' THEN [value] END),0) [fund_certificate_dealing]
            FROM [result] CROSS APPLY (
                values ([result].[stock_type], [result].[B_HNX_val_period], 'B_HNX_val_period', 1),
                        ([result].[stock_type], [result].[B_HSX_val_period], 'B_HSX_val_period', 2),
                        ([result].[stock_type], [result].[B_HNX_val_ytd], 'B_HNX_val_ytd', 3),
                        ([result].[stock_type], [result].[B_HSX_val_ytd], 'B_HSX_val_ytd', 4),
                        ([result].[stock_type], [result].[S_HNX_val_period], 'S_HNX_val_period', 5),
                        ([result].[stock_type], [result].[S_HSX_val_period], 'S_HSX_val_period', 6),
                        ([result].[stock_type], [result].[S_HNX_val_ytd], 'S_HNX_val_ytd', 7),
                        ([result].[stock_type], [result].[S_HSX_val_ytd], 'S_HSX_val_ytd', 8),
                        ([result].[stock_type], [result].[T_HNX_val_period], 'T_HNX_val_period', 9),
                        ([result].[stock_type], [result].[T_HSX_val_period], 'T_HSX_val_period', 10),
                        ([result].[stock_type], [result].[T_HNX_val_ytd], 'T_HNX_val_ytd', 11),
                        ([result].[stock_type], [result].[T_HSX_val_ytd], 'T_HSX_val_ytd', 12)
                ) vals([stock_type], [value], [b_s_time], [ordering])
            GROUP BY vals.[b_s_time]
            ORDER BY MAX([ordering])
        """,
        connect_DWH_CoSo
    )

    #################################################################################################
    #################################################################################################
    #################################################################################################

    # Write to Báo cáo phí chuyển khoản
    file_name = f'Báo cáo tình hình HĐKD (Biểu II.6) {period}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book
    worksheet = workbook.add_worksheet(period)
    worksheet.hide_gridlines(option=2)
    # set column width
    worksheet.set_column('A:A',8)
    worksheet.set_column('B:B',16)
    worksheet.set_column('C:N',12)
    worksheet.set_row(7,47)
    worksheet.set_row(8,47)
    worksheet.set_row(9,47)
    worksheet.set_row(10,47)
    title_format = workbook.add_format(
        {
            'bold':True,
            'font_name':'Times New Roman',
            'font_size':13,
            'align':'center',
        }
    )
    dvt_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'italic':True,
            'align':'right',
        }
    )
    supheader_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'align':'center',
            'valign':'vcenter',
            'border':1,
        }
    )
    header_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'valign':'vcenter',
            'border':1,
        }
    )
    stt_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'valign':'vcenter',
            'border':1,
        }
    )
    index_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'text_wrap':True,
            'valign':'top',
            'border':1
        }
    )
    supindex_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'text_wrap':True,
            'valign':'top',
            'border':1
        }
    )
    value_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'valign':'vcenter',
            'border':1
        }
    )
    worksheet.merge_range('A1:N1',f'Báo Cáo Tình Hình HĐKD - Biểu II.6 - {period}',title_format)
    worksheet.merge_range('A3:A5','STT',header_format)
    worksheet.merge_range('B3:B5','Loại chứng khoán',header_format)
    worksheet.merge_range('C3:F3','Tổng mua',supheader_format)
    worksheet.merge_range('G3:J3','Tổng bán',supheader_format)
    worksheet.merge_range('K3:N3','Tổng mua + bán',supheader_format)
    worksheet.merge_range('C4:D4','Trong kỳ',header_format)
    worksheet.merge_range('E4:F4','Lũy kế từ đầu năm',header_format)
    worksheet.merge_range('G4:H4','Trong kỳ',header_format)
    worksheet.merge_range('I4:J4','Lũy kế từ đầu năm',header_format)
    worksheet.merge_range('K4:L4','Trong kỳ',header_format)
    worksheet.merge_range('M4:N4','Lũy kế từ đầu năm',header_format)
    worksheet.write_row('C5',['HNX','HSX']*6,header_format)
    worksheet.write_row('A6',np.arange(14)+1,header_format)
    worksheet.write_column('A7',np.arange(9)+1,stt_format)
    idx_list = [
        'A. Nhà đầu tư',
        '1. Giao dịch cổ phiếu của nhà đầu tư trong nước',
        '2. Giao dịch cổ phiếu của nhà đầu tư nước ngoài',
        '3. Giao dịch chứng chỉ quỹ của nhà đầu tư trong nước',
        '4. Giao dịch chứng chỉ quỹ của nhà đầu tư nước ngoài',
        'B. Tự doanh',
        '1. Cổ phiếu',
        '2. Trái phiếu',
        '3. Chứng chỉ quỹ',
    ]
    for row,idx in enumerate(idx_list):
        if idx in ['A. Nhà đầu tư','B. Tự doanh']:
            fmt = supindex_format
        else:
            fmt = index_format
        worksheet.write(row+6,1,idx,fmt)
    worksheet.write_row('C7',[None]*12,value_format)
    worksheet.write_row('C12',[None]*12,value_format)
    dvt = 'mil'
    if dvt=='bil':
        div = 1e9
        dvt_text = 'Đơn vị tính: tỷ đồng'
    elif dvt=='mil':
        div = 1e6
        dvt_text = 'Đơn vị tính: triệu đồng'
    elif dvt=='k':
        div = 1e3
        dvt_text = 'Đơn vị tính: nghìn đồng'
    elif dvt=='unit':
        div = 1
        dvt_text = 'Đơn vị tính: đồng'
    else:
        raise ValueError("dvt must be either 'bil', 'mil', 'k', 'unit'")
    worksheet.merge_range('L2:N2',dvt_text,dvt_format)
    worksheet.write_row('C8',result['stock_domestic'].values/div,value_format)
    worksheet.write_row('C9',result['stock_foreign'].values/div,value_format)
    worksheet.write_row('C10',result['fund_certificate_domestic'].values/div,value_format)
    worksheet.write_row('C11',result['fund_certificate_foreign'].values/div,value_format)
    worksheet.write_row('C13',result['stock_dealing'].values/div,value_format)
    worksheet.write_row('C14',result['bond_dealing'].values/div,value_format)
    worksheet.write_row('C15',result['fund_certificate_dealing'].values/div, value_format)
    writer.close()

    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
