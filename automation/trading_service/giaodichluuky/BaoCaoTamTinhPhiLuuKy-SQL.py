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

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):  # dept_folder from import
        os.mkdir(join(dept_folder,folder_name,period))

    result = pd.read_sql(
        f"""
            WITH [rse9992] AS (
                SELECT
                    [relationship].[branch_id],
                    SUM([depository_fee].[fee_amount]) [fee_amount]
                FROM [depository_fee]
                FULL OUTER JOIN [relationship] 
                ON [relationship].[sub_account] = [depository_fee].[sub_account]
                AND [relationship].[date] = [depository_fee].[date]
                WHERE [depository_fee].[date] BETWEEN '{start_date}' AND '{end_date}'
                GROUP BY [branch_id]
            ),
            [res] AS (
                SELECT
                    CASE 
                        WHEN [branch].[branch_id] = '0001' THEN N'HQ'
                        WHEN [branch].[branch_id] = '0101' THEN N'Quận 3'
                        WHEN [branch].[branch_id] = '0102' THEN N'PMH'
                        WHEN [branch].[branch_id] = '0104' THEN N'Q7'
                        WHEN [branch].[branch_id] = '0105' THEN N'TB'
                        WHEN [branch].[branch_id] = '0116' THEN N'P.QLTK1'
                        WHEN [branch].[branch_id] = '0111' THEN N'InB1'
                        WHEN [branch].[branch_id] = '0113' THEN N'IB'
                        WHEN [branch].[branch_id] = '0201' THEN N'Hà Nội'
                        WHEN [branch].[branch_id] = '0202' THEN N'TX'
                        WHEN [branch].[branch_id] = '0301' THEN N'Hải Phòng'
                        WHEN [branch].[branch_id] = '0117' THEN N'Quận 1'
                        WHEN [branch].[branch_id] = '0118' THEN N'P.QLTK3'
                        WHEN [branch].[branch_id] = '0119' THEN N'InB2'
                    END [Tên Chi Nhánh],
                    [branch].[branch_id],
                    [rse9992].[fee_amount]
                FROM [rse9992]
                FULL OUTER JOIN [branch] ON [rse9992].[branch_id] = [branch].[branch_id]
            )
            SELECT
                ROW_NUMBER() OVER (ORDER BY [branch_id]) [STT],
                [res].[Tên Chi Nhánh],
                [res].[branch_id] [Mã Chi Nhánh],
                ISNULL([res].[fee_amount],0) [Phí Lưu Ký]
            FROM [res]
            WHERE [res].[Tên Chi Nhánh] IS NOT NULL
            ORDER BY [Mã Chi Nhánh]
        """,
        connect_DWH_CoSo
    )

    ######################################################################################
    ######################################################################################
    ######################################################################################

    table_title = f'PHÍ LƯU KÝ {period}'
    # Write to Excel
    file_name = f'Báo cáo phí tạm tính phí lưu ký {period}.xlsx'
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
    worksheet.set_column('B:B',21)
    worksheet.set_column('C:D',18)

    title_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'align':'center',
        }
    )
    header_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'align':'center',
            'border':1,
        }
    )
    stt_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'border':1,
        }
    )
    tenchinhanh_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'border':1
        }
    )
    machinhanh_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'align':'center',
            'border':1
        }
    )
    philuuky_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'border':1
        }
    )
    tong_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'align':'center',
            'border':1,
        }
    )
    tongphiluuky_format = workbook.add_format(
        {
            'font_name':'Times New Roman',
            'font_size':12,
            'bold':True,
            'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            'border':1,
        }
    )
    sum_fee = result['Phí Lưu Ký'].sum()
    worksheet.merge_range('A1:D1',table_title,title_format)
    for col_num,col_name in enumerate(result.columns):
        worksheet.write(2,col_num,col_name,header_format)
    for row in range(result.shape[0]):
        worksheet.write(row+3,0,result.iloc[row,0],stt_format)
        worksheet.write(row+3,1,result.iloc[row,1],tenchinhanh_format)
        worksheet.write(row+3,2,result.iloc[row,2],machinhanh_format)
        worksheet.write(row+3,3,result.iloc[row,3],philuuky_format)
    tong_row = result.shape[0]+3
    worksheet.merge_range(tong_row,0,tong_row,2,'Tổng',tong_format)
    worksheet.write(tong_row,3,sum_fee,tongphiluuky_format)
    writer.close()

    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
