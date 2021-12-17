from reporting_tool.trading_service.thanhtoanbutru import *

# DONE
def run(
        periodicity:str,
        run_time=None,
):
    start = time.time()
    info = get_info(periodicity,run_time)
    start_date = info['start_date']
    end_date = info['end_date']
    period = info['period']
    folder_name = info['folder_name']

    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir((join(dept_folder,folder_name,period)))

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viết Query ---------------------
    table = pd.read_sql(
        f"""
        SELECT
            [cash_balance].[date], 
            [relationship].[account_code], 
            [relationship].[sub_account], 
            [account].[customer_name], 
            [cash_balance].[transaction_id],
            [cash_balance].[remark],
            [cash_balance].[increase],
            [cash_balance].[decrease],
            [cash_balance].[creator],
            [cash_balance].[approver]
        FROM 
            [cash_balance]
        LEFT JOIN 
            [relationship] 
        ON 
            [relationship].[sub_account] = [cash_balance].[sub_account] AND [relationship].[date] = [cash_balance].[date]
        LEFT JOIN 
            [account] 
        ON 
            [account].[account_code] = [relationship].[account_code]
        WHERE 
            [cash_balance].[date] BETWEEN '{start_date}' AND '{end_date}'
        AND 
            [relationship].[account_code] != '022P002222'
        """,
        connect_DWH_CoSo
    )
    table.sort_values(['date','account_code','sub_account'],inplace=True)
    # Lấy tổng số dư đầu kỳ
    opening_balance = pd.read_sql(
        f"""
        SELECT 
            SUM(opening_balance) AS sum_opening_balance 
        FROM 
            [sub_account_deposit] 
        WHERE 
            [sub_account_deposit].[date] = '{start_date}'
        """,
        connect_DWH_CoSo
    )

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File Excel ---------------------
    # Write file BÁO CÁO PHÁT SINH GIAO DỊCH TIỀN
    footer_date = bdate(end_date,1).split('-')
    start_date = dt.datetime.strptime(start_date,"%Y/%m/%d").strftime("%d-%m-%Y")
    end_date = dt.datetime.strptime(end_date,"%Y/%m/%d").strftime("%d-%m-%Y")
    file_name = 'Báo cáo phát sinh giao dịch tiền.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options': {'nan_inf_to_errors': True}}
    )
    workbook = writer.book

    ###################################################
    ###################################################
    ###################################################

    # ------------- Viet sheet -------------
    # Format
    company_name_format = workbook.add_format(
        {
            'bold': True,
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    company_format = workbook.add_format(
        {
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    empty_row_format = workbook.add_format(
        {
            'bottom': 1,
            'valign': 'top',
            'font_size': 10,
            'font_name': 'Arial',
        }
    )
    sheet_title_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 16,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    from_to_format = workbook.add_format(
        {
            'italic': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    headers_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 12,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    sub_header_empty_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    so_du_dau_cuoi_cong_ps_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
        }
    )
    num_right_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'top',
            'font_size': 10,
            'font_name': 'Times New Roman'
        }
    )
    date_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'top',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'num_format': 'dd/mm/yyyy'
        }
    )
    text_left_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'top',
            'font_size': 10,
            'font_name': 'Times New Roman'
        }
    )
    text_center_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'top',
            'font_size': 10,
            'font_name': 'Times New Roman'
        }
    )
    money_inc_dec_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'top',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'num_format': '#,##0'
        }
    )
    money_dau_cuoi_ky_cong_ps = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'num_format': '#,##0'
        }
    )
    footer_dmy_format = workbook.add_format(
        {
            'italic': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman',
        }
    )
    footer_text_format = workbook.add_format(
        {
            'bold': True,
            'italic': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    empty_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
        }
    )
    headers = [
        'STT',
        'Ngày',
        'Số tài khoản',
        'Số tiểu khoản',
        'Tên khách hàng',
        'Mã nghiệp vụ',
        'Tên nghiệp vụ',
        'Tăng tiền',
        'Giảm tiền',
        'Người lập',
        'Người duyệt',
    ]

    # --------- sheet BAO CAO CAN LAM ---------
    worksheet = workbook.add_worksheet(f'{period}')
    worksheet.hide_gridlines(option=2)

    # Insert Phu Hung Logo
    worksheet.insert_image('A1','./img/phs_logo.png',{'x_scale':0.62,'y_scale':0.71})

    # Set Column Width and Row Height
    worksheet.set_column('A:A',7)
    worksheet.set_column('B:B',11)
    worksheet.set_column('C:C',14)
    worksheet.set_column('D:D',15)
    worksheet.set_column('E:E',41)
    worksheet.set_column('F:F',10)
    worksheet.set_column('G:G',67)
    worksheet.set_column('H:I',19)
    worksheet.set_column('J:K',13)
    worksheet.set_row(5,21)
    worksheet.set_row(9,32)

    # row of Cộng phát sinh and Số dư tiền cuối kỳ
    sum_row = table.shape[0] + 12
    so_du_tien_cuoi_ky_start_row = table.shape[0] + 13

    worksheet.merge_range('C1:I1',CompanyName,company_name_format)
    worksheet.merge_range('C2:I2',CompanyAddress,company_format)
    worksheet.merge_range('C3:I3',CompanyPhoneNumber,company_format)
    worksheet.merge_range('A6:K6','BÁO CÁO PHÁT SINH GIAO DỊCH TIỀN',sheet_title_format)
    worksheet.merge_range('A7:K7',f'Từ ngày {start_date} đến ngày {end_date}',from_to_format)
    worksheet.merge_range('A11:G11','Số dư tiền đầu kỳ',so_du_dau_cuoi_cong_ps_format)
    worksheet.merge_range('H11:I11',opening_balance['sum_opening_balance'],money_dau_cuoi_ky_cong_ps)
    worksheet.merge_range('J11:K11','',sub_header_empty_format)
    worksheet.merge_range(
        f'A{sum_row}:G{sum_row}',
        'Cộng phát sinh',
        so_du_dau_cuoi_cong_ps_format
    )
    worksheet.merge_range(
        f'A{so_du_tien_cuoi_ky_start_row}:G{so_du_tien_cuoi_ky_start_row}',
        'Số dư tiền cuối kỳ',
        so_du_dau_cuoi_cong_ps_format
    )
    calc_increase_decrease = table['increase'].sum() - table['decrease'].sum()
    worksheet.merge_range(
        f'H{so_du_tien_cuoi_ky_start_row}:I{so_du_tien_cuoi_ky_start_row}',
        opening_balance['sum_opening_balance'] + calc_increase_decrease,
        money_dau_cuoi_ky_cong_ps
    )
    worksheet.merge_range(
        f'J{so_du_tien_cuoi_ky_start_row}:K{so_du_tien_cuoi_ky_start_row}',
        '',
        empty_format
    )
    worksheet.merge_range(
        f'A{so_du_tien_cuoi_ky_start_row+3}:D{so_du_tien_cuoi_ky_start_row+3}',
        'Người lập',
        footer_text_format
    )
    worksheet.merge_range(
        f'I{so_du_tien_cuoi_ky_start_row+2}:K{so_du_tien_cuoi_ky_start_row+2}',
        f'Ngày{footer_date[2]} tháng {footer_date[1]} năm {footer_date[0]}',
        footer_dmy_format
    )
    worksheet.merge_range(
        f'I{so_du_tien_cuoi_ky_start_row+3}:K{so_du_tien_cuoi_ky_start_row+3}',
        'Người duyệt',
        footer_text_format
    )
    # write row, column
    worksheet.write_row(
        'A4',
        [''] * len(headers),
        empty_row_format
    )
    worksheet.write_row('A10', headers, headers_format)
    worksheet.write_column('A12',np.arange(table.shape[0])+1,num_right_format)
    worksheet.write_column('B12',table['date'], date_format)
    worksheet.write_column('C12',table['account_code'],text_center_format)
    worksheet.write_column('D12',table['sub_account'],text_center_format)
    worksheet.write_column('E12',table['customer_name'],text_left_format)
    worksheet.write_column('F12',table['transaction_id'],text_center_format)
    worksheet.write_column('G12',table['remark'],text_left_format)
    worksheet.write_column('H12',table['increase'],money_inc_dec_format)
    worksheet.write_column('I12',table['decrease'],money_inc_dec_format)
    worksheet.write_column('J12',table['creator'],text_center_format)
    worksheet.write_column('K12',table['approver'],text_center_format)
    worksheet.write(f'H{sum_row}',table['increase'].sum(),money_dau_cuoi_ky_cong_ps)
    worksheet.write(f'I{sum_row}',table['decrease'].sum(),money_dau_cuoi_ky_cong_ps)
    worksheet.write(f'J{sum_row}','',empty_format)
    worksheet.write(f'K{sum_row}','',empty_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py', '')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time() - start, 1)}s')