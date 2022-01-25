from reporting.trading_service.thanhtoanbutru import *


def run(
        start_date: str,
        end_date: str,
):
    start = time.time()
    folder_name = 'BaoCaoNgay'
    date_character = ['/', '-', '.']

    # create folder
    if not os.path.isdir(join(dept_folder, folder_name)):  # dept_folder from import
        os.mkdir(join(dept_folder, folder_name))

    start_fixed = ''
    end_fixed = ''
    for date_char in date_character:
        if date_char in end_date:
            if not os.path.isdir(join(dept_folder, folder_name, end_date.replace(date_char, '.'))):
                os.mkdir((join(dept_folder, folder_name, end_date.replace(date_char, '.'))))
                end_fixed = end_fixed + end_date.replace(date_char, '.')
            else:
                end_fixed = end_fixed + end_date.replace(date_char, '.')
        if date_char in start_date:
            start_fixed = start_fixed + start_date.replace(date_char, '.')

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viết Query ---------------------
    nop_rut_query = pd.read_sql(
        f"""
            SELECT 
            date, 
            account.customer_name, 
            bank, bank_account, 
            sub_account.account_code, 
            transaction_id, 
            amount, 
            status
            FROM money_in_out_transfer
            LEFT JOIN sub_account ON money_in_out_transfer.sub_account = sub_account.sub_account
            LEFT JOIN account ON sub_account.account_code = account.account_code
            WHERE (money_in_out_transfer.bank IN ('OCB'))
            AND (money_in_out_transfer.transaction_id IN ('6692', '6693'))
            AND (date BETWEEN '{start_date}' AND '{end_date}')
        """,
        connect_DWH_CoSo
    )

    # chia query thành 2 query theo bank
    nop_rut_query_ocb = nop_rut_query.loc[nop_rut_query['bank'] == 'OCB'].copy()
    nop_rut_query_ocb = nop_rut_query_ocb.sort_values(by=['date'])

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File ---------------------
    # Write file excel Bao cao doi chieu file ngan hang
    file_name = f'Xác nhận nhập rút OCB.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder, folder_name, end_fixed, file_name),
        engine='xlsxwriter',
        engine_kwargs={'options': {'nan_inf_to_errors': True}}
    )

    workbook = writer.book

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet sheet ---------------------
    # Format
    header_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 11
        }
    )
    text_left_format = workbook.add_format(
        {
            'border': 1,
            'text_wrap': True,
            'align': 'left',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 9
        }
    )
    text_center_format = workbook.add_format(
        {
            'border': 1,
            'text_wrap': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 9
        }
    )
    file_title_header_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'valign': 'vtop',
            'font_name': 'Times New Roman',
            'font_size': 11
        }
    )
    file_title_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': 'Cambria',
            'font_size': 15
        }
    )
    sub_title_format = workbook.add_format(
        {
            'align': 'center',
            'valign': 'vbottom',
            'font_name': 'Times New Roman',
            'font_size': 11
        }
    )
    date_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'num_format': 'dd/mm/yyyy',
            'font_name': 'Times New Roman',
            'font_size': 9
        }
    )
    date_xuat_file_format = workbook.add_format(
        {
            'italic': True,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'num_format': 'dd/mm/yyyy',
            'font_name': 'Times New Roman',
            'font_size': 11
        }
    )
    amount_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 9,
            'num_format': '#,##0'
        }
    )
    footer_1_format = workbook.add_format(
        {
            'align': 'center',
            'valign': 'vbottom',
            'font_name': 'Cambria',
            'font_size': 12
        }
    )
    footer_2_format = workbook.add_format(
        {
            'align': 'center',
            'valign': 'vbottom',
            'font_name': 'Cambria',
            'font_size': 12
        }
    )
    footer_3_format = workbook.add_format(
        {
            'align': 'center',
            'font_name': 'Cambria',
            'font_size': 12
        }
    )

    # Column headers
    headers = [
        'Ngày',
        'Mã GD',
        'HỌ VÀ TÊN',
        'TK NH',
        'TK CK',
        'SỐ TIỀN',
        'NỘP TIỀN',
        'RÚT TIỀN',
        'XN BÊN NH',
        'GHI CHÚ',
    ]

    # Sheet lọc theo ngày
    title_file_name = 'GIẤY XÁC NHẬN ĐỒNG Ý CHO CÁC KHÁCH HÀNG RÚT TIỀN \nCHUYỂN KHOẢN HOẶC NỘP TIỀN'
    sub_title_name = 'TP.HCM, Ngày ... Tháng 11 Năm 2021'
    footer_1_name = 'XÁC NHẬN CỦA OCB - PHẠM VĂN HAI'
    footer_2_name = 'XÁC NHẬN CỦA CTY CK PHÚ HƯNG'
    footer_3_name = 'NGƯỜI LẬP'
    footer_4_name = 'TRƯỞNG ĐƠN VỊ'
    footer_5_name = 'CAO ĐAN THANH'
    footer_6_name = 'TRẦN THU TRANG'

    start_date_format = dt.strptime(start_fixed, "%Y.%m.%d").strftime('%d/%m/%Y')
    end_date_format = dt.strptime(end_fixed, "%Y.%m.%d").strftime('%d/%m/%Y')
    ngay_xuat_file = f'from {start_date_format} to {end_date_format}'

    sheet_loc_theo_ngay = workbook.add_worksheet('lọc theo ngày')

    # Set Column Width and Row Height
    sheet_loc_theo_ngay.set_column('A:A', 12)
    sheet_loc_theo_ngay.set_column('B:B', 0)
    sheet_loc_theo_ngay.set_column('C:C', 25)
    sheet_loc_theo_ngay.set_column('D:D', 18)
    sheet_loc_theo_ngay.set_column('E:E', 13)
    sheet_loc_theo_ngay.set_column('F:F', 0)
    sheet_loc_theo_ngay.set_column('G:G', 14)
    sheet_loc_theo_ngay.set_column('H:H', 14)
    sheet_loc_theo_ngay.set_column('I:I', 12)
    sheet_loc_theo_ngay.set_column('J:J', 12)

    sheet_loc_theo_ngay.set_default_row(23)
    sheet_loc_theo_ngay.set_row(0, 15)
    sheet_loc_theo_ngay.set_row(1, 15.75)
    sheet_loc_theo_ngay.set_row(2, 15)
    sheet_loc_theo_ngay.set_row(3, 15)
    sheet_loc_theo_ngay.set_row(4, 24)
    sheet_loc_theo_ngay.set_row(5, 19.5)

    last_row_of_df = nop_rut_query_ocb.shape[0]
    sheet_loc_theo_ngay.set_row(last_row_of_df + 8, 15.75)
    sheet_loc_theo_ngay.set_row(last_row_of_df + 9, 15.75)
    sheet_loc_theo_ngay.set_row(last_row_of_df + 10, 15.75)
    sheet_loc_theo_ngay.set_row(last_row_of_df + 15, 15.75)

    # Insert image OCB
    sheet_loc_theo_ngay.insert_image('A1', './img/ocb_img.png')

    footer_start_row = nop_rut_query_ocb.shape[0] + 10
    sheet_loc_theo_ngay.merge_range('D2:F2', ngay_xuat_file, date_xuat_file_format)
    sheet_loc_theo_ngay.merge_range('G1:J1', xhcn_title, file_title_header_format)
    sheet_loc_theo_ngay.merge_range('G2:J2', xhcn_2_title, file_title_header_format)
    sheet_loc_theo_ngay.merge_range('A5:J6', title_file_name, file_title_format)
    sheet_loc_theo_ngay.merge_range('H4:J4', sub_title_name, sub_title_format)
    sheet_loc_theo_ngay.merge_range(f'A{footer_start_row}:C{footer_start_row}', footer_1_name, footer_1_format)
    sheet_loc_theo_ngay.merge_range(f'H{footer_start_row}:J{footer_start_row}', footer_2_name, footer_2_format)
    sheet_loc_theo_ngay.merge_range(f'C{footer_start_row+1}:D{footer_start_row+1}', footer_4_name, footer_2_format)
    sheet_loc_theo_ngay.merge_range(f'G{footer_start_row+1}:H{footer_start_row+1}', footer_3_name, footer_2_format)
    sheet_loc_theo_ngay.merge_range(f'I{footer_start_row+1}:J{footer_start_row+1}', footer_4_name, footer_2_format)
    sheet_loc_theo_ngay.merge_range(f'G{footer_start_row+6}:H{footer_start_row+6}', footer_5_name, footer_2_format)
    sheet_loc_theo_ngay.merge_range(f'I{footer_start_row+6}:J{footer_start_row+6}', footer_6_name, footer_2_format)

    # Thêm df Nộp tiền và Rút tiền
    nop_rut_query_ocb['NopTien'] = nop_rut_query_ocb.loc[nop_rut_query_ocb['transaction_id'] == '6692', 'amount']
    nop_rut_query_ocb['NopTien'].fillna('', inplace=True)
    nop_rut_query_ocb['RutTien'] = nop_rut_query_ocb.loc[nop_rut_query_ocb['transaction_id'] == '6693', 'amount']
    nop_rut_query_ocb['RutTien'].fillna('', inplace=True)

    sheet_loc_theo_ngay.write_row('A8', headers, header_format)

    sheet_loc_theo_ngay.write_column(
        'A9',
        nop_rut_query_ocb['date'],
        date_format
    )
    sheet_loc_theo_ngay.write_column(
        'B9',
        nop_rut_query_ocb['transaction_id'],
        text_center_format
    )
    sheet_loc_theo_ngay.write_column(
        'C9',
        nop_rut_query_ocb['customer_name'],
        text_left_format
    )
    sheet_loc_theo_ngay.write_column(
        'D9',
        nop_rut_query_ocb['bank_account'],
        text_center_format
    )
    sheet_loc_theo_ngay.write_column(
        'E9',
        nop_rut_query_ocb['account_code'],
        text_center_format
    )
    sheet_loc_theo_ngay.write_column(
        'F9',
        nop_rut_query_ocb['amount'],
        text_center_format
    )
    sheet_loc_theo_ngay.write_column(
        'G9',
        nop_rut_query_ocb['NopTien'],
        amount_format
    )
    sheet_loc_theo_ngay.write_column(
        'H9',
        nop_rut_query_ocb['RutTien'],
        amount_format
    )
    sheet_loc_theo_ngay.write_column(
        'I9',
        ['ĐỒNG Ý'] * nop_rut_query_ocb.shape[0],
        text_left_format
    )
    sheet_loc_theo_ngay.write_column(
        'J9',
        [''] * nop_rut_query_ocb.shape[0],
        text_center_format
    )
    sheet_loc_theo_ngay.write(f'A{footer_start_row + 1}', footer_3_name, footer_3_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py', '')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time() - start, 1)}s')
