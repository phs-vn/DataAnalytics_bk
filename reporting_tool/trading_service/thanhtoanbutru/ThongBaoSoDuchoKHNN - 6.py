from reporting_tool.trading_service.thanhtoanbutru import *


def run(
        periodicity: str,
        run_time=None,
):
    start = time.time()
    info = get_info(periodicity, run_time)
    period = info['period']
    end_date = info['end_date']
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder, folder_name)):
        os.mkdir(join(dept_folder, folder_name))
    if not os.path.isdir(join(dept_folder, folder_name, period)):
        os.mkdir((join(dept_folder, folder_name, period)))

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viết Query ---------------------
    so_du_TK_query = pd.read_sql(
        f"""
            select top 10 *
            from rmr1062
        """,
        connect_DWH_CoSo
    )

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File Excel ---------------------
    # Write file ROD0040
    eom = dt.datetime.strptime(end_date, "%Y/%m/%d").strftime("%d.%m.%Y")
    f_name = f'Số dư TK FII ngày {eom} - gửi IT.xlsx'
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
    headers_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman'
        }
    )
    ten_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'top',
            'font_size': 11,
            'font_name': 'Calibri'
        }
    )
    account_and_email_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri'
        }
    )
    money_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri',
            'num_format': '#,##0'
        }
    )
    date_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'num_format': 'mmmm d, yyyy'
        }
    )
    headers = [
        'Ngày',
        'Số TK',
        'Tên',
        'Số tiền',
        'Email'
    ]

    # --------- sheet Cá nhân ---------
    sheet_ca_nhan = workbook.add_worksheet('Cá nhân')
    sheet_ca_nhan.set_tab_color('#FF0000')
    # Set Column Width and Row Height
    sheet_ca_nhan.set_column('A:A', 17)
    sheet_ca_nhan.set_column('B:B', 10.14)
    sheet_ca_nhan.set_column('C:C', 30)
    sheet_ca_nhan.set_column('D:D', 37.86)
    sheet_ca_nhan.set_column('E:E', 13.57)

    sheet_ca_nhan.write_row('A2', headers, headers_format)
    sheet_ca_nhan.write_column(
        'A3',
        so_du_TK_query['date'],
        date_format
    )
    sheet_ca_nhan.write_column(
        'B3',
        so_du_TK_query['account_code'],
        account_and_email_format
    )
    sheet_ca_nhan.write_column(
        'C3',
        so_du_TK_query['customer_name'],
        ten_format
    )
    sheet_ca_nhan.write_column(
        'D3',
        so_du_TK_query[''],
        money_format
    )
    sheet_ca_nhan.write_column(
        'E3',
        [''] * so_du_TK_query.shape[0],
        account_and_email_format
    )

    # --------- sheet Tổ chức ---------
    ten_to_chuc_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    account_and_email_to_chuc_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman'
        }
    )
    money_to_chuc_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'num_format': '#,##0'
        }
    )
    sheet_to_chuc = workbook.add_worksheet('Tổ chức')
    sheet_to_chuc.set_tab_color('#FFFF00')
    # Set Column Width and Row Height
    sheet_to_chuc.set_column('A:A', 17.14)
    sheet_to_chuc.set_column('B:B', 11.43)
    sheet_to_chuc.set_column('C:C', 57.71)
    sheet_to_chuc.set_column('D:D', 37.86)
    sheet_to_chuc.set_column('E:E', 13.14)
    sheet_to_chuc.write_row('A2', headers, headers_format)


    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py', '')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time() - start, 1)}s')



