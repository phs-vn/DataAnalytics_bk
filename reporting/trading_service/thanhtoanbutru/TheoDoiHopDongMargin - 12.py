from reporting.trading_service.thanhtoanbutru import *


def run(
        periodicity: str,
        start_date: str,
        end_date: str,
        run_time=None,
):
    date_character = ['/', '-', '.']

    start = time.time()
    info = get_info(periodicity, run_time)
    # start_date = info['start_date']
    # end_date = info['end_date']
    period = info['period']
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder, folder_name)):  # dept_folder from import
        os.mkdir(join(dept_folder, folder_name))
    if not os.path.isdir(join(dept_folder, folder_name, period)):
        os.mkdir((join(dept_folder, folder_name, period)))

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viết Query ---------------------
    tracking_margin_query = pd.read_sql(
        f"""
            SELECT
            DISTINCT 
            branch.branch_name, 
            account.account_code, 
            account.customer_name, 
            new_sub_account.open_date,
            account.contract_number_margin
            FROM new_sub_account
            LEFT JOIN relationship ON relationship.sub_account = new_sub_account.sub_account
            LEFT JOIN account ON account.account_code = relationship.account_code
            LEFT JOIN branch ON branch.branch_id = relationship.branch_id
            WHERE open_date between '{start_date}' and '{end_date}'
            AND relationship.date = '{end_date}'
            AND account.contract_number_margin IS NOT NULL
            ORDER BY open_date
        """,
        connect_DWH_CoSo
    )

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File Excel ---------------------
    # Write file excel Bao cao doi chieu file ngan hang
    for date_char in date_character:
        if date_char in start_date and date_char in end_date:
            start_date = start_date.replace(date_char, '/')
            end_date = end_date.replace(date_char, '/')

    start_date = dt.datetime.strptime(start_date, "%Y/%m/%d").strftime("%d-%m")
    end_date = dt.datetime.strptime(end_date, "%Y/%m/%d").strftime("%d-%m")
    f_name = ''
    if start_date == end_date:
        f_name = f_name + f'Theo dõi hợp đồng Margin {end_date}.xlsx'
    else:
        f_name = f_name + f'Theo dõi hợp đồng Margin từ {start_date} đến {end_date}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder, folder_name, period, f_name),
        engine='xlsxwriter',
        engine_kwargs={'options': {'nan_inf_to_errors': True}}
    )
    workbook = writer.book

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet sheet ---------------------
    # Format
    header_left_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'bg_color': '#92D050'
        }
    )
    header_left_yellow_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'bg_color': '#FFFF00'
        }
    )
    header_left_wrap_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'left',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_size': 11,
            'font_name': 'Times New Roman',
            'bg_color': '#92D050'
        }
    )
    header_right_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'bg_color': '#92D050'
        }
    )
    header_center_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'bg_color': '#92D050'
        }
    )
    num_right_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Times New Roman'
        }
    )
    text_left_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Times New Roman'
        }
    )
    date_mo_hop_dong_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'num_format': 'dd/mm/yyyy'
        }
    )
    so_hd_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Times New Roman',
        }
    )
    headers = [
        'No',
        '',
        'Branch',
        'A/c No.',
        'Customer Name',
        'Ngày mở hợp đồng trên hệ thống',
        'Số HĐ',
        'Contract Number (Số HĐ)',
        'Ngày Chi nhánh giao',
        'Người giao',
        'Ngày nhận',
        'Người nhận',
        'Ngày trả chi nhánh',
        'Người trả',
    ]

    #  ----------------- Viết Sheet theo dõi Hợp đồng Margin -----------------
    # --------- sheet1 ---------
    sheet1 = workbook.add_worksheet('Sheet1')
    # header dùng để ghi đè
    no = 'No'
    ngay_mo_hd = 'Ngày mở hợp đồng trên hệ thống'
    so_hd = 'Số HĐ'
    ngay_cn_giao = 'Ngày Chi nhánh giao'
    nguoi_giao = 'Người giao'
    nguoi_tra = 'Người trả'

    # Set Column Width and Row Height
    sheet1.set_column('A:A', 5.22)
    sheet1.set_column('B:B', 25.22)
    sheet1.set_column('C:C', 10.67)
    sheet1.set_column('D:D', 14.11)
    sheet1.set_column('E:E', 42.33)
    sheet1.set_column('F:F', 19.11)
    sheet1.set_column('G:G', 15.22)
    sheet1.set_column('H:L', 31.67)
    sheet1.set_column('M:M', 25.33)
    sheet1.set_column('N:N', 20.56)
    sheet1.set_row(0, 27.6)

    # Xử lý Dataframe
    branch = {
        'Cầu Giấy': 'CG',
        'Chi nhánh Q7': 'Q7',
        'Hà Nội': 'HN',
        'Hải Phòng': 'HP',
        'Institutional Business 01': 'INB',
        'Internet Broker': 'IB',
        'Phú Mỹ Hưng': 'PMH',
        'Quận 3': 'Q3',
        'Tân Bình': 'TB',
        'Thanh Xuân': 'HNTX',
        'Phòng Quản lý tài khoản - 01': 'AMD1',
        'Phòng Quản lý tài khoản - 02': 'AMD2',
        'Institutional Business 02': 'INB2',
        'Chi nhánh Quận 1': 'Q1N'
    }
    tracking_margin_query = tracking_margin_query.sort_values(by=['open_date'])  # sort values of open_date column
    for idx in tracking_margin_query.index:
        branch_query = tracking_margin_query.loc[idx, 'branch_name']
        contract_number_margin = tracking_margin_query.loc[idx, 'contract_number_margin']
        if branch_query in branch.keys():
            tracking_margin_query.loc[idx, 'branch_abbre'] = branch.get(branch_query)
        char_in_contract = ['-', '_']
        for char in char_in_contract:
            if char in contract_number_margin:
                tracking_margin_query.loc[idx, 'so_hd'] = contract_number_margin.split(char)[1].replace('M', '')

    # write column and row
    sheet1.write_row('A1', headers, header_left_format)
    sheet1.write('A1', no, header_right_format)
    sheet1.write('F1', ngay_mo_hd, header_left_wrap_format)
    sheet1.write('G1', so_hd, header_center_format)
    sheet1.write('I1', ngay_cn_giao, header_left_yellow_format)
    sheet1.write('J1', nguoi_giao, header_left_yellow_format)
    sheet1.write('N1', nguoi_tra, header_center_format)

    sheet1.write_column(
        'A2',
        [int(i) for i in np.arange(tracking_margin_query.shape[0]) + 1],
        num_right_format
    )
    sheet1.write_column(
        'B2',
        tracking_margin_query['branch_name'],
        text_left_format
    )
    sheet1.write_column(
        'C2',
        tracking_margin_query['branch_abbre'],
        text_left_format
    )
    sheet1.write_column(
        'D2',
        tracking_margin_query['account_code'],
        text_left_format
    )
    sheet1.write_column(
        'E2',
        tracking_margin_query['customer_name'],
        text_left_format
    )
    sheet1.write_column(
        'F2',
        tracking_margin_query['open_date'],
        date_mo_hop_dong_format
    )
    sheet1.write_column(
        'G2',
        tracking_margin_query['so_hd'],
        so_hd_format
    )
    sheet1.write_column(
        'H2',
        tracking_margin_query['contract_number_margin'],
        text_left_format
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

