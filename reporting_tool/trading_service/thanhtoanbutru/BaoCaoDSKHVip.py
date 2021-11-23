from reporting_tool.trading_service.thanhtoanbutru import *


def run(
        periodicity: str,
        start_date: str,  # 2015-01-01
        end_date: str,    # 2021-10-31
        run_time=None,
):
    start = time.time()
    info = get_info(periodicity, run_time)
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
    vip_phs_query = pd.read_sql(
        f"""
            SELECT
            DISTINCT
            customer_information.sub_account,
            account.account_code, 
            account.customer_name, 
            branch.branch_name, 
            account.date_of_birth, 
            customer_information.contract_type,
            customer_information_change.date_of_change,
            customer_information_change.time_of_change,
            customer_information_change.date_of_approval,
            customer_information_change.time_of_approval, 
            broker.broker_id, broker.broker_name, 
            customer_information_change.change_content
            FROM customer_information
            LEFT JOIN relationship ON customer_information.sub_account=relationship.sub_account
            LEFT JOIN account ON relationship.account_code=account.account_code
            LEFT JOIN broker ON relationship.broker_id=broker.broker_id
            LEFT JOIN branch ON relationship.branch_id=branch.branch_id
            LEFT JOIN customer_information_change ON customer_information_change.account_code=relationship.account_code
            WHERE (customer_information.contract_type LIKE N'%GOLD%' 
            OR customer_information.contract_type LIKE N'%SILV%' 
            OR customer_information.contract_type LIKE N'%VIPCN%')
            AND relationship.date='2021-10-31'
            AND (customer_information_change.date_of_change BETWEEN '{start_date}' AND '{end_date}')
            AND customer_information_change.change_content = 'Loai hinh hop dong'
        """,
        connect_DWH_CoSo
    )
    vip_phs_query = vip_phs_query.sort_values(by=['date_of_birth'])

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File ---------------------
    # Write file excel Bao cao doi chieu file ngan hang
    file_name = f'DH KH VIP SINH NHAT.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder, folder_name, period, file_name),
        engine='xlsxwriter',
        engine_kwargs={'options': {'nan_inf_to_errors': True}}
    )

    workbook = writer.book

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet sheet ---------------------
    # Format
    sheet_title_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 18
        }
    )
    sub_title_format = workbook.add_format(
        {
            'bold': True,
            'italic': True,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 11,
            'color': '#FF0000',
            'bg_color': '#FFFFCC'
        }
    )
    header_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 11,
            'bg_color': '#65FF65'
        }
    )
    no_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 11,
        }
    )
    header_format_tong_theo_cn = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 11,
            'bg_color': '#92D050'
        }
    )
    list_customer_vip_row_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 11,
            'color': '#FF0000',
            'bg_color': '#FFFF00',
        }
    )
    ngay_hieu_luc_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 11,
            'color': '#FF0000',
        }
    )
    list_customer_vip_col_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 11,
            'bg_color': '#FFCCFF',
        }
    )
    date_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'num_format': 'dd/mm/yyyy',
            'font_name': 'Times New Roman',
            'font_size': 11
        }
    )
    text_left_format = workbook.add_format(
        {
            'border': 1,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'left',
            'font_name': 'Times New Roman',
            'font_size': 11
        }
    )
    num_left_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 11
        }
    )
    num_center_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 11
        }
    )
    text_right_format = workbook.add_format(
        {
            'border': 1,
            'valign': 'vcenter',
            'align': 'right',
            'font_name': 'Times New Roman',
            'font_size': 11
        }
    )
    hide_column_num_right_format = workbook.add_format(
        {
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 11
        }
    )
    headers = [
        'CTT',
        'No',
        'Account',
        'Name',
        'Location',
        'Birthday',
        'LIST CUSTOMER VIP',
        'Ngày hiệu lực đầu tiên',
        'Approve date',
        'Review date',
        'Mã môi giới quản lý tài khoản',
        'Tên môi giới quản lý tài khoản',
        'Note',
    ]
    headers_tong_theo_cn = [
        'Location',
        'GOLD PHS',
        'SILV PHS',
        'VIP BRANCH',
        'SUM',
        'NOTE'
    ]

    #  Viết Sheet VIP PHS
    sheet_vip_phs = workbook.add_worksheet('VIP PHS')

    # Content of sheet
    sheet_title_name = 'UPDATED LIST OF COMPANY VIP TO 25.07.2021'
    sub_title_name = 'Chỉ tặng quà sinh nhật cho Gold PHS & Silv PHS'

    # Groupby and condition
    vip_phs_groupby_H = (vip_phs_query.groupby('account_code')['date_of_change'].min())
    vip_phs_query['min_date'] = vip_phs_query['account_code'].map(vip_phs_groupby_H)
    vip_phs_groupby_N = (vip_phs_query.groupby('account_code')['time_of_change'].min())
    vip_phs_query['min_time'] = vip_phs_query['account_code'].map(vip_phs_groupby_N)
    vip_phs_groupby_H_approval = (vip_phs_query.groupby('account_code')['date_of_approval'].max())
    vip_phs_query['max_date'] = vip_phs_query['account_code'].map(vip_phs_groupby_H_approval)
    vip_phs_groupby_N_approval = (vip_phs_query.groupby('account_code')['time_of_approval'].max())
    vip_phs_query['max_time'] = vip_phs_query['account_code'].map(vip_phs_groupby_N_approval)

    # loc duplicate row
    vip_phs_query = vip_phs_query.loc[vip_phs_query['time_of_approval'] == vip_phs_query['max_time']]
    vip_phs_query = vip_phs_query.drop_duplicates('customer_name')
    # xuất cột contract_type
    mapper = vip_phs_query['contract_type'].apply(
        lambda x: 'SILV PHS' if 'SILV' in x else ('GOLD PHS' if 'GOLD' in x else 'VIP Branch'))
    vip_phs_query['contract_type'] = mapper

    def last_day_of_month(date_value):
        return date_value.replace(day=calendar.monthrange(date_value.year, date_value.month)[1])

    # điều kiện phân biệt giữa SILV, GOLD và VIP Branch
    # for idx in vip_phs_query.index:
    #     contract_type = vip_phs_query.loc[idx, 'contract_type']
    #     date_present = dt.now()
    #     if contract_type == 'SILV PHS' or contract_type == 'GOLD PHS':
    #         if 1 <= date_present.month < 6:
    #             last_day_of_June = last_day_of_month(dt(date_present.year, 6, 1).date())
    #             vip_phs_query.loc[idx, 'review_date'] = last_day_of_June.strftime("%d/%m/%Y")
    #         elif 6 <= date_present.month < 12:
    #             last_day_of_Dec = last_day_of_month(dt(date_present.year, 12, 1).date())
    #             vip_phs_query.loc[idx, 'review_date'] = last_day_of_Dec.strftime("%d/%m/%Y")
    #         else:
    #             last_day_of_June_next_year = last_day_of_month(dt(date_present.year + 1, 6, 1).date())
    #             vip_phs_query.loc[idx, 'review_date'] = last_day_of_June_next_year.strftime("%d/%m/%Y")
    #     elif contract_type == 'VIP Branch':
    #         if 1 <= date_present.month < 3:
    #             last_day_of_Mar = last_day_of_month(dt(date_present.year, 3, 1).date())
    #             vip_phs_query.loc[idx, 'review_date'] = last_day_of_Mar.strftime("%d/%m/%Y")
    #         elif 3 <= date_present.month < 6:
    #             last_day_of_June = last_day_of_month(dt(date_present.year, 6, 1).date())
    #             vip_phs_query.loc[idx, 'review_date'] = last_day_of_June.strftime("%d/%m/%Y")
    #         elif 6 <= date_present.month < 9:
    #             last_day_of_Sep = last_day_of_month(dt(date_present.year + 1, 9, 1).date())
    #             vip_phs_query.loc[idx, 'review_date'] = last_day_of_Sep.strftime("%d/%m/%Y")
    #         else:
    #             last_day_of_Mar_next_year = last_day_of_month(dt(date_present.year + 1, 3, 1).date())
    #             vip_phs_query.loc[idx, 'review_date'] = last_day_of_Mar_next_year.strftime("%d/%m/%Y")

    # ko phân biệt SILV, GOLD và VIP Branch
    for idx in vip_phs_query.index:
        date_present = dt.datetime.now()
        if 1 <= date_present.month < 6:
            last_day_of_June = last_day_of_month(dt.datetime(date_present.year, 6, 1).date())
            vip_phs_query.loc[idx, 'review_date'] = last_day_of_June.strftime("%d/%m/%Y")
        elif 6 <= date_present.month < 12:
            last_day_of_Dec = last_day_of_month(dt.datetime(date_present.year, 12, 1).date())
            vip_phs_query.loc[idx, 'review_date'] = last_day_of_Dec.strftime("%d/%m/%Y")
        else:
            last_day_of_June_next_year = last_day_of_month(dt.datetime(date_present.year + 1, 6, 1).date())
            vip_phs_query.loc[idx, 'review_date'] = last_day_of_June_next_year.strftime("%d/%m/%Y")

    # Set Column Width and Row Height
    sheet_vip_phs.set_column('A:A', 0)
    sheet_vip_phs.set_column('B:B', 3.56)
    sheet_vip_phs.set_column('C:C', 10.71)
    sheet_vip_phs.set_column('D:D', 30.57)
    sheet_vip_phs.set_column('E:E', 26.29)
    sheet_vip_phs.set_column('F:F', 9.43)
    sheet_vip_phs.set_column('G:G', 12.86)
    sheet_vip_phs.set_column('H:H', 11)
    sheet_vip_phs.set_column('I:I', 13)
    sheet_vip_phs.set_column('J:J', 11.86)
    sheet_vip_phs.set_column('K:K', 11.43)
    sheet_vip_phs.set_column('L:L', 65.86)
    sheet_vip_phs.set_column('M:M', 13.78)
    sheet_vip_phs.set_column('N:N', 21.71)
    sheet_vip_phs.set_column('O:S', 13.71)

    sheet_vip_phs.set_row(0, 22.8)
    sheet_vip_phs.set_row(1, 39)
    sheet_vip_phs.set_row(2, 21.8)
    sheet_vip_phs.set_row(3, 71.3)

    # merge row
    sheet_vip_phs.merge_range('B1:K1', sheet_title_name, sheet_title_format)
    sheet_vip_phs.merge_range('C2:M2', sub_title_name, sub_title_format)
    sheet_vip_phs.merge_range('C3:K3', '')

    # write row and write column
    sheet_vip_phs.write_row('A4', headers, header_format)
    sheet_vip_phs.write('B4', 'No', no_format)
    sheet_vip_phs.write('G4', 'LIST CUSTOMER VIP', list_customer_vip_row_format)
    sheet_vip_phs.write('H4', 'Ngày hiệu lực đầu tiên', ngay_hieu_luc_format)
    sheet_vip_phs.write_column(
        'A5',
        [1] * vip_phs_query.shape[0],
        hide_column_num_right_format
    )
    sheet_vip_phs.write_column(
        'B5',
        [f'{i}' for i in np.arange(vip_phs_query.shape[0]) + 1],
        text_right_format
    )
    sheet_vip_phs.write_column(
        'C5',
        vip_phs_query['account_code'],
        num_left_format
    )
    sheet_vip_phs.write_column(
        'D5',
        vip_phs_query['customer_name'],
        text_left_format
    )
    sheet_vip_phs.write_column(
        'E5',
        vip_phs_query['branch_name'],
        text_left_format
    )
    sheet_vip_phs.write_column(
        'F5',
        vip_phs_query['date_of_birth'],
        date_format
    )
    sheet_vip_phs.write_column(
        'G5',
        vip_phs_query['contract_type'],
        list_customer_vip_col_format
    )
    sheet_vip_phs.write_column(
        'H5',
        vip_phs_query['min_date'],
        date_format
    )
    sheet_vip_phs.write_column(
        'I5',
        vip_phs_query['max_date'],
        date_format
    )
    sheet_vip_phs.write_column(
        'J5',
        vip_phs_query['review_date'],
        date_format
    )
    sheet_vip_phs.write_column(
        'K5',
        vip_phs_query['broker_id'],
        num_center_format
    )
    sheet_vip_phs.write_column(
        'L5',
        vip_phs_query['broker_name'],
        num_left_format
    )
    sheet_vip_phs.write_column(
        'M5',
        [''] * vip_phs_query.shape[0],
        text_left_format
    )

    #  Viết Sheet TONG THEO CN
    # Format
    sum_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'valign': 'vbottom',
            'align': 'right',
            'font_size': 11,
            'font_name': 'Calibri'
        }
    )
    sum_title_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'valign': 'vbottom',
            'align': 'left',
            'font_size': 11,
            'font_name': 'Calibri'
        }
    )
    location_format = workbook.add_format(
        {
            'border': 1,
            'valign': 'vbottom',
            'align': 'left',
            'font_size': 11,
            'font_name': 'Calibri'
        }
    )
    num_format = workbook.add_format(
        {
            'border': 1,
            'valign': 'vbottom',
            'align': 'right',
            'font_size': 11,
            'font_name': 'Calibri'
        }
    )
    note_format = workbook.add_format(
        {
            'border': 1,
            'valign': 'vbottom',
            'align': 'left',
            'font_size': 11,
            'font_name': 'Calibri'
        }
    )
    sheet_tong_theo_cn = workbook.add_worksheet('TONG THEO CN')
    # Set Column Width and Row Height
    sheet_tong_theo_cn.set_column('A:A', 28.22)
    sheet_tong_theo_cn.set_column('B:B', 13.78)
    sheet_tong_theo_cn.set_column('C:C', 15.22)
    sheet_tong_theo_cn.set_column('D:F', 17.78)

    vip_phs_query['CTT'] = [1] * vip_phs_query.shape[0]
    tong_theo_cn_groupby = vip_phs_query.groupby(['branch_name', 'contract_type'])['CTT'].sum().unstack()
    tong_theo_cn_groupby['GOLD PHS'].fillna(0, inplace=True)
    tong_theo_cn_groupby['SILV PHS'].fillna(0, inplace=True)
    tong_theo_cn_groupby['VIP Branch'].fillna(0, inplace=True)

    sheet_tong_theo_cn.write_row('A1', headers_tong_theo_cn, header_format_tong_theo_cn)
    sheet_tong_theo_cn.write_column(
        'A2',
        tong_theo_cn_groupby.index,
        location_format
    )
    sheet_tong_theo_cn.write_column(
        'B2',
        tong_theo_cn_groupby['GOLD PHS'],
        num_format
    )
    sheet_tong_theo_cn.write_column(
        'C2',
        tong_theo_cn_groupby['SILV PHS'],
        num_format
    )
    sheet_tong_theo_cn.write_column(
        'D2',
        tong_theo_cn_groupby['VIP Branch'],
        num_format
    )
    sum_row = tong_theo_cn_groupby['GOLD PHS'] + tong_theo_cn_groupby['SILV PHS'] + tong_theo_cn_groupby['VIP Branch']
    sheet_tong_theo_cn.write_column(
        'E2',
        sum_row,
        sum_format
    )
    note_col = [''] * tong_theo_cn_groupby.shape[0]
    sheet_tong_theo_cn.write_column(
        'F2',
        note_col,
        note_format,
    )

    start_sum_row = tong_theo_cn_groupby.shape[0] + 2
    sheet_tong_theo_cn.write(
        f'A{start_sum_row}',
        'SUM',
        sum_title_format
    )
    sheet_tong_theo_cn.write(
        f'B{start_sum_row}',
        str(int(tong_theo_cn_groupby['GOLD PHS'].sum())),
        sum_format
    )
    sheet_tong_theo_cn.write(
        f'C{start_sum_row}',
        str(int(tong_theo_cn_groupby['SILV PHS'].sum())),
        sum_format
    )
    sheet_tong_theo_cn.write(
        f'D{start_sum_row}',
        str(int(tong_theo_cn_groupby['VIP Branch'].sum())),
        sum_format
    )
    sheet_tong_theo_cn.write(
        f'E{start_sum_row}',
        str(int(sum_row.values.sum())),
        sum_format
    )
    sheet_tong_theo_cn.write(
        f'F{start_sum_row}',
        0,
        sum_format
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
