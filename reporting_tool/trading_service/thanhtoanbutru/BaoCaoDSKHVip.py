from reporting_tool.trading_service.thanhtoanbutru import *


def run(
        periodicity: str,
        run_time=None,
):
    start = time.time()
    info = get_info(periodicity, run_time)
    start_date = info['start_date']
    end_date = info['end_date']
    period = info['period']
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder, folder_name, period)):  # dept_folder from import
        os.mkdir(join(dept_folder, folder_name, period))

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viết Query ---------------------

    vip_phs_query = pd.read_sql(
        f"""
            SELECT
            account.account_code, 
            account.customer_name, 
            branch.branch_name, 
            account.date_of_birth, 
            customer_information.contract_type,
            customer_information_change.date_of_change,
            customer_information_change.date_of_approval, 
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
            AND customer_information_change.change_content = N'Loại hình hợp đồng'
        """,
        connect_DWH_CoSo
    )

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
            'font_size': 11
        }
    )
    tang_qua_sinh_nhat_format = workbook.add_format(
        {
            'italic': True,
            'underline': True,
            'font_name': 'Times New Roman',
            'font_size': 11
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
    text_center_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
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
        'Địa chỉ thực',
        'SĐT thực'
    ]

    #  Viết Sheet VIP PHS
    sheet_vip_phs = workbook.add_worksheet('VIP PHS')

    # Content of sheet
    sheet_title_name = 'UPDATED LIST OF COMPANY VIP TO 25.07.2021'
    sub_title_name = 'Chỉ tặng quà sinh nhật cho Gold PHS & Silv PHS'
    tang_qua_sinh_nhat_name = ''
    tmp = 'tặng quà sinh nhật'
    if tmp in sub_title_name:
        tang_qua_sinh_nhat_name = tang_qua_sinh_nhat_name + tmp

    # Set Column Width and Row Height
    sheet_vip_phs.set_column('A:A', 0)
    sheet_vip_phs.set_column('B:B', 3.56)
    sheet_vip_phs.set_column('C:C', 10.71)
    sheet_vip_phs.set_column('D:D', 34.14)
    sheet_vip_phs.set_column('E:E', 15.57)
    sheet_vip_phs.set_column('F:F', 11.43)
    sheet_vip_phs.set_column('G:G', 12.86)
    sheet_vip_phs.set_column('H:H', 11)
    sheet_vip_phs.set_column('I:I', 13.14)
    sheet_vip_phs.set_column('J:J', 13.78)
    sheet_vip_phs.set_column('K:K', 16.14)
    sheet_vip_phs.set_column('L:L', 28.43)
    sheet_vip_phs.set_column('M:M', 13.78)
    sheet_vip_phs.set_column('N:N', 21.71)
    sheet_vip_phs.set_column('O:S', 13.71)
    sheet_vip_phs.set_row(0, 22.8)
    sheet_vip_phs.set_row(1, 39)
    sheet_vip_phs.set_row(2, 21.8)
    sheet_vip_phs.set_row(3, 71.3)

    sheet_vip_phs.merge_range('B1:K1', sheet_title_name, sheet_title_format)
    sheet_vip_phs.merge_range('C2:M2', sub_title_name, sub_title_format)
    sheet_vip_phs.merge_range('C3:K3', '')
    sheet_vip_phs.write_row('A4', headers, header_format)
    sheet_vip_phs.write('G4', 'LIST CUSTOMER VIP', list_customer_vip_row_format)
    sheet_vip_phs.write('H4', 'Ngày hiệu lực đầu tiên', ngay_hieu_luc_format)

    # Groupby and loc
    vip_phs_groupby = vip_phs_query.groupby(vip_phs_query['date_of_change']).agg('min')
    mapper = vip_phs_groupby['contract_type'].apply(lambda x: 'SILV PHS' if 'SILV' in x
                                                                     else ('GOLD PHS' if 'GOLD' in x else 'VIP Branch'))
    vip_phs_groupby['contract_type'] = mapper
    sheet_vip_phs.write_column(
        'A5',
        ['1'] * vip_phs_groupby.shape[0],
        hide_column_num_right_format
    )
    sheet_vip_phs.write_column(
        'B5',
        [f'{i}' for i in np.arange(vip_phs_groupby.shape[0]) + 1],
        text_right_format
    )
    sheet_vip_phs.write_column(
        'C5',
        vip_phs_groupby['account_code'],
        num_left_format
    )
    sheet_vip_phs.write_column(
        'D5',
        vip_phs_groupby['customer_name'],
        text_left_format
    )
    sheet_vip_phs.write_column(
        'E5',
        vip_phs_groupby['branch_name'],
        text_left_format
    )
    sheet_vip_phs.write_column(
        'F5',
        vip_phs_groupby['date_of_birth'],
        date_format
    )
    sheet_vip_phs.write_column(
        'G5',
        vip_phs_groupby['contract_type'],
        text_center_format
    )
    sheet_vip_phs.write_column(
        'H5',
        vip_phs_groupby['date_of_approval'],
        date_format
    )
    sheet_vip_phs.write_column(
        'I5',
        vip_phs_groupby['date_of_approval'],
        date_format
    )
    sheet_vip_phs.write_column(
        'K5',
        vip_phs_groupby['broker_id'],
        num_center_format
    )
    sheet_vip_phs.write_column(
        'L5',
        vip_phs_groupby['broker_name'],
        num_left_format
    )
    sheet_vip_phs.write_column(
        'M5',
        [''] * vip_phs_groupby.shape[0],
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
