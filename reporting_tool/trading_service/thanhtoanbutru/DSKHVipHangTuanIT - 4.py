"""
    1. Báo cáo Tuần - weekly
    2. Table: VCF0051 - customer_information
    3. Bỏ phần lấy mail từ SW
    4. Bên SS thường chạy vào T6 hoặc Chủ nhật (Từ khi chi nhánh báo batch cuối ngày tới chủ nhật)
    5. TK bị lệch: 022C040943   PHẠM THỊ HỒNG LIÊN
        a. Trên báo cáo mẫu
            022C040943	0101003744	PHẠM THỊ HỒNG LIÊN  MR - KHTN NOR 60% - Margin.PIA 13.5%
            022C040943	0101003743	PHẠM THỊ HỒNG LIÊN	Thường - KHTN - PIA 13.5%
        b. Trên Database SQL
            022C040943	0101003744	PHẠM THỊ HỒNG LIÊN	MR - KHTN SILV 65% - OD 0,15% - Margin.PIA 12.5% - MR 1 (KgTC)
            022C040943	0101003743	PHẠM THỊ HỒNG LIÊN	Thường - KHTN - OD 0,15% - PIA 12.5%

"""

from reporting_tool.trading_service.thanhtoanbutru import *


def run(
        periodicity: str,
        date: str,
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
    dskh_vip_query = pd.read_sql(
        f"""
            SELECT
            relationship.account_code as account_code, 
            account.customer_name as customer_name, 
            customer_information.contract_type as contract_type
            FROM relationship
            LEFT JOIN customer_information 
            ON customer_information.sub_account=relationship.sub_account
            LEFT JOIN account 
            ON account.account_code = relationship.account_code
            WHERE relationship.date='{date}'
            AND (contract_type LIKE N'%GOLD%' 
            OR contract_type LIKE N'%SILV%' 
            OR contract_type LIKE N'%VIPCN%')
            ORDER BY contract_type DESC
        """,
        connect_DWH_CoSo
    )
    # xuất cột contract_type
    mapper = dskh_vip_query['contract_type'].apply(
        lambda x: 'SILVER PHS' if 'SILV' in x else ('GOLD PHS' if 'GOLD' in x else 'VIP BRANCH'))
    dskh_vip_query['contract_type'] = mapper

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File ---------------------
    # Write file excel Bao cao doi chieu file ngan hang
    # điều kiện ngày
    date_character = ['/', '-', '.']
    for date_char in date_character:
        if date_char in date:
            date = date.replace(date_char, '/')

    weekend = dt.datetime.strptime(date, "%Y/%m/%d").strftime("%d.%m.%Y")

    file_name = f'Báo cáo DS KH VIP hàng tuần cho IT ngày {weekend}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder, folder_name, period, file_name),
        engine='xlsxwriter',
        engine_kwargs={'options': {'nan_inf_to_errors': True}}
    )

    workbook = writer.book

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viết sheet ---------------------
    # Format
    sheet_title_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'valign': 'vbottom',
            'font_name': 'Calibri',
            'font_size': 20
        }
    )
    header_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Calibri',
            'font_size': 11,
            'bg_color': '#2F75B5',
            'color': '#FFFFFF'
        }
    )
    no_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Calibri',
            'font_size': 11,
        }
    )
    text_left_format = workbook.add_format(
        {
            'border': 1,
            'valign': 'vcenter',
            'align': 'left',
            'font_name': 'Calibri',
            'font_size': 11
        }
    )
    text_center_format = workbook.add_format(
        {
            'border': 1,
            'valign': 'vcenter',
            'align': 'center',
            'font_name': 'Calibri',
            'font_size': 11
        }
    )
    headers = [
        'NO',
        'ACCOUNT',
        'NAME',
        'LIST CUSTOMER VIP'
    ]

    #  Name of Sheet VCF0051
    sheet_vip_phs = workbook.add_worksheet('VCF0051')
    sheet_vip_phs.hide_gridlines(option=2)

    # Content in sheet
    sheet_title_name = f'UPDATED LIST OF COMPANY VIP TO {weekend}'

    # Set Column Width and Row Height
    sheet_vip_phs.set_column('A:A', 11.57)
    sheet_vip_phs.set_column('B:B', 13.14)
    sheet_vip_phs.set_column('C:C', 34.14)
    sheet_vip_phs.set_column('D:D', 24.14)

    sheet_vip_phs.set_row(1, 26.25)
    sheet_vip_phs.set_row(4, 25.5)

    # merge row
    sheet_vip_phs.merge_range('A2:D2', sheet_title_name, sheet_title_format)

    # write row and write column
    sheet_vip_phs.write_row('A5', headers, header_format)
    sheet_vip_phs.write_column(
        'A6',
        [f'{i}' for i in np.arange(dskh_vip_query.shape[0]) + 1],
        no_format
    )
    sheet_vip_phs.write_column(
        'B6',
        dskh_vip_query['account_code'],
        text_left_format
    )
    sheet_vip_phs.write_column(
        'C6',
        dskh_vip_query['customer_name'],
        text_left_format
    )
    sheet_vip_phs.write_column(
        'D6',
        dskh_vip_query['contract_type'],
        text_center_format
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
