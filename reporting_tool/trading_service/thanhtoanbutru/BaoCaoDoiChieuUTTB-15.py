"""
    1. daily
    2. table: trading_record, cash_balance
"""
from reporting_tool.trading_service.thanhtoanbutru import *


def run(
        periodicity: str,
        start_date: str,  # 2021-11-18
        end_date: str,    # 2021-11-22
        run_time=None,
):
    start = time.time()
    info = get_info(periodicity, run_time)
    # start_date = info['start_date']
    # end_date = info['end_date']
    period = info['period']
    folder_name = info['folder_name']
    date_character = ['/', '-', '.']

    # create folder
    if not os.path.isdir(join(dept_folder, folder_name)):  # dept_folder from import
        os.mkdir(join(dept_folder, folder_name))
    if not os.path.isdir(join(dept_folder, folder_name, period)):
        os.mkdir((join(dept_folder, folder_name, period)))

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viết Query ---------------------
    doi_chieu_uttb_query = pd.read_sql(
        f"""
            SELECT
            trading_record.date,
            MAX(relationship.account_code) as account_code,
            MAX(relationship.sub_account) as sub_account,
            MAX(account.customer_name) as customer_name,
            SUM(trading_record.value) as value,
            SUM(cash_balance.decrease) as decrease
            from trading_record
            LEFT JOIN relationship ON relationship.sub_account = trading_record.sub_account
            LEFT JOIN account ON account.account_code = relationship.account_code
            LEFT JOIN cash_balance ON cash_balance.sub_account = relationship.sub_account
            WHERE trading_record.date BETWEEN '{start_date}' AND '{end_date}'
            AND cash_balance.date BETWEEN '{start_date}' AND '{end_date}'
            AND relationship.date = '{end_date}'
            AND trading_record.type_of_order = 'S'
            AND cash_balance.remark like N'%Hoàn trả ƯTTB%'
            GROUP BY trading_record.date, relationship.sub_account
            ORDER BY trading_record.date ASC
        """,
        connect_DWH_CoSo,
    )

    # ------------- Xử lý Dataframe -------------
    def converter(x):
        if x == dt.datetime.strptime(start_date, '%Y-%m-%d'):
            result = 'T-2'
        elif x == dt.datetime.strptime(end_date, '%Y-%m-%d'):
            result = 'T0'
        else:
            result = 'T-1'
        return result

    # thay đổi giá trị trong cột date thành các giá trị trong hàm converter
    doi_chieu_uttb_query.date = doi_chieu_uttb_query.date.map(converter)
    # thêm các cột account_code, customer_name vào df doi_chieu_uttb_groupby
    account = doi_chieu_uttb_query[['sub_account', 'account_code']].drop_duplicates().set_index('sub_account').squeeze()
    customer = doi_chieu_uttb_query[['sub_account', 'customer_name']].drop_duplicates().set_index(
        'sub_account').squeeze()
    # tạo dataframe mới bằng cách group by df lấy từ sql
    uttb_groupby = doi_chieu_uttb_query.groupby(['date', 'sub_account'])[['value', 'decrease']].sum()
    df_uttb = uttb_groupby.unstack(level=0)
    df_uttb.fillna(0, inplace=True)  # thay các dòng có giá trị NaN bằng 0
    df_uttb = df_uttb.drop([('decrease', 'T-1'), ('decrease', 'T-2')], axis=1)
    df_uttb['account_code'] = account
    df_uttb['customer_name'] = customer

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File Excel ---------------------
    # Write file BÁO CÁO ĐỐI CHIẾU LÃI TIỀN GỬI PHÁT SINH TRÊN TÀI KHOẢN KHÁCH HÀNG
    for date_char in date_character:
        if date_char in start_date and date_char in end_date:
            start_date = start_date.replace(date_char, '/')
            end_date = end_date.replace(date_char, '/')
    footer_date = bdate(end_date, 1).split('-')
    start_date = dt.datetime.strptime(start_date, "%Y/%m/%d").strftime("%d-%m")
    end_date = dt.datetime.strptime(end_date, "%Y/%m/%d").strftime("%d-%m")
    f_name = f'Báo cáo đối chiếu UTTB {end_date}.xlsx'
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
    stt_row_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    stt_col_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'top',
            'font_size': 10,
            'font_name': 'Times New Roman'
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
    text_left_wrap_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'text_wrap': True,
            'valign': 'top',
            'font_size': 10,
            'font_name': 'Times New Roman'
        }
    )
    money_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'top',
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
            'font_size': 12,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    headers = [
        'STT',
        'Số tài khoản',
        'Số tiểu khoản',
        'Tên khách hàng',
        'Giá trị tiền bán T-2',
        'Giá trị tiền bán T-1',
        'Giá trị tiền bán T0',
        'Tiền Hoàn trả UTTB T0',
        'Tổng giá trị tiền bán có thể ứng',
        'Tiền đã ứng',
        'Tiền còn có thể ứng',
        'Bất Thường'
    ]

    companyAddress = 'Tầng 3, CR3-03A, 109 Tôn Dật Tiên, phường Tân Phú, Quận 7, Thành phố Hồ Chí Minh'
    sheet_title_name = 'BÁO CÁO ĐỐI CHIẾU UTTB'
    sub_title_name = f'Ngày {end_date}'

    # --------- sheet BAO CAO CAN LAM ---------
    sheet_bao_cao_can_lam = workbook.add_worksheet('BAO CAO CAN LAM')

    # Insert phu hung picture
    sheet_bao_cao_can_lam.insert_image('A1', './img/phu_hung.png', {'x_scale': 0.65, 'y_scale': 0.71})

    # Set Column Width and Row Height
    sheet_bao_cao_can_lam.set_column('A:A', 6.29)
    sheet_bao_cao_can_lam.set_column('B:B', 12.14)
    sheet_bao_cao_can_lam.set_column('C:C', 13)
    sheet_bao_cao_can_lam.set_column('D:D', 30)
    sheet_bao_cao_can_lam.set_column('E:F', 11.86)
    sheet_bao_cao_can_lam.set_column('G:G', 14.71)
    sheet_bao_cao_can_lam.set_column('H:H', 14.29)
    sheet_bao_cao_can_lam.set_column('I:I', 15.71)
    sheet_bao_cao_can_lam.set_column('J:J', 11.71)
    sheet_bao_cao_can_lam.set_column('K:K', 14.86)
    sheet_bao_cao_can_lam.set_column('L:L', 15.86)

    # merge row
    sheet_bao_cao_can_lam.merge_range('C1:L1', CompanyName, company_name_format)
    sheet_bao_cao_can_lam.merge_range('C2:L2', companyAddress, company_format)
    sheet_bao_cao_can_lam.merge_range('C3:L3', CompanyPhoneNumber, company_format)
    sheet_bao_cao_can_lam.merge_range('A7:L7', sheet_title_name, sheet_title_format)
    sheet_bao_cao_can_lam.merge_range('A8:L8', sub_title_name, from_to_format)
    sum_start_row = df_uttb.shape[0] + 13
    sheet_bao_cao_can_lam.merge_range(
        f'A{sum_start_row}:D{sum_start_row}',
        'Tổng',
        headers_format
    )
    footer_start_row = sum_start_row + 2
    sheet_bao_cao_can_lam.merge_range(
        f'J{footer_start_row}:L{footer_start_row}',
        f'Ngày {footer_date[2]} tháng {footer_date[1]} năm {footer_date[0]}',
        footer_dmy_format
    )
    sheet_bao_cao_can_lam.merge_range(
        f'J{footer_start_row + 1}:L{footer_start_row + 1}',
        'Người duyệt',
        footer_text_format
    )
    sheet_bao_cao_can_lam.merge_range(
        f'A{footer_start_row + 1}:C{footer_start_row + 1}',
        'Người lập',
        footer_text_format
    )

    # write row, column
    sheet_bao_cao_can_lam.write_row(
        'A4',
        [''] * len(headers),
        empty_row_format
    )
    sheet_bao_cao_can_lam.write_row('A11', headers, headers_format)
    sheet_bao_cao_can_lam.write_row(
        'A12',
        [f'({i})' for i in np.arange(len(headers)) + 1],
        stt_row_format,
    )
    sheet_bao_cao_can_lam.write_row(
        f'E{sum_start_row}',
        [''] * 8,
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'A13',
        [int(i) for i in np.arange(df_uttb.shape[0]) + 1],
        stt_col_format
    )
    sheet_bao_cao_can_lam.write_column(
        'B13',
        df_uttb['account_code'],
        text_left_format
    )
    sheet_bao_cao_can_lam.write_column(
        'C13',
        df_uttb.index,
        text_left_format
    )
    sheet_bao_cao_can_lam.write_column(
        'D13',
        df_uttb['customer_name'],
        text_left_wrap_format
    )
    sheet_bao_cao_can_lam.write_column(
        'E13',
        df_uttb[('value', 'T-2')],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'F13',
        df_uttb[('value', 'T-1')],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'G13',
        df_uttb[('value', 'T0')],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'H13',
        df_uttb['decrease', 'T0'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'I13',
        [''] * df_uttb.shape[0],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'J13',
        [''] * df_uttb.shape[0],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'K13',
        [''] * df_uttb.shape[0],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'L13',
        [''] * df_uttb.shape[0],
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
