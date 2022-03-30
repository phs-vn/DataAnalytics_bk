from automation.trading_service.thanhtoanbutru import *


def run(
        run_time=None
):
    start = time.time()
    info = get_info('daily', run_time)
    period = info['period']
    folder_name = info['folder_name']

    if not os.path.isdir(join(dept_folder, folder_name, period)):
        os.mkdir((join(dept_folder, folder_name, period)))

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viết Query ---------------------
    query_TA = pd.read_sql(
        f"""
            SELECT
                [trading_record].[date],
                [relationship].[account_code],
                [trading_record].[sub_account],
                [trading_record].[value],
                [trading_record].[fee]
            FROM
                [trading_record]
            LEFT JOIN [relationship]
            ON [relationship].[sub_account] = [trading_record].[sub_account]
            WHERE
                [trading_record].[date] BETWEEN '2022-01-01' AND '2022-03-29'
                AND [relationship].[date] = '2022-03-29'
            ORDER BY
                [date]
        """,
        connect_DWH_CoSo
    )

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File Excel ---------------------
    # Write file ROD0040
    f_name = f'ROD0040 (01012022-29032022).xlsx'
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
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 12,
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
    headers = [
        'STT',
        'Ngày',
        'Số tài khoản',
        'Giá trị giao dịch',
        'Phí giao dịch',
        'Số tiểu khoản'
    ]

    # --------- sheet BAO CAO CAN LAM ---------
    sheet_bao_cao_can_lam = workbook.add_worksheet('BAO CAO CAN LAM')
    # Set Column Width and Row Height
    sheet_bao_cao_can_lam.set_column('A:A', 9)
    sheet_bao_cao_can_lam.set_column('B:B', 12)
    sheet_bao_cao_can_lam.set_column('C:C', 12)
    sheet_bao_cao_can_lam.set_column('D:D', 15)
    sheet_bao_cao_can_lam.set_column('E:E', 15)
    sheet_bao_cao_can_lam.set_column('F:F', 15)
    sheet_bao_cao_can_lam.write_row('A1', headers, headers_format)
    sheet_bao_cao_can_lam.write_column(
        'A2',
        np.arange(query_TA.shape[0]) + 1,
        stt_col_format
    )
    sheet_bao_cao_can_lam.write_column(
        'B2',
        query_TA['date'],
        date_format
    )
    sheet_bao_cao_can_lam.write_column(
        'C2',
        query_TA['account_code'],
        text_left_format
    )
    sheet_bao_cao_can_lam.write_column(
        'D2',
        query_TA['value'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'E2',
        query_TA['fee'],
        money_format
    )
    sheet_bao_cao_can_lam.write_column(
        'F2',
        query_TA['sub_account'],
        money_format
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



