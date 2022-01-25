"""
1. monthly
2. qui tắc đầu kỳ, cuối kỳ
    Tháng 11:
        + số đầu kỳ: lấy RLN0006 cuối tháng 10 (là số đầu kỳ tháng 11)
        + số cuối kỳ : lấy RLN0006 cuối tháng 11
3. hỏi về cách lấy dữ liệu đầu kỳ (cuối kỳ tháng trước) trong bảng VCF0051
4. Dữ liệu bị lệch khá nhiều so với dữ liệu trên flex
5. Cách để thêm danh sách KH nợ xấu vào báo cáo (Set cứng hay sao)
"""
from reporting.trading_service.thanhtoanbutru import *


def run(
        start_date: str,  # 2021-11-01
        end_date: str,  # 2021-11-30
        run_time=None,
):
    start = time.time()
    info = get_info('monthly', run_time)
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

    # --------------------- Viết Query và xử lý dataframe ---------------------
    # query RLN0006
    day_dau_ki = bdate(start_date, -1)  # tính đầu kỳ tháng 11
    query_sum_du_no_goc = pd.read_sql(
        f"""
            SELECT
            date,
            CASE margin_outstanding.type WHEN N'Ứng trước cổ tức' THEN 'UTCT'
                                         WHEN N'Trả chậm' THEN 'DP'
                                         WHEN N'Margin' THEN 'MR'
                                         WHEN N'Bảo lãnh' THEN 'BL'
                                         ELSE margin_outstanding.type
                                         END AS type,
            SUM(principal_outstanding) AS principal_outstanding
            FROM margin_outstanding
            WHERE date = '{day_dau_ki}' or date = '{end_date}'
            GROUP BY date, type
            ORDER BY date, CASE WHEN type = N'Ứng trước cổ tức' THEN '1'
                                WHEN type = N'Trả chậm' THEN '2'
                                WHEN type = N'Margin' THEN '3'
                                ELSE type END ASC
        """,
        connect_DWH_CoSo,
    )
    sum_du_no_goc_dau_ky = query_sum_du_no_goc.loc[query_sum_du_no_goc['date'] == f'{day_dau_ki}']
    sum_du_no_goc_cuoi_ky = query_sum_du_no_goc.loc[query_sum_du_no_goc['date'] == f'{end_date}']

    # Tính tổng dư nợ gốc trong kỳ
    trong_ky = sum_du_no_goc_cuoi_ky['principal_outstanding'].values - sum_du_no_goc_dau_ky[
        'principal_outstanding'].values
    du_no_goc_trong_ky = pd.DataFrame(trong_ky, columns=['values'], index=['UTCT', 'DP', 'MR', 'BL', 'Cầm cố '])

    # query VCF0051
    query_total_account = pd.read_sql(
        f"""
            SELECT relationship.date, customer_information.sub_account, contract_type, status
            FROM relationship
            LEFT JOIN customer_information 
            ON customer_information.sub_account = relationship.sub_account 
            LEFT JOIN account
            ON account.account_code = relationship.account_code 
            WHERE (relationship.date = '{day_dau_ki}' OR relationship.date = '{end_date}')
            AND customer_information.status IN ('A','B')
            AND (contract_type NOT LIKE N'Tự doanh' and contract_type NOT LIKE N'Thường%')
            ORDER BY date, contract_type
        """,
        connect_DWH_CoSo,
    )

    mapper = query_total_account['contract_type'].apply(
        lambda x: 'VIPCN - T1'
        if ('MR 1' in x and 'VIPCN' in x)
        else ('SILV - T1' if ('SILV' in x and 'MR 1' in x)
              else ('GOLD - T0' if ('GOLD' in x and 'MR 0' in x)
                    else ('GOLD - T2' if ('GOLD'  in x and 'MR 2' in x)
                          else 'NOR'
                          )
                    )
              )
    )
    mapper_2 = query_total_account['contract_type'].apply(
        lambda x: 'VIP CN - DP +3'
        if ('VIPCN' in x and 'DP 3' in x)
        else ('SILV DP+4' if ('SILV' in x and 'DP 4' in x)
              else ('GOLD DP +5' if ('GOLD' in x and 'DP 5' in x)
                    else ('Normal DP +2' if ('NOR' in x and 'DP 2' in x)
                          else 'NotDP'
                          )
                    )
              )
    )
    mapper_3 = query_total_account['contract_type'].apply(
        lambda x: 'SILVER' if 'SILV' in x else ('GOLDEN' if 'GOLD' in x else ('VIPCN' if 'VIPCN' in x else 'NOR')))

    query_total_account['TK VIP'] = mapper
    query_total_account['TK DP'] = mapper_2
    query_total_account['all_TK'] = mapper_3

    amount_TK_dau_ky = query_total_account.loc[query_total_account['date'] == day_dau_ki]
    loc_TK_VIP_dau_ky = amount_TK_dau_ky.groupby('TK VIP').size()
    loc_TK_DP_dau_ky = amount_TK_dau_ky.groupby('TK DP').size()
    loc_all_TK_dau_ky = amount_TK_dau_ky.groupby('all_TK').size()

    df_tk_vip_dau_ky = pd.DataFrame(loc_TK_VIP_dau_ky,
                                    index=["VIPCN - T1", "SILV - T1", "GOLD - T0", "GOLD - T2", "NOR"],
                                    columns=['num_account']
                                    ).drop('NOR')
    df_tk_dp_dau_ky = pd.DataFrame(loc_TK_DP_dau_ky,
                                   index=["Normal DP +2", "VIP CN - DP +3", "SILV DP+4", "GOLD DP +5", "NotDP"],
                                   columns=['num_account']
                                   ).drop('NotDP')
    df_tk_all_dau_ky = pd.DataFrame(loc_all_TK_dau_ky,
                                    index=["VIPCN", "SILVER", "GOLDEN", "NOR"],
                                    columns=['num_account']
                                    ).drop('NOR')
    amount_TK_cuoi_ky = query_total_account.loc[query_total_account['date'] == end_date]
    loc_TK_VIP_cuoi_ky = amount_TK_cuoi_ky.groupby('TK VIP').size()
    loc_TK_DP_cuoi_ky = amount_TK_cuoi_ky.groupby('TK DP').size()
    loc_all_TK_cuoi_ky = amount_TK_cuoi_ky.groupby('all_TK').size()

    df_tk_vip_cuoi_ky = pd.DataFrame(loc_TK_VIP_cuoi_ky,
                                     index=["VIPCN - T1", "SILV - T1", "GOLD - T0", "GOLD - T2", "NOR"],
                                     columns=['num_account']
                                     ).drop('NOR')
    df_tk_dp_cuoi_ky = pd.DataFrame(loc_TK_DP_cuoi_ky,
                                    index=["Normal DP +2", "VIP CN - DP +3", "SILV DP+4", "GOLD DP +5", "NotDP"],
                                    columns=['num_account']
                                    ).drop('NotDP')
    df_tk_all_cuoi_ky = pd.DataFrame(loc_all_TK_cuoi_ky,
                                     index=["VIPCN", "SILVER", "GOLDEN", "NOR"],
                                     columns=['num_account']
                                     ).drop('NOR')

    # Query which account has branch_id = 0111
    inb_01_query = pd.read_sql(
        f"""
            select
            margin_outstanding.account_code as account_code, 
            case margin_outstanding.type
                      when N'Trả chậm' then 'DP'
                      when N'Margin' then 'MR'
                      when N'Bảo lãnh' then 'BL'
                      end as type, 
            principal_outstanding
            from margin_outstanding
            left join relationship on relationship.account_code = margin_outstanding.account_code
            left join account on account.account_code = relationship.account_code
            where margin_outstanding.date = '{end_date}'
            and relationship.date = '{end_date}'
            and relationship.branch_id = '0111'
            order by case when type = 'MR' then '1'
                          when type = 'DP' then '2'
                          when type = 'BL' then '3'
                          END ASC
        """,
        connect_DWH_CoSo
    )
    inb_01_query = inb_01_query.drop_duplicates()
    inb_groupby = inb_01_query.groupby(['type']).sum()
    sum_du_no_goc_cuoi_ky = sum_du_no_goc_cuoi_ky.set_index('type')
    sum_du_no_goc_cuoi_ky['INB-01'] = inb_groupby['principal_outstanding']
    sum_du_no_goc_cuoi_ky['ty_le'] = (
            (sum_du_no_goc_cuoi_ky['INB-01'] / sum_du_no_goc_cuoi_ky['principal_outstanding']) * 100)
    sum_du_no_goc_cuoi_ky.fillna('', inplace=True)

    # Query số lượng tài khoản margin INB-01
    query_amount_account_INB = pd.read_sql(
        f"""
            SELECT customer_information.sub_account, contract_type, customer_information.status, branch.branch_name
            FROM relationship
            LEFT JOIN customer_information 
            ON customer_information.sub_account = relationship.sub_account 
            LEFT JOIN branch
            ON branch.branch_id = relationship.branch_id
            WHERE relationship.date = '{end_date}'
            AND customer_information.status IN ('A','B')
            AND (contract_type LIKE '%MR%')
            AND branch_name = 'Institutional Business 01'
        """,
        connect_DWH_CoSo
    )

    ###################################################
    ###################################################
    ###################################################

    # --------------------- Viet File Excel ---------------------
    # Write file BÁO CÁO SỐ LIỆU THỐNG KÊ DVTC THÁNG
    for date_char in date_character:
        if date_char in start_date and date_char in end_date:
            start_date = start_date.replace(date_char, '/')
            end_date = end_date.replace(date_char, '/')
    som = dt.datetime.strptime(start_date, "%Y/%m/%d").strftime("%m-%Y")
    f_name = f'Số liệu thống kê DVTC tháng {som}.xlsx'
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
    sheet_title_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 14,
            'font_name': 'Arial',
            'bg_color': '#FFFF00',
        }
    )
    INB_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'bg_color': '#FFFF00',
            'color': '#00B050'
        }
    )
    ty_le_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
        }
    )
    num_ty_le_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'num_format': '#,##0.00'
        }
    )
    header_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'bg_color': '#FFC000'
        }
    )
    no_xau_header_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'bg_color': '#00B050'
        }
    )
    cuoi_ky_no_xau_values_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'color': '#FF0000',
            'num_format': '#,##0'
        }
    )
    normal_money_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'num_format': '#,##0'
        }
    )
    trong_ky_money_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'num_format': '#,##0.00;(#,##0.00)'
        }
    )
    trong_ky_num_account_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'num_format': '#,##0;(#,##0)'
        }
    )
    tong_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
        }
    )
    sum_money_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'num_format': '#,##0'
        }
    )
    text_left_wrap_text_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'text_wrap': True
        }
    )
    SSC_name_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'text_wrap': True,
            'bg_color': '#E0FFC1'
        }
    )
    SSC_empty_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'bg_color': '#E0FFC1'
        }
    )
    SSC_value_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'bg_color': '#E0FFC1'
        }
    )
    TK_margin_empty_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'bg_color': '#FFCCFF'
        }
    )

    text_left_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
        }
    )
    text_center_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
        }
    )
    text_center_wrap_text_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'text_wrap': True
        }
    )

    # --------- sheet BAO CAO CAN LAM ---------
    sheet_DVTC = workbook.add_worksheet('DVTC')
    sheet_title_name = 'TOÀN CÔNG TY'
    INB_name = 'INB - 01'
    ty_le = 'TỶ LỆ %'
    TK_margin = 'Tổng số lượng TK Margin (Active & Block)'
    TK_vip = 'Tổng số lượng TK VIP miễn lãi vay MR (Active & Block)'
    TK_use_DP = 'Tổng số lượng TK sử dụng hạn mức DP (Active &Block)'
    total_account_vipcn = 'Tổng số lượng TK VIPCN'
    total_account_silv = 'Tổng số lượng TK SILVER'
    total_account_gold = 'Tổng số lượng TK GOLDEN'
    header = [
        '`',
        'Loại',
        'Đầu kỳ',
        'Trong kỳ',
        'Cuối kỳ',
        'Nợ xấu',
        'Ghi chú'
    ]
    # Set Column Width and Row Height
    sheet_DVTC.set_column('A:A', 30.43)
    sheet_DVTC.set_column('B:B', 15)
    sheet_DVTC.set_column('C:C', 21.14)
    sheet_DVTC.set_column('D:D', 18.86)
    sheet_DVTC.set_column('E:E', 19.14)
    sheet_DVTC.set_column('F:F', 18)
    sheet_DVTC.set_column('G:G', 37)
    sheet_DVTC.set_column('H:H', 16.14)
    sheet_DVTC.set_column('I:I', 10.14)

    sheet_DVTC.merge_range('A1:G1', sheet_title_name, sheet_title_format)
    sheet_DVTC.merge_range('H1:H2', INB_name, INB_format)
    sheet_DVTC.merge_range('I1:I2', ty_le, ty_le_format)
    sheet_DVTC.write_row('A2', header, header_format)
    sheet_DVTC.write('F2', 'Nợ xấu', no_xau_header_format)
    sheet_DVTC.write_column('B3', sum_du_no_goc_dau_ky['type'], text_center_format)
    sheet_DVTC.write_column(
        'C3',
        sum_du_no_goc_dau_ky['principal_outstanding'],
        normal_money_format
    )
    sheet_DVTC.write_column(
        'E3',
        sum_du_no_goc_cuoi_ky['principal_outstanding'],
        cuoi_ky_no_xau_values_format
    )
    sheet_DVTC.write('B5', 'MR', TK_margin_empty_format)

    sum_row = sum_du_no_goc_dau_ky.shape[0] + 3
    sheet_DVTC.merge_range(f'A3:A{sum_row}', 'Tổng Dư nợ gốc', text_left_wrap_text_format)
    sheet_DVTC.write(f'A{sum_row}', '', text_left_wrap_text_format)
    sheet_DVTC.write(f'B{sum_row}', 'TỔNG', tong_format)
    sheet_DVTC.write(
        f'C{sum_row}',
        sum_du_no_goc_dau_ky['principal_outstanding'].sum(),
        sum_money_format
    )
    sheet_DVTC.write(
        f'D{sum_row}',
        du_no_goc_trong_ky['values'].abs().sum(),
        sum_money_format
    )
    du_no_goc_trong_ky['values'] = du_no_goc_trong_ky['values'].replace(0, '-')
    sheet_DVTC.write_column(
        'D3',
        du_no_goc_trong_ky['values'],
        trong_ky_money_format
    )
    sheet_DVTC.write(
        f'E{sum_row}',
        sum_du_no_goc_cuoi_ky['principal_outstanding'].sum(),
        sum_money_format
    )
    sheet_DVTC.write(
        f'A{sum_row + 1}',
        'Dư nợ báo cáo SSC',
        SSC_name_format
    )
    sheet_DVTC.write(f'B{sum_row + 1}', '', SSC_empty_format)
    sheet_DVTC.write(f'C{sum_row + 1}', '', SSC_value_format)
    sheet_DVTC.write(f'D{sum_row + 1}', '', SSC_value_format)
    sheet_DVTC.write(f'E{sum_row + 1}', '', SSC_value_format)
    sheet_DVTC.write(f'F{sum_row + 1}', '', SSC_value_format)

    # Tổng số lượng TK Margin (Active & Block)
    sheet_DVTC.write(f'A{sum_row + 2}', TK_margin, text_left_wrap_text_format)
    sheet_DVTC.write(f'B{sum_row + 2}', '', TK_margin_empty_format)
    sheet_DVTC.write(f'C{sum_row + 2}', amount_TK_dau_ky.shape[0], normal_money_format)
    sheet_DVTC.write(
        f'D{sum_row + 2}',
        amount_TK_cuoi_ky.shape[0] - amount_TK_dau_ky.shape[0],
        trong_ky_num_account_format
    )
    sheet_DVTC.write(f'E{sum_row + 2}', amount_TK_cuoi_ky.shape[0], cuoi_ky_no_xau_values_format)
    sheet_DVTC.write(
        f'H{sum_row + 2}',
        query_amount_account_INB.shape[0],
        cuoi_ky_no_xau_values_format
    )
    sheet_DVTC.write(
        f'I{sum_row + 2}',
        (query_amount_account_INB.shape[0]/amount_TK_cuoi_ky.shape[0]) * 100,
        num_ty_le_format
    )

    # Tổng số lượng TK VIP miễn lãi vay MR  (Active &Block)
    sheet_DVTC.merge_range(f'A{sum_row + 3}:A{sum_row + 6}', TK_vip, text_center_wrap_text_format)
    sheet_DVTC.write_column(f'B{sum_row + 3}', df_tk_vip_dau_ky.index, text_left_format)
    sheet_DVTC.write_column(f'C{sum_row + 3}', df_tk_vip_dau_ky['num_account'], normal_money_format)
    sheet_DVTC.write_column(
        f'D{sum_row + 3}',
        (df_tk_vip_cuoi_ky['num_account'] - df_tk_vip_dau_ky['num_account']).abs(),
        trong_ky_num_account_format
    )
    sheet_DVTC.write_column(f'E{sum_row + 3}', df_tk_vip_cuoi_ky['num_account'], cuoi_ky_no_xau_values_format)

    # Tổng số lượng TK sử dụng hạn mức DP  (Active &Block)
    sheet_DVTC.merge_range(f'A{sum_row + 7}:A{sum_row + 10}', TK_use_DP, text_center_wrap_text_format)
    sheet_DVTC.write_column(f'B{sum_row + 7}', df_tk_dp_dau_ky.index, text_left_format)
    sheet_DVTC.write_column(f'C{sum_row + 7}', df_tk_dp_dau_ky['num_account'], normal_money_format)
    sheet_DVTC.write_column(
        f'D{sum_row + 7}',
        df_tk_dp_cuoi_ky['num_account'] - df_tk_dp_dau_ky['num_account'],
        trong_ky_num_account_format
    )
    sheet_DVTC.write_column(f'E{sum_row + 7}', df_tk_dp_cuoi_ky['num_account'], cuoi_ky_no_xau_values_format)

    # Tổng số lượng TK VIPCN, TK SILVER, TK GOLDEN
    sheet_DVTC.write(f'A{sum_row + 11}', total_account_vipcn, text_left_format)
    sheet_DVTC.write(f'A{sum_row + 12}', total_account_silv, text_left_format)
    sheet_DVTC.write(f'A{sum_row + 13}', total_account_gold, text_left_format)
    sheet_DVTC.write_column(f'C{sum_row + 11}', df_tk_all_dau_ky['num_account'], normal_money_format)
    sheet_DVTC.write_column(
        f'D{sum_row + 11}',
        df_tk_all_cuoi_ky['num_account'] - df_tk_all_dau_ky['num_account'],
        trong_ky_num_account_format
    )
    sheet_DVTC.write_column(
        f'E{sum_row + 11}',
        df_tk_all_cuoi_ky['num_account'],
        cuoi_ky_no_xau_values_format
    )

    sheet_DVTC.write_column(f'B{sum_row + 11}', [''] * 3, text_center_format)

    # Ghi chú column
    sheet_DVTC.write_column(
        'G3',
        [''] * (sum_row + df_tk_vip_dau_ky.shape[0] + df_tk_dp_dau_ky.shape[0] + df_tk_all_dau_ky.shape[0]),
        text_center_format
    )
    sheet_DVTC.write(f'G{sum_row + 5}', 'HPX', text_left_format)

    # INB-01 column
    sheet_DVTC.write_column(
        'H3',
        sum_du_no_goc_cuoi_ky['INB-01'],
        cuoi_ky_no_xau_values_format
    )
    # Tỷ lệ column
    sheet_DVTC.write_column(
        'I3',
        sum_du_no_goc_cuoi_ky['ty_le'],
        num_ty_le_format
    )
    # DANH SÁCH KH NỢ XẤU

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py', '')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time() - start, 1)}s')
