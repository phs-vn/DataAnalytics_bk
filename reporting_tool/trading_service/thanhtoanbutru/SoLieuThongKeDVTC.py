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
import pandas as pd

from reporting_tool.trading_service.thanhtoanbutru import *


def run(
        periodicity: str,
        start_date: str,  # 2021-11-01
        end_date: str,  # 2021-11-30
        run_time=None,
):
    start = time.time()
    info = get_info(periodicity, run_time)
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
    day_dau_ki = bdate(start_date, -1)
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
            SELECT sub_account, contract_type, status
            FROM customer_information
            WHERE (contract_type NOT LIKE N'Tự doanh' AND contract_type NOT LIKE N'Thường%')
            AND (status = 'A' OR status = 'B')
            ORDER BY contract_type
        """,
        connect_DWH_CoSo,
    )
    mapper = query_total_account['contract_type'].apply(
        lambda x: 'VIPCN - T1'
        if ('MR 1' and 'VIPCN') in x
        else ('SILV - T1' if ('SILV' and 'MR 1') in x
              else ('GOLD - T0' if ('GOLD' and 'MR 0') in x
                    else ('GOLD - T2' if ('GOLD' and 'MR 2') in x
                          else 'NOR'
                          )
                    )
              )
    )
    mapper_2 = query_total_account['contract_type'].apply(
        lambda x: 'VIP CN - DP +3'
        if ('VIPCN' and 'DP 3') in x
        else ('SILV DP+4' if ('SILV' and 'DP 4') in x
              else ('GOLD DP +5' if ('GOLD' and 'DP 5') in x
                    else ('Normal DP +2' if ('NOR' and 'DP 2') in x
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

    loc_TK_VIP = query_total_account.groupby('TK VIP').size()
    loc_TK_DP = query_total_account.groupby('TK DP').size()
    loc_all_TK = query_total_account.groupby('all_TK').size()

    df_tk_vip = pd.DataFrame(loc_TK_VIP,
                             index=["VIPCN - T1", "SILV - T1", "GOLD - T0", "GOLD - T2", "NOR"],
                             columns=['num_account']
                             ).drop('NOR')
    df_tk_dp = pd.DataFrame(loc_TK_DP,
                            index=["Normal DP +2", "VIP CN - DP +3", "SILV DP+4", "GOLD DP +5", "NotDP"],
                            columns=['num_account']
                            ).drop('NotDP')
    df_tk_all = pd.DataFrame(loc_all_TK,
                             index=["VIPCN", "SILVER", "GOLDEN", "NOR"],
                             columns=['num_account']
                             ).drop('NOR')

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
    cuoi_ky_num_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'color': '#FF0000',
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
    num_right_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
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
    sheet_DVTC.write_column('B5', 'MR', TK_margin_empty_format)

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
    sheet_DVTC.write(f'A{sum_row + 2}', TK_margin, text_left_wrap_text_format)
    sheet_DVTC.write(f'B{sum_row + 2}', '', TK_margin_empty_format)
    sheet_DVTC.write(f'C{sum_row + 2}', '', num_right_format)
    sheet_DVTC.write(f'D{sum_row + 2}', '', num_right_format)
    sheet_DVTC.write(f'E{sum_row + 2}', query_total_account.shape[0], cuoi_ky_num_format)
    sheet_DVTC.merge_range(f'A{sum_row + 3}:A{sum_row + 6}', TK_vip, text_center_wrap_text_format)
    sheet_DVTC.write_column(f'B{sum_row + 3}', df_tk_vip.index, text_left_format)
    sheet_DVTC.write_column(f'E{sum_row + 3}', df_tk_vip['num_account'], cuoi_ky_num_format)
    sheet_DVTC.merge_range(f'A{sum_row + 7}:A{sum_row + 10}', TK_use_DP, text_center_wrap_text_format)
    sheet_DVTC.write_column(f'B{sum_row + 7}', df_tk_dp.index, text_left_format)
    sheet_DVTC.write_column(f'E{sum_row + 7}', df_tk_dp['num_account'], cuoi_ky_num_format)
    sheet_DVTC.write(f'A{sum_row + 11}', total_account_vipcn, text_left_format)
    sheet_DVTC.write(f'A{sum_row + 12}', total_account_silv, text_left_format)
    sheet_DVTC.write(f'A{sum_row + 13}', total_account_gold, text_left_format)
    sheet_DVTC.write_column(f'E{sum_row + 11}', df_tk_all['num_account'], cuoi_ky_num_format)
    sheet_DVTC.write_column(f'B{sum_row + 11}', [''] * 3, text_center_format)
    sheet_DVTC.write_column(
        f'G3',
        [''] * (sum_row + df_tk_vip.shape[0] + df_tk_dp.shape[0] + df_tk_all.shape[0]),
        text_center_format
    )
    sheet_DVTC.write(f'G{sum_row + 5}', 'HPX', text_left_format)
    sheet_DVTC.write_column(
        f'H3',
        [''] * (sum_row + df_tk_vip.shape[0] + df_tk_dp.shape[0] + df_tk_all.shape[0]),
        num_right_format
    )
    sheet_DVTC.write_column(
        f'I3',
        [''] * (sum_row + df_tk_vip.shape[0] + df_tk_dp.shape[0] + df_tk_all.shape[0]),
        num_right_format
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
