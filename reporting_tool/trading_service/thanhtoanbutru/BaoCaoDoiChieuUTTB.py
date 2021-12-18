"""
    1. daily
    2. table: trading_record, cash_balance
    3. Tiền Hoàn trả UTTB T0: mã giao dịch 8851
    4. Tiền đã ứng: mã giao dịch - 1153
    5. có thể dùng bảng RCI0015 thay cho RCF1002 (chị Tuyết)
    6. Số tiền UTTB nhận và phí UTTB truyền vào cột ngày nào (T-2, T-1 hay T0) dựa vào
    phần ngày đầu tiên nằm trong diễn giải
    7. Phần tử Ngày đầu tiên luôn luôn bé hơn phần tử ngày thứ 2, ko có trường hợp ngược lại (Chị Thu Anh)
    8. Giá trị của cột Số tiền UTTB nhận trong RCF1002 chưa thực hiện trừ phí UTTB
    9. TK bị lệch:
                                                Số tiền UTTB KH     Phí UTTB ngày T0
                                                đã nhận ngày T0
        022C336888	0117002702	NGUYỄN THỊ TÀI	14,008,058,561	    7,786,581
                                Giá trị tiền bán T-2
                                                    Thuế Cổ Tức
        022C019179	0104001345	PHẠM THANH TÙNG     20,000
        (đã check, do file mẫu sai)

"""
from reporting_tool.trading_service.thanhtoanbutru import *

# DONE
def run(
        run_time=None,
):
    start = time.time()
    info = get_info('daily',run_time)
    period = info['period']
    t0_date = info['end_date'].replace('/','-')
    t1_date = bdate(t0_date,-1)
    t2_date = bdate(t0_date,-2)
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir((join(dept_folder,folder_name,period)))

    ###################################################
    ###################################################
    ###################################################

    target = pd.read_pickle('df154.pickle')

    t0_wildcard = f'{t0_date[-2:]}_{t0_date[5:7]}_{t0_date[:4]}'
    t1_wildcard = f'{t1_date[-2:]}_{t1_date[5:7]}_{t1_date[:4]}'
    t2_wildcard = f'{t2_date[-2:]}_{t2_date[5:7]}_{t2_date[:4]}'
    table = pd.read_sql(
        f"""
        WITH 
        [i] AS (
            SELECT
                [relationship].[sub_account],
                [relationship].[account_code],
                [account].[customer_name]
                FROM
                    [relationship]
                LEFT JOIN
                    [account]
                ON
                    [relationship].[account_code] = [account].[account_code]
                WHERE
                    [relationship].[date] = '{t0_date}'
        ),
        [o] AS (
            SELECT
                CASE 
                    WHEN [trading_record].[date] = '{t2_date}' THEN 't2'
                    WHEN [trading_record].[date] = '{t1_date}' THEN 't1'
                    WHEN [trading_record].[date] = '{t0_date}' THEN 't0'
                END [date],
                [trading_record].[sub_account],
                [trading_record].[value] [value],
                [trading_record].[fee] [fee],
                [trading_record].[tax_of_selling] [sell_tax],
                [trading_record].[tax_of_share_dividend] [dividend_tax]
            FROM 
                [trading_record]
            WHERE
                [trading_record].[date] BETWEEN '{t2_date}' AND '{t0_date}'
            AND 
                [trading_record].[type_of_order] = 'S'
        ),
        [c] AS (
            SELECT 
                CASE 
                    WHEN [cash_balance].[date] = '{t2_date}' THEN 't2'
                    WHEN [cash_balance].[date] = '{t1_date}' THEN 't1'
                    WHEN [cash_balance].[date] = '{t0_date}' THEN 't0'
                END AS [date],
                [cash_balance].[sub_account],
                [cash_balance].[transaction_id],
                [cash_balance].[remark],
                [cash_balance].[increase],
                [cash_balance].[decrease]
            FROM
                [cash_balance]
            WHERE
                [cash_balance].[date] BETWEEN '{t2_date}' AND '{t0_date}'
            AND 
                [cash_balance].[transaction_id] IN ('1153','8851')
        )
        SELECT 
            [i].[sub_account],
            [i].[account_code],
            [i].[customer_name],
            [o_t2].[value] [value_t2],
            [o_t2].[fee] [fee_t2],
            [o_t2].[sell_tax] [sell_tax_t2],
            [o_t2].[dividend_tax] [dividend_tax_t2],
            [o_t1].[value] [value_t1],
            [o_t1].[fee] [fee_t1],
            [o_t1].[sell_tax] [sell_tax_t1],
            [o_t1].[dividend_tax] [dividend_tax_t1],
            [o_t0].[value] [value_t0],
            [o_t0].[fee] [fee_t0],
            [o_t0].[sell_tax] [sell_tax_t0],
            [o_t0].[dividend_tax] [dividend_tax_t0],
            [p_t0].[payback_uttb_t0] [payback_uttb_t0],
            [d_t2].[advanced_amount_t2] [advanced_amount_t2],
            [d_t2].[advanced_fee_t2] [advanced_fee_t2],
            [d_t1].[advanced_amount_t1] [advanced_amount_t1],
            [d_t1].[advanced_fee_t1] [advanced_fee_t1],
            [d_t0].[advanced_amount_t0] [advanced_amount_t0],
            [d_t0].[advanced_fee_t0] [advanced_fee_t0]
        FROM
            [i]
        
        FULL OUTER JOIN (
            SELECT
                [o].[sub_account],
                SUM([o].[value]) [value],
                SUM([o].[fee]) [fee],
                SUM([o].[sell_tax]) [sell_tax],
                SUM([o].[dividend_tax]) [dividend_tax]
            FROM
                [o]
            WHERE
                [o].[date] = 't2'
            GROUP BY
                [o].[sub_account]
        ) [o_t2]
        ON 
            [i].[sub_account] = [o_t2].[sub_account]
        
        FULL OUTER JOIN (
            SELECT
                [o].[sub_account],
                SUM([o].[value]) [value],
                SUM([o].[fee]) [fee],
                SUM([o].[sell_tax]) [sell_tax],
                SUM([o].[dividend_tax]) [dividend_tax]
            FROM
                [o]
            WHERE
                [o].[date] = 't1'
            GROUP BY 
                [o].[sub_account]
        ) [o_t1]
        ON 
            [i].[sub_account] = [o_t1].[sub_account]
        
        FULL OUTER JOIN (
            SELECT
                [o].[sub_account],
                SUM([o].[value]) [value],
                SUM([o].[fee]) [fee],
                SUM([o].[sell_tax]) [sell_tax],
                SUM([o].[dividend_tax]) [dividend_tax]
            FROM
                [o]
            WHERE
                [o].[date] = 't0'
            GROUP BY 
                [o].[sub_account]
        ) [o_t0]
        ON 
            [i].[sub_account] = [o_t0].[sub_account]
        
        FULL OUTER JOIN (
            SELECT
                [c].[sub_account],
                SUM([c].[decrease]) [payback_uttb_t0]
            FROM
                [c]
            WHERE
                [c].[transaction_id] = '8851' AND [c].[date] = 't0'
            GROUP BY
                [c].[sub_account]
        ) [p_t0]
        ON
            [i].[sub_account] = [p_t0].[sub_account]
        
        FULL OUTER JOIN (
            SELECT
                [c].[sub_account],
                SUM(ISNULL([c].[increase],0)-ISNULL([c].[decrease],0)) [advanced_amount_t2],
                SUM([c].[decrease]) [advanced_fee_t2]
            FROM 
                [c]
            WHERE
                [c].[transaction_id] = '1153' AND [c].[remark] LIKE N'Phí%GD {t2_wildcard}%'
            GROUP BY
                [c].[sub_account]
        ) [d_t2]
        ON
            [i].[sub_account] = [d_t2].[sub_account]
        
        FULL OUTER JOIN (
            SELECT
                [c].[sub_account],
                SUM(ISNULL([c].[increase],0)-ISNULL([c].[decrease],0)) [advanced_amount_t1],
                SUM([c].[decrease]) [advanced_fee_t1]
            FROM 
                [c]
            WHERE
                [c].[transaction_id] = '1153' AND [c].[remark] LIKE N'Phí%GD {t1_wildcard}%'
            GROUP BY
                [c].[sub_account]
        ) [d_t1]
        ON
            [i].[sub_account] = [d_t1].[sub_account]
        
        FULL OUTER JOIN (
            SELECT
                [c].[sub_account],
                SUM(ISNULL([c].[increase],0)-ISNULL([c].[decrease],0)) [advanced_amount_t0],
                SUM([c].[decrease]) [advanced_fee_t0]
            FROM
                [c]
            WHERE
                [c].[transaction_id] = '1153' AND [c].[remark] LIKE N'Phí%GD {t0_wildcard}%'
            GROUP BY
                [c].[sub_account]
        ) [d_t0]
        ON
            [i].[sub_account] = [d_t0].[sub_account]
        """,
        connect_DWH_CoSo,
        index_col='sub_account',
    ).dropna(thresh=3).fillna(0)

    able_to_advance_t1 = table['value_t1'] - table['fee_t1'] - table['sell_tax_t1'] - table['dividend_tax_t1']
    able_to_advance_t0 = table['value_t0'] - table['fee_t0'] - table['sell_tax_t0'] - table['dividend_tax_t0']
    table['available_to_advance'] = able_to_advance_t1 + able_to_advance_t0

    advanced_amount = table['advanced_amount_t1'] + table['advanced_fee_t1'] + table['advanced_amount_t0'] + table['advanced_fee_t0']
    table['remaining_advance'] = table['available_to_advance'] - advanced_amount

    able_to_advance_t2 = table['value_t2'] - table['fee_t2'] - table['sell_tax_t2'] - table['dividend_tax_t2']
    check_1 = table['payback_uttb_t0'] > able_to_advance_t2
    check_2 = advanced_amount > table['available_to_advance']
    check_3 = table['remaining_advance'] < 0
    total_check = check_1 | check_2 | check_3
    table.loc[total_check,'check'] = 'Bất thường'
    table.sort_values('check',ascending=False,inplace=True)
    table['check'].fillna('',inplace=True)
    count_abonormal = total_check.sum()

    table = table.reset_index()[[
        'account_code',
        'sub_account',
        'customer_name',
        'value_t2',
        'fee_t2',
        'sell_tax_t2',
        'dividend_tax_t2',
        'value_t1',
        'fee_t1',
        'sell_tax_t1',
        'dividend_tax_t1',
        'value_t0',
        'fee_t0',
        'sell_tax_t0',
        'dividend_tax_t0',
        'payback_uttb_t0',
        'available_to_advance',
        'advanced_amount_t2',
        'advanced_fee_t2',
        'advanced_amount_t1',
        'advanced_fee_t1',
        'advanced_amount_t0',
        'advanced_fee_t0',
        'remaining_advance',
        'check',
    ]]

    ###################################################
    ###################################################
    ###################################################

    file_name = f'Báo cáo Đối chiếu UTTB {t0_date.replace("-",".")}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options': {'nan_inf_to_errors': True}}
    )
    workbook = writer.book
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
    company_info_format = workbook.add_format(
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
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
        }
    )
    sheet_title_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 14,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    sub_title_date_format = workbook.add_format(
        {
            'italic': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 10,
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
            'font_size': 10,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    text_center_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman'
        }
    )
    text_left_format = workbook.add_format( # for customer name only
        {
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman'
        }
    )
    money_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )
    sum_money_format = workbook.add_format(
        {
            'bold': True,
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )
    footer_dmy_format = workbook.add_format(
        {
            'italic': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
        }
    )
    footer_text_format = workbook.add_format(
        {
            'bold': True,
            'italic': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    sheet_title_name = 'BÁO CÁO ĐỐI CHIẾU UTTB'
    eod_sub = dt.datetime.strptime(t0_date,"%Y-%m-%d").strftime("%d/%m/%Y")
    sub_title_name = f'Ngày {eod_sub}'
    worksheet = workbook.add_worksheet(f'{period}')
    worksheet.hide_gridlines(option=2)
    worksheet.insert_image('A1','./img/phs_logo.png',{'x_scale':0.65,'y_scale': 0.71})
    worksheet.set_column('A:A',6)
    worksheet.set_column('B:C',13)
    worksheet.set_column('D:D',23)
    worksheet.set_column('E:Y',12)
    worksheet.set_column('Z:Z',10)

    # merge row
    worksheet.merge_range('D1:M1',CompanyName,company_name_format)
    worksheet.merge_range('D2:M2',CompanyAddress,company_info_format)
    worksheet.merge_range('D3:M3',CompanyPhoneNumber,company_info_format)
    worksheet.merge_range('A7:Z7',sheet_title_name,sheet_title_format)
    worksheet.merge_range('A8:Z8',sub_title_name,sub_title_date_format)
    worksheet.merge_range('A10:A11','STT',headers_format)
    worksheet.merge_range('B10:B11','Số tài khoản',headers_format)
    worksheet.merge_range('C10:C11','Số tiểu khoản',headers_format)
    worksheet.merge_range('D10:D11','Tên khách hàng',headers_format)
    worksheet.merge_range('E10:H11','Giá trị tiền bán T-2',headers_format)
    worksheet.merge_range('I10:L10','Giá trị tiền bán T-1',headers_format)
    worksheet.merge_range('M10:P10','Giá trị tiền bán T0',headers_format)
    worksheet.merge_range('Q10:Q11','Tiền Hoàn trả UTTB T0',headers_format)
    worksheet.merge_range('R10:R11','Tổng giá trị tiền bán có thể ứng',headers_format)
    worksheet.merge_range('S10:X10','Tiền đã ứng',headers_format)
    worksheet.merge_range('Y10:Y11','Tổng tiền còn có thể ứng',headers_format)
    worksheet.merge_range('Z10:Z11','Bất Thường',headers_format)
    sub_headers_1 = [
        'Giá trị tiền bán',
        'Phí bán',
        'Thuế bán',
        'Thuế Cổ Tức'
    ]
    sub_headers_2 = [
        'Số tiền UTTB KH đã nhận ngày T-2',
        'Phí UTTB ngày T-2',
        'Số tiền UTTB KH đã nhận ngày T-1',
        'Phí UTTB ngày T-1',
        'Số tiền UTTB KH đã nhận ngày T0',
        'Phí UTTB ngày T0'
    ]
    worksheet.write_row('E11',sub_headers_1*3,headers_format)
    worksheet.write_row('S11',sub_headers_2,headers_format)
    sum_start_row = table.shape[0]+13
    worksheet.merge_range(f'A{sum_start_row}:D{sum_start_row}','Tổng',headers_format)
    footer_start_row = sum_start_row+2

    footer_date = bdate(t0_date,1).split('-')
    worksheet.merge_range(
        f'U{footer_start_row}:Z{footer_start_row}',
        f'Ngày {footer_date[2]} tháng {footer_date[1]} năm {footer_date[0]}',
        footer_dmy_format
    )
    worksheet.merge_range(
        f'U{footer_start_row+1}:Z{footer_start_row+1}',
        'Người duyệt',
        footer_text_format
    )
    worksheet.merge_range(
        f'A{footer_start_row+1}:C{footer_start_row+1}',
        'Người lập',
        footer_text_format
    )
    worksheet.write_row('A4',['']*26,empty_row_format)
    stt_headers = [
        '(1)','(2)','(3)','(4)',
        '(5a)','(5b)','(5c)','(5d)',
        '(6a)','(6b)','(6c)','(6d)',
        '(7a)','(7b)','(7c)','(7d)',
        '(8)','(9)',
        '(10a)','(10b)','(10c)','(10d)','(10e)','(10f)',
        '(11)','(12)',
    ]
    worksheet.write_row('A12',stt_headers,headers_format)
    worksheet.write_column('A13',np.arange(table.shape[0])+1,text_center_format)
    for col,col_name in enumerate(table.columns):
        if col_name in ['account_code','sub_account','check']:
            fmt = text_center_format
        elif col_name == 'customer_name':
            fmt = text_left_format
        else:
            fmt = money_format
        worksheet.write_column(12,col+1,table[col_name],fmt)
    worksheet.write_row(f'E{sum_start_row}',table.iloc[:,3:-1].sum(),sum_money_format)
    worksheet.write(f'Z{sum_start_row}',count_abonormal,sum_money_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py', '')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time() - start, 1)}s')