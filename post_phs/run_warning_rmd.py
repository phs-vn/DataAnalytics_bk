import sys
sys.path.extend([r'C:\\Users\\hiepdang\\PycharmProjects\\DataAnalytics',
                 r'C:/Users/hiepdang/PycharmProjects/DataAnalytics'])

from post_phs import *

now = dt.date.today()
run_time = now.strftime('%Y%m%d')

destination_dir = r'\\192.168.10.101\phs-storge-2018' \
                  r'\RiskManagementDept\RMD_Data' \
                  r'\Luu tru van ban\RMC Meeting 2018' \
                  r'\00. Meeting minutes\Warning List'

file_name = f'{run_time}.xlsx'

file_path = fr'{destination_dir}\{file_name}'

# consecutively scan on 2 exchanges
hose_table = post.risk_alert(True,'HOSE','all')
hnx_table = post.risk_alert(True,'HNX','all')

# export to excel
result = pd.concat([hose_table,hnx_table]).sort_values(
    [
        'Consecutive Floor Days',
        'Consecutive Illiquidity Days'
    ],
    ascending=False
)
if not result.empty:
    writer = pd.ExcelWriter(
        file_path,
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book
    worksheet = workbook.add_worksheet('Warning List')
    worksheet.hide_gridlines(option=2)
    # format
    header_format = workbook.add_format({
        'bold': True,
        'font_name': 'Arial',
        'border': 1,
        'text_wrap': True,
        'align': 'center',
        'valign': 'bottom',
    })
    text_format = workbook.add_format({
        'font_name': 'Arial',
        'border': 1,
    })
    n_floor_format = workbook.add_format({
        'font_name': 'Arial',
        'num_format': '0',
        'border': 1,
    })
    n_illiquidity_format = workbook.add_format({
        'font_name': 'Arial',
        'num_format': '0',
        'border': 1,
    })
    volume_change_format = workbook.add_format({
        'font_name': 'Arial',
        'num_format': '0%',
        'border': 1,
    })
    dos_format = workbook.add_format({
        'font_name': 'Arial',
        'num_format': '0.00',
        'border': 1,
    })
    n_floor_stress_format = workbook.add_format({
        'font_name': 'Arial',
        'num_format': '0',
        'bg_color': '#FF8787',
        'font_color': '#640000',
        'border': 1,
    })
    n_illiquidity_stress_format = workbook.add_format({
        'font_name': 'Arial',
        'num_format': '0',
        'bg_color': '#FF8787',
        'font_color': '#640000',
        'border': 1,
    })
    volume_change_stress_format = workbook.add_format({
        'font_name': 'Arial',
        'num_format': '0%',
        'bg_color': '#FF8787',
        'font_color': '#640000',
        'border': 1,
    })
    dos_stress_format = workbook.add_format({
        'font_name': 'Arial',
        'num_format': '0.00',
        'bg_color': '#FF8787',
        'font_color': '#640000',
        'border': 1,
    })
    worksheet.set_column('A:A',7)
    worksheet.set_column('B:F',14)
    # write
    worksheet.write('A1',result.index.name,header_format)
    worksheet.write_row('B1',result.columns,header_format)
    worksheet.write_column('A2',result.index,text_format)
    worksheet.write_column('B2',result['Exchange'],text_format)
    for row in range(result.shape[0]):
        n_floor = result.loc[result.index[row],'Consecutive Floor Days']
        n_illiquidity = result.loc[result.index[row],'Consecutive Illiquidity Days']
        volume_change = result.loc[result.index[row],'Volume Change']
        dos_3m = result.loc[result.index[row],'Days of Sale (3M Avg)']
        if n_floor >= 2:
            fmt = n_floor_stress_format
        else:
            fmt = n_floor_format
        worksheet.write(row+1,2,n_floor,fmt)
        if n_illiquidity >= 1.5:
            fmt = n_illiquidity_stress_format
        else:
            fmt = n_illiquidity_format
        worksheet.write(row+1,3,n_illiquidity,fmt)
        if volume_change <= -0.3:
            fmt = volume_change_stress_format
        else:
            fmt = volume_change_format
        worksheet.write(row+1,4,volume_change,fmt)
        if dos_3m >= 1.5:
            fmt = dos_stress_format
        else:
            fmt = dos_format
        worksheet.write(row+1,5,dos_3m,fmt)

    writer.close()

# Note: Trong trường hợp không có cổ phiếu giảm sàn 2 phiên sẽ không xuất file
