from request.stock import *
from warning import warning_list_RMD

def run():

  now = dt.date.today()
  run_time = now.strftime('%Y%m%d')

  destination_dir = r'\\192.168.10.101\phs-storge-2018' \
                    r'\RiskManagementDept\RMD_Data' \
                    r'\Luu tru van ban\RMC Meeting 2018' \
                    r'\00. Meeting minutes\Warning List'

  file_name = f'{run_time}.xlsx'
  file_path = fr'{destination_dir}\{file_name}'

  # consecutively scan on 2 exchanges
  hose_table = warning_list_RMD.run(True,'HOSE','all')
  hnx_table = warning_list_RMD.run(True,'HNX','all')

  # export to excel
  result = pd.concat([hose_table,hnx_table]).sort_values(
    [
      'Consecutive Floor Days',
      '1M Illiquidity Days'
    ],
    ascending=False,
    ignore_index=True,
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
      'bold'     :True,
      'font_name':'Times New Roman',
      'font_size':8,
      'border'   :1,
      'text_wrap':True,
      'align'    :'center',
      'valign'   :'vcenter',
    })
    text_format = workbook.add_format({
      'font_name':'Times New Roman',
      'font_size':11,
      'align'    :'center',
      'border'   :1,
    })
    rate_format = workbook.add_format({
      'font_name' :'Times New Roman',
      'font_size' :11,
      'num_format':'0',
      'border'    :1,
      'align'     :'center',
    })
    price_format = workbook.add_format({
      'font_name' :'Times New Roman',
      'font_size' :11,
      'num_format':'#,##0',
      'border'    :1,
    })
    volume_format = workbook.add_format({
      'font_name' :'Times New Roman',
      'font_size' :11,
      'num_format':'#,##0',
      'border'    :1,
    })
    total_room_format = workbook.add_format({
      'font_name' :'Times New Roman',
      'font_size' :11,
      'font_color':'#FF0000',
      'bold'      :True,
      'num_format':'#,##0',
      'border'    :1,
    })
    integer_format = workbook.add_format({
      'font_name' :'Times New Roman',
      'font_size' :11,
      'num_format':'0',
      'border'    :1,
    })
    n_floor_format = workbook.add_format({
      'font_name' :'Times New Roman',
      'font_size' :11,
      'num_format':'0',
      'border'    :1,
    })
    n_illiquidity_format = workbook.add_format({
      'font_name' :'Times New Roman',
      'font_size' :11,
      'num_format':'0',
      'border'    :1,
    })
    pct_volume_format = workbook.add_format({
      'font_name' :'Times New Roman',
      'font_size' :11,
      'num_format':'0%',
      'border'    :1,
    })
    float_format = workbook.add_format({
      'font_name' :'Times New Roman',
      'font_size' :11,
      'num_format':'0.00',
      'border'    :1,
    })
    n_floor_stress_format = workbook.add_format({
      'font_name' :'Times New Roman',
      'font_size' :11,
      'num_format':'0',
      'bg_color'  :'#FF8787',
      'font_color':'#640000',
      'border'    :1,
    })
    n_illiquidity_stress_format = workbook.add_format({
      'font_name' :'Times New Roman',
      'font_size' :11,
      'num_format':'0',
      'bg_color'  :'#FF8787',
      'font_color':'#640000',
      'border'    :1,
    })
    pct_volume_1m_stress_format = workbook.add_format({
      'font_name' :'Times New Roman',
      'font_size' :11,
      'num_format':'0%',
      'bg_color'  :'#FF8787',
      'font_color':'#640000',
      'border'    :1,
    })
    room_on_avg_vol_3m_stress_format = workbook.add_format({
      'font_name' :'Times New Roman',
      'font_size' :11,
      'num_format':'0.00',
      'bg_color'  :'#FF8787',
      'font_color':'#640000',
      'border'    :1,
    })
    worksheet.set_column('A:A',3)
    worksheet.set_column('B:E',6)
    worksheet.set_column('C:C',7)  # overwrite previous line
    worksheet.set_column('F:I',11)
    worksheet.set_column('J:Q',10)
    # write
    worksheet.write('A1','No.',header_format)
    worksheet.write_row('B1',result.columns,header_format)
    worksheet.write_column('A2',np.arange(result.shape[0])+1,text_format)
    worksheet.write_column('B2',result['Stock'],text_format)
    worksheet.write_column('C2',result['Exchange'],text_format)
    worksheet.write_column('D2',result['Tỷ lệ vay KQ (%)'],integer_format)
    worksheet.write_column('E2',result['Tỷ lệ vay TC (%)'],integer_format)
    worksheet.write_column('F2',result['Giá vay / Giá TSĐB tối đa (VND)'],price_format)
    worksheet.write_column('G2',result['General Room'],volume_format)
    worksheet.write_column('H2',result['Special Room'],volume_format)
    worksheet.write_column('I2',result['Total Room'],total_room_format)
    worksheet.write_column('K2',result['Last day Volume'],volume_format)
    worksheet.write_column('M2',result['1W Avg. Volume'],volume_format)
    worksheet.write_column('N2',result['1M Avg. Volume'],volume_format)
    worksheet.write_column('O2',result['3M Avg. Volume'],volume_format)
    alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    for row in range(result.shape[0]):
      idx = result.index[row]
      n_floor = result.loc[idx,'Consecutive Floor Days']
      volume_change_1m = result.loc[idx,'% Last day volume / 1M Avg.']
      room_on_avg_vol_3m = result.loc[idx,'Approved Room / Avg. Liquidity 3 months']
      n_illiquidity = result.loc[idx,'1M Illiquidity Days']
      if n_floor >= 2:
        fmt = n_floor_stress_format
      else:
        fmt = n_floor_format
      worksheet.write(row+1,alphabet.index('J'),n_floor,fmt)  # cot J
      if volume_change_1m <= -0.3:
        fmt = pct_volume_1m_stress_format
      else:
        fmt = pct_volume_format
      worksheet.write(row+1,alphabet.index('L'),volume_change_1m,fmt)  # cot L
      if room_on_avg_vol_3m >= 1.5:
        fmt = room_on_avg_vol_3m_stress_format
      else:
        fmt = float_format
      worksheet.write(row+1,alphabet.index('P'),room_on_avg_vol_3m,fmt)  # cot P
      if n_illiquidity >= 1.5:
        fmt = n_illiquidity_stress_format
      else:
        fmt = n_illiquidity_format
      worksheet.write(row+1,alphabet.index('Q'),n_illiquidity,fmt)  # cot Q

    writer.close()

  # Note: Trong trường hợp không có cổ phiếu thỏa điều kiện sẽ không xuất file
