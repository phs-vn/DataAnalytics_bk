"""
BC này không thể chạy lùi ngày
(tuy nhiên có thể viết lại để chạy lùi ngày do đã bắt đầu lưu VCF0051)
"""

from reporting.trading_service.thanhtoanbutru import *


# DONE
def run():
  start = time.time()
  info = get_info('daily',None)
  end_date = info['end_date']
  period = info['period']
  folder_name = info['folder_name']
  run_time = dt.datetime.now()

  # create folder
  if not os.path.isdir(join(dept_folder,folder_name,period)):
    os.mkdir(join(dept_folder,folder_name,period))

  ###################################################
  ###################################################
  ###################################################

  vip_phs = pd.read_sql(
    f"""
        
        """,
    connect_DWH_CoSo,
    index_col='account_code'
  )
  # Groupby and condition
  vip_phs['effective_date'] = vip_phs.groupby('account_code')['date_of_change'].min()
  vip_phs['approved_date'] = vip_phs.groupby('account_code')['date_of_approval'].max()
  condition_1 = vip_phs['date_of_change'] == vip_phs['effective_date']
  condition_2 = vip_phs['date_of_approval'] == vip_phs['approved_date']
  vip_phs = vip_phs.loc[(condition_1|condition_2)]

  vip_phs = vip_phs[[
    'birth_month',
    'customer_name',
    'branch_name',
    'date_of_birth',
    'contract_type',
    'broker_id',
    'broker_name',
    'effective_date',
    'approved_date',
  ]]

  # convert cột contract_type
  def mapper(x):
    if 'SILV' in x:
      result = 'SILV PHS'
    elif 'GOLD' in x:
      result = 'GOLD PHS'
    else:
      result = 'VIP Branch'
    return result

  vip_phs['contract_type'] = vip_phs['contract_type'].map(mapper)
  vip_phs.drop_duplicates(keep='last',inplace=True)

  if run_time.month < 6:
    vip_phs['review_date'] = dt.datetime(run_time.year,6,30)
  elif run_time.month < 12:
    vip_phs['review_date'] = dt.datetime(run_time.year,12,31)
  else:
    vip_phs['review_date'] = dt.datetime(run_time.year+1,6,30)

  branch_groupby_table = vip_phs.groupby(['branch_name','contract_type'])['customer_name'].count().unstack()
  branch_groupby_table.fillna(0,inplace=True)

  # --------------------- Viet File ---------------------

  file_name = f'Danh sách KH VIP phái sinh.xlsx'
  writer = pd.ExcelWriter(
    join(dept_folder,folder_name,period,file_name),
    engine='xlsxwriter',
    engine_kwargs={'options':{'nan_inf_to_errors':True}}
  )
  workbook = writer.book

  ###################################################
  ###################################################
  ###################################################

  sheet_title_format = workbook.add_format(
    {
      'bold'     :True,
      'align'    :'center',
      'valign'   :'vcenter',
      'text_wrap':True,
      'font_name':'Times New Roman',
      'font_size':18
    }
  )
  sub_title_format = workbook.add_format(
    {
      'bold'     :True,
      'italic'   :True,
      'align'    :'center',
      'valign'   :'vcenter',
      'text_wrap':True,
      'font_name':'Times New Roman',
      'font_size':11,
      'color'    :'#FF0000',
      'bg_color' :'#FFFFCC'
    }
  )
  header_format = workbook.add_format(
    {
      'border'   :1,
      'bold'     :True,
      'align'    :'center',
      'valign'   :'vcenter',
      'text_wrap':True,
      'font_name':'Times New Roman',
      'font_size':11,
      'bg_color' :'#65FF65'
    }
  )
  no_format = workbook.add_format(
    {
      'border'   :1,
      'bold'     :True,
      'align'    :'center',
      'valign'   :'vcenter',
      'text_wrap':True,
      'font_name':'Times New Roman',
      'font_size':11,
    }
  )
  header_format_tong_theo_cn = workbook.add_format(
    {
      'border'   :1,
      'bold'     :True,
      'align'    :'center',
      'valign'   :'vcenter',
      'font_name':'Times New Roman',
      'font_size':11,
      'bg_color' :'#92D050'
    }
  )
  list_customer_vip_row_format = workbook.add_format(
    {
      'border'   :1,
      'bold'     :True,
      'align'    :'center',
      'valign'   :'vcenter',
      'text_wrap':True,
      'font_name':'Times New Roman',
      'font_size':11,
      'color'    :'#FF0000',
      'bg_color' :'#FFFF00',
    }
  )
  ngay_hieu_luc_format = workbook.add_format(
    {
      'border'   :1,
      'bold'     :True,
      'align'    :'center',
      'valign'   :'vcenter',
      'text_wrap':True,
      'font_name':'Times New Roman',
      'font_size':11,
      'color'    :'#FF0000',
    }
  )
  list_customer_vip_col_format = workbook.add_format(
    {
      'border'   :1,
      'align'    :'center',
      'valign'   :'vcenter',
      'text_wrap':True,
      'font_name':'Times New Roman',
      'font_size':11,
      'bg_color' :'#FFCCFF',
    }
  )
  date_format = workbook.add_format(
    {
      'border'    :1,
      'align'     :'right',
      'valign'    :'vcenter',
      'num_format':'dd/mm/yyyy',
      'font_name' :'Times New Roman',
      'font_size' :11
    }
  )
  text_left_format = workbook.add_format(
    {
      'border'   :1,
      'text_wrap':True,
      'valign'   :'vcenter',
      'align'    :'left',
      'font_name':'Times New Roman',
      'font_size':11
    }
  )
  text_center_format = workbook.add_format(
    {
      'border'   :1,
      'text_wrap':True,
      'valign'   :'vcenter',
      'align'    :'center',
      'font_name':'Times New Roman',
      'font_size':11
    }
  )
  num_left_format = workbook.add_format(
    {
      'border'    :1,
      'align'     :'left',
      'valign'    :'vcenter',
      'font_name' :'Times New Roman',
      'font_size' :11,
      'num_format':'#,##0',
    }
  )
  num_center_format = workbook.add_format(
    {
      'border'    :1,
      'align'     :'center',
      'valign'    :'vcenter',
      'font_name' :'Times New Roman',
      'font_size' :11,
      'num_format':'#,##0'
    }
  )
  headers = [
    'No',
    'Tháng sinh',
    'Account',
    'Name',
    'Location',
    'Birthday',
    'LIST CUSTOMER VIP',
    'Ngày hiệu lực đầu tiên',
    'Approve date',
    'Review date',
    'Mã MG quản lý tài khoản',
    'Tên MG quản lý tài khoản',
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
  vip_phs_sheet = workbook.add_worksheet('VIP PHS')
  vip_phs_sheet.hide_gridlines(option=2)

  # Content of sheet
  sheet_title_name = f'UPDATED LIST OF COMPANY VIP TO {run_time.day}.{run_time.month}.{run_time.year}'
  sub_title_name = 'Chỉ tặng quà sinh nhật cho Gold PHS & Silv PHS'

  # Set Column Width and Row Height
  vip_phs_sheet.set_column('A:A',4)
  vip_phs_sheet.set_column('B:C',13)
  vip_phs_sheet.set_column('D:D',40)
  vip_phs_sheet.set_column('E:E',26)
  vip_phs_sheet.set_column('F:F',11)
  vip_phs_sheet.set_column('G:G',13)
  vip_phs_sheet.set_column('H:H',11)
  vip_phs_sheet.set_column('I:I',11)
  vip_phs_sheet.set_column('J:J',12)
  vip_phs_sheet.set_column('K:K',11)
  vip_phs_sheet.set_column('L:L',40)
  vip_phs_sheet.set_column('M:M',20)
  vip_phs_sheet.set_row(0,23)
  vip_phs_sheet.set_row(1,27)
  vip_phs_sheet.set_row(2,12)
  vip_phs_sheet.set_row(3,50)

  # merge row
  vip_phs_sheet.merge_range('A1:M1',sheet_title_name,sheet_title_format)
  vip_phs_sheet.merge_range('A2:M2',sub_title_name,sub_title_format)

  # write row and write column
  vip_phs_sheet.write_row('A4',headers,header_format)
  vip_phs_sheet.write('A4','No',no_format)
  vip_phs_sheet.write('G4','LIST CUSTOMER VIP',list_customer_vip_row_format)
  vip_phs_sheet.write('H4','Ngày hiệu lực đầu tiên',ngay_hieu_luc_format)
  vip_phs_sheet.write_column('A5',np.arange(vip_phs.shape[0])+1,num_center_format)
  vip_phs_sheet.write_column('B5',vip_phs['birth_month'],text_center_format)
  vip_phs_sheet.write_column('C5',vip_phs.index,num_left_format)
  vip_phs_sheet.write_column('D5',vip_phs['customer_name'],text_left_format)
  vip_phs_sheet.write_column('E5',vip_phs['branch_name'],text_left_format)
  vip_phs_sheet.write_column('F5',vip_phs['date_of_birth'],date_format)
  vip_phs_sheet.write_column('G5',vip_phs['contract_type'],list_customer_vip_col_format)
  vip_phs_sheet.write_column('H5',vip_phs['effective_date'],date_format)
  vip_phs_sheet.write_column('I5',vip_phs['approved_date'],date_format)
  vip_phs_sheet.write_column('J5',vip_phs['review_date'],date_format)
  vip_phs_sheet.write_column('K5',vip_phs['broker_id'],text_center_format)
  vip_phs_sheet.write_column('L5',vip_phs['broker_name'],text_left_format)
  vip_phs_sheet.write_column('M5',['']*vip_phs.shape[0],text_left_format)

  ##############################################
  ##############################################
  ##############################################

  #  Viết Sheet TONG THEO CN
  sum_format = workbook.add_format(
    {
      'border'    :1,
      'bold'      :True,
      'valign'    :'vcenter',
      'align'     :'right',
      'font_size' :11,
      'font_name' :'Calibri',
      'num_format':'#,##0'
    }
  )
  sum_title_format = workbook.add_format(
    {
      'border'   :1,
      'bold'     :True,
      'valign'   :'vcenter',
      'align'    :'left',
      'font_size':11,
      'font_name':'Calibri'
    }
  )
  location_format = workbook.add_format(
    {
      'border'   :1,
      'valign'   :'vcenter',
      'align'    :'left',
      'font_size':11,
      'font_name':'Calibri'
    }
  )
  num_format = workbook.add_format(
    {
      'border'    :1,
      'valign'    :'vcenter',
      'align'     :'right',
      'font_size' :11,
      'font_name' :'Calibri',
      'num_format':'#,##0'
    }
  )
  note_format = workbook.add_format(
    {
      'border'   :1,
      'valign'   :'vcenter',
      'align'    :'left',
      'font_size':11,
      'font_name':'Calibri'
    }
  )
  branch_groupby_sheet = workbook.add_worksheet('TONG THEO CN')
  branch_groupby_sheet.hide_gridlines(option=2)
  # Set Column Width and Row Height
  branch_groupby_sheet.set_column('A:A',28)
  branch_groupby_sheet.set_column('B:B',14)
  branch_groupby_sheet.set_column('C:C',15)
  branch_groupby_sheet.set_column('D:F',18)

  branch_groupby_sheet.write_row('A1',headers_tong_theo_cn,header_format_tong_theo_cn)
  branch_groupby_sheet.write_column('A2',branch_groupby_table.index,location_format)
  branch_groupby_sheet.write_column('B2',branch_groupby_table['GOLD PHS'],num_format)
  branch_groupby_sheet.write_column('C2',branch_groupby_table['SILV PHS'],num_format)
  branch_groupby_sheet.write_column('D2',branch_groupby_table['VIP Branch'],num_format)
  sum_row = branch_groupby_table['GOLD PHS']+branch_groupby_table['SILV PHS']+branch_groupby_table['VIP Branch']
  branch_groupby_sheet.write_column('E2',sum_row,sum_format)
  note_col = ['']*branch_groupby_table.shape[0]
  branch_groupby_sheet.write_column('F2',note_col,note_format,)
  start_sum_row = branch_groupby_table.shape[0]+2
  branch_groupby_sheet.write(f'A{start_sum_row}','SUM',sum_title_format)
  branch_groupby_sheet.write(
    f'B{start_sum_row}',
    branch_groupby_table['GOLD PHS'].sum(),
    sum_format
  )
  branch_groupby_sheet.write(
    f'C{start_sum_row}',
    branch_groupby_table['SILV PHS'].sum(),
    sum_format
  )
  branch_groupby_sheet.write(
    f'D{start_sum_row}',
    branch_groupby_table['VIP Branch'].sum(),
    sum_format
  )
  branch_groupby_sheet.write(f'E{start_sum_row}',sum_row.values.sum(),sum_format)
  branch_groupby_sheet.write(f'F{start_sum_row}','',sum_format)

  ###########################################################################
  ###########################################################################
  ###########################################################################

  writer.close()

  if __name__ == '__main__':
    print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
  else:
    print(f"{__name__.split('.')[-1]} ::: Finished")
  print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')
