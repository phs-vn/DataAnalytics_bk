import sys
sys.path.extend([r'C:\\Users\\hiepdang\\PycharmProjects\\DataAnalytics',
                 r'C:/Users/hiepdang/PycharmProjects/DataAnalytics'])

from post_phs import *

now = date.today()
run_time = now.strftime('%Y%m%d')

destination_dir = r'\\192.168.10.101\phs-storge-2018' \
                  r'\RiskManagementDept\RMD_Data' \
                  r'\Luu tru van ban\RMC Meeting 2018' \
                  r'\00. Meeting minutes\Co Phieu Giam San'

file_name = f'{run_time}_floor.xlsx'

file_path = fr'{destination_dir}\{file_name}'

# consecutively scan on 2 exchanges
hose_table = post.fc_consecutive('floor',True,2,'HOSE','all')
hnx_table = post.fc_consecutive('floor',True,2,'HNX','all')

# export to excel
result = pd.concat([hose_table,hnx_table])
if not result.empty:
    result.to_excel(file_path, sheet_name='Cổ phiếu giảm sàn')
    # format excel file
    wb = xw.Book(file_path)
    wb.sheets[0].autofit()
    wb.save()
    wb.close()

# Note: Trong trường hợp không có cổ phiếu giảm sàn 2 phiên sẽ không xuất file
