from request.stock import *
from news_collector import newsts

def run(num_hours:int=48):

  """
  This method runs all functions in module news_collector.newsts (try till success)
  and exports all resulted DataFrames to a single excel file in the specified folder
  for daily usage of TS. This function is called in a higher-level module and
  automatically run on a daily basis

  :param num_hours: number of hours in the past that's in our concern

  :return: None
  """

  now = dt.date.today()
  date_string = now.strftime('%Y%m%d')

  path = r'C:\Users\hiepdang\Shared Folder\Trading Service\News Update'
  file_name = f'{date_string}_news.xlsx'
  file_path = fr'{path}\{file_name}'

  while True:
    try:
      try:
        vsd_TCPH = newsts.vsd.tinTCPH(num_hours)
        break
      except newsts.vsd.ignored_exceptions:
        vsd_TCPH = pd.DataFrame()
        continue
    except newsts.NoNewsFound:
      vsd_TCPH = pd.DataFrame()
      break

  while True:
    try:
      try:
        vsd_TVLK = newsts.vsd.tinTVLK(num_hours)
        break
      except newsts.vsd.ignored_exceptions:
        vsd_TVLK = pd.DataFrame()
        continue
    except newsts.NoNewsFound:
      vsd_TVLK = pd.DataFrame()
      break

  while True:
    try:
      try:
        hnx_tintuso = newsts.hnx.tintuso(num_hours)
        break
      except newsts.hnx.ignored_exceptions:
        hnx_tintuso = pd.DataFrame()
        continue
    except newsts.NoNewsFound:
      hnx_tintuso = pd.DataFrame()
      break

  while True:
    try:
      try:
        hose_tinTCNY = newsts.hose.tinTCNY(num_hours)
        break
      except newsts.hose.ignored_exceptions:
        hose_tinTCNY = pd.DataFrame()
        continue
    except newsts.NoNewsFound:
      hose_tinTCNY = pd.DataFrame()
      break

  while True:
    try:
      try:
        hose_tinCW = newsts.hose.tinCW(num_hours)
        break
      except newsts.hose.ignored_exceptions:
        hose_tinCW = pd.DataFrame()
        continue
    except newsts.NoNewsFound:
      hose_tinCW = pd.DataFrame()
      break

  with pd.ExcelWriter(file_path) as writer:
    vsd_TCPH.to_excel(writer,sheet_name='vsd_tinTCPH',index=False)
    vsd_TVLK.to_excel(writer,sheet_name='vsd_tinTVLK',index=False)
    hnx_tintuso.to_excel(writer,sheet_name='hnx_tintuso',index=False)
    hose_tinTCNY.to_excel(writer,sheet_name='hose_tinTCNY',index=False)
    hose_tinCW.to_excel(writer,sheet_name='hose_tinCW',index=False)

  wb = xw.Book(file_path)

  wb.sheets['vsd_tinTCPH'].autofit()
  wb.sheets['vsd_tinTCPH'].range("C:D").column_width = 50
  wb.sheets['vsd_tinTCPH'].range("A:XFD").api.WrapText = True
  wb.sheets['vsd_tinTCPH'].range("A:XFD").api.Cells.VerticalAlignment = 2

  wb.sheets['vsd_tinTVLK'].autofit()
  wb.sheets['vsd_tinTVLK'].range("C:C").column_width = 80
  wb.sheets['vsd_tinTVLK'].range("A:XFD").api.WrapText = True
  wb.sheets['vsd_tinTVLK'].range("A:XFD").api.Cells.VerticalAlignment = 2

  wb.sheets['hnx_tintuso'].autofit()
  wb.sheets['hnx_tintuso'].range("C:C").column_width = 80
  wb.sheets['hnx_tintuso'].range("A:XFD").api.WrapText = True
  wb.sheets['hnx_tintuso'].range("A:XFD").api.Cells.VerticalAlignment = 2

  wb.sheets['hose_tinTCNY'].autofit()
  wb.sheets['hose_tinTCNY'].range("C:D").column_width = 90
  wb.sheets['hose_tinTCNY'].range("A:XFD").api.WrapText = True
  wb.sheets['hose_tinTCNY'].range("A:XFD").api.Cells.VerticalAlignment = 2

  wb.sheets['hose_tinCW'].autofit()
  wb.sheets['hose_tinCW'].range("B:B").column_width = 45
  wb.sheets['hose_tinCW'].range("C:C").column_width = 90
  wb.sheets['hose_tinCW'].range("D:E").column_width = 20
  wb.sheets['hose_tinCW'].range("A:XFD").api.WrapText = True
  wb.sheets['hose_tinCW'].range("A:XFD").api.Cells.VerticalAlignment = 2

  wb.save()
  wb.close()