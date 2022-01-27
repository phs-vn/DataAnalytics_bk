from request import *
from datawarehouse import *


def update(
    date:str,
    n:int=1,
):

    """
    This function get data from Securities Service Division and insert to [DWH-CoSo].[dbo].[de083]

    :param date: date to insert YYYY-MM-DD
    :param n: number of historical days to insert a overwrite old data
    """

    updateDates = pd.read_sql(
        f"""
        SELECT TOP {n-1}
            [d].[date]
        FROM (SELECT DISTINCT [de083].[date] FROM [de083]) [d]
        ORDER BY [d].[date] DESC
        """,
        connect_DWH_CoSo
    ).squeeze()

    if not updateDates.empty:
        updateDates = updateDates.dt.strftime('%Y.%m.%d').to_list()
    else:
        updateDates = []

    insertDates = [f'{date[:4]}.{date[5:7]}.{date[-2:]}'] + updateDates
    SQLdateset = iterable_to_sqlstring(insertDates)
    DELETE(connect_DWH_CoSo,'de083',f'WHERE [de083].[date] IN {SQLdateset}')

    insertDatesSQL = [tuple(d.split('.')) for d in insertDates]
    root = r'C:\Users\hiepdang\Shared Folder\Trading Service\Report\GiaoDichLuuKy\FilesFromVSD\SoDuChungKhoan'
    for d in insertDatesSQL:
        year, month, day = d
        folder = join(root,year,f'THÁNG {month}',f'{day}.{month}.{year}C')
        files = [f for f in listdir(folder) if f[:2] in ('BO','DC','HN','HO','UP') and f.split('.')[-1]=='csv']
        for file in files:
            df = pd.read_csv(
                join(folder,file),
                delimiter=';',
                usecols=[10,13,14],
                header=None,
                names=['account_code','amount','ticker']
            )
            df.insert(0,'date',dt.datetime(int(year),int(month),int(day)))
            df = df[[
                'date',
                'account_code',
                'ticker',
                'amount',
            ]]
            INSERT(connect_DWH_CoSo,'de083',df)



def init():

    """
    This function get ALL data from Securities Service Division and insert to [DWH-CoSo].[dbo].[de083]
    Used to initiate data into the empy [DWH-CoSo].[dbo].[de083]
    """

    root = r'C:\Users\hiepdang\Shared Folder\Trading Service\Report\GiaoDichLuuKy\FilesFromVSD\SoDuChungKhoan'
    for yearFolder in listdir(root):
        for monthFolder in listdir(join(root,yearFolder)):
            for dateFolder in listdir(join(root,yearFolder,monthFolder)):
                year = dateFolder.split('.')[2].replace('C','')
                month = dateFolder.split('.')[1]
                day = dateFolder.split('.')[0]
                update(f'{year}-{month}-{day}')
