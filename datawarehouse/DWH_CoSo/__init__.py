from datawarehouse import connect_DWH_CoSo, EXEC
from request import *

def UPDATE(*tables):

    """
    This function EXEC the corresponding stored procedures to update
    the specified tables

    :param tables: names of tables. If no tables specified, run all tables
    """

    if tables == ():
        EXEC(connect_DWH_CoSo,'spRunCoSo')
    else:
        today = dt.datetime.now().strftime('%Y-%m-%d')
        for table in tables:
            EXEC(connect_DWH_CoSo,f'sp{table}',FrDate=today,ToDate=today)
