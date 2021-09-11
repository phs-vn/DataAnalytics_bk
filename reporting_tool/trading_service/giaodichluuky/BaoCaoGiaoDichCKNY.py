from reporting_tool.trading_service.giaodichluuky import *

def run(
        periodicity:str,
        run_time=None,
):
    start = time.time()
    info = get_info(periodicity, run_time)
    start_date = info['start_date']
    begin_of_year = f'{start_date[:4]}/01/01'
    end_date = info['end_date']
    period = info['period']
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder, folder_name, period)):  # dept_folder from import
        os.mkdir(join(dept_folder, folder_name, period))

    trading_record = pd.read_sql(
        "SELECT sub_account, exchange, type_of_account, type_of_asset, type_of_order, value "
        "FROM trading_record "
        f"WHERE date BETWEEN '{begin_of_year}' AND '{end_date}' "
        f"AND settlement_period IN (1,2);",
        connect,
        index_col='sub_account',
    )
    account = pd.read_sql(
        "SELECT sub_account, account_code FROM sub_account;",
        connect,
        index_col='sub_account',
    ).squeeze()
    broker = pd.read_sql(
        "SELECT account_code, broker_id FROM account",
        connect,
        index_col='account_code',
    ).squeeze()
    branch_id = pd.read_sql(
        "SELECT broker_id, branch_id FROM broker",
        connect,
        index_col='broker_id',
    ).squeeze()
    branch_name = pd.read_sql(
        "SELECT branch_id, branch_name FROM branch;",
        connect,
        index_col='branch_id',
    ).squeeze()

    trading_record['exchange'].replace('UPCOM','HNX',inplace=True)
    domestic = trading_record['type_of_account'].str.endswith('trong nước')
    foreign = trading_record['type_of_account'].str.endswith('nước ngoài')
    tudoanh = trading_record['type_of_account'].inna() # chờ IT fix loi se sua thanh .str.endswith('tu doanh')
    trading_record.loc[domestic] = 'trong nước'
    trading_record.loc[foreign] = 'nước ngoài'
    trading_record.loc[tudoanh] = 'tự doanh'
    trading_record.insert(
        3,
        'type',
        trading_record['type_of_account']+' - '+trading_record['type_of_asset']
    )


