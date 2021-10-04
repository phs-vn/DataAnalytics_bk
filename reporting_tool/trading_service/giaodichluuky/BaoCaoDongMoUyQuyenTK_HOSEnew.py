from reporting_tool.trading_service.giaodichluuky import *

def run(
        periodicity:str,
        run_time=None,
):
    start = time.time()
    info = get_info(periodicity,run_time)
    run_time = info['run_time']
    start_date = info['start_date']
    end_date = info['end_date']
    period = info['period']
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):  # dept_folder from import
        os.mkdir(join(dept_folder,folder_name,period))

    period_account = pd.read_sql(
        "SELECT "
        "account_type,"
        "account_code, "
        "customer_name, "
        "nationality, "
        "address, "
        "customer_id_number, "
        "date_of_issue, "
        "place_of_issue, "
        "date_of_open, "
        "date_of_close "
        "FROM account "
        f"WHERE (date_of_open BETWEEN '{start_date}' AND '{end_date}') "
        f"OR (date_of_close BETWEEN '{start_date}' AND '{end_date}') ",
        connect_DWH_CoSoo,
        index_col='account_code',
    )
    count_account = pd.read_sql(
        "SELECT "
        "COUNT('account_type') AS count, "
        "account_type "
        "FROM account "
        "GROUP BY account_type",
        connect_DWH_CoSoo,
        index_col='account_type',
    )
    contract_type = pd.read_sql(
        "SELECT "
        "customer_information.sub_account, "
        "sub_account.account_code, "
        "customer_information.contract_code, "
        "customer_information.contract_type "
        "FROM customer_information "
        "LEFT JOIN sub_account ON customer_information.sub_account=sub_account.sub_account",
        connect_DWH_CoSoo,
        index_col='sub_account',
    )
    customer_information_change = pd.read_sql(
        "SELECT * "
        "FROM customer_information_change "
        f"WHERE change_date BETWEEN '{start_date}' AND '{end_date}'",
        connect_DWH_CoSoo,
        index_col='account_code',
    )
    authorization = pd.read_sql(
        "SELECT * "
        "FROM [authorization] "
        f"WHERE date_of_authorization BETWEEN '{start_date}' AND '{end_date}' "
        f"AND authorized_person_id = '155/GCNTVLK' "
        f"AND scope_of_authorization IS NOT NULL "
        f"AND scope_of_authorization <> 'I,IV,V' ",
        connect_DWH_CoSoo,
        index_col='account_code',
    )
    # Loai cac uy quyen duoc mo moi chi de dang ky uy quyen them (rule ben DVKH)
    drop_account = pd.read_sql(
        "SELECT account_code "
        "FROM authorization_change "
        f"WHERE date_of_change BETWEEN '{start_date}' AND '{end_date}' "
        "AND new_end_date IS NOT NULL",
        connect_DWH_CoSoo,
        index_col='account_code'
    )
    authorization_change = pd.read_sql(
        "SELECT * "
        "FROM authorization_change "
        f"WHERE date_of_change BETWEEN '{start_date}' AND '{end_date}'",
        connect_DWH_CoSoo,
        index_col='account_code'
    )
    authorization = authorization.loc[~authorization.index.isin(drop_account.index)]
    authorization['scope_of_authorization'] = 'I,II,IV,V,VII,IX,X'

    mapper = lambda x: 'Thường' if x.startswith('Thường') else 'Ký Quỹ'
    contract_type['contract_type'] = contract_type['contract_type'].map(mapper)

    margin_account = contract_type.loc[contract_type['contract_type']=='Ký Quỹ','account_code']
    period_account.loc[period_account.index.isin(margin_account),'remark'] = 'TKKQ'
    period_account['remark'].fillna('',inplace=True)
    period_account.loc[period_account['account_type'].str.startswith('Cá nhân'),'entity_type'] = 'CN'
    period_account.loc[period_account['account_type'].str.startswith('Tổ chức'),'entity_type'] = 'TC'
    account_open = period_account.loc[period_account['date_of_close'].isnull()]
    account_close = period_account.loc[period_account['date_of_close'].notnull()]

    # Cuoi ky
    closing_ind_domestic = count_account.loc['Cá nhân trong nước','count']
    closing_ins_domestic = count_account.loc['Tổ chức trong nước','count']
    closing_ind_foreign = count_account.loc['Cá nhân nước ngoài','count']
    closing_ins_foreign = count_account.loc['Tổ chức nước ngoài','count']
    closing_totaL_domestic = closing_ind_domestic + closing_ins_domestic
    closing_totaL_foreign = closing_ind_foreign + closing_ins_foreign
    closing_total = closing_totaL_domestic + closing_totaL_foreign

    # Tinh bien dong TK
    open_ind_domestic = (account_open['account_type']=='Cá nhân trong nước').sum()
    open_ins_domestic = (account_open['account_type']== 'Tổ chức trong nước').sum()
    open_ind_foreign = (account_open['account_type']=='Cá nhân nước ngoài').sum()
    open_ins_foreign = (account_open['account_type']=='Tổ chức nước ngoài').sum()
    close_ind_domestic =  (account_close['account_type']=='Cá nhân trong nước').sum()
    close_ins_domestic = (account_close['account_type']=='Tổ chức trong nước').sum()
    close_ind_foreign = (account_close['account_type']=='Cá nhân nước ngoài').sum()
    close_ins_foreign = (account_close['account_type']=='Tổ chức nước ngoài').sum()
    open_total_domestic = open_ind_domestic + open_ins_domestic
    open_total_foreign = open_ind_foreign + open_ins_foreign
    close_total_domestic = close_ind_domestic + close_ins_domestic
    close_total_foreign = close_ind_foreign + close_ins_foreign
    open_total = open_total_domestic + open_total_foreign
    close_total = close_total_domestic + close_total_foreign

