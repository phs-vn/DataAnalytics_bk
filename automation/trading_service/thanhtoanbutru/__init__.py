from automation.trading_service import *

dept_folder = r'C:\Users\namtran\Share Folder\Trading Service\Report\ThanhToanBuTru'


def getBranchBroker(
        x: str,
        d: dt.datetime
) -> tuple:
    """
    This function return branch name and broker name given an account code or sub account
    :param x: either account_code or sub_auccount
    :param d: date to check the relationship
    """

    sqlStatement = f"""
        SELECT DISTINCT
            [branch].[branch_name], 
            [broker].[broker_name]
        FROM [relationship]
        LEFT JOIN [branch] ON [branch].[branch_id] = [relationship].[branch_id]
        LEFT JOIN [broker] ON [broker].[broker_id] = [relationship].[broker_id]
        WHERE [relationship].[date] = '{d.strftime('%Y-%m-%d')}'
    """

    if re.findall('[A-Z]', x):  # account code
        sqlStatement += f" AND [relationship].[account_code] = '{x}'"
    else:
        sqlStatement += f" AND [relationship].[sub_account] = '{x}'"

    res = tuple(pd.read_sql(sqlStatement, connect_DWH_CoSo).squeeze())
    return res
