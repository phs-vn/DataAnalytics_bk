import pandas as pd

from request import *


def INSERT(
    conn,
    table:str,
    df:pd.DataFrame,
):
    """
    This function INSERT a pd.DataFrame to a particular [db].[table].
    Must make sure the order / data type of pd.DataFrame align with [db].[table]

    :param conn: connect object of the Database
    :param table: name of the table in the Database
    :param df: inserted pd.DataFrame

    :return: None
    """

    sqlStatement = f"INSERT INTO [{table}] VALUES ({','.join(['?']*df.shape[1])})"
    cursor = conn.cursor()
    cursor.fast_executemany = True
    cursor.executemany(sqlStatement,df.values.tolist())
    cursor.commit()
    cursor.close()


def DELETE(
    conn,
    table:str,
    where:str,
):
    """
    This function DELETE entire rows from a [db].[table] given a paticular WHERE clause.
    If WHERE = '', it completely clears all data from the [db].[table]

    :param conn: connect object of the Database
    :param table: name of the table in the Database
    :param where: WHERE clause in DELETE statement
    """
    if where == '':
        where = ''
    else:
        if not where.startswith('WHERE'):
            where = 'WHERE ' + where
    sqlStatement = f"DELETE FROM [{table}] {where}"
    cursor = conn.cursor()
    cursor.fast_executemany = True
    cursor.execute(sqlStatement)
    cursor.commit()
    cursor.close()

