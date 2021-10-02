from function_phs import *

connect = pyodbc.connect(
    f'Driver={driver};'
    f'Server={server};'
    f'Database={database};'
    f'uid={user_id};'
    f'pwd={user_password}'
)
TableNames = pd.read_sql(
    'SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES',
    connect
)
