from automation.trading_service import *

driver = '{SQL Server}'
server = 'SRV-RPT'
database = 'RiskDb'
user_id = 'namtran'
user_password = open(r'C:\Users\namtran\Desktop\pass_word.txt').readline()

connect = pyodbc.connect(
    f'Driver={driver};'
    f'Server={server};'
    f'Database={database};'
    f'uid={user_id};'
    f'pwd={user_password}'
)

TableNames = pd.read_sql(
    'SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES',connect
)
