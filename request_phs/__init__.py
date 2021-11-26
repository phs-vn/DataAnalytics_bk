from function_phs import *

# Risk Database Information
driver_RMD = '{SQL Server}'
server_RMD = 'SRV-RPT'
db_RMD = 'RiskDb'
id_RMD = 'namtran'
password_RMD = open(r"C:\Users\namtran\Desktop\pass_word.txt", "r").read()

connect_RMD = pyodbc.connect(
    f'Driver={driver_RMD};'
    f'Server={server_RMD};'
    f'Database={db_RMD};'
    f'uid={id_RMD};'
    f'pwd={password_RMD}'
)
TableNames_DWH_CoSo = pd.read_sql(
    'SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES',
    connect_RMD,
)

# DWH-CoSo Database Information
driver_DWH_CoSo = '{SQL Server}'
server_DWH_CoSo = 'SRV-RPT'
db_DWH_CoSo = 'DWH-CoSo'
id_DWH_CoSo = 'namtran'
password_DWH_CoSo = 'nam!tran@2021'

connect_DWH_CoSo = pyodbc.connect(
    f'Driver={driver_DWH_CoSo};'
    f'Server={server_DWH_CoSo};'
    f'Database={db_DWH_CoSo};'
    f'uid={id_DWH_CoSo};'
    f'pwd={password_DWH_CoSo}'
)
TableNames_RMD = pd.read_sql(
    'SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES',
    connect_DWH_CoSo
)

# DWH-PhaiSinh Database Information
driver_DWH_PhaiSinh = '{SQL Server}'
server_DWH_PhaiSinh = 'SRV-RPT'
db_DWH_PhaiSinh = 'DWH-PhaiSinh'
id_DWH_PhaiSinh = 'namtran'
password_DWH_PhaiSinh = 'nam!tran@2021'

connect_DWH_PhaiSinh = pyodbc.connect(
    f'Driver={driver_DWH_PhaiSinh};'
    f'Server={server_DWH_PhaiSinh};'
    f'Database={db_DWH_PhaiSinh};'
    f'uid={id_DWH_PhaiSinh};'
    f'pwd={password_DWH_PhaiSinh}'
)
TableNames_DWH_PhaiSinh = pd.read_sql(
    'SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES',
    connect_DWH_PhaiSinh
)