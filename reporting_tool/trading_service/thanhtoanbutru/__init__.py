from reporting_tool.trading_service import *

dept_folder = join(realpath(dirname(dirname(__file__))), 'output', 'ThanhToanBuTru')


class CustomError(Exception):
    pass
