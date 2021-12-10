from reporting_tool.trading_service.thanhtoanbutru import *


def run(
        periodicity: str,
        end_date: str,
        run_time=None,
):
    start = time.time()
    info = get_info(periodicity, run_time)
    period = info['period']
    folder_name = info['folder_name']
    date_character = ['/', '-', '.']

    # create folder
    if not os.path.isdir(join(dept_folder, folder_name)):
        os.mkdir(join(dept_folder, folder_name))
    if not os.path.isdir(join(dept_folder, folder_name, period)):
        os.mkdir((join(dept_folder, folder_name, period)))