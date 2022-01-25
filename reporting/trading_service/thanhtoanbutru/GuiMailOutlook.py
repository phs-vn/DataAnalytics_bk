from reporting.trading_service.thanhtoanbutru import *


def run(
        periodicity: str,
        run_time=None,

):
    info = get_info(periodicity, run_time)
    period = info['period']
    file_name = 'check_chenh_lech.txt'

    outlook = Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mapi = outlook.GetNamespace("MAPI")

    for account in mapi.Accounts:
        print(f"Account {account.DeliveryStore.DisplayName} is being logged")

    mail.To = 'huynam177199@gmail.com'
    mail.Subject = 'Cảnh báo sự chênh lệch của báo cáo số 8 và báo cáo số 18'
    mail.Body = "Dear Mrs, \n" \
                "\nGửi đính kèm file cảnh báo sự chênh lệch giữa 2 báo cáo số 18 và báo cáo số 8\n" \
                "\nBest Regards and Thanks."
    mail.Attachments.Add(join(dept_folder, 'Attachment', period, file_name))
    mail.CC = ''

    n = int(input())
    if n > 0:
        mail.Send()
        print('Send email successfully!')
    else:
        raise "You can't send email!"
