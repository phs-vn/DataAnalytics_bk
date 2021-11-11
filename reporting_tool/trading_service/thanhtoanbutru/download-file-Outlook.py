from reporting_tool.trading_service.thanhtoanbutru import *


def run(
        periodicity: str,
        run_time=None,
):
    start = time.time()
    info = get_info(periodicity, run_time)
    start_date = info['start_date']
    end_date = info['end_date']
    period = info['period']

    outlook = Dispatch('outlook.application')
    mapi = outlook.GetNamespace("MAPI")

    for account in mapi.Accounts:
        print(f"Account {account.DeliveryStore.DisplayName} is being logged")  # print which account is being logged

    inbox = mapi.GetDefaultFolder(6)  # 6 = Inbox (without mails from the sub_folder)
    messages = inbox.Items

    # Save the email attachment to the below directory
    if not os.path.isdir(join(dept_folder, 'Attachment')):  # dept_folder from import
        os.mkdir(join(dept_folder, 'Attachment'))
    if not os.path.isdir(join(dept_folder, 'Attachment', period)):
        os.mkdir((join(dept_folder, 'Attachment', period)))

    try:
        for message in messages:
            try:
                emails = ['@ocb', '@eximbank', '@phs']
                extension_img = ('.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.gif')
                EIB_form = ['PHU HUNG', 'EIB', 'PH']
                if message.Class == 43:  # Type of email is MailItem, if it = 2
                    check_Email_type = message.SenderEmailType
                    check_date = message.ReceivedTime.date() == dt.strptime(start_date, '%Y/%m/%d').date()
                    # check_date = dt.now() - timedelta(days=1)
                    if check_Email_type == 'EX' and check_date:  # message != Ex -> not automatically replied message
                        check_Sender_name = message.Sender.GetExchangeUser().PrimarySmtpAddress
                        if any(email in check_Sender_name for email in emails):
                            for attachment in message.Attachments:
                                check_file_extension = attachment.FileName.endswith(extension_img)
                                check_file_name = (EIB in attachment.FileName for EIB in EIB_form)
                                if not check_file_extension and check_file_name:
                                    dir_name = 'EIB'
                                    if not os.path.isdir(join(dept_folder, 'Attachment', period, dir_name)):
                                        os.mkdir((join(dept_folder, 'Attachment', period, dir_name)))
                                        attachment.SaveASFile(os.path.join(dept_folder,
                                                                           'Attachment',
                                                                           period,
                                                                           dir_name,
                                                                           attachment.FileName))
                                elif not check_file_extension and not check_file_name:
                                    dir_name = 'OCB'
                                    if not os.path.isdir(join(dept_folder, 'Attachment', period, dir_name)):
                                        os.mkdir((join(dept_folder, 'Attachment', period, dir_name)))
                                        attachment.SaveASFile(os.path.join(dept_folder,
                                                                           'Attachment',
                                                                           period,
                                                                           dir_name,
                                                                           attachment.FileName))

            except Exception as e:
                raise CustomError("error when saving the attachment:" + str(e))
    except Exception as e:
        raise CustomError("error when processing emails messages:" + str(e))
