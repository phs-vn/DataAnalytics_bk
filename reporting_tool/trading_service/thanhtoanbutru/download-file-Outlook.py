from reporting_tool.trading_service.thanhtoanbutru import *


def run():
    outlook = Dispatch('outlook.application')
    mapi = outlook.GetNamespace("MAPI")

    for account in mapi.Accounts:
        print(f"Account {account.DeliveryStore.DisplayName} is being logged")  # print which account is being logged

    inbox = mapi.GetDefaultFolder(6)  # 6 = Inbox (without mails from the sub_folder)
    messages = inbox.Items

    # Save the email attachment to the below directory
    if not os.path.isdir(join(dept_folder, 'Attachment')):  # dept_folder from import
        os.mkdir(join(dept_folder, 'Attachment'))
    if not os.path.isdir(join(dept_folder, 'Attachment', dt.now().strftime("%Y-%m-%d"))):
        os.mkdir((join(dept_folder, 'Attachment', dt.now().strftime("%Y-%m-%d"))))

    try:
        for message in messages:
            try:
                emails = ['@ocb', '@eximbank', '@phs']
                extension_img = ('.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.gif')

                if message.Class == 43:  # Type of email is MailItem, if it = 2
                    check_Email_type = message.SenderEmailType
                    check_date = message.ReceivedTime.date() == dt.now().date()
                    if check_Email_type == 'EX' and check_date:  # message != Ex -> not automatically replied message
                        check_Sender_name = message.Sender.GetExchangeUser().PrimarySmtpAddress
                        if any(email in check_Sender_name for email in emails):
                            for attachment in message.Attachments:
                                if not attachment.FileName.endswith(extension_img):
                                    attachment.SaveASFile(os.path.join(dept_folder,
                                                                       'Attachment',
                                                                       dt.now().strftime("%Y-%m-%d"),
                                                                       attachment.FileName))

            except Exception as e:
                raise CustomError("error when saving the attachment:" + str(e))
    except Exception as e:
        raise CustomError("error when processing emails messages:" + str(e))
