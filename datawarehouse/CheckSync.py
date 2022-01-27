from request import *


class Check:

    """
    Được chạy vào:
        - Mon-Fri: 12:15, 18:00, 21:00
        - Sat-Sun: 01:00
    """

    check_time = dt.datetime.now()

    ignored_Tables = [
        'v_sqlrun',
        'ExcludeDays',
        'ExecTaskLog',
        'bank_account_list',
        'de083',
    ]

    def __init__(self, db):

        if db == 'DWH-CoSo':
            conn = connect_DWH_CoSo
            self.db_Tables = TableNames_DWH_CoSo.squeeze()
            self._prefix = '[DWH-CoSo]'
            self.description = 'RunCoSo'
            self.since = dt.datetime.now()-dt.timedelta(minutes=45)

        elif db == 'DWH-PhaiSinh':
            conn = connect_DWH_PhaiSinh
            self.db_Tables = TableNames_DWH_PhaiSinh.squeeze()
            self._prefix = '[DWH-PhaiSinh]'
            self.description = 'RunPhaiSinh'
            self.since = dt.datetime.now()-dt.timedelta(minutes=90)
        else:
            raise ValueError('The module currently checks DWH-CoSo or DWH-PhaiSinh only')

        self.db_Tables = self.db_Tables.loc[~self.db_Tables.isin(self.ignored_Tables)]

        # check xem 45p vừa rồi có chạy chưa
        run_check = pd.read_sql(
            f"""
            SELECT 
                CAST([ExecTaskLog].[EXEC_DATE] AS DATE) [DATE],
                CAST([ExecTaskLog].[EXEC_DATE] AS TIME) [TIME]
            FROM [ExecTaskLog]
            WHERE CAST([ExecTaskLog].[EXEC_DATE] AS DATE) = '{self.check_time.strftime("%Y-%m-%d")}'
                AND CAST([ExecTaskLog].[EXEC_DATE] AS TIME) >= '{self.since.strftime("%H:%M:%S")}'
                AND [ExecTaskLog].[STATUS] = 'START'
            """,
            conn,
        ).squeeze()

        if run_check.empty: # nếu chưa chạy
            self.missing_Tables = self.db_Tables.to_frame()
        else: # nếu chạy rồi
            # check xem nếu chạy rồi thì có đủ bảng chưa
            self.run_Tables = pd.read_sql(
                f"""
                SELECT 
                    [ExecTaskLog].[DESCRIPTION]
                FROM [ExecTaskLog]
                WHERE CAST([ExecTaskLog].[EXEC_DATE] AS DATE) = '{self.check_time.strftime("%Y-%m-%d")}'
                    AND [ExecTaskLog].[STATUS] = 'OK'
                    AND CAST([ExecTaskLog].[EXEC_DATE] AS TIME) >= (
                        SELECT MAX(CAST([ExecTaskLog].[EXEC_DATE] AS TIME))
                        FROM [ExecTaskLog]
                        WHERE CAST([ExecTaskLog].[EXEC_DATE] AS DATE) = '{self.check_time.strftime("%Y-%m-%d")}'
                            AND [ExecTaskLog].[STATUS] = 'START'
                            AND [ExecTaskLog].[DESCRIPTION] = '{self.description}'
                    )
                """,
                conn,
            ).squeeze()
            self.missing_Tables = self.db_Tables.loc[~self.db_Tables.isin(self.run_Tables)].to_frame()

        self.missing_Tables.columns = ['Missing Tables']
        self.missing_Tables = self._prefix + '.[' + self.missing_Tables + ']'
        self.html_table = self.missing_Tables.to_html(index=False)
        self.html_table = self.html_table.replace('<tr>','<tr align="center">')  # center columns
        self.html_table = self.html_table.replace('border="1"','border="1" style="border-collapse:collapse"')  # make thinner borders


    def send_mail(self):

        if self.missing_Tables.empty:
            content = f"""
                <p style="font-family:Times New Roman; font-size:100%"><i>
                    ĐỦ SỐ LƯỢNG BẢNG
                </i></p>
            """
        else:
            content = self.html_table

        body = f"""
        <html>
            <head></head>
            <body>
                {content}
                <p style="font-family:Times New Roman; font-size:90%"><i>
                    -- Generated by Reporting System
                </i></p>
            </body>
        </html>
        """

        outlook = Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mapi = outlook.GetNamespace("MAPI")

        for account in mapi.Accounts:
            print(f"Account {account.DeliveryStore.DisplayName} is being logged")

        mail.To = 'tupham@phs.vn'
        mail.CC = 'hiepdang@phs.vn'
        mail.Subject = f"{self._prefix} Missing Tables {self.check_time.strftime('%Y-%m-%d %H:%M:%S')}"
        mail.HTMLBody = body
        mail.Send()

