import os

EIB_form = ['PHU HUNG', 'EIB']
files = os.listdir(
    os.path.join('D:/DataAnalytics/reporting_tool/trading_service/output/ThanhToanBuTru', 'Attachment', '2021.11.11'))

eib_file = ''
ocb_file = ''
for file in files:
    for EIB in EIB_form:
        if EIB in file:
            eib_file = file
        else:
            ocb_file = file
print(eib_file, ' ,', ocb_file)
