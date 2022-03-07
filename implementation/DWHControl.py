from implementation import TaskMonitor

# @TaskMonitor
def DWH_CoSoCheckSyncUpdate():
    from datawarehouse import CheckSyncUpdate as CoSoPhaiSinhCheckSyncUpdate
    CoSoPhaiSinhCheckSyncUpdate.DB('DWH-CoSo').SendMailAndExec()

@TaskMonitor
def DWH_PhaiSinhCheckSyncUpdate():
    from datawarehouse import CheckSyncUpdate as CoSoPhaiSinhCheckSyncUpdate
    CoSoPhaiSinhCheckSyncUpdate.DB('DWH-PhaiSinh').SendMailAndExec()