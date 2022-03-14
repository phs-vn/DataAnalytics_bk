from implementation import TaskMonitor


@TaskMonitor
def DWH_CoSo_Update_All():
    from datawarehouse.DWH_CoSo import UPDATE
    UPDATE()


@TaskMonitor
def DWH_PhaiSinh_Update_All():
    from datawarehouse.DWH_PhaiSinh import UPDATE
    UPDATE()


@TaskMonitor
def DWH_ThiTruong_Update_DanhSachMa():
    from datawarehouse.DWH_ThiTruong.DanhSachMa import update as Update_DanhSachMa
    Update_DanhSachMa()


@TaskMonitor
def DWHThiTruongUpdate_DuLieuGiaoDichNgay():
    from datawarehouse.DWH_ThiTruong.DuLieuGiaoDichNgay import update as Update_DuLieuGiaoDichNgay
    today = dt.datetime.now()
    Update_DuLieuGiaoDichNgay(today,today)


@TaskMonitor
def DWHThiTruongUpdate_TinChungKhoan():
    from datawarehouse.DWH_ThiTruong.TinChungKhoan import update as Update_TinChungKhoan
    Update_TinChungKhoan(24)

