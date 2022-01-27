"""
This module simply brings functions from reporting package to implementation package
to ensure consistent workflow among projects.
"""

from implementation import TaskMonitor
from request.stock import *

from reporting.trading_service.giaodichluuky import BaoCaoCheckGia
from reporting.trading_service.giaodichluuk``y import BaoCaoDanhSachChungKhoanSSI
from reporting.trading_service.giaodichluuky import BaoCaoDongMoUyQuyenTK_HNX
from reporting.trading_service.giaodichluuky import BaoCaoDongMoUyQuyenTK_HOSEnew
from reporting.trading_service.giaodichluuky import BaoCaoDongMoUyQuyenTK_HOSEold
from reporting.trading_service.giaodichluuky import BaoCaoHoatDongLuuKyNDTNN
from reporting.trading_service.giaodichluuky import BaoCaoTamTinhPhiChuyenKhoan
from reporting.trading_service.giaodichluuky import BaoCaoTamTinhPhiGiaoDich
from reporting.trading_service.giaodichluuky import BaoCaoTamTinhPhiLuuKy
from reporting.trading_service.giaodichluuky import BaoCaoTienGuiSSC
from reporting.trading_service.giaodichluuky import BaoCaoTinhHinhHDKD

from reporting.trading_service.thanhtoanbutru import BaoCaoDoiChieuLaiTienGuiPhatSinh_p1
from reporting.trading_service.thanhtoanbutru import BaoCaoDoiChieuLaiTienGuiPhatSinh_p2
from reporting.trading_service.thanhtoanbutru import BaoCaoDoiChieuSoDuTaiKhoanKhachHang
from reporting.trading_service.thanhtoanbutru import BaoCaoDoiChieuTTBTTienMuaBanChungKhoan
from reporting.trading_service.thanhtoanbutru import BaoCaoDoiChieuUTTB
from reporting.trading_service.thanhtoanbutru import BaoCaoDoiChieuVaImportEIB
from reporting.trading_service.thanhtoanbutru import BaoCaoDoiChieuVaImportOCB
from reporting.trading_service.thanhtoanbutru import BaoCaoDuLieuGuiKiemToan
from reporting.trading_service.thanhtoanbutru import BaoCaoFileGhiAmGuiKSNB
from reporting.trading_service.thanhtoanbutru import BaoCaoFileImportTKNganHangBuoiSang
from reporting.trading_service.thanhtoanbutru import BaoCaoKHNNCoSoDuNho
from reporting.trading_service.thanhtoanbutru import BaoCaoPhatSinhGiaoDichTien
from reporting.trading_service.thanhtoanbutru import BaoCaoPSTheoDoiTKKHCanChuyenTienTuVSDVeTKGD
from reporting.trading_service.thanhtoanbutru import BaoCaoReviewKHVIPCoSo
from reporting.trading_service.thanhtoanbutru import BaoCaoSoDuTienKyQuyPhaiSinh
from reporting.trading_service.thanhtoanbutru import BaoCaoSoLieuThongKeDVTC
from reporting.trading_service.thanhtoanbutru import BaoCaoUngTienROS
from reporting.trading_service.thanhtoanbutru import BaoCaoXacNhanNopRutEIBOCB
from reporting.trading_service.thanhtoanbutru import DanhSachKHVIPCoSo
from reporting.trading_service.thanhtoanbutru import DanhSachKHVIPPhaiSinh

@TaskMonitor
def imp_BaoCaoCheckGia():
    BaoCaoCheckGia.run()

@TaskMonitor
def imp_BaoCaoDanhSachChungKhoanSSI():
    BaoCaoDanhSachChungKhoanSSI.run()

@TaskMonitor
def imp_BaoCaoDongMoUyQuyenTK_HNX():
    BaoCaoDongMoUyQuyenTK_HNX.run()

@TaskMonitor
def imp_BaoCaoDongMoUyQuyenTK_HOSEnew():
    BaoCaoDongMoUyQuyenTK_HOSEnew.run()

@TaskMonitor
def imp_BaoCaoDongMoUyQuyenTK_HOSEold():
    BaoCaoDongMoUyQuyenTK_HOSEold.run()

@TaskMonitor
def imp_BaoCaoHoatDongLuuKyNDTNN():
    BaoCaoHoatDongLuuKyNDTNN.run()

@TaskMonitor
def imp_BaoCaoTamTinhPhiChuyenKhoan():
    BaoCaoTamTinhPhiChuyenKhoan.run()

@TaskMonitor
def imp_BaoCaoTamTinhPhiGiaoDich():
    BaoCaoTamTinhPhiGiaoDich.run()

@TaskMonitor
def imp_BaoCaoTamTinhPhiLuuKy():
    BaoCaoTamTinhPhiLuuKy.run()

@TaskMonitor
def imp_BaoCaoTienGuiSSC():
    BaoCaoTienGuiSSC.run()

@TaskMonitor
def imp_BaoCaoTinhHinhHDKD():
    BaoCaoTinhHinhHDKD.run()

@TaskMonitor
def imp_BaoCaoDoiChieuLaiTienGuiPhatSinh_p1():
    BaoCaoDoiChieuLaiTienGuiPhatSinh_p1.run(dt.datetime.now()-dt.timedelta(days=1))

@TaskMonitor
def imp_BaoCaoDoiChieuLaiTienGuiPhatSinh_p2():
    BaoCaoDoiChieuLaiTienGuiPhatSinh_p2.run()

@TaskMonitor
def imp_BaoCaoDoiChieuSoDuTaiKhoanKhachHang():
    BaoCaoDoiChieuSoDuTaiKhoanKhachHang.run()

@TaskMonitor
def imp_BaoCaoDoiChieuTTBTTienMuaBanChungKhoan():
    BaoCaoDoiChieuTTBTTienMuaBanChungKhoan.run()

@TaskMonitor
def imp_BaoCaoDoiChieuUTTB():
    BaoCaoDoiChieuUTTB.run()

@TaskMonitor
def imp_BaoCaoDoiChieuVaImportEIB():
    BaoCaoDoiChieuVaImportEIB.run()

@TaskMonitor
def imp_BaoCaoDoiChieuVaImportOCB():
    BaoCaoDoiChieuVaImportOCB.run()

@TaskMonitor
def imp_BaoCaoDuLieuGuiKiemToan():
    BaoCaoDuLieuGuiKiemToan.run()

@TaskMonitor
def imp_BaoCaoFileGhiAmGuiKSNB():
    BaoCaoFileGhiAmGuiKSNB.run()

@TaskMonitor
def imp_BaoCaoFileImportTKNganHangBuoiSang():
    BaoCaoFileImportTKNganHangBuoiSang.run()

@TaskMonitor
def imp_BaoCaoKHNNCoSoDuNho():
    BaoCaoKHNNCoSoDuNho.run()

@TaskMonitor
def imp_BaoCaoPhatSinhGiaoDichTien():
    BaoCaoPhatSinhGiaoDichTien.run()

@TaskMonitor
def imp_BaoCaoPSTheoDoiTKKHCanChuyenTienTuVSDVeTKGD():
    BaoCaoPSTheoDoiTKKHCanChuyenTienTuVSDVeTKGD.run()

@TaskMonitor
def imp_BaoCaoReviewKHVIPCoSo():
    BaoCaoReviewKHVIPCoSo.run()

@TaskMonitor
def imp_BaoCaoSoDuTienKyQuyPhaiSinh():
    BaoCaoSoDuTienKyQuyPhaiSinh.run()

@TaskMonitor
def imp_BaoCaoSoLieuThongKeDVTC():
    BaoCaoSoLieuThongKeDVTC.run()

@TaskMonitor
def imp_BaoCaoUngTienROS():
    BaoCaoUngTienROS.run()

@TaskMonitor
def imp_BaoCaoXacNhanNopRutEIBOCB():
    BaoCaoXacNhanNopRutEIBOCB.run()

@TaskMonitor
def imp_DanhSachKHVIPCoSo():
    DanhSachKHVIPCoSo.run()

@TaskMonitor
def imp_DanhSachKHVIPPhaiSinh():
    DanhSachKHVIPPhaiSinh.run()


"""
Cần viết tool email báo có chênh lệch trong các BC đối chiếu 
(as decorators that wrap functions from reporting package)
"""
