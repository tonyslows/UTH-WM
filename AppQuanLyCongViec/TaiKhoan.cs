using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Markup;

namespace AppQuanLyCongViec
{
    public class TaiKhoan
    {
        private string tenTK;

        public string TenTK
        {
            get => tenTK;
            set => tenTK = value;  
        }

        private string matKhau;
        public string MatKhau
        {
            get => matKhau;
            set => matKhau = value;
        }

        public enum LoaiTK
        { 
            admin,
            manager,
            employee
        }

        private LoaiTK loaiTaiKhoan;
        public LoaiTK LoaiTaiKhoan
        {
            get => loaiTaiKhoan; 
            set => loaiTaiKhoan = value;
        }

        private string tenHienThi;

        public string TenHienThi
        {
            get
            {
                switch (loaiTaiKhoan)
                {
                    case LoaiTK.admin:
                        tenHienThi = "Admin";
                        break;
                    case LoaiTK.manager:
                        tenHienThi = "Manager";
                        break;
                    default:
                        tenHienThi = "Nhân viên";
                        break;
                }
                return tenHienThi;
            }
            set => tenHienThi = value;
        }

        public TaiKhoan(string tentaikhoan, string matkhau, LoaiTK loaitaikhoan)
        {
            this.tenTK = tentaikhoan;
            this.matKhau = matkhau; 
            this.loaiTaiKhoan = loaitaikhoan; 
        }
    }
}
