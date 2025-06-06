using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace AppQuanLyCongViec
{
    public class DanhSachTaiKhoan
    {
        private static DanhSachTaiKhoan instance;
        public static DanhSachTaiKhoan Instance 
        { 
            get
            {
                if(instance == null)
                    instance = new DanhSachTaiKhoan(); 
                return instance; 
            }
            set => instance = value; 
        }

        List<TaiKhoan> listTaiKhoan;

        public List<TaiKhoan> ListTaiKhoan
        {
            get => listTaiKhoan; 
            set => listTaiKhoan = value;   
        }
        
        DanhSachTaiKhoan()
        {
            listTaiKhoan = new List<TaiKhoan>();
            listTaiKhoan.Add(new TaiKhoan("anhkhai","123456",TaiKhoan.LoaiTK.admin));
            listTaiKhoan.Add(new TaiKhoan("thanhphu", "1234567", TaiKhoan.LoaiTK.manager));
            listTaiKhoan.Add(new TaiKhoan("quocdung", "12345678", TaiKhoan.LoaiTK.employee));
            listTaiKhoan.Add(new TaiKhoan("nhatanh", "123456789", TaiKhoan.LoaiTK.employee));

        }
    }
}
