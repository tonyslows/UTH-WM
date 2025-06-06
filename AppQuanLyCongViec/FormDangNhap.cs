using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AppQuanLyCongViec
{
    public partial class FormDangNhap : Form
    {
        List<TaiKhoan> listTaiKhoan = DanhSachTaiKhoan.Instance.ListTaiKhoan;
        public FormDangNhap()
        {
            InitializeComponent();
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnDangNhap_Click(object sender, EventArgs e)
        {
            if (check_ID(txb_dangNhap.Text, txb_matKhau.Text))
            {
                FormChuongTrinh f = new FormChuongTrinh();
                f.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("Sai tên tài khoản hoặc mật khẩu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txb_dangNhap.Focus();
            }
        }

        bool check_ID(string tentaikhoan, string matkhau)
        {
            for(int i =0; i < listTaiKhoan.Count; i++)
            {
                if (tentaikhoan == listTaiKhoan[i].TenTK && matkhau == listTaiKhoan[i].MatKhau)
                {
                    Const.TaiKhoan = listTaiKhoan[i];
                    return true;
                }
            }

            return false;
        }
    }
}
