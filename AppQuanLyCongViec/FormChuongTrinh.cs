using AppQuanLyCongViec;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AppQuanLyCongViec
{
    public partial class FormChuongTrinh : Form
    {
        public FormChuongTrinh()
        {
            InitializeComponent();
        }

        db _db = new db();
        private void cbb_listSheet_DropDown(object sender, EventArgs e)
        {
            cbb_listSheet.Items.Clear();
            _db.docFile_Excel();
            foreach (DataTable tb in _db.ds.Tables)
            {
                string strName = tb.TableName; // tb.TableName.ToUpper();
                if (strName.Contains("ds") || strName.Contains("DANHSACH"))
                {
                    continue;
                }
                else
                {
                    cbb_listSheet.Items.Add(strName);
                }
            }
        }

        private void cbb_listSheet_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                dgv_cv.DataSource = _db.ds.Tables[cbb_listSheet.Text];
                hienThiNV();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void hienThiNV()
        {
            lv_dsNV.Items.Clear();
            int stt = 0;
            foreach (DataRow rw in _db.ds.Tables["dsNV"].Rows)
            {
                stt++;
                ListViewItem it = new ListViewItem();
                it.Text = stt.ToString();
                it.SubItems.Add(rw[0].ToString());
                lv_dsNV.Items.Add(it);
            }
        }

        private void layMaCV()
        {
            try
            {
                string strNgay = "";
                strNgay = dtp_ngayCV.Text;
                strNgay = strNgay.Replace("/", "");

                if (cbb_soLuongCV.Text == "0")
                {
                    txb_maCV.Text = strNgay + "_" + cbb_sttCV.Text;
                }
                else
                {
                    txb_maCV.Text = strNgay + "_" + cbb_sttCV.Text + "." + cbb_soLuongCV.Text;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void dtp_ngayCV_ValueChanged(object sender, EventArgs e)
        {
            layMaCV();
        }

        private void cbb_sttCV_SelectedValueChanged(object sender, EventArgs e)
        {
            layMaCV();
        }

        private void cbb_soLuongCV_SelectedValueChanged(object sender, EventArgs e)
        {
            layMaCV();
        }

        int nr = 0;
        private void dgv_cv_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (dgv_cv.CurrentCell.ColumnIndex > 0)
                {
                    int r = dgv_cv.CurrentRow.Index;
                    nr = r;
                    txb_maCV.Text = dgv_cv.Rows[r].Cells[1].Value.ToString();
                    txb_tenCV.Text = dgv_cv.Rows[r].Cells[2].Value.ToString();
                    txb_ndCV.Text = dgv_cv.Rows[r].Cells[3].Value.ToString();
                    
                    int[] dateColumns = { 5, 6, 7 }; 
                    foreach (int colIndex in dateColumns)
                    {
                        object cellValue = dgv_cv.Rows[r].Cells[colIndex].Value;
                        DateTime dateValue;

                        if (cellValue != null && DateTime.TryParse(cellValue.ToString(), out dateValue))
                        {
                            dgv_cv.Rows[r].Cells[colIndex].Value = dateValue.ToString("HH:mm");
                        }
                    }

                    cbb_ketQua.Text = dgv_cv.Rows[r].Cells[8].Value.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        void PhanQuyen()
        {
            switch (Const.TaiKhoan.LoaiTaiKhoan)
            {
                case TaiKhoan.LoaiTK.manager:
                    btn_qltk.Enabled = false;
                    break;
                case TaiKhoan.LoaiTK.employee:
                    btn_qltk.Enabled = btn_themData.Enabled = btn_Xoa.Enabled = btn_xuatBaoCao.Enabled = btn_moBaoCao.Enabled = false;
                    dgv_cv.ReadOnly = true;
                    break;
                default:
                    break;
            }
            txb_loaiTK.Text = Const.TaiKhoan.TenHienThi;
            txb_loaiTK.ReadOnly = true;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                PhanQuyen();
                dateTime_batdau.Format = DateTimePickerFormat.Custom;
                dateTime_batdau.CustomFormat = "hh:mm";
                dateTime_ketthuc.Format = DateTimePickerFormat.Custom;
                dateTime_ketthuc.CustomFormat = "hh:mm";
                dgv_cv.Columns[5].DefaultCellStyle.Format = "hh:mm";
                dgv_cv.Columns[6].DefaultCellStyle.Format = "hh:mm";
                dgv_cv.Columns[7].DefaultCellStyle.Format = "hh:mm";
                dgv_cv.CellValueChanged += dgv_cv_CellValueChanged;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void btn_themData_Click(object sender, EventArgs e)
        {
            try
            {
                string nameDb = cbb_listSheet.Text;
                string ma = "", ten = "", noiDung = "", nhanVien = "", batDau = "", ketThuc = "", tongThoiGian = "", ketQua = "";

                ma = txb_maCV.Text;
                ten = txb_tenCV.Text;
                noiDung = txb_ndCV.Text;
                nhanVien = "";
                tongThoiGian = "";
                ketQua = cbb_ketQua.Text;

                //lay nhan vien
                foreach (ListViewItem item in lv_dsNV.Items)
                {
                    if (item.Checked)
                    {
                        if (nhanVien != "")
                        {
                            nhanVien = nhanVien + "; " + item.SubItems[1].Text;
                        }
                        else
                        {
                            nhanVien = item.SubItems[1].Text;
                        }
                    }
                }

                //lay thoi gian
                //DateTime la kieu du lieu lay mot thoi diem cu the trong thoi gian
                //TimeSpan la kieu du lieu lay mot khoang thoi gian
                DateTime t1 = dateTime_batdau.Value;
                DateTime t2 = dateTime_ketthuc.Value;
                batDau = t1.ToString("HH:mm");
                ketThuc = t2.ToString("HH:mm");
                TimeSpan t = new TimeSpan();
                t = t2 - t1;
                tongThoiGian = t.ToString(@"hh\:mm");


                _db.themCongViec(nameDb, bl_vitri, nr, ma, ten, noiDung, nhanVien, batDau, ketThuc, tongThoiGian, ketQua);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        bool bl_vitri = false;
        private void chk_vitriluu_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_vitriluu.Checked)
            {
                bl_vitri = true;
            }
            else
            {
                bl_vitri = false;
            }
        }

        private void dgv_cv_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (dgv_cv.CurrentCell.ColumnIndex > 0)
            {
                string nameDb = cbb_listSheet.Text;
                string ma = dgv_cv.CurrentRow.Cells[1].Value.ToString();
                string ten = dgv_cv.CurrentRow.Cells[2].Value.ToString();
                string noiDung = dgv_cv.CurrentRow.Cells[3].Value.ToString();
                string nhanVien = dgv_cv.CurrentRow.Cells[4].Value.ToString();
                string batDau = dgv_cv.CurrentRow.Cells[5].Value.ToString();
                string ketThuc = dgv_cv.CurrentRow.Cells[6].Value.ToString();
                string tongThoiGian = dgv_cv.CurrentRow.Cells[7].Value.ToString();
                string ketQua = dgv_cv.CurrentRow.Cells[8].Value.ToString();


                _db.suaCongViec(nameDb, nr, ma, ten, noiDung, nhanVien, batDau, ketThuc, tongThoiGian, ketQua);
            }

        }

        private void btn_Xoa_Click(object sender, EventArgs e)
        {
            try
            {
                bool bldem = false;
                foreach (DataGridViewRow rw in dgv_cv.Rows)
                {
                    bool blSel = Convert.ToBoolean(rw.Cells[0].Value);
                    if (blSel)
                    {
                        // xoa
                        string ma = rw.Cells[1].Value.ToString();
                        // add vao ham xoa data
                        _db.xoaCongViec(cbb_listSheet.Text, ma);
                        bldem = true;
                    }
                }
                if (bldem)
                {
                    MessageBox.Show("Xoá thành công", "Thông báo");
                }
                else
                {
                    MessageBox.Show("Chưa có công việc nào được chọn", "Thông báo");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void btn_moBaoCao_Click(object sender, EventArgs e)
        {
            _db._openBaoCao();
        }

        private void btn_xuatBaoCao_Click(object sender, EventArgs e)
        {
            _db.xuatBaoCao(dgv_cv);
        }

        private void btn_thoatCT_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void txb_timKiem_TextChanged(object sender, EventArgs e)
        {
            (dgv_cv.DataSource as DataTable).DefaultView.RowFilter =
            string.Format("TenCV LIKE '%{0}%' OR MaCV LIKE '%{0}%'", txb_timKiem.Text);
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://ut.edu.vn");
        }

        public bool isThoat = true;
        private void btn_dangXuat_Click(object sender, EventArgs e)
        {
            isThoat = false;
            this.Close();
            FormDangNhap f = new FormDangNhap();
            f.Show();
        }

        private void FormChuongTrinh_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (isThoat)  
                Application.Exit();
        }

        private void FormChuongTrinh_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(isThoat)
            {
                if(MessageBox.Show("Bạn có thật sự muốn thoát chương trình?", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) != System.Windows.Forms.DialogResult.OK)
                    e.Cancel = true;
            }
        }

        private void dgv_cv_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0) 
            {
                int colBatDau = 5;  
                int colKetThuc = 6; 
                int colTongThoiGian = 7; 

                if (e.ColumnIndex == colBatDau || e.ColumnIndex == colKetThuc)
                {
                    tinhTongTG(e.RowIndex, colBatDau, colKetThuc, colTongThoiGian);
                }
            }
        }

        private void tinhTongTG(int rowIndex, int colBatDau, int colKetThuc, int colTongThoiGian)
        {
            DataGridViewRow row = dgv_cv.Rows[rowIndex];

            if (row.Cells[colBatDau].Value != null && row.Cells[colKetThuc].Value != null)
            {
                DateTime batDau, ketThuc;
                if (DateTime.TryParse(row.Cells[colBatDau].Value.ToString(), out batDau) &&
                    DateTime.TryParse(row.Cells[colKetThuc].Value.ToString(), out ketThuc))
                {
                    TimeSpan duration = ketThuc - batDau;
                    if (duration.TotalMinutes < 0) duration = TimeSpan.Zero;

                    row.Cells[colTongThoiGian].Value = duration.ToString(@"hh\:mm");
                }
            }
        }
    }
}