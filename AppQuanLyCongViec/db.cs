using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDataReader;
using Microsoft.Office.Interop.Excel;
using ex = Microsoft.Office.Interop.Excel;

namespace AppQuanLyCongViec
{
    public class db
    {

        #region  biến
        string strPath_database = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data\data.xlsx");
        string strPath_baoCao = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"HoSo\MauBaoCao.xlsx");

        public System.Data.DataSet ds;
        public string dataHienHanh()
        {
            //code doc de lay data hien hanh
            return strPath_database;
        }
        #endregion

        #region  hàm 

        public void docFile_Excel()
        {
            try
            {
                string _file = strPath_database;    // duong dan cua data
                IExcelDataReader reader;
                if (_file.Contains(".xls") || _file.Contains(".xlsm"))
                {
                    using (var stream = File.Open(_file, FileMode.Open, FileAccess.Read))
                    {
                        reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                        ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true
                            }
                        });
                        reader.Close();
                    }
                }
                else if (_file.Contains(".xls"))
                {
                    using (var stream = File.Open(_file, FileMode.Open, FileAccess.Read))
                    {
                        reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                        ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true
                            }
                        });
                        reader.Close();
                    }

                }
                else
                {
                    MessageBox.Show("Hãy chọn file Excel ?");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        public void themCongViec(string tenSheet, bool bl, int rw, string ma, string ten, string noiDung, string nhanVien,
            string batDau, string ketThuc, string tongThoiGian, string ketQua)
        {
            string fullNameData = dataHienHanh();
            ex.Application exApp = new ex.Application();
            ex.Workbook wbdata = exApp.Application.Workbooks.Open(fullNameData, ReadOnly: false);
            ex.Worksheet sh = wbdata.Sheets[tenSheet];
            int tongDong = sh.Range["a:a"].Rows.Count;
            int lr = sh.Cells[tongDong, 1].end[ex.XlDirection.xlUp].row + 1;
            _off_Scr(exApp);
            // code
            if (ktCongViec(sh,lr,ma))
            {
                MessageBox.Show("Mã hiệu công việc đã tồn tại !", "Thông báo");
            }
            else
            {
                if (bl)
                {
                    //lưu vào vị trí bất kỳ
                    #region MyRegion
                    rw = rw + 3;
                    sh.Rows[rw + ":" + rw].insert();
                    ex.Range rng = sh.Range["a" + rw];
                    rng.Value2 = ma;
                    rng.Offset[0, 1].Value2 = ten;
                    rng.Offset[0, 2].Value2 = noiDung;
                    rng.Offset[0, 3].Value2 = nhanVien;
                    rng.Offset[0, 4].Value2 = batDau;
                    rng.Offset[0, 5].Value2 = ketThuc;
                    rng.Offset[0, 6].Value2 = tongThoiGian;
                    rng.Offset[0, 7].Value2 = ketQua;  
                    MessageBox.Show("Hoàn thành !", "Thông báo");
                    #endregion

                }
                else
                {
                    //lưu vào vị trí cuối cùng (tự động)
                    #region MyRegion
                    ex.Range rng = sh.Range["a" + lr];
                    rng.Value2 = ma;
                    rng.Offset[0, 1].Value2 = ten;
                    rng.Offset[0, 2].Value2 = noiDung;
                    rng.Offset[0, 3].Value2 = nhanVien;
                    rng.Offset[0, 4].Value2 = batDau;
                    rng.Offset[0, 5].Value2 = ketThuc;
                    rng.Offset[0, 6].Value2 = tongThoiGian;
                    rng.Offset[0, 7].Value2 = ketQua;
                    MessageBox.Show("Hoàn thành !", "Thông báo");
                    #endregion
                }
            }
            _on_Scr(exApp);
            wbdata.Save();
            // giai phong tai nguyen la cac trinh ex chay ngam
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wbdata);
            exApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(exApp);
        }

        public bool ktCongViec(ex.Worksheet sh, int lr, string maCV)
        {
            bool bl = false;
            for(int i = 2; i < lr; i++)
            {
                if (sh.Cells[i,1].text == maCV)
                {
                    bl = true;
                    rwAC = i;
                    break;
                }
            }
            return bl;
        }

        public void suaCongViec(string tenSheet, int rw, string ma, string ten, string noiDung, string nhanVien,
            string batDau, string ketThuc, string tongThoiGian, string ketQua)
        {
            string fullNameData = dataHienHanh();
            ex.Application exApp = new ex.Application();
            ex.Workbook wbdata = exApp.Application.Workbooks.Open(fullNameData, ReadOnly: false);
            ex.Worksheet sh = wbdata.Sheets[tenSheet];
            int tongDong = sh.Range["a:a"].Rows.Count;
            int lr = sh.Cells[tongDong, 1].end[ex.XlDirection.xlUp].row + 1;
            _off_Scr(exApp);
            // code
            if (ktCongViec(sh, lr, ma))
            {
                //lưu vào vị trí bất kỳ
                #region MyRegion
                rw = rw + 2;
                ex.Range rng = sh.Range["a" + rw];
                rng.Value2 = ma;
                rng.Offset[0, 1].Value2 = ten;
                rng.Offset[0, 2].Value2 = noiDung;
                rng.Offset[0, 3].Value2 = nhanVien;
                rng.Offset[0, 4].Value2 = batDau;
                rng.Offset[0, 5].Value2 = ketThuc;
                rng.Offset[0, 6].Value2 = tongThoiGian;
                rng.Offset[0, 7].Value2 = ketQua;
                MessageBox.Show("Đã cập nhật !", "Thông báo");
                #endregion
            }
            else
            {
                MessageBox.Show("Mã hiệu công việc không tồn tồn tại !", "Thông báo");
            }
            _on_Scr(exApp);
            wbdata.Save();
            // giai phong tai nguyen la cac trinh ex chay ngam
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wbdata);
            exApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(exApp);
        }

        int rwAC = 0;
        public void xoaCongViec(string tenSheet, string ma)
        {
            string fullNameData = dataHienHanh();
            if (fullNameData != "")
            {
                ex.Application exApp = new ex.Application();
                ex.Workbook wbdata = exApp.Application.Workbooks.Open(fullNameData, ReadOnly: false);
                ex.Worksheet sh = wbdata.Sheets[tenSheet];
                int tongDong = sh.Range["a:a"].Rows.Count;
                int lr = sh.Cells[tongDong, 1].end[ex.XlDirection.xlUp].row + 1;
                _off_Scr(exApp);

                if(ktCongViec(sh,lr,ma))
                {
                    sh.Rows[rwAC + ":" + rwAC].delete();
                }
                /* else
                 {
                     MessageBox.Show("Không tồn tại mã công việc");
                 }*/
                _on_Scr(exApp);
                wbdata.Save();
                // giai phong tai nguyen la cac trinh ex chay ngam
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wbdata);
                exApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(exApp);
            }   
        }

        public void _off_Scr(ex.Application exApp)
        {
            exApp.ScreenUpdating = false;
            exApp.EnableEvents = false;
            exApp.DisplayAlerts = false;
            exApp.Calculation = XlCalculation.xlCalculationManual;
        }

        public void _on_Scr(ex.Application exApp)
        {
            exApp.ScreenUpdating = true;
            exApp.EnableEvents = true;
            exApp.DisplayAlerts = true;
            exApp.Calculation = XlCalculation.xlCalculationAutomatic;
        }

        ex.Application exApp = new ex.Application();
        ex.Workbook wbBaoCao;
        public void _openBaoCao()
        {
            if(File.Exists(strPath_baoCao))
            {
                wbBaoCao = exApp.Application.Workbooks.Open(strPath_baoCao, ReadOnly: true);
                exApp.Visible = true;
            }
            else
            {
                MessageBox.Show("Không cáo mẫu báo cáo", "Thông báo");
            }
        }

        public void xuatBaoCao(DataGridView dgv)
        {
            try
            {
                if (wbBaoCao != null)
                {
                    // vong for
                    ex.Worksheet shBC = wbBaoCao.Sheets["BaoCao"];
                    int rd = 6, stt = 1;
                    _off_Scr(exApp);
                    int totalRw = shBC.Range["b:b"].Rows.Count;
                    int lr = shBC.Cells[totalRw, 2].End[ex.XlDirection.xlUp].Row + 1;
                    if(lr >= rd)
                    {
                        shBC.Rows[rd + ":" + lr].delete();
                    }
                    foreach (DataGridViewRow rw in dgv.Rows)
                    {
                        bool bl = Convert.ToBoolean(rw.Cells[0].Value);
                        if (bl)
                        {
                            Range rngA = shBC.Range["a" + rd];
                            rngA.Value = stt;
                            rngA.Offset[0, 1].Value = rw.Cells[1].Value.ToString();
                            rngA.Offset[0, 2].Value = rw.Cells[2].Value.ToString();
                            rngA.Offset[0, 3].Value = rw.Cells[3].Value.ToString();
                            rngA.Offset[0, 4].Value = rw.Cells[4].Value.ToString();

                            // timeBatDau
                            if (rw.Cells[5].Value != null && DateTime.TryParse(rw.Cells[5].Value.ToString(), out DateTime time1))
                            {
                                rngA.Offset[0, 5].Value = time1.ToString("hh:mm");
                            }
                            else
                            {
                                rngA.Offset[0, 5].Value = "";
                            }
                            // timeKetThuc
                            if (rw.Cells[6].Value != null && DateTime.TryParse(rw.Cells[6].Value.ToString(), out DateTime time2))
                            {
                                rngA.Offset[0, 6].Value = time2.ToString("hh:mm");
                            }
                            else
                            {
                                rngA.Offset[0, 6].Value = "";
                            }
                            // tongThoiGian
                            if (rw.Cells[7].Value != null && DateTime.TryParse(rw.Cells[7].Value.ToString(), out DateTime tongTG))
                            {
                                rngA.Offset[0, 7].Value = tongTG.ToString("hh:mm");
                            }
                            else
                            {
                                rngA.Offset[0, 7].Value = "";
                            }

                            rngA.Offset[0, 8].Value = rw.Cells[8].Value.ToString();
                            shBC.Range["a" + rd + ":L" + rd].WrapText = true;


                            stt++;
                            rd++;
                        }
                    }
                    _on_Scr(exApp);
                }
                else
                {
                    MessageBox.Show("Hãy mở mẫu báo cáo", "Thông báo");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                _on_Scr(exApp);
            }
        }

        
        #endregion
    }
}