using ExportDocApi.Common;
using ExportDocApi.Models;
using GemBox.Spreadsheet;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ExportDocApi.Controllers
{
    public class PassengerController : BaseController
    {
        // GET: Passenger
        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public JsonResult GetList(string tungay, string denngay, string so_giay_to, string quoc_tich, string so_hieu, string noi_di, string noi_den, int is_trong_diem = -1, int pageNum = 1, int pageSize = 100)
        {
            try
            {
                if (string.IsNullOrEmpty(tungay) || string.IsNullOrEmpty(denngay))
                {
                    return Json(new { status = 400, description = "Vui lòng chọn khoảng ngày" }, JsonRequestBehavior.AllowGet);
                }
                //Xử lý dữ liệu
                MySqlConnection conn = new MySqlConnection();
                conn.ConnectionString = ConfigKey.ConnectionString;
                conn.Open();

                List<string> lstTD = GetListHKTrongDiem(conn);

                string table = "chuyenbay_hanhkhach";
                string infoSelect = "IDCHUYENBAY,FLIGHTDATE,SOHIEU,MADATCHO, HO, TENDEM, TEN, GIOITINH, QUOCTICH, NGAYSINH, LOAIGIAYTO, SOGIAYTO, NOICAP, MANOIDI, MANOIDEN, NOIDI, NOIDEN, HANHLY, SYS_DATE";
                string from = " from " + table;

                string where = string.Empty;
                if (!string.IsNullOrEmpty(tungay))
                {
                    where += " where flightdate between '" + tungay + "' and '" + denngay + "'";
                }

                if (!string.IsNullOrEmpty(noi_di))
                {
                    where += " and NOIDI = '" + noi_di + "'";
                }
                if (!string.IsNullOrEmpty(noi_den))
                {
                    where += " and NOIDEN = '" + noi_den + "'";
                }

                if (!string.IsNullOrEmpty(quoc_tich))
                {
                    where += " and QUOCTICH = '" + quoc_tich + "'";
                }
                if (!string.IsNullOrEmpty(so_giay_to))
                {
                    where += " and SOGIAYTO = '" + so_giay_to + "'";
                }

                if (is_trong_diem == 1)
                {
                    where += " and SOGIAYTO in ('" + string.Join("','", lstTD) + "')";
                }
                if (is_trong_diem == 0)
                {
                    where += " and SOGIAYTO not in ('" + string.Join("','", lstTD) + "')";
                }

                if (!string.IsNullOrEmpty(so_hieu))
                {
                    var arr = so_hieu.Split(',');
                    so_hieu = "'"+string.Join("','", arr)+"'";
                    where += " and SOHIEU in (" + so_hieu + ")";
                }

                string queryCount = "select count(*)" + from + where;
                var cmdCount = new MySqlCommand(queryCount, conn);

                int totalRows = int.Parse(cmdCount.ExecuteScalar().ToString());

                int skip = (pageNum - 1) * pageSize;
                string query = "select " + infoSelect + from + where + $" limit {skip},{pageSize}";
                var cmd = new MySqlCommand(query, conn)
                {
                    CommandTimeout = int.MaxValue
                };
                var dr = cmd.ExecuteReader();

                var lstHK = new List<ChuyenBay_HanhKhach_ExportDto>();
                if (dr.HasRows)
                {
                    DataTable dt = new DataTable();
                    dt.Load(dr);
                    lstHK = (from rw in dt.AsEnumerable()
                             select new ChuyenBay_HanhKhach_ExportDto()
                             {
                                 IDChuyenBay = Convert.ToString(rw["IDCHUYENBAY"]),
                                 NgayBay = Convert.ToString(rw["FLIGHTDATE"]),
                                 SoHieu = Convert.ToString(rw["SOHIEU"]),
                                 SoGhe = Convert.ToString(rw["MADATCHO"]),
                                 Ho = Convert.ToString(rw["HO"]),
                                 TenDem = Convert.ToString(rw["TENDEM"]),
                                 Ten = Convert.ToString(rw["TEN"]),
                                 GioiTinh = Convert.ToString(rw["GIOITINH"]),
                                 QuocTich = Convert.ToString(rw["QUOCTICH"]),
                                 NgaySinh = Convert.ToString(rw["NGAYSINH"]),
                                 SoGiayTo = Convert.ToString(rw["SOGIAYTO"]),
                                 LoaiGiayTo = Convert.ToString(rw["LOAIGIAYTO"]),
                                 NoiCap = Convert.ToString(rw["NOICAP"]),
                                 MaNoiDi = Convert.ToString(rw["MANOIDI"]),
                                 MaNoiDen = Convert.ToString(rw["MANOIDEN"]),
                                 NoiDi = Convert.ToString(rw["NOIDI"]),
                                 NoiDen = Convert.ToString(rw["NOIDEN"]),
                                 HanhLy = Convert.ToString(rw["HANHLY"]),
                                 SYS_DATE = Convert.ToString(rw["SYS_DATE"]),
                                 IsTheoDoi = lstTD.Contains(Convert.ToString(rw["SOGIAYTO"]))
                             }).OrderBy(x => x.FlightDate).ThenBy(x => x.IDChuyenBay).ThenBy(x => x.HoTen).ToList();
                }
                conn.Close();
                return Json(new { status = 200, data = lstHK, totalRow =  totalRows}, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { status = 500, description = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        private List<string> GetListHKTrongDiem(MySqlConnection conn)
        {
            string query = "select so_giay_to from canhbao_hanhkhach ";
            var cmd = new MySqlCommand(query, conn);
            var dr = cmd.ExecuteReader();
            var lstSGT = new List<string>();
            if (dr.HasRows)
            {
                DataTable dt = new DataTable();
                dt.Load(dr);
                lstSGT = (from rw in dt.AsEnumerable()
                         select Convert.ToString(rw["so_giay_to"])).ToList();
            }
            return lstSGT;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="tungay"></param>
        /// <param name="denngay"></param>
        /// <param name="so_giay_to"></param>
        /// <param name="quoc_tich"></param>
        /// <param name="noi_di"></param>
        /// <param name="noi_den"></param>
        /// <param name="is_trong_diem">0:Không theo dõi; 1:Theo dõi</param>
        /// <returns></returns>
        [HttpGet]
        public JsonResult ExportExcel(string tungay, string denngay, string so_giay_to, string quoc_tich, string so_hieu, string noi_di, string noi_den, int is_trong_diem = -1)
        {
            try
            {
                if (string.IsNullOrEmpty(tungay) || string.IsNullOrEmpty(denngay))
                {
                    return Json(new { status = 400, description = "Vui lòng chọn khoảng ngày" }, JsonRequestBehavior.AllowGet);
                }

                SpreadsheetInfo.SetLicense(ConfigKey.KeyGemBoxSpreadsheet);
                var workbook = ExcelFile.Load(Server.MapPath("/FileTemp/DS_HanhKhach.xlsx"));
                var workSheet = workbook.Worksheets[0];
                //Xử lý dữ liệu
                MySqlConnection conn = new MySqlConnection();
                conn.ConnectionString = ConfigKey.ConnectionString;
                conn.Open();

                List<string> lstTD = GetListHKTrongDiem(conn);
                string table = "chuyenbay_hanhkhach";
                string infoSelect = "IDCHUYENBAY,FLIGHTDATE,SOHIEU,MADATCHO, HO, TENDEM, TEN, GIOITINH, QUOCTICH, NGAYSINH, LOAIGIAYTO, SOGIAYTO, NOICAP, MANOIDI, MANOIDEN, NOIDI, NOIDEN, HANHLY, SYS_DATE";
                string from = " from " + table;

                string where = string.Empty;
                if (!string.IsNullOrEmpty(tungay))
                {
                    where += " where flightdate between '" + tungay + "' and '" + denngay + "'";
                }

                if (!string.IsNullOrEmpty(noi_di))
                {
                    where += " and NOIDI = '" + noi_di + "'";
                }
                if (!string.IsNullOrEmpty(noi_den))
                {
                    where += " and NOIDEN = '" + noi_den + "'";
                }

                if (!string.IsNullOrEmpty(quoc_tich))
                {
                    where += " and QUOCTICH = '" + quoc_tich + "'";
                }
                if (!string.IsNullOrEmpty(so_giay_to))
                {
                    where += " and SOGIAYTO = '" + so_giay_to + "'";
                }

                if (is_trong_diem == 1)
                {
                    where += " and SOGIAYTO in ('" + string.Join("','", lstTD) + "')";
                }
                if (is_trong_diem == 0)
                {
                    where += " and SOGIAYTO not in ('" + string.Join("','", lstTD) + "')";
                }

                if (!string.IsNullOrEmpty(so_hieu))
                {
                    var arr = so_hieu.Split(',');
                    so_hieu = "'" + string.Join("','", arr) + "'";
                    where += " and SOHIEU in (" + so_hieu + ")";
                }

                string query = "select " + infoSelect + from + where;
                var cmd = new MySqlCommand(query, conn)
                {
                    CommandTimeout = int.MaxValue
                };
                var dr = cmd.ExecuteReader();

                var lstHK = new List<ChuyenBay_HanhKhach_ExportDto>();
                if (dr.HasRows)
                {
                    DataTable dt = new DataTable();
                    dt.Load(dr);
                    lstHK = (from rw in dt.AsEnumerable()
                             select new ChuyenBay_HanhKhach_ExportDto()
                             {
                                 IDChuyenBay = Convert.ToString(rw["IDCHUYENBAY"]),
                                 NgayBay = Convert.ToString(rw["FLIGHTDATE"]),
                                 SoHieu = Convert.ToString(rw["SOHIEU"]),
                                 SoGhe = Convert.ToString(rw["MADATCHO"]),
                                 Ho = Convert.ToString(rw["HO"]),
                                 TenDem = Convert.ToString(rw["TENDEM"]),
                                 Ten = Convert.ToString(rw["TEN"]),
                                 GioiTinh = Convert.ToString(rw["GIOITINH"]),
                                 QuocTich = Convert.ToString(rw["QUOCTICH"]),
                                 NgaySinh = Convert.ToString(rw["NGAYSINH"]),
                                 SoGiayTo = Convert.ToString(rw["SOGIAYTO"]),
                                 LoaiGiayTo = Convert.ToString(rw["LOAIGIAYTO"]),
                                 NoiCap = Convert.ToString(rw["NOICAP"]),
                                 MaNoiDi = Convert.ToString(rw["MANOIDI"]),
                                 MaNoiDen = Convert.ToString(rw["MANOIDEN"]),
                                 NoiDi = Convert.ToString(rw["NOIDI"]),
                                 NoiDen = Convert.ToString(rw["NOIDEN"]),
                                 HanhLy = Convert.ToString(rw["HANHLY"]),
                                 SYS_DATE = Convert.ToString(rw["SYS_DATE"]),
                                 IsTheoDoi = lstTD.Contains(Convert.ToString(rw["SOGIAYTO"]))
                             }).OrderBy(x => x.FlightDate).ThenBy(x => x.IDChuyenBay).ThenBy(x => x.HoTen).ToList();
                }

                int row = 5;
                int stt = 1;
                if (lstHK.Any())
                {
                    foreach (var item in lstHK)
                    {
                        workSheet.Cells["A" + row].SetValue(stt);
                        workSheet.Cells["B" + row].SetValue(item.NgayBayTxt);
                        workSheet.Cells["C" + row].SetValue(item.SoHieu);
                        workSheet.Cells["D" + row].SetValue(item.SoGhe);
                        workSheet.Cells["E" + row].SetValue(item.HoTen);
                        workSheet.Cells["F" + row].SetValue(item.GioiTinhTxt);
                        workSheet.Cells["G" + row].SetValue(item.QuocTich);
                        workSheet.Cells["H" + row].SetValue(item.NgaySinhTxt);
                        workSheet.Cells["I" + row].SetValue(item.LoaiGiayTo);
                        workSheet.Cells["J" + row].SetValue(item.SoGiayTo);
                        workSheet.Cells["K" + row].SetValue(item.NoiCap);
                        workSheet.Cells["L" + row].SetValue(item.NoiDi);
                        workSheet.Cells["M" + row].SetValue(item.NoiDen);
                        workSheet.Cells["N" + row].SetValue(item.HanhLy);

                        if (item.IsTheoDoi)
                        {
                            var rowCr = workSheet.Cells.GetSubrange("A" + row, "N" + row);
                            rowCr.Style.Font.Color = SpreadsheetColor.FromName(ColorName.Red);
                        }

                        stt++;
                        row++;
                    }
                }
                conn.Close();
                string fTuNgay = General.ConvertStringToDate(tungay).ToString("dd/MM/yyyy");
                string fDenNgay = General.ConvertStringToDate(denngay).ToString("dd/MM/yyyy");
                string fIsTrongDiem = is_trong_diem == 1 ? "Là đối tượng theo dõi" : (is_trong_diem == 0 ? "Không phải đối tượng theo dõi" : "");

                string filter = string.Format("(Từ ngày: {0} đến {1}; Số giấy tờ: {2}; Quốc tịch: {3}; {4}; Số hiệu: {5}; Nơi đi:{6}; Nơi đến:{7})", fTuNgay, fDenNgay, so_giay_to, quoc_tich, fIsTrongDiem, so_hieu, noi_di, noi_den);
                workSheet.Cells["A2"].SetValue(filter);

                var range = workSheet.Cells.GetSubrange("A4", "N" + (row - 1));
                range.Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);
                // xuất tài liệu thành tệp tin
                string handle = Guid.NewGuid().ToString();
                var stream = new MemoryStream();
                workbook.Save(stream, SaveOptions.XlsxDefault);
                stream.Position = 0;
                System.Web.HttpContext.Current.Cache.Insert(handle, stream.ToArray());
                byte[] data = stream.ToArray() as byte[];
                //return File(data, "application/octet-stream", "documentDemo.docx");
                return Json(new { status = 200, fileguid = handle, filename = "DS_HanhKhach_" + tungay.Replace("-", "") + "_" + denngay.Replace("-", "") + (string.IsNullOrEmpty(so_hieu) ? "" : "_" + so_hieu) + ".xlsx" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { status = 500, description = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }
    }
}