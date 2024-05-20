using ExportDocApi.Common;
using ExportDocApi.Models;
using GemBox.Spreadsheet;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace ExportDocApi.Controllers
{
    public class FlightCrewController : BaseController
    {
        // GET: FlightCrew
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="tungay"></param>
        /// <param name="denngay"></param>
        /// <param name="so_giay_to"></param>
        /// <param name="quoc_tich"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult ExportExcel(string tungay, string denngay, string so_giay_to, string quoc_tich, string so_hieu)
        {
            try
            {
                if (string.IsNullOrEmpty(tungay) || string.IsNullOrEmpty(denngay))
                {
                    return Json(new { status = 400, description = "Vui lòng chọn khoảng ngày" }, JsonRequestBehavior.AllowGet);
                }

                SpreadsheetInfo.SetLicense(ConfigKey.KeyGemBoxSpreadsheet);
                var workbook = ExcelFile.Load(Server.MapPath("/FileTemp/DS_ToBay.xlsx"));
                var workSheet = workbook.Worksheets[0];
                //Xử lý dữ liệu
                MySqlConnection conn = new MySqlConnection();
                conn.ConnectionString = ConfigKey.ConnectionString;
                conn.Open();

                string table = "chuyenbay_tobay";


                string infoSelect = "IDCHUYENBAY, IDTOBAY, HO, TENDEM, TEN, SOGIAYTO, LOAIGIAYTO, GIOITINH, QUOCTICH, SOHIEU, NGAYBAY, NGAYSINH, SYS_DATE";

                StringBuilder condition = new StringBuilder();
                condition.Append(string.Format("SYS_DATE between '{0}' and '{1} 23:59:59' ", tungay, denngay));
                if (!string.IsNullOrEmpty(so_giay_to))
                {
                    condition.Append(string.Format("and SOGIAYTO = '{0}' ", so_giay_to));
                }
                if (!string.IsNullOrEmpty(quoc_tich))
                {
                    condition.Append(string.Format("and QUOCTICH = '{0}' ", quoc_tich));
                }
                if (!string.IsNullOrEmpty(so_hieu))
                {
                    condition.Append(string.Format("and SOHIEU = '{0}' ", so_hieu));
                }
                condition.Append("order by IDTOBAY");
                string query = string.Format("select {0} from {1} where {2}", infoSelect, table, condition.ToString());

                var cmd = new MySqlCommand(query, conn);
                var dr = cmd.ExecuteReader();

                var lstTB = new List<ChuyenBay_ToBay_ExportDto>();
                if (dr.HasRows)
                {
                    DataTable dt = new DataTable();
                    dt.Load(dr);
                    lstTB = (from rw in dt.AsEnumerable()
                             select new ChuyenBay_ToBay_ExportDto()
                             {
                                 IdToBay = Convert.ToString(rw["IDTOBAY"]),
                                 IdChuyenBay = Convert.ToString(rw["IDCHUYENBAY"]),
                                 SoHieu = Convert.ToString(rw["SOHIEU"]),
                                 NgayBay = Convert.ToString(rw["NGAYBAY"]),
                                 Ho = Convert.ToString(rw["HO"]),
                                 TenDem = Convert.ToString(rw["TENDEM"]),
                                 Ten = Convert.ToString(rw["TEN"]),
                                 GioiTinh = Convert.ToString(rw["GIOITINH"]),
                                 QuocTich = Convert.ToString(rw["QUOCTICH"]),
                                 NgaySinh = Convert.ToString(rw["NGAYSINH"]),
                                 SoGiayTo = Convert.ToString(rw["SOGIAYTO"]),
                                 LoaiGiayTo = Convert.ToString(rw["LOAIGIAYTO"])
                             }).ToList();
                }

                int row = 5;
                int stt = 1;
                if (lstTB.Any())
                {
                    foreach (var item in lstTB)
                    {
                        string ngayBay = item.NgayBayTxt == "01/01/0001" ? string.Empty : item.NgayBayTxt;
                        workSheet.Cells["A" + row].SetValue(stt);
                        workSheet.Cells["B" + row].SetValue(item.SoHieu);
                        workSheet.Cells["C" + row].SetValue(ngayBay);
                        workSheet.Cells["D" + row].SetValue(item.HoTen);
                        workSheet.Cells["E" + row].SetValue(item.GioiTinhTxt);
                        workSheet.Cells["F" + row].SetValue(item.QuocTich);
                        workSheet.Cells["G" + row].SetValue(item.NgaySinhTxt);
                        workSheet.Cells["H" + row].SetValue(item.LoaiGiayTo);
                        workSheet.Cells["I" + row].SetValue(item.SoGiayTo);
                        stt++;
                        row++;
                    }
                }

                string fTuNgay = General.ConvertStringToDate(tungay).ToString("dd/MM/yyyy");
                string fDenNgay = General.ConvertStringToDate(denngay).ToString("dd/MM/yyyy");

                string filter = string.Format("(Từ ngày: {0} đến {1}; Số giấy tờ: {2}; Quốc tịch: {3}; Số hiệu: {4})", fTuNgay, fDenNgay, so_giay_to, quoc_tich, so_hieu);
                workSheet.Cells["A2"].SetValue(filter);

                var range = workSheet.Cells.GetSubrange("A4", "I" + (row - 1));
                range.Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);
                // xuất tài liệu thành tệp tin
                string handle = Guid.NewGuid().ToString();
                var stream = new MemoryStream();
                workbook.Save(stream, SaveOptions.XlsxDefault);
                stream.Position = 0;
                System.Web.HttpContext.Current.Cache.Insert(handle, stream.ToArray());
                byte[] data = stream.ToArray() as byte[];
                //return File(data, "application/octet-stream", "documentDemo.docx");
                return Json(new { status = 200, fileguid = handle, filename = "DS_ToBay_" + tungay.Replace("-", "") + "_" + denngay.Replace("-", "") + (string.IsNullOrEmpty(so_hieu) ? "" : "_" + so_hieu) + ".xlsx" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { status = 500, description = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }
    }
}