using ExportDocApi.Common;
using ExportDocApi.Models;
using GemBox.Spreadsheet;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ExportDocApi.Controllers
{
    public class PassengerConnectingController : BaseController
    {
        // GET: PassengerConnecting
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public JsonResult ExportExcel(string tungay, string denngay, string so_giay_to, string quoc_tich, string so_hieu, int is_trong_diem = -1)
        {
            try
            {
                if (string.IsNullOrEmpty(tungay) || string.IsNullOrEmpty(denngay))
                {
                    return Json(new { status = 400, description = "Vui lòng chọn khoảng ngày" }, JsonRequestBehavior.AllowGet);
                }

                SpreadsheetInfo.SetLicense(ConfigKey.KeyGemBoxSpreadsheet);
                var workbook = ExcelFile.Load(Server.MapPath("/FileTemp/DS_HanhKhach_NoiChuyen.xlsx"));
                var workSheet = workbook.Worksheets[0];
                //Xử lý dữ liệu
                MySqlConnection conn = new MySqlConnection();
                conn.ConnectionString = ConfigKey.ConnectionString;
                conn.Open();

                string table = "hanhkhach_noichuyen";
                string tableJoin = "canhbao_hanhkhach";

                string infoSelect = string.Format("{0}.ID_CHUYENBAY,{0}.SOHIEU,MADATCHO, {0}.HO, {0}.TENDEM, {0}.TEN, {0}.GIOITINH, {0}.QUOCTICH, {0}.NGAYSINH, {0}.LOAIGIAYTO, {0}.SOGIAYTO, {0}.NOICAP, {0}.MANOIDI, {0}.MANOIDEN, {0}.HANHLY,{1}.is_trong_diem as GhiChu", table, tableJoin);
                string query = string.Format("select {0} from {1}", infoSelect, table);
                query += string.Format(" left join {0} on {0}.so_giay_to = {1}.SOGIAYTO AND {0}.loai_giay_to = {1}.LOAIGIAYTO", tableJoin, table);
                query += string.Format(" WHERE {0}.`FLIGHTDATE` BETWEEN '{1}' AND '{2}'", table, tungay, denngay);
                
                if (!string.IsNullOrEmpty(so_giay_to))
                {
                    query += " and SOGIAYTO = '" + so_giay_to + "'";
                }
                if (!string.IsNullOrEmpty(quoc_tich))
                {
                    query += " and QUOCTICH = '" + quoc_tich + "'";
                }
                if (!string.IsNullOrEmpty(so_hieu))
                {
                    query += " and SOHIEU = '" + so_hieu + "'";
                }

                var cmd = new MySqlCommand(query, conn);
                var dr = cmd.ExecuteReader();

                var lstHK = new List<HanhKhach_NoiChuyen_ExportDto>();
                if (dr.HasRows)
                {
                    DataTable dt = new DataTable();
                    dt.Load(dr);
                    lstHK = (from rw in dt.AsEnumerable()
                             select new HanhKhach_NoiChuyen_ExportDto()
                             {
                                 //IDHanhKhach = Convert.ToString(rw["IDHANHKHACH"]),
                                 SoHieu = Convert.ToString(rw["SOHIEU"]),
                                 MaDatCho = Convert.ToString(rw["MADATCHO"]),
                                 Ho = Convert.ToString(rw["HO"]),
                                 TenDem = Convert.ToString(rw["TENDEM"]),
                                 Ten = Convert.ToString(rw["TEN"]),
                                 GioiTinh = Convert.ToString(rw["GIOITINH"]),
                                 QuocTich = Convert.ToString(rw["QUOCTICH"]),
                                 NgaySinh = Convert.ToString(rw["NGAYSINH"]),
                                 SoGiayTo = Convert.ToString(rw["SOGIAYTO"]),
                                 LoaiGiayTo = Convert.ToString(rw["LOAIGIAYTO"]),
                                 NoiCap = Convert.ToString(rw["NOICAP"]),
                                 NoiDi = Convert.ToString(rw["MANOIDI"]),
                                 NoiDen = Convert.ToString(rw["MANOIDEN"]),
                                 HanhLy = Convert.ToString(rw["HANHLY"]),
                                 GhiChu = Convert.ToString(rw["GhiChu"])
                             }).OrderBy(x => x.SoHieu).ThenBy(x=>x.HoVaTen).ToList();
                }

                int row = 4;
                int stt = 1;
                if (lstHK.Any())
                {
                    foreach (var item in lstHK)
                    {
                        workSheet.Cells["A" + row].SetValue(stt);
                        workSheet.Cells["B" + row].SetValue(item.SoHieu);
                        workSheet.Cells["C" + row].SetValue(item.MaDatCho);
                        workSheet.Cells["D" + row].SetValue(item.HoVaTen);
                        workSheet.Cells["E" + row].SetValue(item.GioiTinhTxt);
                        workSheet.Cells["F" + row].SetValue(item.QuocTich);
                        workSheet.Cells["G" + row].SetValue(item.NgaySinhTxt);
                        workSheet.Cells["H" + row].SetValue(item.LoaiGiayTo);
                        workSheet.Cells["I" + row].SetValue(item.SoGiayTo);
                        workSheet.Cells["J" + row].SetValue(item.NoiCap);
                        workSheet.Cells["K" + row].SetValue(item.NoiDi);
                        workSheet.Cells["L" + row].SetValue(item.NoiDen);
                        workSheet.Cells["M" + row].SetValue(item.HanhLy);
                        workSheet.Cells["N" + row].SetValue(item.GhiChuTxt);
                        stt++;
                        row++;
                    }
                }

                string fTuNgay = General.ConvertStringToDate(tungay).ToString("dd/MM/yyyy");
                string fDenNgay = General.ConvertStringToDate(denngay).ToString("dd/MM/yyyy");

                string filter = string.Format("(Từ ngày: {0} đến {1}; Số giấy tờ: {2}; Quốc tịch: {3}; Số hiệu: {4})", fTuNgay, fDenNgay, so_giay_to, quoc_tich, so_hieu);
                workSheet.Cells["A2"].SetValue(filter);

                var range = workSheet.Cells.GetSubrange("A3", "N" + (row - 1));
                range.Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);
                // xuất tài liệu thành tệp tin
                string handle = Guid.NewGuid().ToString();
                var stream = new MemoryStream();
                workbook.Save(stream, SaveOptions.XlsxDefault);
                stream.Position = 0;
                System.Web.HttpContext.Current.Cache.Insert(handle, stream.ToArray());
                byte[] data = stream.ToArray() as byte[];
                //return File(data, "application/octet-stream", "documentDemo.docx");
                return Json(new { status = 200, fileguid = handle, filename = "DS_HanhKhach_NoiChuyen_" + tungay.Replace("-", "") + "_" + denngay.Replace("-", "") + (string.IsNullOrEmpty(so_hieu) ? "" : "_" + so_hieu) + ".xlsx" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { status = 500, description = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }
    }
}