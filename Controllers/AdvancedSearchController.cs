using ExportDocApi.Common;
using ExportDocApi.Models;
using GemBox.Spreadsheet;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace ExportDocApi.Controllers
{
    public class AdvancedSearchController : BaseController
    {
        // GET: AdvancedSearch
        public ActionResult Index()
        {
            return View();
        }

        public async Task<JsonResult> GetList(string tungay, string denngay, string xnc, string noi_di, string noi_den, int solan, string so_hieu)
        {
            try
            {
                if (string.IsNullOrEmpty(tungay) || string.IsNullOrEmpty(denngay))
                {
                    return Json(new { status = 400, description = "Vui lòng chọn khoảng ngày" }, JsonRequestBehavior.AllowGet);
                }

                var lstHK = await GetListHanhKhach(tungay, denngay, xnc, noi_di, noi_den, solan, so_hieu);

                var json = Json(new { status = 200, data = lstHK }, JsonRequestBehavior.AllowGet);
                json.MaxJsonLength = int.MaxValue;
                return json;
            }
            catch (Exception ex)
            {
                return Json(new { status = 500, description = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public async Task<JsonResult> ExportExcel(string tungay, string denngay, string xnc, string noi_di, string noi_den, int solan, string so_hieu)
        {
            try
            {
                if (string.IsNullOrEmpty(tungay) || string.IsNullOrEmpty(denngay))
                {
                    return Json(new { status = 400, description = "Vui lòng chọn khoảng ngày" }, JsonRequestBehavior.AllowGet);
                }

                SpreadsheetInfo.SetLicense(ConfigKey.KeyGemBoxSpreadsheet);
                var workbook = ExcelFile.Load(Server.MapPath("/FileTemp/ExportAdvancedSearch.xlsx"));
                var workSheet = workbook.Worksheets[0];

                var lstHK = await GetListHanhKhach(tungay, denngay, xnc, noi_di, noi_den, solan, so_hieu);

                int row = 4;
                int stt = 1;
                if (lstHK.Any())
                {
                    foreach (var item in lstHK)
                    {
                        workSheet.Cells["A" + row].SetValue(stt);
                        workSheet.Cells["B" + row].SetValue(item.HoTen);
                        workSheet.Cells["C" + row].SetValue(item.GioiTinh);
                        workSheet.Cells["D" + row].SetValue(item.QuocTich);
                        workSheet.Cells["E" + row].SetValue(item.SoGiayTo);
                        workSheet.Cells["F" + row].SetValue(item.LoaiGiayTo);
                        workSheet.Cells["G" + row].SetValue(item.NgaySinhTxt);
                        workSheet.Cells["H" + row].SetValue(item.SoLan);
                        workSheet.Cells["I" + row].SetValue(item.CB);
                        workSheet.Cells["J" + row].SetValue(item.HanhLy);
                        stt++;
                        row++;
                    }
                }

                string fTuNgay = tungay.Replace("-", "/");
                string fDenNgay = denngay.Replace("-", "/");

                string filter = string.Format("(Từ ngày: {0} đến {1}; Số hiệu: {2}; Nơi đi:{3}; Nơi đến:{4})", fTuNgay, fDenNgay, so_hieu, noi_di, noi_den);
                workSheet.Cells["A2"].SetValue(filter);

                var range = workSheet.Cells.GetSubrange("A3", "J" + (row - 1));
                range.Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);
                string handle = Guid.NewGuid().ToString();
                var stream = new MemoryStream();
                workbook.Save(stream, SaveOptions.XlsxDefault);
                stream.Position = 0;
                System.Web.HttpContext.Current.Cache.Insert(handle, stream.ToArray());
                byte[] data = stream.ToArray() as byte[];
                return Json(new { status = 200, fileguid = handle, filename = "DS_HanhKhach_" + tungay.Replace("-", "") + "_" + denngay.Replace("-", "") + (string.IsNullOrEmpty(so_hieu) ? "" : "_" + so_hieu) + ".xlsx" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { status = 500, description = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }


        private async Task<List<AdvancedSearchResultDto>> GetListHanhKhach(string tungay, string denngay, string xnc, string noi_di, string noi_den, int solan, string so_hieu)
        {
            MySqlConnection conn = new MySqlConnection();
            conn.ConnectionString = ConfigKey.ConnectionString;
            conn.Open();
            string tuNgay = General.ConvertStringToDate(tungay, "dd-MM-yyyy").ToString("yyyy-MM-dd");
            string denNgay = General.ConvertStringToDate(denngay, "dd-MM-yyyy").ToString("yyyy-MM-dd");
            if (xnc == "xc")
            {
                noi_di = "HAN";
            }
            if (xnc == "nc")
            {
                noi_den = "HAN";
            }
            string query = string.Format("call SP_GetListHanhKhach('{0}','{1}','{2}','{3}','{4}')", tuNgay, denNgay, noi_di, noi_den, so_hieu);

            var cmd = new MySqlCommand(query, conn)
            {
                CommandTimeout = int.MaxValue
            };
            var dr = cmd.ExecuteReader();

            var lstHK = new List<AdvancedSearchResultDto>();
            if (dr.HasRows)
            {
                while (await dr.ReadAsync())
                {
                    AdvancedSearchResultDto obj = new AdvancedSearchResultDto
                    {
                        IdHanhKhach = Convert.ToString(dr["IDHANHKHACH"]),
                        Ho = Convert.ToString(dr["HO"]),
                        TenDem = Convert.ToString(dr["TENDEM"]),
                        Ten = Convert.ToString(dr["TEN"]),
                        SoGiayTo = Convert.ToString(dr["SOGIAYTO"]),
                        LoaiGiayTo = Convert.ToString(dr["LOAIGIAYTO"]),
                        NgaySinh = Convert.ToString(dr["NGAYSINH"]),
                        GioiTinh = Convert.ToString(dr["GIOITINH"]),
                        QuocTich = Convert.ToString(dr["QUOCTICH"]),
                        SoHieu = Convert.ToString(dr["SOHIEU"]),
                        NgayBay = Convert.ToString(dr["FLIGHTDATE"]),
                        HanhLy = Convert.ToString(dr["HANHLY"])
                    };

                    lstHK.Add(obj);
                }
            }
            solan--;
            var lstHK_Grouped = lstHK.GroupBy(x => x.SoGiayTo).Where(x => x.Count() > solan);
            var lstResult = lstHK_Grouped.Select(x => x.OrderByDescending(y => y.FlightDate).ThenByDescending(y => y.IdHanhKhach)).Select(x => new AdvancedSearchResultCount
            {
                Data = x.FirstOrDefault(),
                Count = x.Count()
            }).ToList();
            lstResult.ForEach(x => x.Data.SoLan = x.Count);
            
            return lstResult.Select(x=>x.Data).OrderByDescending(x=>x.SoLan).ToList();
        }

        private List<string> GetListCB(string strCB)
        {
            var array = strCB.Split(';');
            return array.Select(x => GetCB(x)).ToList();
        }
        private string GetCB(string cb)
        {
            try
            {
                int index = cb.IndexOf('/');
                string sh = cb.Substring(0, index);
                string ngay = cb.Substring(index + 1);
                return string.Format("{0} ngày {1}", sh, General.ConvertDate(ngay));
            }
            catch (Exception ex)
            {
                return string.Empty;
            }

        }
    }
}