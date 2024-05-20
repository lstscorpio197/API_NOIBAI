using ExportDocApi.Common;
using ExportDocApi.Models;
using GemBox.Document;
using MySql.Data.MySqlClient;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Mvc;

namespace ExportDocApi.Controllers
{
    public class WarningController : BaseController
    {
        // GET: Warning
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult ExportDocument(string tungay, string denngay, string ho_va_ten, string so_giay_to, string so_hieu)
        {
            try
            {
                if (string.IsNullOrEmpty(tungay))
                {
                    return Json(new { status = 400, description = "Vui lòng chọn khoảng ngày" }, JsonRequestBehavior.AllowGet);
                }

                ComponentInfo.SetLicense("DOVJ-6B74-HTFU-1VVQ");
                var document = DocumentModel.Load(Server.MapPath("/FileTemp/Mau02.docx"));
                //Xử lý dữ liệu
                MySqlConnection conn = new MySqlConnection();
                conn.ConnectionString = ConfigKey.ConnectionString;
                conn.Open();
                string query = "select * from warning_history where true";
                if (!string.IsNullOrEmpty(tungay))
                {
                    query += " and ngay_canh_bao >= '" + tungay + "'";
                }
                if (!string.IsNullOrEmpty(denngay))
                {
                    query += " and ngay_canh_bao <= '" + denngay + "'";
                }
                if (!string.IsNullOrEmpty(ho_va_ten))
                {
                    query += " and ho_va_ten like '%" + ho_va_ten + "%'";
                }
                if (!string.IsNullOrEmpty(so_giay_to))
                {
                    query += " and so_giay_to like '%" + so_giay_to + "%'";
                }
                if (!string.IsNullOrEmpty(so_hieu))
                {
                    query += " and so_hieu = '" + so_hieu + "'";
                }
                query += " order by ngay_canh_bao desc";
                var cmd = new MySqlCommand(query, conn);
                var dr = cmd.ExecuteReader();
                var lstCB = new List<Warning_History>();
                if (dr.HasRows)
                {
                    DataTable dt = new DataTable();
                    dt.Load(dr);
                    lstCB = (from rw in dt.AsEnumerable()
                             select new Warning_History()
                             {
                                 id = Convert.ToInt32(rw["id"]),
                                 ho = Convert.ToString(rw["ho"]),
                                 ten_dem = Convert.ToString(rw["ten_dem"]),
                                 ten = Convert.ToString(rw["ten"]),
                                 ho_va_ten = Convert.ToString(rw["ho_va_ten"]),
                                 quoc_tich = Convert.ToString(rw["quoc_tich"]),
                                 so_giay_to = Convert.ToString(rw["so_giay_to"]),
                                 loai_giay_to = Convert.ToString(rw["loai_giay_to"]),
                                 ngay_canh_bao = Convert.ToString(rw["ngay_canh_bao"]),
                                 so_hieu = Convert.ToString(rw["so_hieu"]),
                                 ngay_bay = Convert.ToString(rw["ngay_bay"]),
                                 hanh_ly = Convert.ToString(rw["hanh_ly"]),
                                 ngay_sinh = Convert.ToString(rw["ngay_sinh"]),
                                 gioi_tinh = Convert.ToString(rw["gioi_tinh"]),
                                 ghi_chu = Convert.ToString(rw["ghi_chu"]),
                             }).ToList();


                }
                if (!lstCB.Any())
                {
                    lstCB.Add(new Warning_History());
                }

                tungay = (DateTime.ParseExact(tungay, "yyyy-MM-dd", CultureInfo.InvariantCulture)).ToString("dd-MM-yyyy");
                denngay = (DateTime.ParseExact(denngay, "yyyy-MM-dd", CultureInfo.InvariantCulture)).ToString("dd-MM-yyyy");
                string ngay = tungay == denngay ? tungay : tungay + " đến " + denngay;

                var merge = new
                {
                    ngay = ngay,
                    so_hieu_head = so_hieu
                };
                document.MailMerge.Execute(merge);


                var tblCB = new DataTable() { TableName = "LstCB" };
                tblCB.Columns.Add("STT", typeof(int));
                tblCB.Columns.Add("ho_va_ten", typeof(string));
                tblCB.Columns.Add("so_giay_to", typeof(string));
                tblCB.Columns.Add("gioi_tinh", typeof(string));
                tblCB.Columns.Add("ngay_sinh", typeof(string));
                tblCB.Columns.Add("quoc_tich", typeof(string));
                tblCB.Columns.Add("dau_hieu", typeof(string));
                tblCB.Columns.Add("chuyen_bay", typeof(string));
                tblCB.Columns.Add("ket_qua", typeof(string));


                for (int i = 0; i < lstCB.Count(); i++)
                {
                    var cb = lstCB[i];
                    string chuyenBay = cb.so_hieu != null ? cb.so_hieu.Replace(" |---| ", "\n") : "";
                    string ngaysinh = string.IsNullOrEmpty(cb.ngay_sinh) ? "" : General.ConvertDate(cb.ngay_sinh);
                    tblCB.Rows.Add(i + 1, cb.ho_va_ten ?? "", cb.so_giay_to ?? "", cb.gioi_tinh ?? "", ngaysinh, cb.quoc_tich ?? "", cb.ghi_chu, chuyenBay, "");
                }
                document.MailMerge.Execute(tblCB);
                // xuất tài liệu thành tệp tin
                string handle = Guid.NewGuid().ToString();
                var stream = new MemoryStream();
                document.Save(stream, SaveOptions.DocxDefault);
                stream.Position = 0;
                System.Web.HttpContext.Current.Cache.Insert(handle, stream.ToArray());
                byte[] data = stream.ToArray() as byte[];
                //return File(data, "application/octet-stream", "documentDemo.docx");
                return Json(new { status = 200, fileguid = handle, filename = "Kết quả xác định trọng điểm " + ngay + ".docx" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { status = 500, description = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult ExportMau01(string id)
        {
            try
            {
                if (string.IsNullOrEmpty(id))
                {
                    return Json(new { status = 400, description = "Vui lòng chọn đối tượng" }, JsonRequestBehavior.AllowGet);
                }

                ComponentInfo.SetLicense("DOVJ-6B74-HTFU-1VVQ");
                var document = DocumentModel.Load(Server.MapPath("/FileTemp/Mau01.docx"));
                //Xử lý dữ liệu
                MySqlConnection conn = new MySqlConnection();
                conn.ConnectionString = ConfigKey.ConnectionString;
                conn.Open();

                var lstId = JsonConvert.DeserializeObject<List<int>>(id);
                string guids = "'" + string.Join("','", lstId) + "'";
                string query = string.Format("select * from canhbao_hanhkhach where id in ({0})", guids);
                var cmd = new MySqlCommand(query, conn);
                var dr = cmd.ExecuteReader();
                var lstCB = new List<Warning_Object>();
                if (dr.HasRows)
                {
                    DataTable dt = new DataTable();
                    dt.Load(dr);
                    lstCB = (from rw in dt.AsEnumerable()
                             select new Warning_Object()
                             {
                                 id = Convert.ToInt32(rw["id"]),
                                 ho = Convert.ToString(rw["ho"]),
                                 ten_dem = Convert.ToString(rw["ten_dem"]),
                                 ten = Convert.ToString(rw["ten"]),
                                 quoc_tich = Convert.ToString(rw["quoc_tich"]),
                                 so_giay_to = Convert.ToString(rw["so_giay_to"]),
                                 loai_giay_to = Convert.ToString(rw["loai_giay_to"]),
                                 ngay_sinh = Convert.ToString(rw["ngay_sinh"]),
                                 gioi_tinh = Convert.ToString(rw["gioi_tinh"]),
                                 ghi_chu = Convert.ToString(rw["ghi_chu"]),
                                 yc_nghiep_vu = Convert.ToString(rw["yc_nghiep_vu"]),
                             }).ToList();


                }
                if (!lstCB.Any())
                {
                    lstCB.Add(new Warning_Object());
                }


                string ngay = DateTime.Today.ToString("dd/MM/yyyy");

                var merge = new
                {
                    ngay = ngay,
                };
                document.MailMerge.Execute(merge);


                var tblCB = new DataTable() { TableName = "LstDs" };
                tblCB.Columns.Add("stt", typeof(int));
                tblCB.Columns.Add("ho_va_ten", typeof(string));
                tblCB.Columns.Add("so_giay_to", typeof(string));
                tblCB.Columns.Add("gioi_tinh", typeof(string));
                tblCB.Columns.Add("ngay_sinh", typeof(string));
                tblCB.Columns.Add("quoc_tich", typeof(string));
                tblCB.Columns.Add("dau_hieu", typeof(string));
                tblCB.Columns.Add("muc_dich", typeof(string));
                tblCB.Columns.Add("yeu_cau", typeof(string));
                tblCB.Columns.Add("khac", typeof(string));


                for (int i = 0; i < lstCB.Count(); i++)
                {
                    var cb = lstCB[i];
                    string ngaysinh = string.IsNullOrEmpty(cb.ngay_sinh) ? "" : General.ConvertDate(cb.ngay_sinh);
                    tblCB.Rows.Add(i + 1, cb.ho_va_ten ?? "", cb.so_giay_to ?? "", cb.gioi_tinh ?? "", ngaysinh, cb.quoc_tich ?? "", cb.ghi_chu, "", cb.yc_nghiep_vu, "");
                }
                document.MailMerge.Execute(tblCB);
                // xuất tài liệu thành tệp tin
                string handle = Guid.NewGuid().ToString();
                var stream = new MemoryStream();
                document.Save(stream, SaveOptions.DocxDefault);
                stream.Position = 0;
                System.Web.HttpContext.Current.Cache.Insert(handle, stream.ToArray());
                byte[] data = stream.ToArray() as byte[];
                //return File(data, "application/octet-stream", "documentDemo.docx");
                return Json(new { status = 200, fileguid = handle, filename = "Phiếu xây dựng thông tin trọng điểm.docx" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { status = 500, description = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        
    }
}