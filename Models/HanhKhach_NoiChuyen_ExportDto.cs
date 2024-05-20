using ExportDocApi.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportDocApi.Models
{
    public class HanhKhach_NoiChuyen_ExportDto
    {
        public string IDHanhKhach { get; set; }
        public string SoHieu { get; set; }
        public string Ho { get; set; }
        public string TenDem { get; set; }
        public string Ten { get; set; }
        public string MaDatCho { get; set; }
        public string GioiTinh { get; set; }
        public string QuocTich { get; set; }
        public string NgaySinh { get; set; }
        public string LoaiGiayTo { get; set; }
        public string SoGiayTo { get; set; }
        public string NoiCap { get; set; }
        public string NoiDi { get; set; }
        public string NoiDen { get; set; }
        public string HanhLy { get; set; }
        public string GhiChu { get; set; }
        public string HoVaTen => string.Format("{0} {1} {2}", Ho, TenDem, Ten);
        public string GhiChuTxt
        {
            get
            {
                switch (GhiChu)
                {
                    case "0":
                        return "Theo dõi";
                    case "1":
                        return "Trọng điểm";
                    default:
                        return string.Empty;
                }
            }
        }
        public string GioiTinhTxt
        {
            get
            {
                switch (GioiTinh)
                {
                    case "M":
                        return "Nam";
                    case "F":
                        return "Nữ";
                    default:
                        return string.Empty;
                }
            }
        }

        public string NgaySinhTxt
        {
            get
            {
                return General.ConvertDate(NgaySinh).Replace("-", "/");
            }
        }
    }
}