using ExportDocApi.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportDocApi.Models
{
    public class ChuyenBay_ToBay_ExportDto
    {
        public string IdToBay { get; set; }
        public string IdChuyenBay { get; set; }
        public string Ho { get; set; }
        public string TenDem { get; set; }
        public string Ten { get; set; }
        public string HoTen => Ho + " " + TenDem + " " + Ten;
        public string SoGiayTo { get; set; }
        public string LoaiGiayTo { get; set; }
        public string NgaySinh { get; set; }
        public string GioiTinh { get; set; }
        public string QuocTich { get; set; }
        public string SoHieu { get; set; }
        public string NgayBay { get; set; }

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

        public string NgayBayTxt
        {
            get
            {
                if (NgayBay.StartsWith("00"))
                    return string.Empty;
                return General.ConvertDate(NgayBay).Replace("-", "/");
            }
        }
    }
}