using ExportDocApi.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportDocApi.Models
{
    public class ChuyenBay_HanhKhach_ExportDto
    {
        public string IDChuyenBay { get; set; }
        public string SoHieu { get; set; }
        public string SoGhe { get; set; }
        public string Ho { get; set; }
        public string TenDem { get; set; }
        public string Ten { get; set; }
        public string HoTen => Ho + " " + TenDem + " " + Ten;
        public string GioiTinh { get; set; }
        public string QuocTich { get; set; }
        public string NgaySinh { get; set; }
        public string SoGiayTo { get; set; }
        public string LoaiGiayTo { get; set; }
        public string NoiCap { get; set; }
        public string NoiDi { get; set; }
        public string NoiDen { get; set; }
        public string HanhLy { get; set; }
        public string NgayBay { get; set; }
        public string IDHanhKhach { get; set; }
        public string MaNoiDi { get; set; }
        public string MaNoiDen { get; set; }
        public string SYS_DATE { get; set; }

        public bool IsTheoDoi { get; set; }
        public string GioiTinhTxt { 
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
                return General.ConvertDate(NgaySinh).Replace("-","/");
            }
        }

        public string NgayBayTxt
        {
            get
            {
                return General.ConvertDate(NgayBay).Replace("-", "/");
            }
        }
        public DateTime? FlightDate => General.ConvertStringToDateTime(NgayBay);
        public DateTime? CreatedTime => General.ConvertStringToDateTime(SYS_DATE);
    }

    
}