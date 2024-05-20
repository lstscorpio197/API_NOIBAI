using ExportDocApi.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportDocApi.Models
{
    public class AdvancedSearchResultDto
    {
        public string IdHanhKhach { get; set; }
        public string Ho { get; set; }
        public string TenDem { get; set; }
        public string Ten { get; set; }
        public string HoTen => string.Format("{0} {1} {2}", Ho, TenDem, Ten);
        public string SoGiayTo { get; set; }
        public string LoaiGiayTo { get; set; } = "P";
        public string GioiTinh { get; set; }
        public string QuocTich { get; set; }
        public string NgaySinh { get; set; }
        public int SoLan { get; set; }
        public string SoHieu { get; set; }
        public string NgayBay { get; set; }
        public string HanhLy { get; set; }
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
                return General.ConvertDate(NgayBay).Replace("-", "/");
            }
        }
        public DateTime? FlightDate => General.ConvertStringToDateTime(NgayBay);

        public string CB => string.Format("{0} ngày {1}", SoHieu, NgayBayTxt);
    }

    public class AdvancedSearchResultCount {
        public AdvancedSearchResultDto Data { get; set; }
        public int Count { get; set; }
    }
}