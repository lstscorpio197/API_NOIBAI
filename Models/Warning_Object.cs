using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportDocApi.Models
{
    public class Warning_Object
    {
        public int id { get; set; }
        public string so_giay_to { get; set; }
        public string loai_giay_to { get; set; }
        public string ho { get; set; }
        public string ten_dem { get; set; }
        public string ten { get; set; }
        public string quoc_tich { get; set; }
        public string ngay_sinh { get; set; }
        public string gioi_tinh { get; set; }
        public string ho_va_ten => ho + " " + ten_dem + " " + ten;
        public string ghi_chu { get; set; }
        public string yc_nghiep_vu { get; set; }
    }
}