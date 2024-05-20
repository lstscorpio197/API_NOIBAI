using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportDocApi.Models
{
    public class Warning_History
    {
        public int id { get; set; }
        public string so_giay_to { get; set; }
        public string loai_giay_to { get; set; }
        public string ho { get; set; }
        public string ten_dem { get; set; }
        public string ten { get; set; }
        public string quoc_tich { get; set; }
        public string ngay_canh_bao { get; set; }
        public string so_hieu { get; set; }
        public string ngay_bay { get; set; }
        public string hanh_ly { get; set; }
        public string ho_va_ten { get; set; }
        public string ngay_sinh { get; set; }
        public string gioi_tinh { get; set; }
        public string ghi_chu { get; set; }
        public string mail_status { get; set; }

    }
}