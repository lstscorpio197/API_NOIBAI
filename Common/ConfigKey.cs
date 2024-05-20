using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Configuration;

namespace ExportDocApi.Common
{
    public class ConfigKey
    {
        public static string ConnectionString = WebConfigurationManager.AppSettings["ConnectionString"].ToString();

        public static string KeyGemBoxDocument = "DOVJ-6B74-HTFU-1VVQ";
        public static string KeyGemBoxSpreadsheet = "ETZX-IT28-33Q6-1HA2";
    }
}