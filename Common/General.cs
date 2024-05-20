using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;

namespace ExportDocApi.Common
{
    public class General
    {
        public static string ConvertDate(string input)
        {
            try
            {
                if (string.IsNullOrEmpty(input))
                {
                    return string.Empty;
                }
                DateTime dateResult;
                if (DateTime.TryParse(input, out dateResult))
                {
                    return dateResult.ToString("dd-MM-yyyy");
                }
                // Chuyển đổi chuỗi string thành đối tượng DateTime
                //var dateResult = DateTime.ParseExact(input, "yyyy-MM-dd HH:mm:ss tt", CultureInfo.InvariantCulture);
                if (DateTime.TryParseExact(input, "yyyy-MM-dd HH:mm:ss tt", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateResult))
                {
                    return dateResult.ToString("dd-MM-yyyy");
                }
                else
                {
                    if (DateTime.TryParseExact(input, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateResult))
                    {
                        return dateResult.ToString("dd-MM-yyyy");
                    }
                    else
                    {
                        return string.Empty;
                    }
                }
            }
            catch (FormatException)
            {
                // Xử lý lỗi và trả về chuỗi trống
                return string.Empty;
            }
        }

        public static DateTime? ConvertStringToDateTime(string input)
        {
            try
            {
                if (string.IsNullOrEmpty(input))
                {
                    return null;
                }
                DateTime dateResult;
                if (DateTime.TryParse(input, out dateResult))
                {
                    return dateResult;
                }
                // Chuyển đổi chuỗi string thành đối tượng DateTime
                //var dateResult = DateTime.ParseExact(input, "yyyy-MM-dd HH:mm:ss tt", CultureInfo.InvariantCulture);
                if (DateTime.TryParseExact(input, "yyyy-MM-dd HH:mm:ss tt", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateResult))
                {
                    return dateResult;
                }
                else
                {
                    if (DateTime.TryParseExact(input, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateResult))
                    {
                        return dateResult;
                    }
                    else
                    {
                        return null;
                    }
                }
            }
            catch (FormatException)
            {
                // Xử lý lỗi và trả về chuỗi trống
                return null;
            }
        }

        public static string ConvertDateToDb(string input)
        {
            try
            {
                if (string.IsNullOrEmpty(input))
                {
                    return "NULL";
                }
                DateTime dateResult;
                if (DateTime.TryParse(input, out dateResult))
                {
                    return dateResult.ToString("yyyy-MM-dd");
                }
                // Chuyển đổi chuỗi string thành đối tượng DateTime
                //var dateResult = DateTime.ParseExact(input, "yyyy-MM-dd HH:mm:ss tt", CultureInfo.InvariantCulture);
                if (DateTime.TryParseExact(input, "yyyy-MM-dd HH:mm:ss tt", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateResult))
                {
                    return dateResult.ToString("yyyy-MM-dd");
                }
                else
                {
                    if (DateTime.TryParseExact(input, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateResult))
                    {
                        return dateResult.ToString("yyyy-MM-dd");
                    }
                    else
                    {
                        return string.Empty;
                    }
                }
            }
            catch (FormatException)
            {
                // Xử lý lỗi và trả về chuỗi trống
                return string.Empty;
            }
        }

        public static DateTime ConvertStringToDate(string date)
        {
            DateTime result = DateTime.ParseExact(date, "yyyy-MM-dd", CultureInfo.InvariantCulture);
            return result;
        }

        public static DateTime ConvertStringToDate(string date, string format = "dd/MM/yyyy")
        {
            DateTime result = DateTime.ParseExact(date, format, CultureInfo.InvariantCulture);
            return result;
        }

        public static string ConvertStringDateToString(string date, string srcFormat = "dd/MM/yyyy", string desFormat = "dd/MM/yyyy")
        {
            try
            {
                DateTime result = DateTime.ParseExact(date, srcFormat, CultureInfo.InvariantCulture);
                return result.ToString(desFormat);
            }
            catch (Exception)
            {
                return string.Empty;
            }
            
        }

        public static int DateDurationCalculator(DateTime date1, DateTime? date2 = null)
        {
            date2 = date2 ?? DateTime.Today;
            TimeSpan ts = date2.Value - date1;
            return ts.Days;
        }
    }
}