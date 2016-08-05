using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MigradorXls
{
    public class DataConvert
    {
        public  DateTime? ExtractDate(string myDate)
        {
            if (!string.IsNullOrEmpty(myDate) && !string.IsNullOrWhiteSpace(myDate))
            {
                DateTime dt;
                var formatStrings = new string[] { "dd/MM/yyyy h:mm:ss","dd/MM/yyyy", "d/M/yyyy",
                    "dd.MM.yyyy h:mm:ss","dd.MM.yyyy", "d.M.yyyy",
                    "dd-MM-yyyy h:mm:ss", "dd-MM-yyyy", "d-M-yyyy" };
                dt = DateTime.ParseExact(myDate.Replace("\t", ""), formatStrings, new CultureInfo("en-US"), DateTimeStyles.None);
                return dt;
            }
            return null;
        }

        public bool convertBoolean(object obj)
        {
            string text = Convert.ToString(obj);
            if (text.Equals("t", StringComparison.OrdinalIgnoreCase) || text.Equals("true", StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }

        public  T GetValue<T>(object value)
        {
            if (value == null || value == DBNull.Value)
                return default(T);
            else
                return (T)value;
        }


    }

    public static class Extension
    {
        public static string ToStringOrEmpty(this Object value)
        {
            return value == null ? "" : value.ToString();
        }
    }
    
}
