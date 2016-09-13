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
        /// <summary>
        /// Metodo que crea un opbjeto del tipo fecha con formatos preestablecidos de un string
        /// </summary>
        public DateTime? ExtractDate(string myDate)
        {
            if (!string.IsNullOrEmpty(myDate) && !string.IsNullOrWhiteSpace(myDate))
            {
                DateTime dt;
                var formatStrings = new string[] { "dd/MM/yyyy h:mm:ss","dd/MM/yyyy HH:mm:ss","dd/MM/yyyy hh:mm:ss","dd/MM/yyyy", "d/M/yyyy",
                    "dd.MM.yyyy h:mm:ss","dd.MM.yyyy", "d.M.yyyy",
                    "dd-MM-yyyy h:mm:ss", "dd-MM-yyyy", "d-M-yyyy" };
                dt = DateTime.ParseExact(myDate.Replace("\t", ""), formatStrings, new CultureInfo("en-US"), DateTimeStyles.None);
                return dt;
            }
            return null;
        }

        /// <summary>
        /// Metodo que toma un string y segun su valor retorna un booleano
        /// </summary>
        public bool convertBoolean(object obj)
        {
            string text = Convert.ToString(obj);
            if (text.Equals("t", StringComparison.OrdinalIgnoreCase) || text.Equals("true", StringComparison.OrdinalIgnoreCase) || text.Equals("1")) return true;
            else return false;
        }

        /// <summary>
        /// Metodo que obtiene valor por defecto del tipo del parametro si este esta vacio o es nulo
        /// </summary>
        public T GetValue<T>(object value)
        {
            if (value == DBNull.Value || string.IsNullOrWhiteSpace(value.ToString()))
                return default(T);
            else
                return (T)value;
        }
    }

    public static class Extension
    {
        /// <summary>
        /// Metodo que devuelve vacio si el string es nulo
        /// </summary>
        public static string ToStringOrEmpty(this object value)
        {
            return value == null ? "" : value.ToString();
        }

        /// <summary>
        /// Metodo que convierte un string en int
        /// </summary>
        public static int FromStringToInt(this object value)
        {
            if (value == DBNull.Value || string.IsNullOrWhiteSpace(value.ToString()))
                return 0;
            else
                return Convert.ToInt32(value);
        }
    }
    
}
