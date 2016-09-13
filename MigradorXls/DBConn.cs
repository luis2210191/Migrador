using LiteDB;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MigradorXls
{
    public class DBConn : IDisposable
    {
        private static DBConn _instance;
        private LiteDatabase _conn;
        private static string dbName = "Colleccion.db";

        /// <summary>
        /// Constructor
        /// </summary>
        private DBConn()
        {

        }

        public static DBConn Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new DBConn();
                    _instance.Init();
                }
                return _instance;
            }
        }

        /// <summary>
        /// Metodo que inicializa todos los valores de la clase
        /// </summary>
        private void Init()
        {
            if (_conn == null)
            {
                _conn = new LiteDatabase(dbName);
            }
        }

        public LiteDatabase Connection
        {
            get
            {
                if (_conn == null)
                {
                    _conn = new LiteDatabase(dbName);
                }
                return _conn;
            }
        }

        /// <summary>
        /// Metodo que asigna la coleccion a consultar en objeto colleccion.db
        /// </summary>
        public LiteCollection<T> Collection<T>() where T : new()
        {
            var name = typeof(T).Name; // Nombre de la clase
                                       //name = name.Substring(0, 1).ToLower() + name; // colocando la primera letra mayuscula. Ejemplo admin -> Admin
            return _conn.GetCollection<T>(name);
        }


        public void Dispose()
        {
            _conn.Dispose();
        }
    }
}
