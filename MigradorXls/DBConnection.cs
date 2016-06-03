using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Npgsql;
using LiteDB;
using MigradorXls;
//using SQLite;

namespace MigradorXls
{
    public partial class DBConnection : Form
    {

        public DBConnection()
        {

            InitializeComponent();
            textBox1.Text = Globals.Host;
            textBox2.Text = Globals.DB;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Globals.Host = textBox1.Text;
                Globals.DB = textBox2.Text;
                Globals.usuario = "Iياجا餐ر爪福غOاب戎ム博ر爪格";
                Globals.pass = "manuganu.15";
                try
                {
                    //string de conexion con las credenciales de postgreSQL
                    string connectionString = @"Host=" + Globals.Host + ";port=" + Globals.port + ";Database=" + Globals.DB + ";User ID=" + Globals.usuario + ";Password=" + Globals.pass + ";";

                    NpgsqlConnection conn = new NpgsqlConnection(connectionString);

                    conn.Open();

                    string sql = @"SELECT org_hijo from admin.cfg_org";
                    //Abriendo la coneccion con npgsql
                    connectionString = @"Host=" + Globals.Host + ";port=" + Globals.port + ";Database=" + Globals.DB + ";User ID=" + Globals.usuario + ";Password=" + Globals.pass + ";";
                    conn = new NpgsqlConnection(connectionString);
                    NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn);
                    conn.Open();
                    Globals.org = dbcmd.ExecuteScalar().ToString();
                    conn.Close();
                    conn.Close();

                    MessageBox.Show("Conexion establecida!");

                    this.Close();
                }
                catch (Exception)
                {
                    MessageBox.Show("Se produjo un error al conectarse a la base de datos con esta informacion", "Atencion", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Revise que todos los campos tenga informacion", "Atencion", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

        }

    }
}

public static class Globals
{
    public static String usuario = "";
    public static String pass = "";
    public static int port = 5432;
    public static String Host = "192.168.1.253";
    public static String DB = "CLAVES COPIA";
    public static String org = "";
    public static String userid = "";
    public static String connectionstring = "";
    public static int pref = 0;


}

public class DBConn : IDisposable
{
    private static DBConn _instance;
    private LiteDatabase _conn;
    private static string dbName = "Colleccion.db";

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
