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

        /// <summary>
        /// Constructor
        /// </summary>
        public DBConnection()
        {

            InitializeComponent();
            textBox1.Text = Globals.Host;
            textBox2.Text = Globals.DB;
        }

        /// <summary>
        /// Evento click del boton para validar conexion a la BD
        /// </summary>
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

