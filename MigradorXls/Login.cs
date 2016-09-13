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
using NpgsqlTypes;
using LiteDB;

namespace MigradorXls
{
    public partial class Login : Form
    {
        //string connectionString = @"Host=192.168.1.254"+";port=5432;Database=migrar;User ID=postgres;Password=TACA8tilo";
        
        public string LoginId;
        
        
        public Login()
        {
            InitializeComponent();
            SetStyle(ControlStyles.UserPaint, true);
            SetStyle(ControlStyles.AllPaintingInWmPaint, true);
            SetStyle(ControlStyles.DoubleBuffer, true);
            DialogResult = DialogResult.Cancel;
            label3.Text += Application.ProductVersion;
            
        }


        /// <summary>
        /// Evento click del boton de inicio de sesion
        /// </summary>
        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                try
                {
                    string connectionString = @"Host=" + Globals.Host + ";port=" + Globals.port + ";Database=" + Globals.DB + ";User ID=" + Globals.usuario + ";Password=" + Globals.pass + ";";
                    string sql;
                    string userId = textBox1.Text;
                    string psw = textBox2.Text;
                    object count = null;
                    String txtxor = xorMsg(psw);
                    psw = Base64Encode(txtxor);
                    NpgsqlConnection conn = new NpgsqlConnection(connectionString);

                    conn.Open();

                    sql = @"SELECT COUNT(*) FROM admin.cfg_usu WHERE codigo='" + userId + "' AND pwd ='" + psw + "'";

                    NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn);

                    // Execute the query and obtain the value of the first column of the first row
                    count = dbcmd.ExecuteScalar();

                    if (count.ToString() == "1")
                    {
                        LoginId = userId;
                        Globals.userid = userId;
                        DialogResult = DialogResult.OK;
                    }
                    else
                    {
                        MessageBox.Show("Combinacion de Usuario y contraseña no es la correcta", "Atencion");
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("No hay conexion valida a la base de datos");
                }

            }

        }

        /// <summary>
        /// Metodo de encriptacion privado
        /// </summary>
        private static string xorMsg(string Msg)
        {
            try
            {
                string Key = "Inn0v4RlZ";
                char[] keys = Key.ToCharArray();
                char[] msg = Msg.ToCharArray();

                int ml = msg.Length;
                int kl = keys.Length;

                char[] newmsg = new char[ml];
                for (int i = 0; i < ml; i++)
                {
                    newmsg[i] = (char)(msg[i] ^ keys[i % kl]);
                }//for i
                msg = null; keys = null;
                return new String(newmsg);
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        /// <summary>
        /// Metodo que encripta con base 64 bits
        /// </summary>
        public static string Base64Encode(string plainText)
        {
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
            return System.Convert.ToBase64String(plainTextBytes);
        }
        /// <summary>
        /// Metodo que desencripta con base 64 bits
        /// </summary>
        public static string Base64Decode(string base64EncodedData)
        {
            var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData);
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }

        /// <summary>
        /// Evento click del boton de configuracion de BD
        /// </summary>
        private void button2_Click(object sender, EventArgs e)
        {
            DBConnection conn = new DBConnection();
            conn.ShowDialog();
        }

        /// <summary>
        /// Evento de presion de tecla para la creacion de un shortcut para la configuracion de la BD
        /// </summary>
        private void Login_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.B)
            {
                DBConnection conn = new DBConnection();
                conn.ShowDialog();
            }
        }
    }
}

