using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Threading.Tasks;
using System.Windows.Forms;
using Npgsql;
using NpgsqlTypes;
using System.Data.Odbc;

namespace MigradorXls
{
    public partial class MigracionBD : Form
    {
        string sql;
        string file;
        string connectionString = @"Host=" + Globals.Host + ";port=" + Globals.port + ";Database=" + Globals.DB + ";User ID=" + Globals.usuario + ";Password=" + Globals.pass + ";";
        int codInterno = 0;
        int count = 0;
        DataTable TAB = new DataTable();
        public MigracionBD()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(comboBox2.Text == "Zonas")
            {
                MigrarZonaBD();
            }
        }

        private void MigrarZonaBD()
        {
            string txtConStr = "DSN=conDBisam";
            //path : The complete path to the folder, where the DBISAM Tables 
            //(i.e. *.dat files) are present.
            OdbcConnection objODBCCon = new OdbcConnection(txtConStr);
            objODBCCon.Open();        
            NpgsqlConnection conn = new NpgsqlConnection(connectionString);
            int i = 0;
            int lenght = TAB.Rows.Count;
            TAB.Clear();
            switch (comboBox2.Text)
            {
                case "Zonas":
                    {
                        A2Zonas(objODBCCon);
                        while (i < lenght)
                        {
                            callbackInsertZona(conn, TAB, i);
                            i++;
                        }
                        break;
                    }
                case "Moneda":
                    {

                        break;
                    }
            }
            
            objODBCCon.Close();
        }

        private void A2Zonas(OdbcConnection objODBCCon)
        {
            string oString = "Select FZ_CODIGO, FZ_DESCRIPCION, FZ_STATUS from Szonas";
            OdbcDataAdapter comm = new OdbcDataAdapter(oString, objODBCCon);
            comm.Fill(TAB);
            
        }

        private void callbackInsertZona(NpgsqlConnection conn, DataTable ROW, int i) 
        {
             codInterno += 1;
             conn.Open();
                sql = @"INSERT INTO admin.gen_zona(org_hijo,cod_interno,codigo,descri,descorta,
                    latitud,longitud,altitud, reg_usu_cc, reg_usu_cu,reg_estatus, 
                    disponible,migrado) VALUES(@orgHijo , @codInterno,
                    @codigo, @descri, @descorta , @logitud, @latitud, 
                    @altitud, @reg_usu_cc , @reg_usu_cu, @regEstatus, @disponible, @migrado)";
                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn);
                dbcmd.Parameters.Add(new NpgsqlParameter("@orghijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codInterno", NpgsqlDbType.Bigint));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codigo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descorta", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@logitud", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@latitud", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@altitud", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cu", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
                dbcmd.Parameters.Add(new NpgsqlParameter("@disponible", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));

                dbcmd.Prepare();

                dbcmd.Parameters[0].Value = Globals.org;
                dbcmd.Parameters[1].Value = codInterno;
                dbcmd.Parameters[2].Value = ROW.Rows[i][0];
                dbcmd.Parameters[3].Value = ROW.Rows[i][1];
                dbcmd.Parameters[4].Value = ROW.Rows[i][1];
                dbcmd.Parameters[5].Value = 0;
                dbcmd.Parameters[6].Value = 0;
                dbcmd.Parameters[7].Value = 0;
                dbcmd.Parameters[8].Value = "INNOVA";
                dbcmd.Parameters[9].Value = "INNOVA";
                dbcmd.Parameters[10].Value = 1;
                dbcmd.Parameters[11].Value = ROW.Rows[i][2];
                dbcmd.Parameters[12].Value = true;

                count +=dbcmd.ExecuteNonQuery();
            conn.Close();
        }

        private void A2Moneda(OdbcConnection objODBCCon)
        {
            string oString = "Select FM_CODE, FM_DESCRIPCION, FM_STATUS, FM_SIMBOLO, FM_FACTOR from Smoneda";
            OdbcDataAdapter comm = new OdbcDataAdapter(oString, objODBCCon);
            comm.Fill(TAB);

        }

        private void callbackInsertMoneda(NpgsqlConnection conn, DataTable ROW, int i)
        {
            codInterno += 1;
            conn.Open();
            sql = @"INSERT INTO admin.gen_moneda(org_hijo,cod_interno,codigo,descri,descorta,simbolo,
                    factor,ant_factor, reg_usu_cc, reg_usu_cu,reg_estatus, 
                    disponible,migrado) VALUES(@orgHijo , @codInterno,
                    @codigo, @descri, @descorta ,@simbolo, @factor, @antFactor, 
                    @reg_usu_cc , @reg_usu_cu, @regEstatus, @disponible, @migrado)";
                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn);
                dbcmd.Parameters.Add(new NpgsqlParameter("@orghijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codInterno", NpgsqlDbType.Bigint));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codigo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descorta", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@simbolo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@factor", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@antFactor", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cu", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
                dbcmd.Parameters.Add(new NpgsqlParameter("@disponible", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));

                dbcmd.Prepare();

                dbcmd.Parameters[0].Value = Globals.org;
                dbcmd.Parameters[1].Value = codInterno;
                dbcmd.Parameters[2].Value = ROW.Rows[i][0];
                dbcmd.Parameters[3].Value = ROW.Rows[i][1];
                dbcmd.Parameters[4].Value = ROW.Rows[i][1];
                dbcmd.Parameters[5].Value = ROW.Rows[i][3];
                dbcmd.Parameters[6].Value = ROW.Rows[i][4];
                dbcmd.Parameters[7].Value = 1;
                dbcmd.Parameters[8].Value = "INNOVA";
                dbcmd.Parameters[9].Value = "INNOVA";
                dbcmd.Parameters[10].Value = 1;
                dbcmd.Parameters[11].Value = ROW.Rows[i][2];
                dbcmd.Parameters[12].Value = true;

                count += dbcmd.ExecuteNonQuery();
            conn.Close();

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            button1.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            DialogResult dr = folderBrowserDialog1.ShowDialog();
            if (dr == DialogResult.OK)
            {
                file = folderBrowserDialog1.SelectedPath;
                button2.Enabled = true;
            }
        }
    }
}
