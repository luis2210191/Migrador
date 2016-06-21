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
using LiteDB;

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
            var db = DBConn.Instance;
            var c = db.Collection<admin>();
            comboBox2.DataSource = c.Find(Query.All()).ToList();
            comboBox2.DisplayMember = "desc";
            comboBox2.ValueMember = "id";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MigrarData();            
        }
        private void MigrarData()
        {
            int i = 0;
            int lenght = TAB.Rows.Count;
            TAB.Clear();
            NpgsqlConnection conn = new NpgsqlConnection(connectionString);
            switch (comboBox2.Text)
            {
                case "Zonas":
                    {
                       
                        while (i < lenght)
                        {
                            callbackInsertZona(conn, TAB, i);
                            i++;
                        }
                        break;
                    }
                case "Moneda":
                    {
                     
                        while (i < lenght)
                        {
                            callbackInsertMoneda(conn, TAB, i);
                            i++;
                        }
                        break;
                    }
                case "Talento":
                    {
                      
                        while (i < lenght)
                        {
                            callbackInsertVendedores(conn, TAB, i);
                        }
                        break;
                    }
            }
        }
        private void ExtractA2Table()
        {
            try
            {
                string txtConStr = "DSN=conDBisam";
                OdbcConnection objODBCCon = new OdbcConnection(txtConStr);
                objODBCCon.Open();
                switch (comboBox2.Text)
                {
                    case "Zonas":
                        {
                            A2Zonas(objODBCCon);
                            dataGridView1.DataSource = TAB;
                            break;
                        }
                    case "Moneda":
                        {
                            A2Moneda(objODBCCon);
                            break;
                        }
                    case "Talento":
                        {
                            A2Vendedores(objODBCCon);
                            break;
                        }
                }

                objODBCCon.Close();
                MessageBox.Show("La Migracion de la tabla "+comboBox2.Text+" se ha realizado","Atencion", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch(Exception E)
            {
                MessageBox.Show(E.Message.ToString());
            }
            
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

        private void A2Vendedores(OdbcConnection objODBCCon)
        {
            string oString = "Select FV_CODIGO, FV_DESCRIPCION, FV_DESCRIPCIONDETALLADA, FV_DIRECCION, FM_STATUS, from Svendedores";
            OdbcDataAdapter comm = new OdbcDataAdapter(oString, objODBCCon);
            comm.Fill(TAB);
        }
        private void callbackInsertVendedores(NpgsqlConnection conn, DataTable ROW, int i)
        {
            //sql = @"INSERT INTO admin.org_talento(org_hijo,cod_interno,
            //            codigo,cedula,descri,es_vendedor,es_cobrador,es_servidor,
            //            es_despachador,fecha_nac,reg_usu_cc,reg_usu_cu,
            //            reg_estatus,disponible,tipo_cont,tipo_pers,cod_zona,migrado, 
            //            rif, descorta, sexo, direc1, cod_depar, porc_retencion, fecha_rif, fecha_ing, observacion )
            //            VALUES(@orgHijo , @codInterno, @codigo, @cedula, @descri, 
            //            @esVendedor,@esCobrador, @esServidor, @esDespachador , @fechaNac, 
            //            @reg_usu_cc , @reg_usu_cu, @regEstatus, @disponible, @tipoCont, @tipoPers, 
            //            @codZona , @migrado, @rif, @descorta, @sexo, @direc1,
            //            @cod_depar, @porc_retencion, @fecharif, @fecha_ing, @observacion)";
            //NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn);


            //dbcmd.Parameters.Add(new NpgsqlParameter("@orghijo", NpgsqlDbType.Varchar));
            //dbcmd.Parameters.Add(new NpgsqlParameter("@codInterno", NpgsqlDbType.Bigint));
            //dbcmd.Parameters.Add(new NpgsqlParameter("@codigo", NpgsqlDbType.Varchar));
            //dbcmd.Parameters.Add(new NpgsqlParameter("@cedula", NpgsqlDbType.Varchar));
            //dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
            //dbcmd.Parameters.Add(new NpgsqlParameter("@esVendedor", NpgsqlDbType.Boolean));
            //dbcmd.Parameters.Add(new NpgsqlParameter("@esCobrador", NpgsqlDbType.Boolean));
            //dbcmd.Parameters.Add(new NpgsqlParameter("@esServidor", NpgsqlDbType.Boolean));
            //dbcmd.Parameters.Add(new NpgsqlParameter("@esDespachador", NpgsqlDbType.Boolean));
            //dbcmd.Parameters.Add(new NpgsqlParameter("@fechaNac", NpgsqlDbType.Date));
            //dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
            //dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cu", NpgsqlDbType.Varchar));
            //dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
            //dbcmd.Parameters.Add(new NpgsqlParameter("@disponible", NpgsqlDbType.Boolean));
            //dbcmd.Parameters.Add(new NpgsqlParameter("@tipoCont", NpgsqlDbType.Varchar));
            //dbcmd.Parameters.Add(new NpgsqlParameter("@tipoPers", NpgsqlDbType.Varchar));
            //dbcmd.Parameters.Add(new NpgsqlParameter("@codZona", NpgsqlDbType.Varchar));
            //dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));
            //dbcmd.Parameters.Add(new NpgsqlParameter("@rif", NpgsqlDbType.Varchar));
            //dbcmd.Parameters.Add(new NpgsqlParameter("@descorta", NpgsqlDbType.Varchar));
            //dbcmd.Parameters.Add(new NpgsqlParameter("@sexo", NpgsqlDbType.Varchar));
            //dbcmd.Parameters.Add(new NpgsqlParameter("@direc1", NpgsqlDbType.Varchar));
            //dbcmd.Parameters.Add(new NpgsqlParameter("@porc_retencion", NpgsqlDbType.Double));
            //dbcmd.Parameters.Add(new NpgsqlParameter("@cod_depar", NpgsqlDbType.Varchar));
            //dbcmd.Parameters.Add(new NpgsqlParameter("@fecharif", NpgsqlDbType.Date));
            //dbcmd.Parameters.Add(new NpgsqlParameter("@fecha_ing", NpgsqlDbType.Date));
            //dbcmd.Parameters.Add(new NpgsqlParameter("@observacion", NpgsqlDbType.Varchar));


            //dbcmd.Prepare();


            //dbcmd.Parameters[0].Value = Globals.org;
            //dbcmd.Parameters[1].Value = codInteno;
            //dbcmd.Parameters[2].Value = ROW.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);
            //dbcmd.Parameters[3].Value = ROW.Cells["cedula"].Value;
            //dbcmd.Parameters[4].Value = ROW.Cells["nombres"].Value;
            //dbcmd.Parameters[5].Value = convertBoolean(ROW.Cells["vendedor"].Value);
            //dbcmd.Parameters[6].Value = convertBoolean(ROW.Cells["cobrador"].Value);
            //dbcmd.Parameters[7].Value = convertBoolean(ROW.Cells["servidor"].Value);
            //dbcmd.Parameters[8].Value = convertBoolean(ROW.Cells["despachador"].Value);
            //dbcmd.Parameters[9].Value = ExtractDate(ROW.Cells["fecha nac"].Value.ToString());
            //dbcmd.Parameters[10].Value = "INNOVA";
            //dbcmd.Parameters[11].Value = "INNOVA";
            //dbcmd.Parameters[12].Value = 1;
            //dbcmd.Parameters[13].Value = convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);

            //var db = DBConn.Instance;
            //var col = db.Collection<Tipos>();
            //dbcmd.Parameters[14].Value = col.Find(x => x.tipo == ROW.Cells["tipo contribuyente"].Value.ToString()).FirstOrDefault().codigo;
            //dbcmd.Parameters[15].Value = col.Find(x => x.tipo == ROW.Cells["tipo persona"].Value.ToString()).FirstOrDefault().codigo;
            //dbcmd.Parameters[16].Value = ROW.Cells["codigo de zona"].Value;
            //dbcmd.Parameters[17].Value = true;
            //dbcmd.Parameters[18].Value = ROW.Cells["rif"].Value.ToString().Replace(" ", string.Empty);
            //dbcmd.Parameters[19].Value = ROW.Cells["apellidos"].Value;
            //dbcmd.Parameters[20].Value = col.Find(x => x.tipo == ROW.Cells["Sexo"].Value.ToString()).FirstOrDefault().codigo;
            //dbcmd.Parameters[21].Value = ROW.Cells["direccion"].Value;
            //dbcmd.Parameters[22].Value = ROW.Cells["porc ret iva"].Value;
            //dbcmd.Parameters[23].Value = ROW.Cells["departamento"].Value;
            //dbcmd.Parameters[24].Value = ExtractDate(ROW.Cells["fecha vcto rif"].Value.ToString());
            //dbcmd.Parameters[25].Value = ExtractDate(ROW.Cells["fecha de ingreso"].Value.ToString());
            //dbcmd.Parameters[26].Value = "ESTA DATA FUE MIGRADA, POR FAVOR VERIFICAR LOS DATOS";

            //count += dbcmd.ExecuteNonQuery();
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
        
        private void button1_Click(object sender, EventArgs e)
        {

            DialogResult dr = folderBrowserDialog1.ShowDialog();
            if (dr == DialogResult.OK)
            {
                file = folderBrowserDialog1.SelectedPath;
                button2.Enabled = true;
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            button2.Enabled = true;
            ExtractA2Table();
        }
    }
}
