using System;
using System.Linq;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Npgsql;
using NpgsqlTypes;
using System.Data.OleDb;
using System.Globalization;
using LiteDB;

namespace MigradorXls
{

    public partial class Main : Form
    {
        string file;
        string sql;
        int count;
        int codInteno = 0;
        double porc1 = 0;
        double porc2 = 0;
        double total;
        int cantidad_items = 0;
        int nro_items = 0;
        int item = 0;
        string reader = "0";
        int status = 0;
        DialogResult result;

        private DataSet DtSet = new DataSet();
        System.Data.OleDb.OleDbConnection Myconnetion;
        System.Data.OleDb.OleDbDataAdapter MyCommand;
        //string de conexion con las credenciales de postgreSQL
        //string connectionString = @"Host=192.168.1.254;port=5432;Database=ROYALSDB;User ID=postgres;Password=TACA8tilo";
        //string connectionString = @"Host=localhost;port=5432;Database=MigrarPrueba;User ID=postgres;Password=postgres";
        string connectionString;
        List<Tipos> listaTipos = new List<Tipos>();
        List<Errores> listaErr = new List<Errores>();
        List<admin> listaAdmin = new List<admin>();
        List<payroll> listaPayRoll = new List<payroll>();
        public Main()
        {

            InitializeComponent();
            //Llamada del formulario de inicio de sesion
            //Creacion de lista que contiene los tipos de personas

            listaTipos.Add(new Tipos("PNR", "02.1"));
            listaTipos.Add(new Tipos("PNNR", "02.2"));
            listaTipos.Add(new Tipos("PJD", "02.3"));
            listaTipos.Add(new Tipos("PJND", "02.4"));
            listaTipos.Add(new Tipos("E", "02.5"));
            listaTipos.Add(new Tipos("Ord", "03.1"));
            listaTipos.Add(new Tipos("Esp", "03.2"));
            listaTipos.Add(new Tipos("For", "03.3"));
            listaTipos.Add(new Tipos("Exe", "03.4"));
            listaTipos.Add(new Tipos("Fin", "03.5"));
            listaTipos.Add(new Tipos("M", "05.1"));
            listaTipos.Add(new Tipos("F", "05.2"));
            listaTipos.Add(new Tipos("NA", "05.3"));
            listaTipos.Add(new Tipos("NORMAL", "11.1"));
            listaTipos.Add(new Tipos("SERIAL", "11.4"));
            listaTipos.Add(new Tipos("LOTE", "11.2"));
            listaTipos.Add(new Tipos("PROPIEDADES", "11.3"));
            listaTipos.Add(new Tipos("Fcxc", "20"));
            listaTipos.Add(new Tipos("Fcxp", "5"));
            listaTipos.Add(new Tipos("NC", "32"));
            listaTipos.Add(new Tipos("ND", "31"));

            var tipoPersona = new[] {
                new { Text = "NATURAL RESIDENTE", Value = "02.1" },
                new { Text = "NATURAL NO RESIDENTE", Value = "02.2" },
                new { Text = "JURIDICA DOMICILIADA", Value = "02.3" },
                new { Text = "JURIDICA NO DIMICILIADA", Value = "02.4" },
                new { Text = "EXTRANJERO", Value = "02.5" }
            };
            //Asignacion de data source y propiedades de visibilidad de campo tipo_pers
            //Creacion de lista que contiene los tipos de contribuyente
            var tipoContribuyente = new[] {
                new { Text = "ORDINARIO", Value = "03.1" },
                new { Text = "FORMAL", Value = "03.2" },
                new { Text = "ESPECIAL", Value = "03.3" },
                new { Text = "CONSUMIDOR FINAL", Value = "03.4" },
                new { Text = "EXENTO", Value = "03.5" }
            };
                       
            listaPayRoll.Add(new payroll("Profesiones", 1));
            listaPayRoll.Add(new payroll("Cargos", 2));

            using (var db = new LiteDatabase("Colleccion.db"))
            {
                var col = db.GetCollection<Tipos>("Tipos");
                col.InsertBulk(listaTipos);
            }

            //Asignacion de data source y propiedades de visibilidad de campo tipo_cont
            sql = @"SELECT org_hijo from admin.cfg_org";
            //Abriendo la coneccion con npgsql
            connectionString = @"Host=" + Globals.Host + ";port=" + Globals.port + ";Database=" + Globals.DB + ";User ID=" + Globals.usuario + ";Password=" + Globals.pass + ";";
            NpgsqlConnection conn = new NpgsqlConnection(connectionString);
            NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn);
            conn.Open();
            Globals.org = dbcmd.ExecuteScalar().ToString();
            conn.Close();
            using (var db = new LiteDatabase("Colleccion.db"))
            {
                var col = db.GetCollection<admin>("Admin");
                var Z = col.Find(Query.All());
                comboBox1.DataSource = Z.ToList();
                comboBox1.DisplayMember = "desc";
                comboBox1.ValueMember = "id";
            }
            
        }
        private void button1_Click(object sender, EventArgs e)
        {
            //seleccion del archivo a cargar
            DialogResult dr = openFileDialog1.ShowDialog();
            if (dr == DialogResult.OK)
            {
                file = openFileDialog1.FileName;
                try
                {
                    //inicio de conexion al archivo .xls
                    Myconnetion = new System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source =" + file + "; Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"");
                    Myconnetion.Open();
                    //carga de informacion del .xls

                    // obtener nombre de la hoja de excel
                    System.Data.DataTable dbSchema = Myconnetion.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    if (dbSchema == null || dbSchema.Rows.Count < 1)
                    {
                        throw new Exception("Error: No se pudo determinar el nombre de la Hoja de Trabajo.");
                    }
                    string firstSheetName = dbSchema.Rows[0]["TABLE_NAME"].ToString();
                    MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [" + firstSheetName + "]", Myconnetion);
                    MyCommand.TableMappings.Add("Table", "TestTable");

                    //Vaciado del DataSet
                    DtSet.Reset();
                    //Llenado del DataSet con resultado de comando ejecutado al .xls
                    MyCommand.Fill(DtSet);
                    //Asignacion del DataSet como origen de datos del DataGridView 
                    dataGridView1.DataSource = DtSet.Tables[0];
                    dataGridView1.Columns["Error"].Visible = true;
                    //Cerrar conexion
                    Myconnetion.Close();
                    button2.Enabled = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Se produjo un error al cargar la informacion. Error: " + ex.Message.ToString());
                }
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            codInteno = 1;
            status = 0;
            listaErr.Add(new Errores("Indeterminacion en division por zero", "22012"));
            listaErr.Add(new Errores("Formato incorrecto en la fecha", "22007"));
            listaErr.Add(new Errores("No se permiten valores nulos", "22004"));
            listaErr.Add(new Errores("Valor numerico fuera del rango establecido", "22003"));
            listaErr.Add(new Errores("Uno de los campos Introducidos posee un valor nulo", "23502"));
            listaErr.Add(new Errores("Violacion de llave foranea", "23503"));
            listaErr.Add(new Errores("Violacion de llave unica", "23505"));
            listaErr.Add(new Errores("Llave foranea invalida", "42830"));

            try
            {

                NpgsqlConnection conn = new NpgsqlConnection(connectionString);
                conn.Open();
                NpgsqlTransaction t = conn.BeginTransaction();
                //Recorriendo el Datagridview e insertando cada valor
                count = 0;
                if (comboBox1.SelectedIndex == 5)//Articulos
                {
                    Exportar_Articulos();
                }
                else if (comboBox1.SelectedIndex == 6)//Servicios
                {

                    Exportar_Servicios();

                }
                else
                {
                    foreach (DataGridViewRow ROW in dataGridView1.Rows)
                    {
                        try
                        {

                            switch (comboBox1.SelectedIndex)
                            {
                                case 0:
                                    {
                                        callbackInsertMoneda(conn, ROW, t);
                                        break;
                                    }
                                case 1:
                                    {

                                        callbackInsertZona(conn, ROW, t);
                                        break;
                                    }
                                case 2:
                                    {
                                        callbackInsertUnidades(conn, ROW, t);
                                        break;
                                    }
                                case 3:
                                    {
                                        callbackInsertDepartamento(conn, ROW, t);
                                        break;
                                    }
                                case 4:
                                    {
                                        callbackInsertImpuesto(conn, ROW, t);
                                        break;
                                    }
                                case 7:
                                    {

                                        callBackInsertVendedores(conn, ROW);
                                        break;
                                    }

                                case 8:
                                    {

                                        callBackInsertClientes(conn, ROW, t);
                                        break;
                                    }
                                case 9:
                                    {
                                        callbackInsertAutorizados(conn, ROW, t);
                                        break;
                                    }
                                case 10:
                                    {
                                        callbackInsertProveedores(conn, ROW, t);
                                        break;
                                    }
                                case 11:
                                    {
                                        if (ROW.Cells["cod cliente"].Value != null)
                                        {
                                            if (ROW.Cells["saldo"].Value.ToString() != "0")
                                            {
                                                callbackInsertCxC(conn, ROW, t);
                                                if (ROW.Cells["cod impuesto1"].Value.ToString() != "")
                                                {
                                                    callbackInsertCxCImp(conn, ROW, t, 1);
                                                }
                                                if (ROW.Cells["cod impuesto2"].Value.ToString() != "")
                                                {
                                                    callbackInsertCxCImp(conn, ROW, t, 2);
                                                }
                                            }
                                        }
                                        break;
                                    }
                                case 12:
                                    {

                                        if (ROW.Cells["cod proveedor"].Value != null)
                                        {
                                            if (ROW.Cells["saldo"].Value.ToString() != "0")
                                            {
                                                callbackInsertCxP(conn, ROW, t);
                                                if (ROW.Cells["cod impuesto1"].Value.ToString() != "")
                                                {
                                                    callbackInsertCxPImp(conn, ROW, t, 1);
                                                }
                                                if (ROW.Cells["cod impuesto2"].Value.ToString() != "")
                                                {
                                                    callbackInsertCxPImp(conn, ROW, t, 2);
                                                }
                                            }
                                        }
                                        break;
                                    }

                                case 14:
                                    {
                                        callbackInsertUsuario(conn, ROW, t);
                                        break;
                                    }

                            }

                        }
                        catch (NpgsqlException ex)
                        {
                            //Mensaje de error en la insercion de datos
                            foreach (var msj in listaErr.Where(s => s.codigo == ex.Code))
                            {
                                ROW.Cells["Error"].Value = msj.Desc;
                            }
                            ROW.DefaultCellStyle.BackColor = Color.Red;
                            count = 0;
                            break;
                            //MessageBox.Show("Hubo un Error en la insercion de datos. Excepcion: " + ex.Message.ToString());
                            // Cambio de color de la fila del DataGridView cuya insercion arrojo una excepcion                       

                        }
                        codInteno++;
                    }

                }
                t.Commit();
                conn.Close();
                Cursor.Current = Cursors.Default;
                MessageBox.Show(count + " Filas se almacenaron correctamente", "Atencion", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception EX)
            {
                MessageBox.Show("Se produjo un error al conectarse a la base de datos\nError: " + EX.Message, "Atencion", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (comboBox1.SelectedIndex == 16)
            {
                //Transacciones trans = new Transacciones();
                //trans.ShowDialog();
            }
            else button1.Enabled = true;
        }

        #region 
        
        /// <summary>
        /// Array de los porcentajes de utilidad
        /// </summary>
        public string[] arrayPU;
        private void Exportar_Servicios()
        {
            NpgsqlConnection conn = new NpgsqlConnection(connectionString);

            conn.Open();
            //Recorriendo el Datagridview e insertando cada valor

            NpgsqlTransaction t = conn.BeginTransaction();

            foreach (DataGridViewRow ROW in dataGridView1.Rows)
            {
                if (ROW.Cells["codigo"].Value != null)
                {
                    arrayPU = new string[]{
                    "30",
                    ROW.Cells["%utilidad1"].Value.ToString(),
                    ROW.Cells["%utilidad2"].Value.ToString(),
                    ROW.Cells["%utilidad3"].Value.ToString(),
                    ROW.Cells["%utilidad4"].Value.ToString(),

                };

                    try
                    {
                        callbackInsertServCategoria(conn, ROW, t);
                        callbackInsertServicios(conn, ROW, t);
                        callbackInsertCategoriaServicio(conn, ROW, t);
                        if (ROW.Cells["cod impuesto1"].Value.ToString() != null)
                        {
                            callbackInsertServicioImpuestos(conn, ROW, t, 1);
                        }
                        for (int i = 0; i < 5; i++)
                        {
                            callbackInsertServicioPrecio(conn, ROW, t, i);
                        }
                    }
                    catch (NpgsqlException ex)
                    {

                        foreach (var msj in listaErr.Where(s => s.codigo == ex.Code))
                        {
                            ROW.Cells["Error"].Value = msj.Desc;
                        }
                        ROW.DefaultCellStyle.BackColor = Color.Red;
                    }
                }

            }
            t.Commit();
            conn.Close();
        }
        private void Exportar_Articulos()
        {
            cantidad_items = 0;
            nro_items = 0;
            item = 0;
            //porcUtil f = new porcUtil((array) => {
            //    Select = true;
            //    arrayPU = array;
            //});
            //f.ShowDialog();



            //Abriendo la coneccion con npgsql
            NpgsqlConnection conn = new NpgsqlConnection(connectionString);

            conn.Open();
            //Recorriendo el Datagridview e insertando cada valor

            NpgsqlTransaction t = conn.BeginTransaction();

            foreach (DataGridViewRow ROW in dataGridView1.Rows)
            {
                if (ROW.Cells["codigo"].Value != null)
                {
                    arrayPU = new string[]{
                    "30",
                    ROW.Cells["%utilidad1"].Value.ToString(),
                    ROW.Cells["%utilidad2"].Value.ToString(),
                    ROW.Cells["%utilidad3"].Value.ToString(),
                    ROW.Cells["%utilidad4"].Value.ToString(),

                };
                    try
                    {
                        callbackInsertCategoria(conn, ROW, t);
                        callbackInsertDeposito(conn, ROW, t);
                        callbackInsertArticulo(conn, ROW, t);
                        callbackInsertCategoriaArticulo(conn, ROW, t);
                        if (ROW.Cells["cod impuesto1"].Value.ToString() != "")
                        {
                            callbackInsertArticuloImpuestos(conn, ROW, t, 1);
                        }
                        if (ROW.Cells["cod impuesto2"].Value.ToString() != "")
                        {
                            callbackInsertArticuloImpuestos(conn, ROW, t, 2);
                        }


                        for (int i = 0; i < 5; i++)
                        {
                            callbackInsertArticuloPrecio(conn, ROW, t, i);
                        }


                    }
                    catch (NpgsqlException ex)
                    {

                        foreach (var msj in listaErr.Where(s => s.codigo == ex.Code))
                        {
                            ROW.Cells["Error"].Value = msj.Desc;
                        }
                        ROW.DefaultCellStyle.BackColor = Color.Red;
                    }
                }
            }
            try
            {
                callbackInsertCargoInventario(conn, t);
            }
            catch (Exception ex)
            {
                //Mensaje de error en la insercion de datos
                MessageBox.Show(ex.ToString());
                //Cambio de color de la fila del DataGridView cuya insercion arrojo una excepcion

            }

            t.Commit();
            conn.Close();
        }
        private void callbackInsertZona(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            if (ROW.Cells["codigo"].Value != null)
            {
                //reader = "0";
                //NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
                //sql = @"select count(*) from admin.gen_zona";
                //dbcmd = new NpgsqlCommand(sql, conn);

                //reader = dbcmd.ExecuteScalar().ToString();
                //if (reader != "0")
                //{
                //    result = MessageBox.Show("Esta tabla ya posee registros \n ¿desea eliminarlos?","Atencion",MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                //    if (result == DialogResult.Yes)
                //    {
                //        sql = @"UPDATE admin.gen_banco SET cod_zona=null; 
                //                ALTER TABLE admin.gen_zona DISABLE TRIGGER tg_01_ft_delete_registro; 
                //                DELETE FROM admin.gen_zona;
                //                ALTER TABLE admin.gen_zona ENABLE TRIGGER tg_01_ft_delete_registro;";
                //        dbcmd = new NpgsqlCommand(sql, conn);
                //        reader = dbcmd.ExecuteScalar().ToString();
                //    }
                //}
                sql = @"INSERT INTO admin.gen_zona(org_hijo,cod_interno,codigo,descri,descorta,
                    latitud,longitud,altitud, reg_usu_cc, reg_usu_cu,reg_estatus, 
                    disponible,migrado) VALUES(@orgHijo , @codInterno,
                    @codigo, @descri, @descorta , @logitud, @latitud, 
                    @altitud, @reg_usu_cc , @reg_usu_cu, @regEstatus, @disponible, @migrado)";
                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
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
                dbcmd.Parameters[1].Value = codInteno;
                dbcmd.Parameters[2].Value = ROW.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[3].Value = ROW.Cells["descripcion"].Value;
                dbcmd.Parameters[4].Value = ROW.Cells["descripcion"].Value;
                dbcmd.Parameters[5].Value = 0;
                dbcmd.Parameters[6].Value = 0;
                dbcmd.Parameters[7].Value = 0;
                dbcmd.Parameters[8].Value = "INNOVA";
                dbcmd.Parameters[9].Value = "INNOVA";
                dbcmd.Parameters[10].Value = 1;
                dbcmd.Parameters[11].Value = convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                dbcmd.Parameters[12].Value = true;

                count += dbcmd.ExecuteNonQuery();
            }
        }

        private void callbackInsertUnidades(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            int estatus = 0;
            if (ROW.Cells["codigo"].Value != null)
            {
                reader = "0";
                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
                sql = @"select count(*) from admin.inv_medida";
                dbcmd = new NpgsqlCommand(sql, conn);

                reader = dbcmd.ExecuteScalar().ToString();
                if (reader != "0" && status == 0)
                {
                    result = MessageBox.Show("Esta tabla ya posee registros \n ¿desea eliminarlos?", "Atencion", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                    if (result == DialogResult.Yes)
                    {
                        try
                        {
                            sql = @"ALTER TABLE admin.inv_medida DISABLE TRIGGER tg_01_ft_delete_registro; 
                                DELETE FROM admin.inv_medida;
                                ALTER TABLE admin.inv_medida ENABLE TRIGGER tg_01_ft_delete_registro;";
                            dbcmd = new NpgsqlCommand(sql, conn);
                            dbcmd.ExecuteNonQuery();
                            status = 1;
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("Los registro se encuentran relacionados y no pueden ser borrados");
                            estatus = 1;
                        }
                    }
                }
                if (estatus == 0)
                {
                    sql = @"INSERT INTO admin.inv_medida(org_hijo,cod_interno,codigo,descri, reg_usu_cc, reg_usu_cu,reg_estatus, 
                    disponible,migrado) VALUES(@orgHijo , @codInterno,
                    @codigo, @descri, @reg_usu_cc , @reg_usu_cu, @regEstatus, @disponible, @migrado)";
                    dbcmd = new NpgsqlCommand(sql, conn, t);
                    dbcmd.Parameters.Add(new NpgsqlParameter("@orghijo", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@codInterno", NpgsqlDbType.Bigint));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@codigo", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cu", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@disponible", NpgsqlDbType.Boolean));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));

                    dbcmd.Prepare();

                    dbcmd.Parameters[0].Value = Globals.org;
                    dbcmd.Parameters[1].Value = codInteno;
                    dbcmd.Parameters[2].Value = ROW.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);
                    dbcmd.Parameters[3].Value = ROW.Cells["descripcion"].Value;
                    dbcmd.Parameters[4].Value = "INNOVA";
                    dbcmd.Parameters[5].Value = "INNOVA";
                    dbcmd.Parameters[6].Value = 1;
                    dbcmd.Parameters[7].Value = convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                    dbcmd.Parameters[8].Value = true;

                    count += dbcmd.ExecuteNonQuery();
                }
            }

        }

        private void callbackInsertDepartamento(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            int estatus = 0;
            if (ROW.Cells["codigo"].Value != null)
            {
                reader = "0";
                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
                sql = @"select count(*) from admin.org_depar";
                dbcmd = new NpgsqlCommand(sql, conn);

                reader = dbcmd.ExecuteScalar().ToString();
                if (reader != "0" && status == 0)
                {
                    result = MessageBox.Show("Esta tabla ya posee registros \n ¿desea eliminarlos?", "Atencion", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                    if (result == DialogResult.Yes)
                    {
                        try
                        {
                            sql = @"DELETE FROM admin.org_depar";
                            dbcmd = new NpgsqlCommand(sql, conn);
                            dbcmd.ExecuteNonQuery();
                            status = 1;
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("Los registro se encuentran relacionados y no pueden ser borrados");
                            estatus = 1;
                        }
                    }
                }
                if (estatus == 0)
                {
                    sql = @"INSERT INTO admin.org_depar(org_hijo,codigo,descri, reg_usu_cc, reg_usu_cu,reg_estatus, 
                    disponible,migrado) VALUES(@orgHijo ,@codigo, @descri, @reg_usu_cc , @reg_usu_cu, @regEstatus,
                    @disponible, @migrado)";
                    dbcmd = new NpgsqlCommand(sql, conn, t);
                    dbcmd.Parameters.Add(new NpgsqlParameter("@orghijo", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@codigo", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cu", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@disponible", NpgsqlDbType.Boolean));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));

                    dbcmd.Prepare();

                    dbcmd.Parameters[0].Value = Globals.org;
                    dbcmd.Parameters[1].Value = ROW.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);
                    dbcmd.Parameters[2].Value = ROW.Cells["descripcion"].Value;
                    dbcmd.Parameters[3].Value = "INNOVA";
                    dbcmd.Parameters[4].Value = "INNOVA";
                    dbcmd.Parameters[5].Value = 1;
                    dbcmd.Parameters[6].Value = convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                    dbcmd.Parameters[7].Value = true;

                    count += dbcmd.ExecuteNonQuery();
                }
            }

        }
        private void callbackInsertMoneda(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            if (ROW.Cells["codigo"].Value != null)
            {
                sql = @"INSERT INTO admin.gen_moneda(org_hijo,cod_interno,codigo,descri,descorta,simbolo,
                    factor,ant_factor, reg_usu_cc, reg_usu_cu,reg_estatus, 
                    disponible,migrado) VALUES(@orgHijo , @codInterno,
                    @codigo, @descri, @descorta ,@simbolo, @factor, @antFactor, 
                    @reg_usu_cc , @reg_usu_cu, @regEstatus, @disponible, @migrado)";
                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
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
                dbcmd.Parameters[1].Value = codInteno;
                dbcmd.Parameters[2].Value = ROW.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[3].Value = ROW.Cells["descripcion"].Value;
                dbcmd.Parameters[4].Value = ROW.Cells["descripcion"].Value;
                dbcmd.Parameters[5].Value = ROW.Cells["simbolo"].Value;
                dbcmd.Parameters[6].Value = ROW.Cells["factor"].Value;
                dbcmd.Parameters[7].Value = 1;
                dbcmd.Parameters[8].Value = "INNOVA";
                dbcmd.Parameters[9].Value = "INNOVA";
                dbcmd.Parameters[10].Value = 1;
                dbcmd.Parameters[11].Value = convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                dbcmd.Parameters[12].Value = true;

                count += dbcmd.ExecuteNonQuery();
            }

        }

        private void callbackInsertAutorizados(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            if (ROW.Cells["codigo"].Value != null)
            {
                sql = @"INSERT INTO admin.ven_cli_autorizado(org_hijo,cli_hijo,codigo,descri,rif,telefono,
                    fax,email, direccion,fecha_nac, disponible,migrado) 
                    VALUES(@orgHijo , @cli_hijo,
                    @codigo, @descri, @rif, @telefono ,@fax, @email,@fechanac, @direccion, 
                    @disponible, @migrado)";
                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
                dbcmd.Parameters.Add(new NpgsqlParameter("@orghijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@cli_hijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codigo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@rif", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@telefono", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@fax", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@email", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@direccion", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@fechanac", NpgsqlDbType.Date));
                dbcmd.Parameters.Add(new NpgsqlParameter("@disponible", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));

                dbcmd.Prepare();

                dbcmd.Parameters[0].Value = Globals.org;
                dbcmd.Parameters[1].Value = ROW.Cells["codigoCliente"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[2].Value = ROW.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[3].Value = ROW.Cells["descripcion"].Value;
                dbcmd.Parameters[4].Value = ROW.Cells["rif"].Value;
                dbcmd.Parameters[5].Value = ROW.Cells["telefono"].Value;
                dbcmd.Parameters[6].Value = ROW.Cells["fax"].Value;
                dbcmd.Parameters[7].Value = ROW.Cells["correo"].Value;
                dbcmd.Parameters[8].Value = ROW.Cells["direccion"].Value;
                dbcmd.Parameters[9].Value = ROW.Cells["fecha_nacimiento"].Value;
                dbcmd.Parameters[10].Value = convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                dbcmd.Parameters[11].Value = true;

                count += dbcmd.ExecuteNonQuery();
            }

        }

        private void callBackInsertVendedores(NpgsqlConnection conn, DataGridViewRow ROW)
        {

            if (ROW.Cells["codigo"].Value != null)
            {
                sql = @"INSERT INTO admin.org_talento(org_hijo,cod_interno,
                        codigo,cedula,descri,es_vendedor,es_cobrador,es_servidor,
                        es_despachador,fecha_nac,reg_usu_cc,reg_usu_cu,
                        reg_estatus,disponible,tipo_cont,tipo_pers,cod_zona,migrado, 
                        rif, descorta, sexo, direc1, cod_depar, porc_retencion, fecha_rif, fecha_ing, observacion )
                        VALUES(@orgHijo , @codInterno, @codigo, @cedula, @descri, 
                        @esVendedor,@esCobrador, @esServidor, @esDespachador , @fechaNac, 
                        @reg_usu_cc , @reg_usu_cu, @regEstatus, @disponible, @tipoCont, @tipoPers, 
                        @codZona , @migrado, @rif, @descorta, @sexo, @direc1,
                        @cod_depar, @porc_retencion, @fecharif, @fecha_ing, @observacion)";
                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn);


                dbcmd.Parameters.Add(new NpgsqlParameter("@orghijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codInterno", NpgsqlDbType.Bigint));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codigo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@cedula", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@esVendedor", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@esCobrador", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@esServidor", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@esDespachador", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@fechaNac", NpgsqlDbType.Date));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cu", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
                dbcmd.Parameters.Add(new NpgsqlParameter("@disponible", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@tipoCont", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@tipoPers", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codZona", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@rif", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descorta", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@sexo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@direc1", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@porc_retencion", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@cod_depar", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@fecharif", NpgsqlDbType.Date));
                dbcmd.Parameters.Add(new NpgsqlParameter("@fecha_ing", NpgsqlDbType.Date));
                dbcmd.Parameters.Add(new NpgsqlParameter("@observacion", NpgsqlDbType.Varchar));


                dbcmd.Prepare();


                dbcmd.Parameters[0].Value = Globals.org;
                dbcmd.Parameters[1].Value = codInteno;
                dbcmd.Parameters[2].Value = ROW.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[3].Value = ROW.Cells["cedula"].Value;
                dbcmd.Parameters[4].Value = ROW.Cells["nombres"].Value;
                dbcmd.Parameters[5].Value = convertBoolean(ROW.Cells["vendedor"].Value);
                dbcmd.Parameters[6].Value = convertBoolean(ROW.Cells["cobrador"].Value);
                dbcmd.Parameters[7].Value = convertBoolean(ROW.Cells["servidor"].Value);
                dbcmd.Parameters[8].Value = convertBoolean(ROW.Cells["despachador"].Value);
                dbcmd.Parameters[9].Value = ExtractDate(ROW.Cells["fecha nac"].Value.ToString());
                dbcmd.Parameters[10].Value = "INNOVA";
                dbcmd.Parameters[11].Value = "INNOVA";
                dbcmd.Parameters[12].Value = 1;
                dbcmd.Parameters[13].Value = convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                foreach (var z in listaTipos.Where(a => a.tipo == ROW.Cells["tipo contribuyente"].Value.ToString()))
                {
                    dbcmd.Parameters[14].Value = z.codigo;
                }

                foreach (var z in listaTipos.Where(a => a.tipo == ROW.Cells["tipo persona"].Value.ToString()))
                {
                    dbcmd.Parameters[15].Value = z.codigo;
                }
                dbcmd.Parameters[16].Value = ROW.Cells["codigo de zona"].Value;
                dbcmd.Parameters[17].Value = true;
                dbcmd.Parameters[18].Value = ROW.Cells["rif"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[19].Value = ROW.Cells["apellidos"].Value;
                foreach (var z in listaTipos.Where(a => a.tipo == ROW.Cells["Sexo"].Value.ToString()))
                {
                    dbcmd.Parameters[20].Value = z.codigo;
                }
                dbcmd.Parameters[21].Value = ROW.Cells["direccion"].Value;
                dbcmd.Parameters[22].Value = ROW.Cells["porc ret iva"].Value;
                dbcmd.Parameters[23].Value = ROW.Cells["departamento"].Value;
                dbcmd.Parameters[24].Value = ExtractDate(ROW.Cells["fecha vcto rif"].Value.ToString());
                dbcmd.Parameters[25].Value = ExtractDate(ROW.Cells["fecha de ingreso"].Value.ToString());
                dbcmd.Parameters[26].Value = "ESTA DATA FUE MIGRADA, POR FAVOR VERIFICAR LOS DATOS";

                count += dbcmd.ExecuteNonQuery();
            }

        }

        private void callBackInsertClientes(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            if (ROW.Cells["código"].Value != null && ROW.Cells["código"].Value.ToString() != "")
            {
                sql = @"INSERT INTO admin.ven_cli(org_hijo,cod_interno,cli_hijo,descri,
                        tipo_cont,tipo_pers,porc_ret_iva,rif,direc1,monto_descuento,es_descuento,
                        es_exento,es_retencion,es_monto,monto_min,monto_max,monto_cred_max,monto_acum,
                        pri_vmonto,ult_vmonto,ult_pmonto,pago_max,pago_adel,longitud, latitud, altitud,
                        pago_prom,monto_cred_min,saldo, reg_usu_cc,reg_usu_cu,reg_estatus,disponible, 
                        migrado,es_datos,es_vip,es_pronto, observacion, tipo_ret_iva) 
                        VALUES(@org_hijo,@cod_interno,@cli_hijo,@descri,@tipo_cont,@tipo_pers,
                        @PorcRetIva,@rif,@direc1,@descuento,@es_descuento,
                        @es_exento,@es_retencion,@es_monto,@monto_min,@monto_max,@monto_cred_max,@monto_acum,
                        @pri_vmonto,@ult_vmonto,@ult_pmonto,@pago_max,@pago_adel,@longitud,@latitud,@altitud,
                        @pago_prom,@montoCredMin,@saldo, @reg_usu_cc,@reg_usu_cu,@reg_estatus,@disponible, 
                        @migrado, @esdatos, @esvip, @espronto, @observacion, @tipo_ret_iva)";


                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
                dbcmd.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@cod_interno", NpgsqlDbType.Bigint));
                dbcmd.Parameters.Add(new NpgsqlParameter("@cli_hijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@tipo_cont", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@tipo_pers", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@rif", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@direc1", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descuento", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@es_descuento", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@es_exento", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@es_retencion", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@es_monto", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@monto_min", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@monto_max", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@monto_cred_max", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@pri_vmonto", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@ult_vmonto", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@ult_pmonto", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@pago_max", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@pago_adel", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@pago_prom", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@saldo", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cu", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_estatus", NpgsqlDbType.Integer));
                dbcmd.Parameters.Add(new NpgsqlParameter("@disponible", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@PorcRetIva", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@monto_acum", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@montoCredMin", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@longitud", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@latitud", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@altitud", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@esdatos", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@esvip", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@espronto", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@observacion", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@tipo_ret_iva", NpgsqlDbType.Varchar));


                dbcmd.Prepare();

                dbcmd.Parameters[0].Value = Globals.org;
                dbcmd.Parameters[1].Value = codInteno;
                dbcmd.Parameters[2].Value = ROW.Cells["código"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[3].Value = ROW.Cells["razon social"].Value;
                foreach (var z in listaTipos.Where(a => a.tipo == ROW.Cells["tipo de contribuyente"].Value.ToString().Replace(" ", string.Empty)))
                {
                    dbcmd.Parameters[4].Value = z.codigo;
                }

                foreach (var z in listaTipos.Where(a => a.tipo == ROW.Cells["tipo de persona"].Value.ToString().Replace(" ", string.Empty)))
                {
                    dbcmd.Parameters[5].Value = z.codigo;
                }
                dbcmd.Parameters[6].Value = ROW.Cells["rif/cedula"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[7].Value = ROW.Cells["dirección"].Value;
                dbcmd.Parameters[8].Value = ROW.Cells["descuento"].Value;
                dbcmd.Parameters[9].Value = false;
                dbcmd.Parameters[10].Value = false;
                dbcmd.Parameters[11].Value = false;
                dbcmd.Parameters[12].Value = false;
                dbcmd.Parameters[13].Value = ROW.Cells["Monto mínimo de una venta"].Value;
                dbcmd.Parameters[14].Value = ROW.Cells["Monto máximo de una venta"].Value;
                dbcmd.Parameters[15].Value = ROW.Cells["Monto máximo de crédito"].Value;
                dbcmd.Parameters[16].Value = ROW.Cells["Monto primera venta"].Value;
                dbcmd.Parameters[17].Value = ROW.Cells["Monto última venta"].Value;
                dbcmd.Parameters[18].Value = ROW.Cells["Monto último pago recibido"].Value;
                dbcmd.Parameters[19].Value = ROW.Cells["Monto máximo de pago recibido"].Value;
                dbcmd.Parameters[20].Value = ROW.Cells["Monto adelanto"].Value;
                dbcmd.Parameters[21].Value = ROW.Cells["Monto promedio de pagos recbidos"].Value;
                dbcmd.Parameters[22].Value = ROW.Cells["Saldo del cliente"].Value;
                dbcmd.Parameters[23].Value = "INNOVA";
                dbcmd.Parameters[24].Value = "INNOVA";
                dbcmd.Parameters[25].Value = 1;
                dbcmd.Parameters[26].Value = convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                dbcmd.Parameters[27].Value = true;
                if (ROW.Cells["tipo de contribuyente"].Value.ToString().Replace(" ", string.Empty) == "Esp")
                {
                    dbcmd.Parameters[28].Value = 75;
                    dbcmd.Parameters[38].Value = "15.2";
                }

                else
                {
                    dbcmd.Parameters[28].Value = 0;
                    dbcmd.Parameters[38].Value = "15.3";
                }
                dbcmd.Parameters[29].Value = ROW.Cells["Monto acumulado de crédito"].Value;
                dbcmd.Parameters[30].Value = ROW.Cells["Monto minimo aceptable de crédito"].Value;
                dbcmd.Parameters[31].Value = 0;
                dbcmd.Parameters[32].Value = 0;
                dbcmd.Parameters[33].Value = 0;
                dbcmd.Parameters[34].Value = true;
                dbcmd.Parameters[35].Value = true;
                dbcmd.Parameters[36].Value = true;
                dbcmd.Parameters[37].Value = "ESTA DATA FUE MIGRADA, POR FAVOR VERIFICAR LOS DATOS";
                count += dbcmd.ExecuteNonQuery();
            }
            codInteno++;
        }

        private void callbackInsertDeposito(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            string reader = "0";
            NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
            sql = @"select count(*) from admin.inv_dep where codigo='" + ROW.Cells["codigo deposito"].Value + "'";
            dbcmd = new NpgsqlCommand(sql, conn);

            reader = dbcmd.ExecuteScalar().ToString();
            if (reader == "0")
            {
                sql = @"INSERT INTO admin.inv_dep(org_hijo,
                        cod_interno,codigo,descri,maximo,minimo,espacio_mq,espacio_vol,espacio_uso,
                        reg_usu_cc,reg_usu_cu,reg_estatus,disponible) 
                        VALUES(@org_hijo , @codInterno, @codigo, @descri,@maximo,@minimo,
                        @espaciomq,@espaciovol,@espaciouso,@reg_usu_cc, @reg_usu_cu, 
                        @regEstatus, @disponible)";


                dbcmd = new NpgsqlCommand(sql, conn, t);

                dbcmd.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codInterno", NpgsqlDbType.Bigint));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codigo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@maximo", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@minimo", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@espaciomq", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@espaciovol", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@espaciouso", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cu", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
                dbcmd.Parameters.Add(new NpgsqlParameter("@disponible", NpgsqlDbType.Boolean));

                dbcmd.Prepare();

                dbcmd.Parameters[0].Value = Globals.org;
                dbcmd.Parameters[1].Value = codInteno;
                dbcmd.Parameters[2].Value = ROW.Cells["codigo deposito"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[3].Value = ROW.Cells["descripcion deposito"].Value;
                dbcmd.Parameters[4].Value = 0;
                dbcmd.Parameters[5].Value = 0;
                dbcmd.Parameters[6].Value = 0;
                dbcmd.Parameters[7].Value = 0;
                dbcmd.Parameters[8].Value = 0;
                dbcmd.Parameters[9].Value = "INNOVA";
                dbcmd.Parameters[10].Value = "INNOVA";
                dbcmd.Parameters[11].Value = 1;
                dbcmd.Parameters[12].Value = convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                count += dbcmd.ExecuteNonQuery();
            }
        }

        private void callbackInsertCategoria(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            string reader = "0";
            NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
            sql = @"select count(*) from admin.inv_cat where cat_hijo ='" + ROW.Cells["codigo categoria"].Value.ToString() + "'";
            dbcmd = new NpgsqlCommand(sql, conn);

            reader = dbcmd.ExecuteScalar().ToString();
            if (reader == "0")
            {
                sql = @"INSERT INTO admin.inv_cat(org_hijo,
                        cod_interno,cat_hijo,descri,reg_usu_cc,reg_usu_cu,
                        reg_estatus,disponible) 
                        VALUES(@org_hijo, @codInterno, @catHijo, @descri,
                        @reg_usu_cc, @reg_usu_cu, @regEstatus, @disponible)";


                dbcmd = new NpgsqlCommand(sql, conn, t);

                dbcmd.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codInterno", NpgsqlDbType.Bigint));
                dbcmd.Parameters.Add(new NpgsqlParameter("@catHijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cu", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
                dbcmd.Parameters.Add(new NpgsqlParameter("@disponible", NpgsqlDbType.Boolean));

                dbcmd.Prepare();

                dbcmd.Parameters[0].Value = Globals.org;
                dbcmd.Parameters[1].Value = codInteno;
                dbcmd.Parameters[2].Value = ROW.Cells["codigo categoria"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[3].Value = ROW.Cells["descripcion categoria"].Value;
                dbcmd.Parameters[4].Value = "INNOVA";
                dbcmd.Parameters[5].Value = "INNOVA";
                dbcmd.Parameters[6].Value = 1;
                dbcmd.Parameters[7].Value = convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                count += dbcmd.ExecuteNonQuery();
            }

        }


        private void callbackInsertCategoriaArticulo(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            string reader = "0";
            NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
            sql = @"select count(*) from admin.inv_cat_art where cat_hijo ='" + ROW.Cells["codigo categoria"].Value.ToString() + "' AND cod_articulo='" + ROW.Cells["codigo"].Value.ToString() + "'";
            dbcmd = new NpgsqlCommand(sql, conn);

            reader = dbcmd.ExecuteScalar().ToString();
            if (reader == "0")
            {
                sql = @"INSERT INTO admin.inv_cat_art(org_hijo,
                        cod_interno,cat_hijo,cod_articulo,reg_usu_cc,reg_estatus,disponible) 
                        VALUES(@org_hijo, @codInterno, @catHijo, @codArticulo,
                        @reg_usu_cc, @regEstatus, @disponible)";


                dbcmd = new NpgsqlCommand(sql, conn, t);

                dbcmd.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codInterno", NpgsqlDbType.Bigint));
                dbcmd.Parameters.Add(new NpgsqlParameter("@catHijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codArticulo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
                dbcmd.Parameters.Add(new NpgsqlParameter("@disponible", NpgsqlDbType.Boolean));

                dbcmd.Prepare();

                dbcmd.Parameters[0].Value = Globals.org;
                dbcmd.Parameters[1].Value = codInteno;
                dbcmd.Parameters[2].Value = ROW.Cells["codigo categoria"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[3].Value = ROW.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[4].Value = "INNOVA";
                dbcmd.Parameters[5].Value = 1;
                dbcmd.Parameters[6].Value = convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                count += dbcmd.ExecuteNonQuery();
            }

        }

        private void callbackInsertCategoriaServicio(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            string reader = "0";
            NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
            sql = @"select count(*) from admin.gen_servicio_cat_serv where cat_hijo ='" + ROW.Cells["codigo categoria"].Value.ToString() + "' AND cod_servicio='" + ROW.Cells["codigo"].Value.ToString() + "'";
            dbcmd = new NpgsqlCommand(sql, conn);

            reader = dbcmd.ExecuteScalar().ToString();
            if (reader == "0")
            {
                sql = @"INSERT INTO admin.gen_servicio_cat_serv(org_hijo,
                        cat_hijo,cod_servicio,reg_estatus,disponible) 
                        VALUES(@org_hijo,@catHijo, @codArticulo, @regEstatus, @disponible)";


                dbcmd = new NpgsqlCommand(sql, conn, t);

                dbcmd.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@catHijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codArticulo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
                dbcmd.Parameters.Add(new NpgsqlParameter("@disponible", NpgsqlDbType.Boolean));

                dbcmd.Prepare();

                dbcmd.Parameters[0].Value = Globals.org;
                dbcmd.Parameters[1].Value = ROW.Cells["codigo categoria"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[2].Value = ROW.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[3].Value = 1;
                dbcmd.Parameters[4].Value = convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                count += dbcmd.ExecuteNonQuery();
            }

        }
        private void callbackInsertServCategoria(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            string reader = "0";
            NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
            sql = @"select count(*) from admin.gen_servicio_cat where cat_hijo ='" + ROW.Cells["codigo categoria"].Value.ToString() + "'";
            dbcmd = new NpgsqlCommand(sql, conn);

            reader = dbcmd.ExecuteScalar().ToString();
            if (reader == "0")
            {
                sql = @"INSERT INTO admin.gen_servicio_cat(org_hijo,
                        cod_interno,cat_hijo,descri,reg_usu_cc,reg_usu_cu,
                        reg_estatus,disponible) 
                        VALUES(@org_hijo, @codInterno, @catHijo, @descri,
                        @reg_usu_cc, @reg_usu_cu, @regEstatus, @disponible)";


                dbcmd = new NpgsqlCommand(sql, conn, t);

                dbcmd.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codInterno", NpgsqlDbType.Bigint));
                dbcmd.Parameters.Add(new NpgsqlParameter("@catHijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cu", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
                dbcmd.Parameters.Add(new NpgsqlParameter("@disponible", NpgsqlDbType.Boolean));

                dbcmd.Prepare();

                dbcmd.Parameters[0].Value = Globals.org;
                dbcmd.Parameters[1].Value = codInteno;
                dbcmd.Parameters[2].Value = ROW.Cells["codigo categoria"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[3].Value = ROW.Cells["descripcion categoria"].Value;
                dbcmd.Parameters[4].Value = "INNOVA";
                dbcmd.Parameters[5].Value = "INNOVA";
                dbcmd.Parameters[6].Value = 1;
                dbcmd.Parameters[7].Value = convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                count += dbcmd.ExecuteNonQuery();
            }

        }
        private void callbackInsertArticulo(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            if (ROW.Cells["codigo"].Value != null)
            {
                string reader = "0";
                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
                sql = @"select count(*) from admin.inv_art where codigo ='" + ROW.Cells["codigo"].Value.ToString() + "'";
                dbcmd = new NpgsqlCommand(sql, conn);

                reader = dbcmd.ExecuteScalar().ToString();
                nro_items++;
                if (reader == "0")
                {
                    sql = @"INSERT INTO admin.inv_art(org_hijo,
                        cod_interno,codigo,descri,cantidad,
                        cant_compro,cant_pedido,cant_consumo,cant_venta,
                        cant_max,cant_min,cant_falla,cant_repos,cant_bulto,
                        es_pos,es_bien,es_activo,es_bulto,es_venc,es_medida,
                        es_peso,es_oferta,es_exento,es_retencion,es_regulado,
                        es_exonerado,es_unico,es_decimal,es_unidad,es_parte,
                        tipo_art,costo,costo_pro,costo_ant,costo_rep,med_peso, 
                        med_alto, med_ancho, med_largo ,med_volumen,
                        reg_usu_cc,reg_usu_cu,reg_estatus,disponible,cod_medida,
                        costo_pro_ant,cantidad_ant,precio,migrado,es_credito_fiscal) 
                        VALUES(@org_hijo , @codInterno, @codigo, @descri,
                        @cantidad, @cantCompro, @cantPedido , @cantConsumo, 
                        @cantVenta, @cantMax, @cantMin, @cantFalla , @cantRepos , @cantBulto, 
                        @esPos, @esBien, @esActivo, @esBulto, @esVenc, 
                        @esMedida, @esPeso, @esOferta, @esExento, @esRetencion, @esRegulado,
                        @esExonerado, @esUnico, @esDecimal, @esUnidad, @esParte , 
                        @tipoArt, @costo,@costoPro,@costoAnt,@costoRep, @medPeso , @medAlto ,
                        @medAncho, @medLargo, @medVolumen,@reg_usu_cc, @reg_usu_cu, @regEstatus, 
                        @disponible, @codMedida,@costoProAnt, @cantidadAnt, @precio, @migrado, @esCreditoFiscal)";

                    //            //dbcmd.CommandText = sql;
                    dbcmd = new NpgsqlCommand(sql, conn, t);

                    dbcmd.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@codInterno", NpgsqlDbType.Bigint));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@codigo", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@cantidad", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@cantCompro", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@cantPedido", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@cantConsumo", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@cantVenta", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@cantMax", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@cantMin", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@cantFalla", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@cantRepos", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@cantBulto", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@esPos", NpgsqlDbType.Boolean));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@esBien", NpgsqlDbType.Boolean));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@esActivo", NpgsqlDbType.Boolean));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@esBulto", NpgsqlDbType.Boolean));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@esVenc", NpgsqlDbType.Boolean));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@esMedida", NpgsqlDbType.Boolean));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@esPeso", NpgsqlDbType.Boolean));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@esOferta", NpgsqlDbType.Boolean));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@esExento", NpgsqlDbType.Boolean));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@esRetencion", NpgsqlDbType.Boolean));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@esRegulado", NpgsqlDbType.Boolean));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@esExonerado", NpgsqlDbType.Boolean));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@esUnico", NpgsqlDbType.Boolean));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@esDecimal", NpgsqlDbType.Boolean));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@esUnidad", NpgsqlDbType.Boolean));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@esParte", NpgsqlDbType.Boolean));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@tipoArt", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@costo", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@costoPro", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@costoAnt", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@costoRep", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@medPeso", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@medAlto", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@medAncho", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@medLargo", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@medVolumen", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cu", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@disponible", NpgsqlDbType.Boolean));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@codMedida", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@costoProAnt", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@cantidadAnt", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@precio", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@esCreditoFiscal", NpgsqlDbType.Boolean));

                    dbcmd.Prepare();

                    dbcmd.Parameters[0].Value = Globals.org; //ORG_HIJO
                    dbcmd.Parameters[1].Value = codInteno;  //COD_INTERNO
                    dbcmd.Parameters[2].Value = ROW.Cells["codigo"].Value.ToString().Replace(" ", string.Empty); //CODIGO
                    dbcmd.Parameters[3].Value = ROW.Cells["descripcion del producto"].Value;    //DESCRI
                    dbcmd.Parameters[4].Value = 0;  //CANTIDAD  
                    dbcmd.Parameters[5].Value = 0;  //CANTIDAD COMPRA
                    dbcmd.Parameters[6].Value = 0;  //CANTIDAD PEDIDO
                    dbcmd.Parameters[7].Value = 0;  //CANTIDAD CONSUMO 
                    dbcmd.Parameters[8].Value = 0;  //CANTIDAD VENTA
                    dbcmd.Parameters[9].Value = 0;  //CANTIDAD MAXIMA
                    dbcmd.Parameters[10].Value = 0; //CANTIDAD MINIMA
                    dbcmd.Parameters[11].Value = 0; //CANTIDAD FALLA
                    dbcmd.Parameters[12].Value = 0; //CANTIDAD REPOSICION
                    dbcmd.Parameters[13].Value = 0; //CANTIDAD BULTO
                    dbcmd.Parameters[14].Value = false; //ES POSESION
                    dbcmd.Parameters[15].Value = false; //ES BIEN
                    dbcmd.Parameters[16].Value = false; //ES ACTIVO
                    dbcmd.Parameters[17].Value = false; //ES BULTO
                    dbcmd.Parameters[18].Value = false; //ES VENCIMIENTO
                    dbcmd.Parameters[19].Value = false; //ES MEDIDA
                    dbcmd.Parameters[20].Value = false; //ES PESO
                    dbcmd.Parameters[21].Value = false; //ES OFERTA
                    dbcmd.Parameters[22].Value = convertBoolean(ROW.Cells["Exento"].Value); //ES EXENTO
                    dbcmd.Parameters[23].Value = false; //ES RETENCION
                    dbcmd.Parameters[24].Value = false; //ES REGULADO
                    dbcmd.Parameters[25].Value = convertBoolean(ROW.Cells["Exonerado"].Value);  //ES EXONERADO
                    dbcmd.Parameters[26].Value = false; //ES UNICO
                    dbcmd.Parameters[27].Value = false; //ES DECIMAL
                    dbcmd.Parameters[28].Value = false; //ES UNIDAD
                    dbcmd.Parameters[29].Value = false; //ES PARTE
                    foreach (var z in listaTipos.Where(a => a.tipo == ROW.Cells["tipo de Articulo"].Value.ToString()))
                    {
                        dbcmd.Parameters[30].Value = z.codigo;  //TIPO ARTICULO
                    }
                    dbcmd.Parameters[31].Value = 0; //COSTO
                    dbcmd.Parameters[32].Value = 0; //COSTO PROMEDIO
                    dbcmd.Parameters[33].Value = 0; //COSTO ANTERIOR
                    dbcmd.Parameters[34].Value = 0; //COSTO REPOSICION
                    dbcmd.Parameters[35].Value = 0; //MEDIDA PESO
                    dbcmd.Parameters[36].Value = 0; //MEDIDA ALTO
                    dbcmd.Parameters[37].Value = 0; //MEDIDA AMCHO
                    dbcmd.Parameters[38].Value = 0; //MEDIDA LARGO
                    dbcmd.Parameters[39].Value = 0; //MEDIDA VOLUMEN
                    dbcmd.Parameters[40].Value = "INNOVA"; //USUARIO QUE REGISTRO
                    dbcmd.Parameters[41].Value = "INNOVA"; //USUARIO QUE MODIFICO
                    dbcmd.Parameters[42].Value = 1; //ESTATUS DE REGISTRO
                    dbcmd.Parameters[43].Value = convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value); //DISPONIBLE
                    dbcmd.Parameters[44].Value = ROW.Cells["codigo de la unidad de medida"].Value; //CODIGO MEDIDA
                    dbcmd.Parameters[45].Value = 0; //COSTO PROMEDIO ANTERIOR
                    dbcmd.Parameters[46].Value = 0; //CANTIDAD ANTERIOR
                    dbcmd.Parameters[47].Value = 0; //PRECIO
                    dbcmd.Parameters[48].Value = true;  //MIGRADO
                    dbcmd.Parameters[49].Value = false; //ES CREDITO FISCAL


                    count += dbcmd.ExecuteNonQuery();
                    total += (Convert.ToDouble(ROW.Cells["costo"].Value) * Convert.ToDouble(ROW.Cells["existencia"].Value));
                    cantidad_items += Convert.ToInt32(ROW.Cells["existencia"].Value);
                }

            }
            codInteno++;
        }

        private void callbackInsertServicios(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            if (ROW.Cells["codigo"].Value != null)
            {

                sql = @"INSERT INTO admin.gen_servicio(org_hijo,
                        cod_interno,codigo,descri,costo,costo_pro,costo_anterior,
                        observacion,reg_usu_cc,reg_usu_cu,reg_estatus,disponible,unidad,
                        migrado) 
                        VALUES(@org_hijo , @codInterno, @codigo, @descri,
                         @costo,@costoPro,@costoAnt,@observacion,@reg_usu_cc, @reg_usu_cu, @regEstatus, 
                        @disponible, @codMedida, @migrado)";


                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);

                dbcmd.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codInterno", NpgsqlDbType.Bigint));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codigo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@costo", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@costoPro", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@costoAnt", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@observacion", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cu", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
                dbcmd.Parameters.Add(new NpgsqlParameter("@disponible", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codMedida", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));


                dbcmd.Prepare();

                dbcmd.Parameters[0].Value = Globals.org; //ORG_HIJO
                dbcmd.Parameters[1].Value = codInteno;  //COD_INTERNO
                dbcmd.Parameters[2].Value = ROW.Cells["codigo"].Value.ToString().Replace(" ", string.Empty); //CODIGO
                dbcmd.Parameters[3].Value = ROW.Cells["descripcion del servicio"].Value;    //DESCRI
                dbcmd.Parameters[4].Value = ROW.Cells["costo"].Value; //COSTO
                dbcmd.Parameters[5].Value = ROW.Cells["costo promedio"].Value; //COSTO PROMEDIO
                dbcmd.Parameters[6].Value = 0; //COSTO ANTERIOR
                dbcmd.Parameters[7].Value = ROW.Cells["descripcion detallada"].Value;
                dbcmd.Parameters[8].Value = "INNOVA"; //USUARIO QUE REGISTRO
                dbcmd.Parameters[9].Value = "INNOVA"; //USUARIO QUE MODIFICO
                dbcmd.Parameters[10].Value = 1; //ESTATUS DE REGISTRO
                dbcmd.Parameters[11].Value = convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value); //DISPONIBLE
                dbcmd.Parameters[12].Value = ROW.Cells["codigo de la unidad de medida"].Value; //CODIGO MEDIDA
                dbcmd.Parameters[13].Value = true;  //MIGRADO


                count += dbcmd.ExecuteNonQuery();

            }
            codInteno++;
        }

        private void callbackInsertServicioImpuestos(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t, int z)
        {
            if (ROW.Cells["codigo"].Value != null)
            {

                sql = @"INSERT INTO admin.gen_servicio_trib(org_hijo,cod_servicio,cod_impuesto, reg_estatus, 
                     migrado) VALUES(@orgHijo , @codServicio,
                    @codImpuesto, @regEstatus,  @migrado)";
                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
                dbcmd.Parameters.Add(new NpgsqlParameter("@orgHijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codServicio", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codImpuesto", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
                dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));


                dbcmd.Prepare();

                dbcmd.Parameters[0].Value = Globals.org;
                dbcmd.Parameters[1].Value = ROW.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[2].Value = ROW.Cells["cod impuesto1"].Value;
                dbcmd.Parameters[3].Value = convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                dbcmd.Parameters[4].Value = true;

                count += dbcmd.ExecuteNonQuery();
            }

        }
        private void callbackInsertImpuesto(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            int estatus = 0;
            if (ROW.Cells["codigo"].Value != null)
            {
                reader = "0";
                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
                sql = @"select count(*) from admin.gen_tributo";
                dbcmd = new NpgsqlCommand(sql, conn);

                reader = dbcmd.ExecuteScalar().ToString();
                if (reader != "0" && status == 0)
                {
                    result = MessageBox.Show("Esta tabla ya posee registros \n ¿desea eliminarlos?", "Atencion", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                    if (result == DialogResult.Yes)
                    {
                        try
                        {
                            sql = @"ALTER TABLE admin.gen_tributo DISABLE TRIGGER tg_01_ft_delete_registro; 
                                DELETE FROM admin.gen_tributo;
                                ALTER TABLE admin.gen_tributo ENABLE TRIGGER tg_01_ft_delete_registro;";
                            dbcmd = new NpgsqlCommand(sql, conn);
                            dbcmd.ExecuteNonQuery();
                            status = 1;
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("Los registro se encuentran relacionados y no pueden ser borrados");
                            estatus = 1;
                        }
                    }
                }
                if (estatus == 0)
                {
                    sql = @"INSERT INTO admin.gen_tributo(org_hijo,
                        cod_interno,codigo,descri,porcentaje,monto,reg_usu_cc,reg_usu_cu,
                        reg_estatus,disponible) 
                        VALUES(@org_hijo, @codInterno, @codigo, @descri,@porcentaje,@monto,
                        @reg_usu_cc, @reg_usu_cu, @regEstatus, @disponible)";


                    dbcmd = new NpgsqlCommand(sql, conn, t);

                    dbcmd.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@codInterno", NpgsqlDbType.Bigint));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@codigo", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@porcentaje", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@monto", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cu", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@disponible", NpgsqlDbType.Boolean));

                    dbcmd.Prepare();

                    dbcmd.Parameters[0].Value = Globals.org;
                    dbcmd.Parameters[1].Value = codInteno;
                    dbcmd.Parameters[2].Value = ROW.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);
                    dbcmd.Parameters[3].Value = ROW.Cells["descripcion"].Value;
                    if (ROW.Cells["es porcentaje"].Value.ToString() == "t")
                    {
                        dbcmd.Parameters[4].Value = ROW.Cells["valor"].Value;
                        dbcmd.Parameters[5].Value = 0;
                    }
                    else
                    {
                        dbcmd.Parameters[4].Value = 0;
                        dbcmd.Parameters[5].Value = ROW.Cells["valor"].Value;
                    }
                    dbcmd.Parameters[6].Value = "INNOVA";
                    dbcmd.Parameters[7].Value = "INNOVA";
                    dbcmd.Parameters[8].Value = 1;
                    dbcmd.Parameters[9].Value = convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);

                    count += dbcmd.ExecuteNonQuery();
                }
            }
        }
        private void callbackInsertArticuloPrecio(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t, int i)
        {
            if (ROW.Cells["codigo"].Value != null)
            {
                string reader = "0";
                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
                sql = @"select count(*) from admin.inv_art_precio where cod_articulo ='" + ROW.Cells["codigo"].Value.ToString() + "' AND cod_precio ='0" + i + "'";
                dbcmd = new NpgsqlCommand(sql, conn);

                reader = dbcmd.ExecuteScalar().ToString();
                if (reader == "0")
                {
                    sql = @"INSERT INTO admin.inv_art_precio(org_hijo,
                        codigo,cod_alterno,cod_articulo,cod_precio,
                        descri,descorta,precio,utilidad,tipo_utilidad,comision,tipo_comision,
                        descuento,tipo_descuento,reg_usu_cc,reg_usu_cu, reg_estatus, disponible, 
                        porc_utilidad,porc_comision, porc_descuento) 
                        VALUES(@org_hijo, @codigo, @codAlterno,  @codArticulo, @codPrecio, 
                        @descri, @descorta, @precio, @utilidad, @tipo_utilidad, @comision, @tipo_comision,
                        @descuento, @tipo_descuento, @reg_usu_cc, @reg_usu_cu,  @reg_estatus,  @disponible, 
                        @porc_utilidad, @porc_comision, @porc_descuento)";

                    //            //dbcmd.CommandText = sql;
                    dbcmd = new NpgsqlCommand(sql, conn, t);

                    dbcmd.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@codigo", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@codAlterno", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@codArticulo", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@codPrecio", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@descorta", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@precio", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@utilidad", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@tipo_utilidad", NpgsqlDbType.Boolean));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@comision", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@tipo_comision", NpgsqlDbType.Boolean));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@descuento", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@tipo_descuento", NpgsqlDbType.Boolean));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cu", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@reg_estatus", NpgsqlDbType.Integer));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@disponible", NpgsqlDbType.Boolean));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@porc_utilidad", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@porc_comision", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@porc_descuento", NpgsqlDbType.Double));

                    dbcmd.Prepare();

                    dbcmd.Parameters[0].Value = Globals.org;
                    dbcmd.Parameters[1].Value = "1";
                    dbcmd.Parameters[2].Value = ROW.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);
                    dbcmd.Parameters[3].Value = ROW.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);
                    dbcmd.Parameters[4].Value = "0" + i;
                    dbcmd.Parameters[5].Value = ROW.Cells["descripcion del producto"].Value;
                    dbcmd.Parameters[6].Value = ROW.Cells["descripcion del producto"].Value;
                    dbcmd.Parameters[7].Value = 0;
                    dbcmd.Parameters[8].Value = 0;
                    dbcmd.Parameters[9].Value = false;
                    dbcmd.Parameters[10].Value = 0;
                    dbcmd.Parameters[11].Value = false;
                    dbcmd.Parameters[12].Value = 0;
                    dbcmd.Parameters[13].Value = false;
                    dbcmd.Parameters[14].Value = "INNOVA";
                    dbcmd.Parameters[15].Value = "INNOVA";
                    dbcmd.Parameters[16].Value = 1;
                    dbcmd.Parameters[17].Value = convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                    dbcmd.Parameters[18].Value = arrayPU[i];
                    dbcmd.Parameters[19].Value = 0;
                    dbcmd.Parameters[20].Value = 0;


                    count += dbcmd.ExecuteNonQuery();
                }
            }
        }

        private void callbackInsertServicioPrecio(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t, int i)
        {
            if (ROW.Cells["codigo"].Value != null)
            {
                sql = @"INSERT INTO admin.gen_servicio_precio(org_hijo,
                        codigo,cod_precio,descri,descorta,precio,porc_utilidad,
                        tipo_utilidad,monto_descuento,tipo_descuento,reg_usu_cc,reg_usu_cu,
                        reg_estatus,porc_descuento, migrado) 
                        VALUES(@org_hijo, @codigo, @codPrecio,@descri, 
                        @descorta, @precio, @porc_utilidad,@tipo_utilidad,
                        @descuento, @tipo_descuento, @reg_usu_cc, @reg_usu_cu,  
                        @regEstatus, @porc_descuento, @migrado)";

                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);

                dbcmd.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codigo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codPrecio", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descorta", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@precio", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@porc_utilidad", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@tipo_utilidad", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descuento", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@tipo_descuento", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cu", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
                dbcmd.Parameters.Add(new NpgsqlParameter("@porc_descuento", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));

                dbcmd.Prepare();

                dbcmd.Parameters[0].Value = Globals.org;
                dbcmd.Parameters[1].Value = ROW.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[2].Value = "0" + i;
                dbcmd.Parameters[3].Value = ROW.Cells["descripcion del servicio"].Value;
                dbcmd.Parameters[4].Value = ROW.Cells["descripcion del servicio"].Value;
                dbcmd.Parameters[5].Value = 0;
                dbcmd.Parameters[6].Value = arrayPU[i];
                dbcmd.Parameters[7].Value = false;
                dbcmd.Parameters[8].Value = 0;
                dbcmd.Parameters[9].Value = false;
                dbcmd.Parameters[10].Value = "INNOVA";
                dbcmd.Parameters[11].Value = "INNOVA";
                dbcmd.Parameters[12].Value = 1;
                dbcmd.Parameters[13].Value = 0;
                dbcmd.Parameters[14].Value = true;


                count += dbcmd.ExecuteNonQuery();
            }
        }
        private void callbackInsertUsuario(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            if (ROW.Cells["cod_interno"].Value != null)
            {
                sql = @"INSERT INTO admin.cfg_usu(org_hijo,cod_interno,codigo,descri,descorta,
                    pwd,pregunta1,pregunta2,respuesta1,respuesta2,perfiles,cod_perfil,conectado, reg_usu_cc, reg_usu_cu,reg_estatus, 
                    disponible,migrado) VALUES(@orgHijo , @codInterno,
                    @codigo, @descri, @descorta , @pwd, @pregunta1, 
                    @pregunta2,@respuesta1,@respuesta2,@perfiles,@cod_perfil,@conectado, 
                    @reg_usu_cc , @reg_usu_cu, @regEstatus, @disponible, @migrado)";
                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
                dbcmd.Parameters.Add(new NpgsqlParameter("@orghijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codInterno", NpgsqlDbType.Bigint));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codigo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descorta", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@pwd", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@pregunta1", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@pregunta2", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@respuesta1", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@respuesta2", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@perfiles", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@cod_perfil", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@conectado", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cu", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
                dbcmd.Parameters.Add(new NpgsqlParameter("@disponible", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));

                dbcmd.Prepare();

                dbcmd.Parameters[0].Value = ROW.Cells["org_hijo"].Value;
                dbcmd.Parameters[1].Value = ROW.Cells["cod_interno"].Value;
                dbcmd.Parameters[2].Value = ROW.Cells["codigo"].Value;
                dbcmd.Parameters[3].Value = ROW.Cells["descri"].Value;
                dbcmd.Parameters[4].Value = ROW.Cells["descorta"].Value;
                dbcmd.Parameters[5].Value = ROW.Cells["pwd"].Value;
                dbcmd.Parameters[6].Value = ROW.Cells["pregunta1"].Value;
                dbcmd.Parameters[7].Value = ROW.Cells["pregunta2"].Value;
                dbcmd.Parameters[8].Value = ROW.Cells["respuesta1"].Value;
                dbcmd.Parameters[9].Value = ROW.Cells["respuesta2"].Value;
                dbcmd.Parameters[10].Value = convertBoolean(ROW.Cells["perfiles"].Value);
                dbcmd.Parameters[11].Value = ROW.Cells["cod_perfil"].Value;
                dbcmd.Parameters[12].Value = convertBoolean(ROW.Cells["conectado"].Value);
                dbcmd.Parameters[13].Value = "INNOVA";
                dbcmd.Parameters[14].Value = "INNOVA";
                dbcmd.Parameters[15].Value = ROW.Cells["reg_estatus"].Value;
                dbcmd.Parameters[16].Value = convertBoolean(ROW.Cells["disponible"].Value);
                dbcmd.Parameters[17].Value = true;

                count += dbcmd.ExecuteNonQuery();
            }
        }
        private void callbackInsertProveedores(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            if (ROW.Cells["código"].Value != null && ROW.Cells["código"].Value.ToString() != "")
            {
                sql = @"INSERT INTO admin.com_prov(org_hijo,cod_interno,prov_hijo,descri,
                        tipo_cont,tipo_pers,rif,direc1,descuento,es_descuento,
                        es_exento,es_retencion,es_monto,monto_min,monto_max,monto_cred,
                        pri_monto,ult_monto,rect_monto,pago_max,pago_ade,
                        pago_prom,saldo, reg_usu_cc,reg_usu_cu,reg_estatus,disponible, migrado,
                        porc_ret_iva, observacion, tipo_ret_iva) 
                        VALUES(@org_hijo,@cod_interno,@prov_hijo,@descri,
                        @tipo_cont,@tipo_pers,@rif,@direc1,@descuento,@es_descuento,
                        @es_exento,@es_retencion,@es_monto,@monto_min,@monto_max,@monto_cred,
                        @pri_monto,@ult_monto,@rect_monto,@pago_max,@pago_ade,
                        @pago_prom,@saldo, @reg_usu_cc,@reg_usu_cu,@reg_estatus,@disponible, 
                        @migrado, @porcretiva, @Observacion, @tipo_ret_iva)";


                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
                dbcmd.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@cod_interno", NpgsqlDbType.Bigint));
                dbcmd.Parameters.Add(new NpgsqlParameter("@prov_hijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@tipo_cont", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@tipo_pers", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@rif", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@direc1", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descuento", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@es_descuento", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@es_exento", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@es_retencion", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@es_monto", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@monto_min", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@monto_max", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@monto_cred", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@pri_monto", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@ult_monto", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@rect_monto", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@pago_max", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@pago_ade", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@pago_prom", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@saldo", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cu", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_estatus", NpgsqlDbType.Integer));
                dbcmd.Parameters.Add(new NpgsqlParameter("@disponible", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@porcretiva", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@Observacion", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@tipo_ret_iva", NpgsqlDbType.Varchar));


                dbcmd.Prepare();

                dbcmd.Parameters[0].Value = Globals.org;
                dbcmd.Parameters[1].Value = codInteno;
                dbcmd.Parameters[2].Value = ROW.Cells["código"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[3].Value = ROW.Cells["razon social"].Value;
                foreach (var z in listaTipos.Where(a => a.tipo == ROW.Cells["tipo de contribuyente"].Value.ToString().Replace(" ", string.Empty)))
                {
                    dbcmd.Parameters[4].Value = z.codigo;
                }

                foreach (var z in listaTipos.Where(a => a.tipo == ROW.Cells["tipo de persona"].Value.ToString().Replace(" ", string.Empty)))
                {
                    dbcmd.Parameters[5].Value = z.codigo;
                }
                dbcmd.Parameters[6].Value = ROW.Cells["rif"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[7].Value = ROW.Cells["dirección"].Value;
                dbcmd.Parameters[8].Value = ROW.Cells["descuento"].Value;
                dbcmd.Parameters[9].Value = false;
                dbcmd.Parameters[10].Value = false;
                dbcmd.Parameters[11].Value = false;
                dbcmd.Parameters[12].Value = false;
                dbcmd.Parameters[13].Value = ROW.Cells["Monto mínimo de una compra"].Value;
                dbcmd.Parameters[14].Value = ROW.Cells["Monto máximo de una compra"].Value;
                dbcmd.Parameters[15].Value = ROW.Cells["Monto de crédito"].Value;
                dbcmd.Parameters[16].Value = ROW.Cells["Monto primera compra"].Value;
                dbcmd.Parameters[17].Value = ROW.Cells["Monto última compra"].Value;
                dbcmd.Parameters[18].Value = ROW.Cells["Monto último pago"].Value;
                dbcmd.Parameters[19].Value = ROW.Cells["Monto máximo de pago recibido"].Value;
                dbcmd.Parameters[20].Value = ROW.Cells["Monto adelanto"].Value;
                dbcmd.Parameters[21].Value = ROW.Cells["Monto promedio de pagos recbidos"].Value;
                dbcmd.Parameters[22].Value = ROW.Cells["Saldo"].Value;
                dbcmd.Parameters[23].Value = "INNOVA";
                dbcmd.Parameters[24].Value = "INNOVA";
                dbcmd.Parameters[25].Value = 1;
                dbcmd.Parameters[26].Value = convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                dbcmd.Parameters[27].Value = true;
                dbcmd.Parameters[28].Value = 75;
                dbcmd.Parameters[29].Value = "ESTA DATA FUE MIGRADA, POR FAVOR VERIFICAR LOS DATOS";
                dbcmd.Parameters[30].Value = "15.2";

                count += dbcmd.ExecuteNonQuery();
            }

        }

        private void callbackInsertArticuloImpuestos(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t, int z)
        {
            if (ROW.Cells["codigo"].Value != null)
            {
                string reader = "0";
                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
                sql = @"select count(*) from admin.inv_art_imp where cod_articulo ='" + ROW.Cells["codigo"].Value.ToString() + "' AND cod_impuesto ='" + ROW.Cells["cod impuesto" + z + ""].Value.ToString() + "'";
                dbcmd = new NpgsqlCommand(sql, conn);

                reader = dbcmd.ExecuteScalar().ToString();
                nro_items++;
                if (reader == "0")
                {

                    sql = @"INSERT INTO admin.inv_art_imp(org_hijo,cod_articulo,cod_impuesto, reg_usu_cc, reg_usu_cu,reg_estatus, 
                    disponible, migrado) VALUES(@orgHijo , @codArticulo,
                    @codImpuesto, @reg_usu_cc , @reg_usu_cu, @regEstatus, @disponible, @migrado)";
                    dbcmd = new NpgsqlCommand(sql, conn, t);
                    dbcmd.Parameters.Add(new NpgsqlParameter("@orgHijo", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@codArticulo", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@codImpuesto", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cu", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@disponible", NpgsqlDbType.Boolean));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));


                    dbcmd.Prepare();

                    dbcmd.Parameters[0].Value = Globals.org;
                    dbcmd.Parameters[1].Value = ROW.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);
                    dbcmd.Parameters[2].Value = ROW.Cells["cod impuesto" + z + ""].Value;
                    dbcmd.Parameters[3].Value = "INNOVA";
                    dbcmd.Parameters[4].Value = "INNOVA";
                    dbcmd.Parameters[5].Value = 1;
                    dbcmd.Parameters[6].Value = convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                    dbcmd.Parameters[7].Value = true;

                    count += dbcmd.ExecuteNonQuery();
                }
            }

        }

        private void callbackInsertAjustePrecio(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            if (ROW.Cells["doc"].Value != null)
            {
                sql = @"INSERT INTO admin.int_ajuste_precio(org_hijo,doc,cod_motivo,descri,cod_autoriza,
                    nomb_autoriza,cod_persona, nomb_persona,cantidad_item,total_precio,total_utilidad, 
                    reg_usu_cc, reg_estatus,nro_items, doc_control, migrado) VALUES(@org_hijo,@doc,@cod_motivo,@descri,@cod_autoriza,
                    @nomb_autoriza,@cod_persona, @nomb_persona,@cantidad_item,@total_precio,@total_utilidad, 
                    @reg_usu_cc, @reg_estatus, @nro_items, @doc_control, @migrado)";
                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
                dbcmd.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@doc", NpgsqlDbType.Integer));
                dbcmd.Parameters.Add(new NpgsqlParameter("@cod_motivo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@cod_autoriza", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@nomb_autoriza", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@cod_persona", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@nomb_persona", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@cantidad_item", NpgsqlDbType.Integer));
                dbcmd.Parameters.Add(new NpgsqlParameter("@total_precio", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@total_utilidad", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_estatus", NpgsqlDbType.Integer));
                dbcmd.Parameters.Add(new NpgsqlParameter("@nro_items", NpgsqlDbType.Integer));
                dbcmd.Parameters.Add(new NpgsqlParameter("@doc_control", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));

                dbcmd.Prepare();

                dbcmd.Parameters[0].Value = ROW.Cells["org_hijo"].Value;
                dbcmd.Parameters[1].Value = ROW.Cells["doc"].Value;
                dbcmd.Parameters[2].Value = ROW.Cells["cod_motivo"].Value;
                dbcmd.Parameters[3].Value = ROW.Cells["descri"].Value;
                dbcmd.Parameters[4].Value = ROW.Cells["cod_autoriza"].Value;
                dbcmd.Parameters[5].Value = ROW.Cells["nomb_autoriza"].Value;
                dbcmd.Parameters[6].Value = ROW.Cells["cod_persona"].Value;
                dbcmd.Parameters[7].Value = ROW.Cells["nomb_persona"].Value;
                dbcmd.Parameters[8].Value = ROW.Cells["cantidad_item"].Value;
                dbcmd.Parameters[9].Value = ROW.Cells["total_precio"].Value;
                dbcmd.Parameters[10].Value = ROW.Cells["total_utilidad"].Value;
                dbcmd.Parameters[11].Value = "INNOVA";
                dbcmd.Parameters[12].Value = ROW.Cells["reg_estatus"].Value;
                dbcmd.Parameters[13].Value = ROW.Cells["nro_items"].Value;
                dbcmd.Parameters[14].Value = ROW.Cells["doc_control"].Value;
                dbcmd.Parameters[15].Value = true;

                count += dbcmd.ExecuteNonQuery();
            }
        }

        private void callbackInsertCargoInventario(NpgsqlConnection conn, NpgsqlTransaction t)
        {


            //insercion de ajuste para el cargo
            sql = @"INSERT INTO admin.int_ajuste_precio(org_hijo,descri,cantidad_item,total_precio,total_utilidad, 
                    reg_usu_cc, reg_estatus, nro_items,migrado) VALUES(@org_hijo,@descri,@cantidad_item,@total_precio,@total_utilidad, 
                    @reg_usu_cc, @reg_estatus, @nro_items, @migrado)";
            NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn);
            dbcmd.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@cantidad_item", NpgsqlDbType.Integer));
            dbcmd.Parameters.Add(new NpgsqlParameter("@total_precio", NpgsqlDbType.Double));
            dbcmd.Parameters.Add(new NpgsqlParameter("@total_utilidad", NpgsqlDbType.Double));
            dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@reg_estatus", NpgsqlDbType.Integer));
            dbcmd.Parameters.Add(new NpgsqlParameter("@nro_items", NpgsqlDbType.Integer));
            dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));

            dbcmd.Prepare();

            dbcmd.Parameters[0].Value = Globals.org;
            dbcmd.Parameters[1].Value = "CARGA INICIAL DE INVENTARIO";
            dbcmd.Parameters[2].Value = cantidad_items;
            dbcmd.Parameters[3].Value = total;
            dbcmd.Parameters[4].Value = "0";
            dbcmd.Parameters[5].Value = "INNOVA";
            dbcmd.Parameters[6].Value = 1;
            dbcmd.Parameters[7].Value = nro_items;
            dbcmd.Parameters[8].Value = true;

            count += dbcmd.ExecuteNonQuery();

            sql = @"SELECT doc from admin.int_ajuste_precio order by fecha_reg desc";
            dbcmd = new NpgsqlCommand(sql, conn);

            string reader = dbcmd.ExecuteScalar().ToString();
            //t = conn.BeginTransaction();
            //Insercion del detalle del ajuste
            foreach (DataGridViewRow ROW2 in dataGridView1.Rows)
            {
                if (ROW2.Cells["codigo"].Value != null)
                {
                    item++;
                    sql = @"INSERT INTO admin.int_ajuste_precio_det(org_hijo,doc,cod_alterno,cod_articulo,
                        costo,costo_promedio,fecha,tipo_ajuste,item, migrado) VALUES(@org_hijo,@doc,
                        @cod_alterno,@cod_articulo,@costo,@costo_promedio,@fecha,@tipo_ajuste,@item,@migrado)";
                    dbcmd = new NpgsqlCommand(sql, conn);
                    dbcmd.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@doc", NpgsqlDbType.Bigint));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@cod_alterno", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@cod_articulo", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@costo", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@costo_promedio", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@fecha", NpgsqlDbType.Date));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@tipo_ajuste", NpgsqlDbType.Integer));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@item", NpgsqlDbType.Integer));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));

                    dbcmd.Prepare();

                    dbcmd.Parameters[0].Value = Globals.org;
                    dbcmd.Parameters[1].Value = Convert.ToInt64(reader);
                    dbcmd.Parameters[2].Value = ROW2.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);
                    dbcmd.Parameters[3].Value = ROW2.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);
                    dbcmd.Parameters[4].Value = ROW2.Cells["costo"].Value;
                    dbcmd.Parameters[5].Value = ROW2.Cells["costo promedio"].Value;
                    dbcmd.Parameters[6].Value = DateTime.Now;
                    dbcmd.Parameters[7].Value = 1;
                    dbcmd.Parameters[8].Value = item;
                    dbcmd.Parameters[9].Value = true;

                    count += dbcmd.ExecuteNonQuery();
                }
            }
            item = 0;
            //insercion del cargo
            sql = @"INSERT INTO admin.int_cargo(org_hijo,cod_terminal,tipo_opera,
                        descri,fecha,cod_motivo, motivo, total, reg_usu_cc, reg_usu_cu, reg_estatus,factor,
                         nro_items, cod_ajuste_precio, migrado) VALUES(@org_hijo,
                        @cod_terminal,@tipo_opera,@descri,@fecha,@cod_motivo, @motivo, @total, @reg_usu_cc, @reg_usu_cu, 
                        @reg_estatus, @factor,@nro_items, @cod_ajuste_precio, @migrado)";
            dbcmd = new NpgsqlCommand(sql, conn);
            dbcmd.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@cod_terminal", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@tipo_opera", NpgsqlDbType.Integer));
            dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@fecha", NpgsqlDbType.Date));
            dbcmd.Parameters.Add(new NpgsqlParameter("@cod_motivo", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@motivo", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@total", NpgsqlDbType.Double));
            dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cu", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@reg_estatus", NpgsqlDbType.Integer));
            dbcmd.Parameters.Add(new NpgsqlParameter("@factor", NpgsqlDbType.Double));
            dbcmd.Parameters.Add(new NpgsqlParameter("@nro_items", NpgsqlDbType.Integer));
            dbcmd.Parameters.Add(new NpgsqlParameter("@cod_ajuste_precio", NpgsqlDbType.Bigint));
            dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));

            dbcmd.Prepare();

            dbcmd.Parameters[0].Value = Globals.org;
            dbcmd.Parameters[1].Value = " ";
            dbcmd.Parameters[2].Value = "28";
            dbcmd.Parameters[3].Value = "CARGA INICIAL DE INVENTARIO";
            dbcmd.Parameters[4].Value = DateTime.Now;
            dbcmd.Parameters[5].Value = "13.4";
            dbcmd.Parameters[6].Value = "CARGA INICIAL";
            dbcmd.Parameters[7].Value = total;
            dbcmd.Parameters[8].Value = "INNOVA";
            dbcmd.Parameters[9].Value = "INNOVA";
            dbcmd.Parameters[10].Value = 1;
            dbcmd.Parameters[11].Value = 0;
            dbcmd.Parameters[12].Value = nro_items;
            dbcmd.Parameters[13].Value = reader;
            dbcmd.Parameters[14].Value = true;

            count += dbcmd.ExecuteNonQuery();
            sql = @"SELECT doc from admin.int_cargo order by fecha_reg desc";
            dbcmd = new NpgsqlCommand(sql, conn);

            reader = dbcmd.ExecuteScalar().ToString();
            //t = conn.BeginTransaction();
            foreach (DataGridViewRow ROW2 in dataGridView1.Rows)
            {
                if (ROW2.Cells["codigo"].Value != null)
                {
                    item++;
                    sql = @"INSERT INTO admin.int_cargo_det(org_hijo, doc, item, cod_alterno,
                                    cod_articulo,descri, cantidad, existencia, existencia_anterior,
                                    costo_anterior, costo_promedio_ant, costo_promedio, cod_dep, costo, total, 
                                    precio_utilidad, descorta, tipo_opera, reg_estatus, tipo_ajuste,
                                    tipo_documento,migrado) VALUES(@org_hijo, @doc, @item, @cod_alterno,
                                    @cod_articulo, @descri, @cantidad, @existencia, @existencia_anterior,
                                    @costo_anterior, @costo_promedio_ant, @costo_promedio, @cod_dep, @costo, @total, 
                                    @precio_utilidad, @descorta, @tipo_opera, @reg_estatus, @tipo_ajuste,
                                    @tipo_documento,@migrado)";
                    dbcmd = new NpgsqlCommand(sql, conn);
                    dbcmd.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@doc", NpgsqlDbType.Bigint));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@item", NpgsqlDbType.Integer));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@cod_alterno", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@cod_articulo", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@cantidad", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@existencia", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@existencia_anterior", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@costo_anterior", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@costo_promedio_ant", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@costo_promedio", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@cod_dep", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@costo", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@total", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@precio_utilidad", NpgsqlDbType.Boolean));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@descorta", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@tipo_opera", NpgsqlDbType.Integer));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@reg_estatus", NpgsqlDbType.Integer));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@tipo_ajuste", NpgsqlDbType.Integer));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@tipo_documento", NpgsqlDbType.Integer));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));

                    dbcmd.Prepare();

                    dbcmd.Parameters[0].Value = Globals.org;
                    dbcmd.Parameters[1].Value = Convert.ToInt64(reader);
                    dbcmd.Parameters[2].Value = item;
                    dbcmd.Parameters[3].Value = ROW2.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);
                    dbcmd.Parameters[4].Value = ROW2.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);
                    dbcmd.Parameters[5].Value = ROW2.Cells["descripcion del producto"].Value;
                    dbcmd.Parameters[6].Value = ROW2.Cells["existencia"].Value;
                    dbcmd.Parameters[7].Value = ROW2.Cells["existencia"].Value;
                    dbcmd.Parameters[8].Value = 0;
                    dbcmd.Parameters[9].Value = 0;
                    dbcmd.Parameters[10].Value = 0;
                    dbcmd.Parameters[11].Value = 0;
                    dbcmd.Parameters[12].Value = ROW2.Cells["codigo deposito"].Value;
                    dbcmd.Parameters[13].Value = ROW2.Cells["costo"].Value;
                    dbcmd.Parameters[14].Value = ((double)ROW2.Cells["costo"].Value * (double)ROW2.Cells["existencia"].Value);
                    dbcmd.Parameters[15].Value = false;
                    dbcmd.Parameters[16].Value = ROW2.Cells["descripcion del producto"].Value;
                    dbcmd.Parameters[17].Value = 28;
                    dbcmd.Parameters[18].Value = 1;
                    dbcmd.Parameters[19].Value = 1;
                    dbcmd.Parameters[20].Value = 10;
                    dbcmd.Parameters[21].Value = true;

                    count += dbcmd.ExecuteNonQuery();
                }
            }
        }

        private void callbackInsertCxC(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            sql = @"INSERT INTO admin.fin_cxc(org_hijo,doc_num,cod_cli,fecha_emi,fecha_ven,factor,
                      total,saldo,saldo_inicial,total_exento,doc_control,reg_usu_cc,
                      reg_estatus,cod_empleado , migrado, cod_impresorafiscal, descri, tipo_opera) VALUES(@org_hijo,@docNum,@codCli,
                      @fechaEmi, @fechaVen, @factor, @total, @saldo, @saldoInicial, @totalEx, 
                      @doc_control,@reg_usu_cc, @reg_estatus, @codEmpleado, @migrado, @cod_impresorafiscal, @descri, @tipoOpera)";
            NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
            dbcmd.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@docNum", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@codCli", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@fechaEmi", NpgsqlDbType.Date));
            dbcmd.Parameters.Add(new NpgsqlParameter("@fechaVen", NpgsqlDbType.Date));
            dbcmd.Parameters.Add(new NpgsqlParameter("@factor", NpgsqlDbType.Double));
            dbcmd.Parameters.Add(new NpgsqlParameter("@total", NpgsqlDbType.Double));
            dbcmd.Parameters.Add(new NpgsqlParameter("@saldo", NpgsqlDbType.Double));
            dbcmd.Parameters.Add(new NpgsqlParameter("@saldoInicial", NpgsqlDbType.Double));
            dbcmd.Parameters.Add(new NpgsqlParameter("@totalEx", NpgsqlDbType.Double));
            dbcmd.Parameters.Add(new NpgsqlParameter("@doc_control", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@reg_estatus", NpgsqlDbType.Integer));
            dbcmd.Parameters.Add(new NpgsqlParameter("@codEmpleado", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));
            dbcmd.Parameters.Add(new NpgsqlParameter("@cod_impresorafiscal", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@tipoOpera", NpgsqlDbType.Integer));
            dbcmd.Prepare();

            dbcmd.Parameters[0].Value = Globals.org;
            dbcmd.Parameters[1].Value = ROW.Cells["numero factura"].Value;
            dbcmd.Parameters[2].Value = ROW.Cells["cod cliente"].Value;
            dbcmd.Parameters[3].Value = ExtractDate(ROW.Cells["fecha emision"].Value.ToString());
            dbcmd.Parameters[4].Value = ExtractDate(ROW.Cells["fecha vencimiento"].Value.ToString());
            dbcmd.Parameters[5].Value = 0;
            dbcmd.Parameters[6].Value = ROW.Cells["monto total"].Value;
            dbcmd.Parameters[7].Value = ROW.Cells["saldo"].Value;
            dbcmd.Parameters[8].Value = ROW.Cells["saldo inicial"].Value;
            dbcmd.Parameters[9].Value = ROW.Cells["monto exento"].Value;
            dbcmd.Parameters[10].Value = ROW.Cells["numero de control"].Value;
            dbcmd.Parameters[11].Value = "INNOVA";
            dbcmd.Parameters[12].Value = 1;
            dbcmd.Parameters[13].Value = ROW.Cells["cod vendedor"].Value;
            dbcmd.Parameters[14].Value = true;
            dbcmd.Parameters[15].Value = ROW.Cells["numero impresora fiscal"].Value;
            dbcmd.Parameters[16].Value = ROW.Cells["descripcion"].Value;
            foreach (var z in listaTipos.Where(a => a.tipo == ROW.Cells["tipo operacion"].Value.ToString()))
            {
                dbcmd.Parameters[17].Value = z.codigo;  //TIPO OPERACION
            }

            count += dbcmd.ExecuteNonQuery();


        }

        private void callbackInsertCxCImp(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t, int z)
        {

            string reader = "0";
            sql = @"SELECT MAX(doc) FROM admin.fin_cxc";
            NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn);

            reader = dbcmd.ExecuteScalar().ToString();

            sql = @"INSERT INTO admin.fin_cxc_imp(org_hijo,porcentaje,cod_impuesto,base, total, doc,reg_estatus, 
                     migrado) VALUES(@orgHijo , @porcentaje,
                    @codImpuesto, @base , @total,@doc, @regEstatus, @migrado)";
            dbcmd = new NpgsqlCommand(sql, conn, t);
            dbcmd.Parameters.Add(new NpgsqlParameter("@orgHijo", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@porcentaje", NpgsqlDbType.Double));
            dbcmd.Parameters.Add(new NpgsqlParameter("@codImpuesto", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@base", NpgsqlDbType.Double));
            dbcmd.Parameters.Add(new NpgsqlParameter("@total", NpgsqlDbType.Double));
            dbcmd.Parameters.Add(new NpgsqlParameter("@doc", NpgsqlDbType.Bigint));
            dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
            dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));


            dbcmd.Prepare();

            dbcmd.Parameters[0].Value = Globals.org;
            dbcmd.Parameters[1].Value = ROW.Cells["porc impuesto" + z + ""].Value;
            dbcmd.Parameters[2].Value = ROW.Cells["cod impuesto" + z + ""].Value;
            dbcmd.Parameters[3].Value = ROW.Cells["base imponible" + z + ""].Value;
            dbcmd.Parameters[4].Value = ROW.Cells["monto total"].Value;
            dbcmd.Parameters[5].Value = reader;
            dbcmd.Parameters[6].Value = 1;
            dbcmd.Parameters[7].Value = true;

            count += dbcmd.ExecuteNonQuery();

        }

        private void callbackInsertCxP(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            sql = @"INSERT INTO admin.fin_cxp(org_hijo,doc_num,cod_prov,fecha_emi,fecha_ven,factor,
                      total,saldo,saldo_inicial,total_exento,doc_control,reg_usu_cc,
                      reg_estatus, migrado, cod_impresorafiscal, descri, tipo_opera) VALUES(@org_hijo,@docNum,@codPro,
                      @fechaEmi, @fechaVen, @factor, @total, @saldo, @saldoInicial, @totalEx, 
                      @doc_control,@reg_usu_cc, @reg_estatus, @migrado, @cod_impresorafiscal, @descri, @tipoOpera)";
            NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
            dbcmd.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@docNum", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@codPro", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@fechaEmi", NpgsqlDbType.Date));
            dbcmd.Parameters.Add(new NpgsqlParameter("@fechaVen", NpgsqlDbType.Date));
            dbcmd.Parameters.Add(new NpgsqlParameter("@factor", NpgsqlDbType.Double));
            dbcmd.Parameters.Add(new NpgsqlParameter("@total", NpgsqlDbType.Double));
            dbcmd.Parameters.Add(new NpgsqlParameter("@saldo", NpgsqlDbType.Double));
            dbcmd.Parameters.Add(new NpgsqlParameter("@saldoInicial", NpgsqlDbType.Double));
            dbcmd.Parameters.Add(new NpgsqlParameter("@totalEx", NpgsqlDbType.Double));
            dbcmd.Parameters.Add(new NpgsqlParameter("@doc_control", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@reg_estatus", NpgsqlDbType.Integer));
            dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));
            dbcmd.Parameters.Add(new NpgsqlParameter("@cod_impresorafiscal", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@tipoOpera", NpgsqlDbType.Integer));
            dbcmd.Prepare();

            dbcmd.Parameters[0].Value = Globals.org;
            dbcmd.Parameters[1].Value = ROW.Cells["numero factura"].Value;
            dbcmd.Parameters[2].Value = ROW.Cells["cod proveedor"].Value;
            dbcmd.Parameters[3].Value = ExtractDate(ROW.Cells["fecha emision"].Value.ToString());
            dbcmd.Parameters[4].Value = ExtractDate(ROW.Cells["fecha vencimiento"].Value.ToString());
            dbcmd.Parameters[5].Value = 0;
            dbcmd.Parameters[6].Value = ROW.Cells["monto total"].Value;
            dbcmd.Parameters[7].Value = ROW.Cells["saldo"].Value;
            dbcmd.Parameters[8].Value = ROW.Cells["saldo inicial"].Value;
            dbcmd.Parameters[9].Value = ROW.Cells["monto exento"].Value;
            dbcmd.Parameters[10].Value = ROW.Cells["numero de control"].Value;
            dbcmd.Parameters[11].Value = "INNOVA";
            dbcmd.Parameters[12].Value = 1;
            dbcmd.Parameters[13].Value = true;
            dbcmd.Parameters[14].Value = ROW.Cells["numero impresora fiscal"].Value;
            dbcmd.Parameters[15].Value = ROW.Cells["descripcion"].Value;
            foreach (var z in listaTipos.Where(a => a.tipo == ROW.Cells["tipo operacion"].Value.ToString()))
            {
                dbcmd.Parameters[16].Value = z.codigo;  //TIPO OPERACION
            }

            count += dbcmd.ExecuteNonQuery();
        }


        private void callbackInsertCxPImp(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t, int z)
        {
            string reader = "0";
            sql = @"SELECT MAX(doc) FROM admin.fin_cxp";
            NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn);

            reader = dbcmd.ExecuteScalar().ToString();

            sql = @"INSERT INTO admin.fin_cxp_imp(org_hijo,porcentaje,cod_impuesto,base, total, doc,reg_estatus, 
                     migrado) VALUES(@orgHijo , @porcentaje,
                    @codImpuesto, @base , @total,@doc, @regEstatus, @migrado)";
            dbcmd = new NpgsqlCommand(sql, conn, t);
            dbcmd.Parameters.Add(new NpgsqlParameter("@orgHijo", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@porcentaje", NpgsqlDbType.Double));
            dbcmd.Parameters.Add(new NpgsqlParameter("@codImpuesto", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@base", NpgsqlDbType.Double));
            dbcmd.Parameters.Add(new NpgsqlParameter("@total", NpgsqlDbType.Double));
            dbcmd.Parameters.Add(new NpgsqlParameter("@doc", NpgsqlDbType.Bigint));
            dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
            dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));


            dbcmd.Prepare();

            dbcmd.Parameters[0].Value = Globals.org;
            dbcmd.Parameters[1].Value = ROW.Cells["porc impuesto" + z + ""].Value;
            dbcmd.Parameters[2].Value = ROW.Cells["cod impuesto" + z + ""].Value;
            dbcmd.Parameters[3].Value = ROW.Cells["base imponible" + z + ""].Value;
            dbcmd.Parameters[4].Value = ROW.Cells["monto total"].Value;
            dbcmd.Parameters[5].Value = reader;
            dbcmd.Parameters[6].Value = 1;
            dbcmd.Parameters[7].Value = true;

            count += dbcmd.ExecuteNonQuery();
        }






        public static DateTime? ExtractDate(string myDate)
        {
            if (!string.IsNullOrEmpty(myDate) && !string.IsNullOrWhiteSpace(myDate))
            {
                DateTime dt;
                var formatStrings = new string[] { "MM/dd/yyyy hh:mm:ss", "MM/d/yyyy" };
                dt = DateTime.ParseExact(myDate, formatStrings, new CultureInfo("en-US"), DateTimeStyles.None);
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
        #endregion

        private void Main_Load(object sender, EventArgs e)
        {

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                using (var db = new LiteDatabase("Colleccion.db"))
                {
                    var col = db.GetCollection<payroll>("PayRoll");
                    var Z = col.Find(Query.All());
                    comboBox1.DataSource = Z.ToList();
                    comboBox1.DisplayMember = "desc";
                    comboBox1.ValueMember = "id";
                }
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                using (var db = new LiteDatabase("Colleccion.db"))
                {
                    var col = db.GetCollection<admin>("Admin");
                    var Z = col.Find(Query.All());
                    comboBox1.DataSource = Z.ToList();
                    comboBox1.DisplayMember = "desc";
                    comboBox1.ValueMember = "id";
                }
            }
        }
    }
    public class Tipos
    {
        public String tipo { get; set; }
        public String codigo { get; set; }
        public Tipos()
        {

        }
        public Tipos(string tipos, string codigos)
        {
            this.tipo = tipos;
            this.codigo = codigos;
        }
    }
    public class admin
    {
        public String desc { get; set; }
        public int Id { get; set; }

        public admin()
        {

        }

        public admin(string descri, int cod)
        {
            this.desc = descri;
            this.Id = cod;
        }

    }
    public class payroll
    {
        public String desc { get; set; }
        public int id { get; set; }
        public payroll()
        {

        }
        public payroll(string descri, int cod)
        {
            this.desc = descri;
            this.id = cod;
        }

    }
    public class Errores
    {
        public String Desc { get; set; }
        public String codigo { get; set; }

        public Errores(string Descs, string codigos)
        {
            this.Desc = Descs;
            this.codigo = codigos;
        }
    }
}


