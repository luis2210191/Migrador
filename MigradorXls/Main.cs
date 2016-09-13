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
        #region Declaraciones
        string file;
        string sql;
        string reader;
        int count;
        int codInteno = 0;
        double total;
        int cantidad_items = 0;
        int nro_items = 0;
        int item = 0;
        int status = 0;
        DialogResult result;
        private DataSet DtSet = new DataSet();
        System.Data.OleDb.OleDbConnection Myconnetion;
        System.Data.OleDb.OleDbDataAdapter MyCommand;
        string connectionString;
        Dictionary<string, DateTime?> adelantos = new Dictionary<string, DateTime?>();
        List<Tipos> listaTipos = new List<Tipos>();
        List<Errores> listaErr = new List<Errores>();
        List<admin> listaAdmin = new List<admin>();
        List<payroll> listaPayRoll = new List<payroll>();
        DataConvert dt = new DataConvert();
        #endregion

        #region Main
        public Main()
        {

            InitializeComponent();
            //Asignacion de data source 
            sql = @"SELECT org_hijo from admin.cfg_org";
            //Abriendo la coneccion con npgsql
            connectionString = @"Host=" + Globals.Host + ";port=" + Globals.port + ";Database=" + Globals.DB + ";User ID=" + Globals.usuario + ";Password=" + Globals.pass + ";";
            NpgsqlConnection conn = new NpgsqlConnection(connectionString);
            NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn);
            conn.Open();
            Globals.org = dbcmd.ExecuteScalar().ToString();
            conn.Close();  
            //Consulta para cargar informacion de las tablas en combobox de tablas         
            var db = DBConn.Instance;
            var c = db.Collection<admin>();
            comboBox1.DataSource = c.Find(Query.All()).ToList();
            comboBox1.DisplayMember = "desc";
            comboBox1.ValueMember = "id";            
        }
        #endregion

        #region Metodos Migracion Admin

        /// <summary>
        /// Array de los porcentajes de utilidad
        /// </summary>
        public double[] arrayPU;

        #region Servicios
        //Metodo que inicia la migracion de los servicios
        /// <summary>
        /// Metodo que inicia el proceso para la migracion de  servicios
        /// </summary>
        private void Exportar_Servicios(NpgsqlConnection conn, NpgsqlTransaction t)
        {

            //Recorriendo el Datagridview e insertando cada valor

            foreach (DataGridViewRow ROW in dataGridView1.Rows)
            {
                if (ROW.Cells["codigo"].Value != null)
                {
                    arrayPU = new double[]{
                    30,
                    porcUtilidad(Convert.ToDouble(ROW.Cells["%utilidad1"].Value)),
                    porcUtilidad(Convert.ToDouble(ROW.Cells["%utilidad2"].Value)),
                    porcUtilidad(Convert.ToDouble(ROW.Cells["%utilidad3"].Value)),
                    porcUtilidad(Convert.ToDouble(ROW.Cells["%utilidad4"].Value)),

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
                        codInteno++;
                    }
                    catch (NpgsqlException ex)
                    {

                        try
                        {
                            var db = DBConn.Instance;
                            var col = db.Collection<Errores>();
                            ROW.Cells["Error"].Value = col.Find(x => x.codigo == ex.Code.ToString()).FirstOrDefault().Desc;
                            ROW.DefaultCellStyle.BackColor = Color.Red;
                            count = 0;
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("Hubo un Error en la insercion de datos. Excepcion: " + ex.Message.ToString());
                        }
                        break;
                    }
                }

            }

        }
        #endregion

        #region Articulos
        //Metodo que inicia la migracion de Articulos
        /// <summary>
        /// Metodo que inicia el proceso para la migracion de articulos
        /// </summary>
        private void Exportar_Articulos(NpgsqlConnection conn, NpgsqlTransaction t)
        {
            cantidad_items = 0;
            nro_items = 0;
            item = 0;
            bool FAIL = false;

            //Recorriendo el Datagridview e insertando cada valor

            foreach (DataGridViewRow ROW in dataGridView1.Rows)
            {
                if (ROW.Cells["codigo"].Value != null)
                {
                    //Llenando el arreglo con los porcentajes de utilidad
                    arrayPU = new double[]{
                    porcUtilidad(Convert.ToDouble(ROW.Cells["%utilidad1"].Value)),
                    porcUtilidad(Convert.ToDouble(ROW.Cells["%utilidad2"].Value)),
                    porcUtilidad(Convert.ToDouble(ROW.Cells["%utilidad3"].Value)),
                    porcUtilidad(Convert.ToDouble(ROW.Cells["%utilidad4"].Value)),
                };
                    try
                    {
                        //Creacion de las categorias
                        callbackInsertCategoria(conn, ROW, t);
                        //Creacion de los depositos
                        callbackInsertDeposito(conn, ROW, t);
                        //Insercion de los articulos
                        callbackInsertArticulo(conn, ROW, t);
                        //Relacion de los articulos con sus respectivas categorias
                        callbackInsertCategoriaArticulo(conn, ROW, t);
                        //Relacion de los Articulos con sus impuestos
                        if (ROW.Cells["cod impuesto1"].Value.ToString() != "")
                        {
                            callbackInsertArticuloImpuestos(conn, ROW, t, 1);
                        }
                        if (!string.IsNullOrWhiteSpace(ROW.Cells["cod impuesto2"].Value.ToString()))
                        {
                            callbackInsertArticuloImpuestos(conn, ROW, t, 2);
                        }
                        //recorrido de cada porcentaje de utilidad y relacionandolo con cada articulo
                        for (int i = 0; i < 4; i++)
                        {
                            callbackInsertArticuloPrecio(conn, ROW, t, i);
                        }
                        codInteno++;

                    }
                    catch (NpgsqlException ex)
                    {

                        try
                        {
                            var db = DBConn.Instance;
                            var col = db.Collection<Errores>();
                            ROW.Cells["Error"].Value = col.Find(x => x.codigo == ex.Code.ToString()).FirstOrDefault().Desc;
                            ROW.DefaultCellStyle.BackColor = Color.Red;
                            count = 0;
                            FAIL = true;
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("Hubo un Error en la insercion de datos. Excepcion: " + ex.Message.ToString());
                        }
                        break;
                    }
                }
            }
            try
            {
                if (!FAIL)
                {
                    //Cargo del inventario previamente cargado
                    callbackInsertCargoInventario(conn, t);

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        try
                        {
                            //Actualizacion de los articulos
                            callbackUpdateArt(conn, row, t);
                        }
                        catch (NpgsqlException ex)
                        {
                            var db = DBConn.Instance;
                            var col = db.Collection<Errores>();
                            row.Cells["Error"].Value = col.Find(x => x.codigo == ex.Code.ToString()).FirstOrDefault().Desc;
                            row.DefaultCellStyle.BackColor = Color.Red;
                            MessageBox.Show(ex.Message);
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                //Mensaje de error en la insercion de datos
                MessageBox.Show(ex.ToString());
                //Cambio de color de la fila del DataGridView cuya insercion arrojo una excepcion

            }

        }

        #endregion

        #region Zonas
        //Metodo de migracion de Zonas a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de las zonas
        /// </summary>
        private void callbackInsertZona(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            if (ROW.Cells["codigo"].Value != null)
            {

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
                dbcmd.Parameters[11].Value = dt.convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                dbcmd.Parameters[12].Value = true;

                count += dbcmd.ExecuteNonQuery();
            }
        }
        #endregion

        #region Unidades
        //Metodo de migracion de Unidades a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de las unidades
        /// </summary>
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
                    dbcmd.Parameters[7].Value = dt.convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                    dbcmd.Parameters[8].Value = true;

                    count += dbcmd.ExecuteNonQuery();
                }
            }

        }
        #endregion

        #region Departamento
        //Metodo de migracion de Departamentos a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de los departamentos
        /// </summary>
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
                    dbcmd.Parameters[6].Value = dt.convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                    dbcmd.Parameters[7].Value = true;

                    count += dbcmd.ExecuteNonQuery();
                }
            }

        }
        #endregion

        #region Moneda
        //Metodo de migracion de Moneda a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de las monedas
        /// </summary>
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
                dbcmd.Parameters[11].Value = dt.convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                dbcmd.Parameters[12].Value = true;

                count += dbcmd.ExecuteNonQuery();
            }

        }
        #endregion

        #region Autorizados
        //Metodo de migracion de Autorizados a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de los autorizados
        /// </summary>
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
                dbcmd.Parameters[10].Value = dt.convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                dbcmd.Parameters[11].Value = true;

                count += dbcmd.ExecuteNonQuery();
            }

        }
        #endregion

        #region Talento
        //Metodo de migracion de Vendedores  a la base de datos de innova (talento)
        /// <summary>
        /// Metodo para migrar  la informacion del talento
        /// </summary>
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
                dbcmd.Parameters[5].Value = dt.convertBoolean(ROW.Cells["vendedor"].Value);
                dbcmd.Parameters[6].Value = dt.convertBoolean(ROW.Cells["cobrador"].Value);
                dbcmd.Parameters[7].Value = dt.convertBoolean(ROW.Cells["servidor"].Value);
                dbcmd.Parameters[8].Value = dt.convertBoolean(ROW.Cells["despachador"].Value);
                dbcmd.Parameters[9].Value = dt.ExtractDate(ROW.Cells["fecha nac"].Value.ToString());
                dbcmd.Parameters[10].Value = "INNOVA";
                dbcmd.Parameters[11].Value = "INNOVA";
                dbcmd.Parameters[12].Value = 1;
                dbcmd.Parameters[13].Value = dt.convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);

                var db = DBConn.Instance;
                var col = db.Collection<Tipos>();
                dbcmd.Parameters[14].Value = col.Find(x => x.tipo == ROW.Cells["tipo contribuyente"].Value.ToString()).FirstOrDefault().codigo;
                dbcmd.Parameters[15].Value = col.Find(x => x.tipo == ROW.Cells["tipo persona"].Value.ToString()).FirstOrDefault().codigo;
                dbcmd.Parameters[16].Value = ROW.Cells["codigo de zona"].Value;
                dbcmd.Parameters[17].Value = true;
                dbcmd.Parameters[18].Value = ROW.Cells["rif"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[19].Value = ROW.Cells["apellidos"].Value;
                dbcmd.Parameters[20].Value = col.Find(x => x.tipo == ROW.Cells["Sexo"].Value.ToString()).FirstOrDefault().codigo;
                dbcmd.Parameters[21].Value = ROW.Cells["direccion"].Value;
                dbcmd.Parameters[22].Value = ROW.Cells["porc ret iva"].Value;
                dbcmd.Parameters[23].Value = ROW.Cells["departamento"].Value;
                dbcmd.Parameters[24].Value = dt.ExtractDate(ROW.Cells["fecha vcto rif"].Value.ToString());
                dbcmd.Parameters[25].Value = dt.ExtractDate(ROW.Cells["fecha de ingreso"].Value.ToString());
                dbcmd.Parameters[26].Value = "ESTA DATA FUE MIGRADA, POR FAVOR VERIFICAR LOS DATOS";

                count += dbcmd.ExecuteNonQuery();
            }

        }
        #endregion

        #region Clientes
        //Metodo de migracion de clientes a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de los clientes
        /// </summary>
        private void callBackInsertClientes(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            if (ROW.Cells["código"].Value != null && ROW.Cells["código"].Value.ToString() != "")
            {
                sql = @"INSERT INTO admin.ven_cli(org_hijo,cod_interno,cli_hijo,descri,
                        tipo_cont,tipo_pers,porc_ret_iva,rif,direc1,monto_descuento,es_descuento,
                        es_exento,es_retencion,es_monto,monto_min,monto_max,monto_cred_max,monto_acum,
                        pri_vmonto,ult_vmonto,ult_pmonto,pago_max,pago_adel,longitud, latitud, altitud,
                        pago_prom,monto_cred_min,saldo, reg_usu_cc,reg_usu_cu,reg_estatus,disponible, 
                        migrado,es_datos,es_vip,es_pronto, observacion, tipo_ret_iva, telefono, email, nomb_persona) 
                        VALUES(@org_hijo,@cod_interno,@cli_hijo,@descri,@tipo_cont,@tipo_pers,
                        @PorcRetIva,@rif,@direc1,@descuento,@es_descuento,
                        @es_exento,@es_retencion,@es_monto,@monto_min,@monto_max,@monto_cred_max,@monto_acum,
                        @pri_vmonto,@ult_vmonto,@ult_pmonto,@pago_max,@pago_adel,@longitud,@latitud,@altitud,
                        @pago_prom,@montoCredMin,@saldo, @reg_usu_cc,@reg_usu_cu,@reg_estatus,@disponible, 
                        @migrado, @esdatos, @esvip, @espronto, @observacion, @tipo_ret_iva,@telefono,@email,@nombPersona)";


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
                dbcmd.Parameters.Add(new NpgsqlParameter("@telefono", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@email", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@nombPersona", NpgsqlDbType.Varchar));


                dbcmd.Prepare();

                dbcmd.Parameters[0].Value = Globals.org;
                dbcmd.Parameters[1].Value = codInteno;
                dbcmd.Parameters[2].Value = ROW.Cells["código"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[3].Value = ROW.Cells["razon social"].Value;
                var db = DBConn.Instance;
                var col = db.Collection<Tipos>();
                dbcmd.Parameters[4].Value = col.Find(x => x.tipo == ROW.Cells["tipo de contribuyente"].Value.ToString().Replace(" ", string.Empty)).FirstOrDefault().codigo;
                dbcmd.Parameters[5].Value = col.Find(x => x.tipo == ROW.Cells["tipo de persona"].Value.ToString().Replace(" ", string.Empty)).FirstOrDefault().codigo;
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
                dbcmd.Parameters[26].Value = dt.convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
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
                dbcmd.Parameters[39].Value = ROW.Cells["telefono"].Value;
                dbcmd.Parameters[40].Value = ROW.Cells["email"].Value;
                dbcmd.Parameters[41].Value = ROW.Cells["nombre de representante"].Value;
                count += dbcmd.ExecuteNonQuery();
            }

        }
        #endregion

        #region Deposito
        //Metodo de migracion de proveedores a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de los depositos
        /// </summary>
        private void callbackInsertDeposito(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            string reader = "0";
            NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
            sql = @"select count(*) from admin.inv_dep where codigo='" + ROW.Cells["codigo deposito"].Value.ToString().Replace(" ", string.Empty) + "'";
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
                dbcmd.Parameters[12].Value = dt.convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                count += dbcmd.ExecuteNonQuery();
            }
        }
        #endregion

        #region Categoria
        //Metodo de migracion de categorias a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de las categorias
        /// </summary>
        private void callbackInsertCategoria(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            string reader = "0";
            NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
            sql = @"select count(*) from admin.inv_cat where cat_hijo ='" + ROW.Cells["codigo categoria"].Value.ToString().Replace(" ", string.Empty) + "'";
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
                dbcmd.Parameters[7].Value = dt.convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                count += dbcmd.ExecuteNonQuery();
            }

        }
        #endregion

        #region Categoria Articulo
        //Metodo de migracion de la relacion entre categorias articulos a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de las categorias por articulos
        /// </summary>
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
                dbcmd.Parameters[6].Value = dt.convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                count += dbcmd.ExecuteNonQuery();
            }

        }
        #endregion

        #region Categoria Servicio
        //Metodo de migracion de la relacion entre categorias servicios a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de las categorias por servicios
        /// </summary>
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
                dbcmd.Parameters[4].Value = dt.convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                count += dbcmd.ExecuteNonQuery();
            }

        }
        #endregion

        #region Categorias por servicio
        //Metodo de migracion de la relacion entre categorias por servicio a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de las categorias de los servicios
        /// </summary>
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
                dbcmd.Parameters[7].Value = dt.convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                count += dbcmd.ExecuteNonQuery();
            }

        }
        #endregion

        #region Articulos
        //Metodo de migracion de articulos a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de los articulos
        /// </summary>
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
                        costo_pro_ant,cantidad_ant,precio,migrado,es_credito_fiscal, tipo_origen, tipo_costo) 
                        VALUES(@org_hijo , @codInterno, @codigo, @descri,
                        @cantidad, @cantCompro, @cantPedido , @cantConsumo, 
                        @cantVenta, @cantMax, @cantMin, @cantFalla , @cantRepos , @cantBulto, 
                        @esPos, @esBien, @esActivo, @esBulto, @esVenc, 
                        @esMedida, @esPeso, @esOferta, @esExento, @esRetencion, @esRegulado,
                        @esExonerado, @esUnico, @esDecimal, @esUnidad, @esParte , 
                        @tipoArt, @costo,@costoPro,@costoAnt,@costoRep, @medPeso , @medAlto ,
                        @medAncho, @medLargo, @medVolumen,@reg_usu_cc, @reg_usu_cu, @regEstatus, 
                        @disponible, @codMedida,@costoProAnt, @cantidadAnt, @precio, @migrado, 
                        @esCreditoFiscal, @tipoOrigen,@tipoCosto)";


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
                    dbcmd.Parameters.Add(new NpgsqlParameter("@tipoOrigen", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@TipoCosto", NpgsqlDbType.Varchar));

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
                    dbcmd.Parameters[22].Value = dt.convertBoolean(ROW.Cells["Exento"].Value); //ES EXENTO
                    dbcmd.Parameters[23].Value = false; //ES RETENCION
                    dbcmd.Parameters[24].Value = false; //ES REGULADO
                    dbcmd.Parameters[25].Value = dt.convertBoolean(ROW.Cells["Exonerado"].Value);  //ES EXONERADO
                    dbcmd.Parameters[26].Value = false; //ES UNICO
                    dbcmd.Parameters[27].Value = true; //ES DECIMAL
                    dbcmd.Parameters[28].Value = false; //ES UNIDAD
                    dbcmd.Parameters[29].Value = false; //ES PARTE
                    var db = DBConn.Instance;
                    var col = db.Collection<Tipos>();
                    dbcmd.Parameters[30].Value = col.Find(x => x.tipo == ROW.Cells["tipo de Articulo"].Value.ToString().Replace(" ", string.Empty)).FirstOrDefault().codigo; //TIPO DE ARTICULO                   
                    if (Convert.ToDouble(ROW.Cells["costo"].Value) > 0 && Convert.ToDouble(ROW.Cells["%utilidad1"].Value) > 0 && Convert.ToDouble(ROW.Cells["existencia"].Value) == 0)
                    {
                        dbcmd.Parameters[31].Value = ROW.Cells["costo"].Value; //COSTO
                        dbcmd.Parameters[32].Value = ROW.Cells["costo promedio"].Value; //COSTO PROMEDIO
                    }
                    else
                    {
                        dbcmd.Parameters[31].Value = 0; //COSTO
                        dbcmd.Parameters[32].Value = 0; //COSTO PROMEDIO
                    }

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
                    dbcmd.Parameters[43].Value = dt.convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value); //DISPONIBLE
                    dbcmd.Parameters[44].Value = ROW.Cells["codigo de la unidad de medida"].Value; //CODIGO MEDIDA
                    dbcmd.Parameters[45].Value = 0; //COSTO PROMEDIO ANTERIOR
                    dbcmd.Parameters[46].Value = 0; //CANTIDAD ANTERIOR
                    dbcmd.Parameters[47].Value = 0; //PRECIO
                    dbcmd.Parameters[48].Value = true;  //MIGRADO
                    dbcmd.Parameters[49].Value = false; //ES CREDITO FISCAL
                    dbcmd.Parameters[50].Value = "10.1";  //TIPO ORIGEN
                    dbcmd.Parameters[51].Value = "12.4"; //TIPO COSTO



                    count += dbcmd.ExecuteNonQuery();
                    total += (Convert.ToDouble(ROW.Cells["costo"].Value) * Convert.ToDouble(ROW.Cells["existencia"].Value));
                    cantidad_items += Convert.ToInt32(ROW.Cells["existencia"].Value);
                }

            }

        }
        #endregion

        #region Servicios
        //Metodo de migracion de servicios a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de los servicios
        /// </summary>
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
                dbcmd.Parameters[11].Value = dt.convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value); //DISPONIBLE
                dbcmd.Parameters[12].Value = ROW.Cells["codigo de la unidad de medida"].Value; //CODIGO MEDIDA
                dbcmd.Parameters[13].Value = true;  //MIGRADO


                count += dbcmd.ExecuteNonQuery();

            }

        }
        #endregion

        #region Servicio Impuestos
        //Metodo de migracion de la relacion entre los servicios y sus impuestos a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de los impuestos en los servicios
        /// </summary>
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
                dbcmd.Parameters[2].Value = ROW.Cells["cod impuesto1"].Value.ToString().Replace(" ", string.Empty); ;
                dbcmd.Parameters[3].Value = dt.convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                dbcmd.Parameters[4].Value = true;

                count += dbcmd.ExecuteNonQuery();
            }

        }
        #endregion

        #region Impuestos
        //Metodo de migracion de impuestos a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de los impuestos
        /// </summary>
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
                    dbcmd.Parameters[9].Value = dt.convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);

                    count += dbcmd.ExecuteNonQuery();
                }
            }
        }
        #endregion

        #region Articulo Precio
        //Metodo de migracion de la relacion entre los articulos y los distintos porcentajes de utilidad a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de los precios por articulos
        /// </summary>
        private void callbackInsertArticuloPrecio(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t, int i)
        {
            if (ROW.Cells["codigo"].Value != null)
            {

                string reader = "0";
                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
                sql = @"select count(*) from admin.inv_art_precio where cod_articulo ='" + ROW.Cells["codigo"].Value.ToString() + "' AND cod_precio ='0" + (i + 1) + "'";
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
                    dbcmd.Parameters[4].Value = "0" + (i + 1);
                    dbcmd.Parameters[5].Value = ROW.Cells["descripcion del producto"].Value;
                    dbcmd.Parameters[6].Value = ROW.Cells["descripcion del producto"].Value;
                    if (Convert.ToDouble(ROW.Cells["costo"].Value) > 0 && arrayPU[i] > 0 && Convert.ToDouble(ROW.Cells["existencia"].Value) == 0)
                    {
                        dbcmd.Parameters[7].Value = preciofinanciero(Convert.ToDouble(ROW.Cells["costo"].Value), arrayPU[i]);
                        dbcmd.Parameters[8].Value = Convert.ToDouble(dbcmd.Parameters[7].Value) - Convert.ToDouble(ROW.Cells["costo"].Value);
                    }
                    else
                    {
                        dbcmd.Parameters[7].Value = 0;
                        dbcmd.Parameters[8].Value = 0;
                    }

                    dbcmd.Parameters[9].Value = false;
                    dbcmd.Parameters[10].Value = 0;
                    dbcmd.Parameters[11].Value = false;
                    dbcmd.Parameters[12].Value = 0;
                    dbcmd.Parameters[13].Value = false;
                    dbcmd.Parameters[14].Value = "INNOVA";
                    dbcmd.Parameters[15].Value = "INNOVA";
                    dbcmd.Parameters[16].Value = 1;
                    dbcmd.Parameters[17].Value = dt.convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                    dbcmd.Parameters[18].Value = porcUtilidad(arrayPU[i]);
                    dbcmd.Parameters[19].Value = 0;
                    dbcmd.Parameters[20].Value = 0;

                    count += dbcmd.ExecuteNonQuery();
                }
            }
        }
        #endregion

        #region Servicio Precio
        //Metodo de migracion de la relacion entre los servicios y los porcentajes de utilidad a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de los precios por servicios
        /// </summary>
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
        #endregion

        #region Usuario
        //Metodo de migracion de usuarios a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de los usuarios
        /// </summary>
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
                dbcmd.Parameters[10].Value = dt.convertBoolean(ROW.Cells["perfiles"].Value);
                dbcmd.Parameters[11].Value = ROW.Cells["cod_perfil"].Value;
                dbcmd.Parameters[12].Value = dt.convertBoolean(ROW.Cells["conectado"].Value);
                dbcmd.Parameters[13].Value = "INNOVA";
                dbcmd.Parameters[14].Value = "INNOVA";
                dbcmd.Parameters[15].Value = ROW.Cells["reg_estatus"].Value;
                dbcmd.Parameters[16].Value = dt.convertBoolean(ROW.Cells["disponible"].Value);
                dbcmd.Parameters[17].Value = true;

                count += dbcmd.ExecuteNonQuery();
            }
        }
        #endregion

        #region Proveedores
        //Metodo de migracion de proveedores a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de los proveedores
        /// </summary>
        private void callbackInsertProveedores(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            if (ROW.Cells["código"].Value != null && ROW.Cells["código"].Value.ToString() != "")
            {
                sql = @"INSERT INTO admin.com_prov(org_hijo,cod_interno,prov_hijo,descri,
                        tipo_cont,tipo_pers,rif,direc1,descuento,es_descuento,
                        es_exento,es_retencion,es_monto,monto_min,monto_max,monto_cred,
                        pri_monto,ult_monto,rect_monto,pago_max,pago_ade,
                        pago_prom,saldo, reg_usu_cc,reg_usu_cu,reg_estatus,disponible, migrado,
                        porc_ret_iva, observacion, tipo_ret_iva, telefono, email, nomb_persona) 
                        VALUES(@org_hijo,@cod_interno,@prov_hijo,@descri,
                        @tipo_cont,@tipo_pers,@rif,@direc1,@descuento,@es_descuento,
                        @es_exento,@es_retencion,@es_monto,@monto_min,@monto_max,@monto_cred,
                        @pri_monto,@ult_monto,@rect_monto,@pago_max,@pago_ade,
                        @pago_prom,@saldo, @reg_usu_cc,@reg_usu_cu,@reg_estatus,@disponible, 
                        @migrado, @porcretiva, @Observacion, @tipo_ret_iva, @telefono,@email,@nombPersona)";


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
                dbcmd.Parameters.Add(new NpgsqlParameter("@telefono", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@email", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@nombPersona", NpgsqlDbType.Varchar));


                dbcmd.Prepare();

                dbcmd.Parameters[0].Value = Globals.org;
                dbcmd.Parameters[1].Value = codInteno;
                dbcmd.Parameters[2].Value = ROW.Cells["código"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[3].Value = ROW.Cells["razon social"].Value;
                var db = DBConn.Instance;
                var col = db.Collection<Tipos>();
                dbcmd.Parameters[4].Value = col.Find(x => x.tipo == ROW.Cells["tipo de contribuyente"].Value.ToString().Replace(" ", string.Empty)).FirstOrDefault().codigo;
                dbcmd.Parameters[5].Value = col.Find(x => x.tipo == ROW.Cells["tipo de persona"].Value.ToString().Replace(" ", string.Empty)).FirstOrDefault().codigo;
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
                dbcmd.Parameters[26].Value = dt.convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                dbcmd.Parameters[27].Value = true;
                dbcmd.Parameters[28].Value = 75;
                dbcmd.Parameters[29].Value = "ESTA DATA FUE MIGRADA, POR FAVOR VERIFICAR LOS DATOS";
                dbcmd.Parameters[30].Value = "15.2";
                dbcmd.Parameters[31].Value = ROW.Cells["telefono"].Value;
                dbcmd.Parameters[32].Value = ROW.Cells["email"].Value;
                dbcmd.Parameters[33].Value = ROW.Cells["nombre de representante"].Value;
                count += dbcmd.ExecuteNonQuery();
            }

        }
        #endregion

        #region Articulo Impuestos
        //Metodo de migracion de la relacion entre articulos y los impuestos aplicados a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de los impuestos por articulos
        /// </summary>
        private void callbackInsertArticuloImpuestos(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t, int z)
        {
            if (ROW.Cells["codigo"].Value != null)
            {
                string reader = "0";
                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
                sql = @"select count(*) from admin.inv_art_imp where cod_articulo ='" + ROW.Cells["codigo"].Value.ToString().Replace(" ", string.Empty) + "' AND cod_impuesto ='" + ROW.Cells["cod impuesto" + z + ""].Value.ToString().Replace(" ", string.Empty) + "'";
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
                    dbcmd.Parameters[2].Value = ROW.Cells["cod impuesto" + z + ""].Value.ToString().Replace(" ", string.Empty); ;
                    dbcmd.Parameters[3].Value = "INNOVA";
                    dbcmd.Parameters[4].Value = "INNOVA";
                    dbcmd.Parameters[5].Value = 1;
                    dbcmd.Parameters[6].Value = dt.convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                    dbcmd.Parameters[7].Value = true;

                    count += dbcmd.ExecuteNonQuery();
                }
            }

        }
        #endregion

        #region Ajuste de Precio
        //Metodo de Ajuste de precios de articulos migrados a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion del ajuste de precios
        /// </summary>
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
        #endregion

        #region Cargo Inventario
        //Metodo para la realziacion del cargo inicial de inventario a la base de datos de innova
        /// <summary>
        /// Metodo para la realizacion del cargo del inventario
        /// </summary>
        private void callbackInsertCargoInventario(NpgsqlConnection conn, NpgsqlTransaction t)
        {

            try
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

                string ajuste = dbcmd.ExecuteScalar().ToString();

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
                dbcmd.Parameters[13].Value = ajuste;
                dbcmd.Parameters[14].Value = true;

                count += dbcmd.ExecuteNonQuery();

                sql = @"SELECT doc from admin.int_cargo order by fecha_reg desc";
                dbcmd = new NpgsqlCommand(sql, conn);

                string cargo = dbcmd.ExecuteScalar().ToString();
                //Insercion del detalle del ajuste
                foreach (DataGridViewRow ROW2 in dataGridView1.Rows)
                {
                    if (ROW2.Cells["codigo"].Value != null)
                    {
                        try
                        {
                            if (ROW2.Cells["existencia"].Value.ToString().Replace(" ", string.Empty) != "0")
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
                                dbcmd.Parameters[1].Value = ajuste;
                                dbcmd.Parameters[2].Value = ROW2.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);
                                dbcmd.Parameters[3].Value = ROW2.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);
                                dbcmd.Parameters[4].Value = ROW2.Cells["costo"].Value;
                                dbcmd.Parameters[5].Value = ROW2.Cells["costo promedio"].Value;
                                dbcmd.Parameters[6].Value = DateTime.Now;
                                dbcmd.Parameters[7].Value = 1;
                                dbcmd.Parameters[8].Value = item;
                                dbcmd.Parameters[9].Value = true;

                                count += dbcmd.ExecuteNonQuery();

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
                                dbcmd.Parameters[1].Value = cargo;
                                dbcmd.Parameters[2].Value = item;
                                dbcmd.Parameters[3].Value = ROW2.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);
                                dbcmd.Parameters[4].Value = ROW2.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);
                                dbcmd.Parameters[5].Value = ROW2.Cells["descripcion del producto"].Value;
                                dbcmd.Parameters[6].Value = ROW2.Cells["existencia"].Value;
                                dbcmd.Parameters[7].Value = ROW2.Cells["existencia"].Value;
                                dbcmd.Parameters[8].Value = 0;
                                dbcmd.Parameters[9].Value = 0;
                                dbcmd.Parameters[10].Value = 0;
                                dbcmd.Parameters[11].Value = ROW2.Cells["costo promedio"].Value;
                                dbcmd.Parameters[12].Value = ROW2.Cells["codigo deposito"].Value.ToString().Replace(" ", string.Empty);
                                dbcmd.Parameters[13].Value = ROW2.Cells["costo promedio"].Value;
                                dbcmd.Parameters[14].Value = (Convert.ToDouble(ROW2.Cells["costo"].Value.ToString().Replace(".", ",")) * (double)ROW2.Cells["existencia"].Value);
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
                        catch (NpgsqlException ex)
                        {
                            MessageBox.Show(ex.Message);
                            var db = DBConn.Instance;
                            var col = db.Collection<Errores>();
                            ROW2.Cells["Error"].Value = col.Find(x => x.codigo == ex.Code.ToString()).FirstOrDefault().Desc;
                            ROW2.DefaultCellStyle.BackColor = Color.Red;
                            count = 0;
                        }
                    }
                }

            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());
            }

        }
        #endregion

        #region Cuentas Contables
        //Metodo de migracion de cuentas contables a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de las cuentas contables
        /// </summary>
        private void callbackInsertCuentasCont(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t, string tabla)
        {
            if (!string.IsNullOrWhiteSpace(Convert.ToString(ROW.Cells["Codigo"].Value)))
            {
                sql = @"INSERT INTO " + tabla + ".acc_cuentas(org_hijo,cuenta_hijo,descri,descorta,es_movimiento,reg_usu_cc,reg_usu_cu,reg_estatus,disponible) VALUES(@org_hijo , @cuentaHijo, @descri,@descorta,@es_movimiento,@reg_usu_cc, @reg_usu_cu, @regEstatus, @disponible)";


                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);

                dbcmd.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@cuentaHijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descorta", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@es_movimiento", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cu", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
                dbcmd.Parameters.Add(new NpgsqlParameter("@disponible", NpgsqlDbType.Boolean));

                dbcmd.Prepare();

                dbcmd.Parameters[0].Value = Globals.org;
                dbcmd.Parameters[1].Value = ROW.Cells["Codigo"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[2].Value = ROW.Cells["Descripcion"].Value;
                dbcmd.Parameters[3].Value = ROW.Cells["Descripcion Corta"].Value;
                dbcmd.Parameters[4].Value = true;
                dbcmd.Parameters[5].Value = "INNOVA";
                dbcmd.Parameters[6].Value = "INNOVA";
                dbcmd.Parameters[7].Value = 1;
                dbcmd.Parameters[8].Value = dt.convertBoolean(ROW.Cells["Estatus (disponibilidad)"].Value);
                count += dbcmd.ExecuteNonQuery();
            }



        }
        #endregion

        #region CxC
        //Metodo de migracion de cuentas por cobrar a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de las cuentas por cobrar
        /// </summary>
        private void callbackInsertCxC(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            sql = @"INSERT INTO admin.fin_cxc(org_hijo,doc_num,cod_cli,fecha_emi,fecha_ven,factor,
                      total,saldo,saldo_inicial,total_exento,doc_control,reg_usu_cc,
                      reg_estatus,cod_empleado , migrado, cod_impresorafiscal, descri, tipo_opera, debito, credito) 
                      VALUES(@org_hijo,@docNum,@codCli,@fechaEmi, @fechaVen, @factor, @total, @saldo, @saldoInicial, @totalEx, 
                      @doc_control,@reg_usu_cc, @reg_estatus, @codEmpleado, @migrado, @cod_impresorafiscal, @descri, @tipoOpera,
                        @debito, @credito)";
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
            dbcmd.Parameters.Add(new NpgsqlParameter("@debito", NpgsqlDbType.Double));
            dbcmd.Parameters.Add(new NpgsqlParameter("@credito", NpgsqlDbType.Double));

            dbcmd.Prepare();

            dbcmd.Parameters[0].Value = Globals.org;
            dbcmd.Parameters[1].Value = ROW.Cells["numero factura"].Value.ToString().Replace(" ", string.Empty);
            dbcmd.Parameters[2].Value = ROW.Cells["cod cliente"].Value.ToString().Replace(" ", string.Empty);
            dbcmd.Parameters[3].Value = dt.ExtractDate(ROW.Cells["fecha emision"].Value.ToString());
            dbcmd.Parameters[4].Value = dt.ExtractDate(ROW.Cells["fecha vencimiento"].Value.ToString());
            dbcmd.Parameters[5].Value = 0;
            dbcmd.Parameters[6].Value = ROW.Cells["monto total"].Value.ToString().Replace(".", ",");
            dbcmd.Parameters[7].Value = ROW.Cells["saldo"].Value.ToString().Replace(".", ",");
            dbcmd.Parameters[8].Value = ROW.Cells["saldo inicial"].Value.ToString().Replace(".", ",");
            dbcmd.Parameters[9].Value = ROW.Cells["monto exento"].Value.ToString().Replace(".", ",");
            dbcmd.Parameters[10].Value = ROW.Cells["numero de control"].Value.ToString().Replace(" ", string.Empty);
            dbcmd.Parameters[11].Value = "INNOVA";
            dbcmd.Parameters[12].Value = 1;
            dbcmd.Parameters[13].Value = ROW.Cells["cod vendedor"].Value.ToString().Replace(" ", string.Empty);
            dbcmd.Parameters[14].Value = true;
            dbcmd.Parameters[15].Value = ROW.Cells["numero impresora fiscal"].Value;
            dbcmd.Parameters[16].Value = ROW.Cells["descripcion"].Value;
            var db = DBConn.Instance;
            var col = db.Collection<Tipos>();
            dbcmd.Parameters[17].Value = col.Find(x => x.tipo == ROW.Cells["tipo operacion"].Value.ToString().Replace(" ", string.Empty)).FirstOrDefault().codigo;
            switch ((string)ROW.Cells["tipo operacion"].Value.ToString().Replace(" ", string.Empty))
            {
                case "Fcxc":
                    {
                        dbcmd.Parameters[18].Value = 0;
                        dbcmd.Parameters[19].Value = ROW.Cells["saldo inicial"].Value.ToString().Replace(".", ",");
                        break;
                    }
                case "NC":
                    {
                        dbcmd.Parameters[18].Value = ROW.Cells["saldo inicial"].Value.ToString().Replace(".", ",");
                        dbcmd.Parameters[19].Value = 0;
                        break;
                    }
                case "ND":
                    {
                        dbcmd.Parameters[18].Value = 0;
                        dbcmd.Parameters[19].Value = ROW.Cells["saldo inicial"].Value.ToString().Replace(".", ",");
                        break;
                    }
            }

            count += dbcmd.ExecuteNonQuery();


        }
        #endregion

        #region CxC Impuestos
        //Metodo de migracion de la relacion entre cuentas por cobrar y los impuestos aplicados a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de las cuentas por  cobrar y sus impuestos
        /// </summary>
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
            dbcmd.Parameters[1].Value = ROW.Cells["porc impuesto" + z + ""].Value.ToString().Replace(".", ",");
            dbcmd.Parameters[2].Value = ROW.Cells["cod impuesto" + z + ""].Value.ToString().Replace(" ", string.Empty);
            dbcmd.Parameters[3].Value = ROW.Cells["base imponible" + z + ""].Value.ToString().Replace(".", ",");
            dbcmd.Parameters[4].Value = ROW.Cells["monto total"].Value.ToString().Replace(".", ",");
            dbcmd.Parameters[5].Value = reader;
            dbcmd.Parameters[6].Value = 1;
            dbcmd.Parameters[7].Value = true;

            count += dbcmd.ExecuteNonQuery();

        }
        #endregion

        #region CxP
        //Metodo de migracion de cuentas por pagar a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de las cuentas por pagar
        /// </summary>
        private void callbackInsertCxP(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            sql = @"INSERT INTO admin.fin_cxp(org_hijo,doc_num,cod_prov,fecha_emi,fecha_ven,factor,
                      total,saldo,saldo_inicial,total_exento,doc_control,reg_usu_cc,
                      reg_estatus, migrado, cod_impresorafiscal, descri, tipo_opera,debito,credito) 
                      VALUES(@org_hijo,@docNum,@codPro,
                      @fechaEmi, @fechaVen, @factor, @total, @saldo, @saldoInicial, @totalEx, 
                      @doc_control,@reg_usu_cc, @reg_estatus, @migrado, @cod_impresorafiscal, @descri, @tipoOpera, @debito, @credito)";
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
            dbcmd.Parameters.Add(new NpgsqlParameter("@debito", NpgsqlDbType.Double));
            dbcmd.Parameters.Add(new NpgsqlParameter("@credito", NpgsqlDbType.Double));
            dbcmd.Prepare();

            dbcmd.Parameters[0].Value = Globals.org;
            dbcmd.Parameters[1].Value = ROW.Cells["numero factura"].Value.ToString().Replace(" ", string.Empty);
            dbcmd.Parameters[2].Value = ROW.Cells["cod proveedor"].Value.ToString().Replace(" ", string.Empty);
            dbcmd.Parameters[3].Value = dt.ExtractDate(ROW.Cells["fecha emision"].Value.ToString());
            dbcmd.Parameters[4].Value = dt.ExtractDate(ROW.Cells["fecha vencimiento"].Value.ToString());
            dbcmd.Parameters[5].Value = 0;
            dbcmd.Parameters[6].Value = ROW.Cells["monto total"].Value;
            dbcmd.Parameters[7].Value = ROW.Cells["saldo"].Value;
            dbcmd.Parameters[8].Value = ROW.Cells["saldo inicial"].Value;
            dbcmd.Parameters[9].Value = ROW.Cells["monto exento"].Value;
            dbcmd.Parameters[10].Value = ROW.Cells["numero de control"].Value.ToString().Replace(" ", string.Empty);
            dbcmd.Parameters[11].Value = "INNOVA";
            dbcmd.Parameters[12].Value = 1;
            dbcmd.Parameters[13].Value = true;
            dbcmd.Parameters[14].Value = ROW.Cells["numero impresora fiscal"].Value;
            dbcmd.Parameters[15].Value = ROW.Cells["descripcion"].Value;
            var db = DBConn.Instance;
            var col = db.Collection<Tipos>();
            dbcmd.Parameters[16].Value = col.Find(x => x.tipo == ROW.Cells["tipo operacion"].Value.ToString().Replace(" ", string.Empty)).FirstOrDefault().codigo;
            switch ((string)ROW.Cells["tipo operacion"].Value.ToString().Replace(" ", string.Empty))
            {
                case "Fcxp":
                    {
                        dbcmd.Parameters[17].Value = ROW.Cells["saldo inicial"].Value;
                        dbcmd.Parameters[18].Value = 0;
                        break;
                    }
                case "NC":
                    {
                        dbcmd.Parameters[17].Value = 0;
                        dbcmd.Parameters[18].Value = ROW.Cells["saldo inicial"].Value;
                        break;
                    }
                case "ND":
                    {
                        dbcmd.Parameters[17].Value = ROW.Cells["saldo inicial"].Value;
                        dbcmd.Parameters[18].Value = 0;
                        break;
                    }
            }
            count += dbcmd.ExecuteNonQuery();
        }
        #endregion

        #region CxP Impuestos
        //Metodo de migracion de la relacion de cuentas por pagar y los impuestos aplicados a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de los impuestos y las cuentas por pagar
        /// </summary>
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
            dbcmd.Parameters[2].Value = ROW.Cells["cod impuesto" + z + ""].Value.ToString().Replace(" ", string.Empty);
            dbcmd.Parameters[3].Value = ROW.Cells["base imponible" + z + ""].Value;
            dbcmd.Parameters[4].Value = ROW.Cells["monto total"].Value;
            dbcmd.Parameters[5].Value = reader;
            dbcmd.Parameters[6].Value = 1;
            dbcmd.Parameters[7].Value = true;

            count += dbcmd.ExecuteNonQuery();
        }
        #endregion

        #region Adelantos (clientes) 
        //Metodo de migracion de Adelantos de clientes a la base de datos de innova 
        /// <summary>
        /// Metodo para migrar la informacion de los adelantos de los clientes
        /// </summary>
        private void callbackInsertAdelantosCli(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            if (ROW.Cells["Codigo del Cliente"].Value != null)
            {
                NpgsqlCommand dbcmd = new NpgsqlCommand();
                if (!adelantos.ContainsKey(ROW.Cells["Codigo del Cliente"].Value.ToString().Replace(" ", string.Empty)))
                {
                    adelantos[ROW.Cells["Codigo del Cliente"].Value.ToString().Replace(" ", string.Empty)] = dt.ExtractDate(ROW.Cells["Fecha adelanto"].Value.ToString());

                    sql = @"INSERT INTO admin.fin_cli_adelanto(org_hijo,saldo,reg_usu_cc,reg_estatus, cli_hijo,
                     migrado) VALUES(@orgHijo ,
                    @saldo, @regusucc , @regEstatus,@clihijo, @migrado)";
                    dbcmd = new NpgsqlCommand(sql, conn, t);
                    dbcmd.Parameters.Add(new NpgsqlParameter("@orgHijo", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@saldo", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@regusucc", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@clihijo", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));


                    dbcmd.Prepare();

                    dbcmd.Parameters[0].Value = Globals.org;
                    dbcmd.Parameters[1].Value = 0;
                    dbcmd.Parameters[2].Value = "INNOVA";
                    dbcmd.Parameters[3].Value = 1;
                    dbcmd.Parameters[4].Value = ROW.Cells["Codigo del Cliente"].Value.ToString().Replace(" ", string.Empty);
                    dbcmd.Parameters[5].Value = true;

                    count += dbcmd.ExecuteNonQuery();
                }


                sql = @"INSERT INTO admin.fin_cli_ade_det(org_hijo,monto,observacion, cli_hijo,
                     fecha,migrado) VALUES(@orgHijo , @monto,
                    @observacion , @clihijo,@fecha, @migrado)";
                dbcmd = new NpgsqlCommand(sql, conn, t);
                dbcmd.Parameters.Add(new NpgsqlParameter("@orgHijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@monto", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@observacion", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@clihijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@fecha", NpgsqlDbType.Date));
                dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));


                dbcmd.Prepare();

                dbcmd.Parameters[0].Value = Globals.org;
                dbcmd.Parameters[1].Value = ROW.Cells["Monto"].Value;
                dbcmd.Parameters[2].Value = ROW.Cells["Observacion"].Value;
                dbcmd.Parameters[3].Value = ROW.Cells["Codigo del Cliente"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[4].Value = dt.ExtractDate(ROW.Cells["Fecha adelanto"].Value.ToString());
                dbcmd.Parameters[5].Value = true;

                count += dbcmd.ExecuteNonQuery();

            }
        }
        #endregion

        #region Adelantos Proveedores
        //Metodo de migracion de adelantos de proveedores a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de los adelantos de los proveedores
        /// </summary>
        private void callbackInsertAdelantosProv(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {

            if (ROW.Cells["Codigo del Proveedor"].Value != null)
            {
                NpgsqlCommand dbcmd = new NpgsqlCommand();
                if (!adelantos.ContainsKey(ROW.Cells["Codigo del Proveedor"].Value.ToString().Replace(" ", string.Empty)))
                {
                    adelantos[ROW.Cells["Codigo del Proveedor"].Value.ToString().Replace(" ", string.Empty)] = dt.ExtractDate(ROW.Cells["Fecha adelanto"].Value.ToString());

                    sql = @"INSERT INTO admin.fin_prov_adelanto(org_hijo,saldo,reg_usu_cc,reg_estatus, prov_hijo,
                     migrado) VALUES(@orgHijo , 
                    @saldo, @regusucc , @regEstatus,@clihijo, @migrado)";
                    dbcmd = new NpgsqlCommand(sql, conn, t);
                    dbcmd.Parameters.Add(new NpgsqlParameter("@orgHijo", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@saldo", NpgsqlDbType.Double));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@regusucc", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@clihijo", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));

                    dbcmd.Prepare();

                    dbcmd.Parameters[0].Value = Globals.org;
                    dbcmd.Parameters[1].Value = 0;
                    dbcmd.Parameters[2].Value = "INNOVA";
                    dbcmd.Parameters[3].Value = 1;
                    dbcmd.Parameters[4].Value = ROW.Cells["Codigo del Proveedor"].Value.ToString().Replace(" ", string.Empty);
                    dbcmd.Parameters[5].Value = true;

                    count += dbcmd.ExecuteNonQuery();
                }


                sql = @"INSERT INTO admin.fin_prov_ade_det(org_hijo,monto,observacion, prov_hijo,
                     fecha,migrado) VALUES(@orgHijo , @monto,
                     @observacion , @clihijo,@fecha, @migrado)";
                dbcmd = new NpgsqlCommand(sql, conn, t);
                dbcmd.Parameters.Add(new NpgsqlParameter("@orgHijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@monto", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@observacion", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@clihijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@fecha", NpgsqlDbType.Date));
                dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));


                dbcmd.Prepare();

                dbcmd.Parameters[0].Value = Globals.org;
                dbcmd.Parameters[1].Value = ROW.Cells["Monto"].Value;
                dbcmd.Parameters[2].Value = ROW.Cells["Observacion"].Value;
                dbcmd.Parameters[3].Value = ROW.Cells["Codigo del Proveedor"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[4].Value = dt.ExtractDate(ROW.Cells["Fecha adelanto"].Value.ToString());
                dbcmd.Parameters[5].Value = true;

                count += dbcmd.ExecuteNonQuery();
            }

        }
        #endregion

        #region Actualizar Articulos
        //Metodo para actualizar articulos migrados a la base de datos de innova
        /// <summary>
        /// Metodo para actualizar la informacion de los articulos luego de realizar el cargo
        /// </summary>
        private void callbackUpdateArt(NpgsqlConnection conn, DataGridViewRow ROW2, NpgsqlTransaction t)
        {
            if (ROW2.Cells["codigo"].Value != null)
            {
                try
                {
                    if (ROW2.Cells["existencia"].Value.ToString().Replace(" ", string.Empty) != "0")
                    {
                        item++;
                        sql = @"UPDATE admin.inv_art SET costo=@costo WHERE codigo=@codigo";
                        NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn);
                        dbcmd.Parameters.Add(new NpgsqlParameter("@costo", NpgsqlDbType.Double));
                        dbcmd.Parameters.Add(new NpgsqlParameter("@codigo", NpgsqlDbType.Varchar));
                        dbcmd.Prepare();
                        dbcmd.Parameters[0].Value = ROW2.Cells["costo"].Value;
                        dbcmd.Parameters[1].Value = ROW2.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);

                        dbcmd.ExecuteNonQuery();
                    }

                }
                catch (NpgsqlException ex)
                {
                    MessageBox.Show(ex.Message);
                    var db = DBConn.Instance;
                    var col = db.Collection<Errores>();
                    ROW2.Cells["Error"].Value = col.Find(x => x.codigo == ex.Code.ToString()).FirstOrDefault().Desc;
                    ROW2.DefaultCellStyle.BackColor = Color.Red;
                    count = 0;

                }


            }
        }
        #endregion

        #endregion

        #region Metodos Migracion Nomina

        #region Profesiones
        //Metodo de migracion de profesiones a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de las profesiones
        /// </summary>
        private void callbackInsertProfesiones(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            if (ROW.Cells["codigo"].Value != null)
            {

                sql = @"INSERT INTO nomina.profesion(org_hijo,codigo,descri,descorta,abreviatura,
                    tipo,reg_usu_cc, reg_usu_cu,reg_estatus, 
                    disponible,migrado) VALUES(@orgHijo , @codigo,
                    @descri, @descorta , @abreviatura, @tipo, @reg_usu_cc , 
                    @reg_usu_cu, @regEstatus, @disponible, @migrado)";
                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
                dbcmd.Parameters.Add(new NpgsqlParameter("@orghijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codigo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descorta", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@abreviatura", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@tipo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cu", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
                dbcmd.Parameters.Add(new NpgsqlParameter("@disponible", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));

                dbcmd.Prepare();

                dbcmd.Parameters[0].Value = Globals.org;
                dbcmd.Parameters[1].Value = ROW.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[2].Value = ROW.Cells["Descripcion de la profesion"].Value;
                dbcmd.Parameters[3].Value = ROW.Cells["Descripcion corta"].Value;
                dbcmd.Parameters[4].Value = ROW.Cells["Abreviatura"].Value;
                dbcmd.Parameters[5].Value = ROW.Cells["Tipo de Profesion"].Value;
                dbcmd.Parameters[6].Value = "INNOVA";
                dbcmd.Parameters[7].Value = "INNOVA";
                dbcmd.Parameters[8].Value = 1;
                dbcmd.Parameters[9].Value = dt.convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                dbcmd.Parameters[10].Value = true;
                count += dbcmd.ExecuteNonQuery();
            }
        }
        #endregion

        #region Cargos
        //Metodo de migracion de cargos a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de la realizacion del cargo
        /// </summary>
        private void callbackInsertCargo(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            if (ROW.Cells["codigo"].Value != null)
            {

                sql = @"INSERT INTO nomina.cargo(org_hijo,codigo,descri,descorta,riego,
                    nivel_prof,reg_usu_cc, reg_usu_cu,reg_estatus, 
                    disponible,migrado) VALUES(@orgHijo , @codigo,
                    @descri, @descorta , @riesgo, @nivelProf, @reg_usu_cc , 
                    @reg_usu_cu, @regEstatus, @disponible, @migrado)";
                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
                dbcmd.Parameters.Add(new NpgsqlParameter("@orghijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codigo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descorta", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@riesgo", NpgsqlDbType.Integer));
                dbcmd.Parameters.Add(new NpgsqlParameter("@nivelProf", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cu", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
                dbcmd.Parameters.Add(new NpgsqlParameter("@disponible", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));

                dbcmd.Prepare();

                dbcmd.Parameters[0].Value = Globals.org;
                dbcmd.Parameters[1].Value = ROW.Cells["Codigo"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[2].Value = ROW.Cells["Descripcion del cargo"].Value;
                dbcmd.Parameters[3].Value = ROW.Cells["Descripcion corta"].Value;
                dbcmd.Parameters[4].Value = ROW.Cells["Escala de riesgo"].Value;
                dbcmd.Parameters[5].Value = ROW.Cells["Nivel de Profesion"].Value;
                dbcmd.Parameters[6].Value = "INNOVA";
                dbcmd.Parameters[7].Value = "INNOVA";
                dbcmd.Parameters[8].Value = 1;
                dbcmd.Parameters[9].Value = dt.convertBoolean(ROW.Cells["estatus (disponibilidad)"].Value);
                dbcmd.Parameters[10].Value = true;
                count += dbcmd.ExecuteNonQuery();
            }
        }
        #endregion

        #region variables
        //Metodo de migracion de variables a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de las variables de nomina
        /// </summary>
        private void callbackInsertVariables(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            if (ROW.Cells["codigo"].Value != null)
            {
                string reader = "0";
                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
                sql = @"select count(*) from nomina.variable where codigo ='" + ROW.Cells["codigo"].Value.ToString() + "'";
                dbcmd = new NpgsqlCommand(sql, conn);
                reader = dbcmd.ExecuteScalar().ToString();

                if (reader == "0")
                {

                    sql = @"INSERT INTO nomina.variable(org_hijo,codigo,
                    descri,descorta,tipo,formula,detalle,reg_usu_cc,
                    reg_usu_cu,reg_estatus, disponible) 
                    VALUES(@orgHijo , @codigo,@descri, @descorta , 
                    @tipo, @formula,@detalle, @reg_usu_cc , 
                    @reg_usu_cu, @regEstatus, @disponible)";
                    dbcmd = new NpgsqlCommand(sql, conn, t);

                    dbcmd.Parameters.Add(new NpgsqlParameter("@orghijo", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@codigo", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@descorta", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@tipo", NpgsqlDbType.Integer));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@formula", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@detalle", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cu", NpgsqlDbType.Varchar));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
                    dbcmd.Parameters.Add(new NpgsqlParameter("@disponible", NpgsqlDbType.Boolean));
                    //dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));

                    dbcmd.Prepare();

                    dbcmd.Parameters[0].Value = Globals.org;
                    dbcmd.Parameters[1].Value = ROW.Cells["Codigo"].Value.ToString().Replace(" ", string.Empty);
                    dbcmd.Parameters[2].Value = ROW.Cells["Descripcion"].Value;
                    dbcmd.Parameters[3].Value = ROW.Cells["Descripcion corta"].Value;
                    var db = DBConn.Instance;
                    var col = db.Collection<Tipos>();
                    dbcmd.Parameters[4].Value = col.Find(x => x.tipo == ROW.Cells["Tipo"].Value.ToString()).FirstOrDefault().codigo;
                    dbcmd.Parameters[5].Value = ROW.Cells["Formula"].Value;
                    dbcmd.Parameters[6].Value = ROW.Cells["Detalle"].Value;
                    dbcmd.Parameters[7].Value = "INNOVA";
                    dbcmd.Parameters[8].Value = "INNOVA";
                    dbcmd.Parameters[9].Value = 1;
                    dbcmd.Parameters[10].Value = dt.convertBoolean(ROW.Cells["Estatus (disponibilidad)"].Value);
                    //dbcmd.Parameters[11].Value = true;
                    count += dbcmd.ExecuteNonQuery();
                }
            }
        }
        #endregion

        #region Detalles de Variables
        //Metodo de migracion de los detalles de las variables de nomina a la base de datos de innova
        /// <summary>
        /// Metodo para migrar  la informacion de los detalles de cada variable de nomina
        /// </summary>
        private void callbackInsertVariablesDet(NpgsqlConnection conn, DataGridViewRow ROW, NpgsqlTransaction t)
        {
            if (ROW.Cells["codigo"].Value != null)
            {

                sql = @"INSERT INTO nomina.variable_det(org_hijo,codigo,
                    monto,valor_min,valor_max,domingo,lunes,martes,miercoles,
                    jueves,viernes,sabado, disponible,fecha) 
                    VALUES(@orgHijo , @codigo, @monto, @valorMin , 
                    @valorMax, @domingo, @lunes, @martes , 
                    @miercoles, @jueves, @viernes, @sabado, @disponible, @fecha)";
                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn, t);
                dbcmd.Parameters.Add(new NpgsqlParameter("@orghijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@codigo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@monto", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@valorMin", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@valorMax", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@domingo", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@lunes", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@martes", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@miercoles", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@jueves", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@viernes", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@sabado", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@disponible", NpgsqlDbType.Boolean));
                dbcmd.Parameters.Add(new NpgsqlParameter("@fecha", NpgsqlDbType.Date));
                //dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));

                dbcmd.Prepare();

                dbcmd.Parameters[0].Value = Globals.org;
                dbcmd.Parameters[1].Value = ROW.Cells["Codigo"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[2].Value = ROW.Cells["Monto"].Value;
                dbcmd.Parameters[3].Value = ROW.Cells["Valor Minimo"].Value;
                dbcmd.Parameters[4].Value = ROW.Cells["Valor Maximo"].Value;
                dbcmd.Parameters[5].Value = dt.convertBoolean(ROW.Cells["Domingo"].Value);
                dbcmd.Parameters[6].Value = dt.convertBoolean(ROW.Cells["Lunes"].Value);
                dbcmd.Parameters[7].Value = dt.convertBoolean(ROW.Cells["Martes"].Value);
                dbcmd.Parameters[8].Value = dt.convertBoolean(ROW.Cells["Miercoles"].Value);
                dbcmd.Parameters[9].Value = dt.convertBoolean(ROW.Cells["Jueves"].Value);
                dbcmd.Parameters[10].Value = dt.convertBoolean(ROW.Cells["Viernes"].Value);
                dbcmd.Parameters[11].Value = dt.convertBoolean(ROW.Cells["Sabado"].Value);
                dbcmd.Parameters[12].Value = dt.convertBoolean(ROW.Cells["Estatus (disponibilidad)"].Value);
                dbcmd.Parameters[13].Value = dt.ExtractDate(ROW.Cells["fecha"].Value.ToString());
                //dbcmd.Parameters[14].Value = true;

                count += dbcmd.ExecuteNonQuery();
            }
        }
        #endregion

        #endregion Nomina

        #region Eventos y Metodos
        //Evento click del boton seleccionar
        /// <summary>
        /// Evento click del boton que consulta la informacion de la tabla seleccionada y la carga en la vista
        /// </summary>
        private void button1_Click(object sender, EventArgs e)
        {
            count = 0;
            //seleccion del archivo a cargar
            DialogResult dr = openFileDialog1.ShowDialog();
            if (dr == DialogResult.OK)
            {
                file = openFileDialog1.FileName;
                try
                {
                    //inicio de conexion al archivo .xls
                    Myconnetion = new System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source =" + file + "; Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1;\"");
                    Myconnetion.Open();
                    //carga de informacion del .xls

                    // obtener nombre de la hoja de excel
                    System.Data.DataTable dbSchema = Myconnetion.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    if (dbSchema == null || dbSchema.Rows.Count < 1)
                    {
                        throw new Exception("Error: No se pudo determinar el nombre de la Hoja de Trabajo.");
                    }
                    string firstSheetName = dbSchema.Rows[0]["TABLE_NAME"].ToString();
                    //Consulta al formato en excel
                    MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [" + firstSheetName + "]", Myconnetion);
                    MyCommand.TableMappings.Add("Table", "TestTable");

                    //Vaciado del DataSet
                    DtSet.Reset();
                    //Llenado del DataSet con resultado de comando ejecutado al .xls
                    MyCommand.Fill(DtSet);
                    //Asignacion del DataSet como origen de datos del DataGridView 
                    dataGridView1.DataSource = DtSet.Tables[0];
                    dataGridView1.Columns["Error"].Visible = true;
                    dataGridView1.Columns["numero"].Visible = true;
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        count++;
                        row.Cells["numero"].Value = count;
                    }

                    //Cerrar conexion
                    Myconnetion.Close();
                    button2.Enabled = true;
                }
                catch (Exception ex)
                {
                    //Captura de excepcion durante las acciones del button1_click
                    MessageBox.Show("Se produjo un error al cargar la informacion. Error: " + ex.Message.ToString());
                }
            }

        }
        //Evento click del boton migrar
        /// <summary>
        /// Evento click del boton que migra la informacion cargada en la vista hacia la base de datos
        /// </summary>
        private void button2_Click(object sender, EventArgs e)
        {
            //Asignacion de valores iniciales de variables necesarias para la migracion
            Cursor.Current = Cursors.WaitCursor;
            codInteno = 1;
            status = 0;
            total = 0;
            adelantos.Clear();
            NpgsqlConnection conn = new NpgsqlConnection(connectionString);
            conn.Open();
            NpgsqlTransaction t = conn.BeginTransaction();
            try
            {
                //Recorriendo el Datagridview e insertando cada valor
                count = 0;
                if (comboBox1.SelectedIndex == 5)//Articulos
                {
                    Exportar_Articulos(conn, t);
                }
                else if (comboBox1.SelectedIndex == 6)//Servicios
                {

                    Exportar_Servicios(conn, t);

                }
                else
                {
                    foreach (DataGridViewRow ROW in dataGridView1.Rows)
                    {
                        try
                        {
                            if (radioButton1.Checked)
                            {
                                //seleccion del metodo de migracion en base a la seleccion del combo box de tablas
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
                                    case 13:
                                        {
                                            callbackInsertAdelantosProv(conn, ROW, t);
                                            break;
                                        }
                                    case 14:
                                        {
                                            callbackInsertAdelantosCli(conn, ROW, t);
                                            break;
                                        }
                                    case 15:
                                        {
                                            callbackInsertCuentasCont(conn, ROW, t, "admin");
                                            callbackInsertCuentasCont(conn, ROW, t, "contab");
                                            break;
                                        }
                                }
                            }
                            else
                            {
                                switch (comboBox1.SelectedIndex)
                                {
                                    case 0:
                                        {
                                            callbackInsertProfesiones(conn, ROW, t);
                                            break;
                                        }
                                    case 1:
                                        {
                                            callbackInsertCargo(conn, ROW, t);
                                            break;
                                        }
                                    case 2:
                                        {
                                            if ((!string.IsNullOrWhiteSpace(((ROW.Cells["Tipo"].Value) ?? "").ToString())))
                                            {
                                                callbackInsertVariables(conn, ROW, t);

                                                if (ROW.Cells["Tipo"].Value.ToString() != "FORMULA")
                                                    if (ROW.Cells["Tipo"].Value.ToString() == "NUMERICA")
                                                        if (((double)ROW.Cells["Monto"].Value >= (double)ROW.Cells["Valor Minimo"].Value && (double)ROW.Cells["Monto"].Value <= (double)ROW.Cells["Valor Maximo"].Value))
                                                            callbackInsertVariablesDet(conn, ROW, t);
                                                        else MessageBox.Show("El monto no puede ser menor al valor minimo o mayor al maximo");
                                                    else callbackInsertVariablesDet(conn, ROW, t);
                                            }


                                            break;
                                        }
                                }
                            }


                        }
                        catch (NpgsqlException ex)
                        {
                            //Mensaje de error en la insercion de datos
                            try
                            {
                                var db = DBConn.Instance;
                                var col = db.Collection<Errores>();
                                ROW.Cells["Error"].Value = col.Find(x => x.codigo == ex.Code.ToString()).FirstOrDefault().Desc;
                                ROW.DefaultCellStyle.BackColor = Color.Red;
                                count = 0;
                            }
                            catch (Exception)
                            {
                                MessageBox.Show("Hubo un Error en la insercion de datos. Excepcion: " + ex.Message.ToString());
                            }
                            break;

                            // Cambio de color de la fila del DataGridView cuya insercion arrojo una excepcion                       

                        }
                        codInteno++;
                    }

                }

                Cursor.Current = Cursors.Default;
                t.Commit();

                MessageBox.Show("Se migraron exitosamente " + count + " registros", "Atencion", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception EX)
            {

                MessageBox.Show("Se produjo un error al conectarse a la base de datos\nError: " + EX.Message, "Atencion", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            conn.Close();
        }

        /// <summary>
        /// Evento de cambio de valor en el combobox de las tablas
        /// </summary>
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (comboBox1.SelectedIndex == 16)
            {
                //Transacciones trans = new Transacciones();
                //trans.ShowDialog();
            }
            else button1.Enabled = true;
        }
        //Metodo para el calculo de precio financiero
        /// <summary>
        /// Metodo de calculo del precio financiero
        /// </summary>
        public double preciofinanciero(double costo, double utilidad)
        {
            return (costo / (100 - utilidad)) * 100;
        }
        //Metodo para asignacion de valor de porc utilidad (validacion)
        public double porcUtilidad(double porc)
        {
            if (porc == 100) return 0;
            else return porc;
            
        }
        //Eventos de cambio en radio button para seleccion de tablas de administrativo y nomina
        /// <summary>
        /// Evento de cambio de seleccion del radio button que determina el sistema (admin o nomina)
        /// </summary>
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                var db = DBConn.Instance;
                var c = db.Collection<payroll>();
                comboBox1.DataSource = c.Find(Query.All()).ToList();
                comboBox1.DisplayMember = "desc";
                comboBox1.ValueMember = "id";
            }
        }

        /// <summary>
        /// Evento de cambio de seleccion del radio button que determina el sistema (admin o nomina)
        /// </summary>
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                var db = DBConn.Instance;
                var c = db.Collection<admin>();
                comboBox1.DataSource = c.Find(Query.All()).ToList();
                comboBox1.DisplayMember = "desc";
                comboBox1.ValueMember = "id";
            }
        }
        //Boton rojo !!! para recorrer el grid y posicionarte sobre un error
        /// <summary>
        /// Evento click del boton que te localiza en un reglon del grid con error (pintado de rojo)
        /// </summary>
        private void button3_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.DefaultCellStyle.BackColor == Color.Red)
                    dataGridView1.CurrentCell = dataGridView1.Rows[row.Index].Cells[0];
            }
        }
        #endregion
    }

    #region Clases
    public class Tipos
    {
        public String tipo { get; set; }
        public String codigo { get; set; }
        /// <summary>
        /// Constructor
        /// </summary>
        public Tipos()
        {

        }
        /// <summary>
        /// Sobrecarga
        /// </summary>
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

        /// <summary>
        /// Constructor
        /// </summary>
        public admin()
        {

        }

        /// <summary>
        /// Sobrecarga del constructor
        /// </summary>
        public admin(string descri, int cod)
        {
            this.desc = descri;
            this.Id = cod;
        }

    }

    public class adminA2
    {
        public String desc { get; set; }
        public int Id { get; set; }

        /// <summary>
        /// Contructor
        /// </summary>
        public adminA2()
        {

        }

        /// <summary>
        /// Sobrecarga del constructor
        /// </summary>
        public adminA2(string descri, int cod)
        {
            this.desc = descri;
            this.Id = cod;
        }

    }

    public class adminSaint
    {
        public String desc { get; set; }
        public int Id { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public adminSaint()
        {

        }

        /// <summary>
        /// Sobrecarga
        /// </summary>
        public adminSaint(string descri, int cod)
        {
            this.desc = descri;
            this.Id = cod;
        }

    }

    public class payroll
    {
        public String desc { get; set; }
        public int id { get; set; }
        /// <summary>
        /// Constructor
        /// </summary>
        public payroll()
        {

        }
        /// <summary>
        /// Sobrecarga del constructor
        /// </summary>
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
        /// <summary>
        /// Constructor
        /// </summary>
        public Errores()
        {

        }
        /// <summary>
        /// Sobrecarga del constructor
        /// </summary>
        public Errores(string Descs, string codigos)
        {
            this.Desc = Descs;
            this.codigo = codigos;
        }
    }

    public class BD
    {
        public string desc { get; set; }
        public int id { get; set; }
        /// <summary>
        /// Constructor
        /// </summary>
        public BD()
        {

        }
        /// <summary>
        /// Sobrecarga
        /// </summary>
        public BD(string Desc, int Id)
        {
            this.desc = Desc;
            this.id = Id;
        }
    }
    #endregion
}


