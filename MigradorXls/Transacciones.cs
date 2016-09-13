using Npgsql;
using NpgsqlTypes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MigradorXls
{
    public partial class Transacciones : Form
    {
        string file;
        string sql;
        string PRUEBA;
        int count;

        private DataSet DtSet = new DataSet();
        System.Data.OleDb.OleDbConnection Myconnetion;
        System.Data.OleDb.OleDbDataAdapter MyCommand;
        //string de conexion con las credenciales de postgreSQL
        string connectionString/*= @"Host=192.168.1.254;port=5432;Database=ROYALSDB;User ID=postgres;Password=TACA8tilo"*/;
        public Transacciones()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Evento que se acciona cuando el valor del combobox cambia
        /// </summary>
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            button1.Enabled = true;
        }

        /// <summary>
        /// Metodo del boton seleccionar que carga la data en la vista
        /// </summary>
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

                    firstSheetName = dbSchema.Rows[1]["TABLE_NAME"].ToString();
                    MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [" + firstSheetName + "]", Myconnetion);
                    MyCommand.TableMappings.Add("Table", "Testtable2");

                    //Vaciado del DataSet
                    //DtSet.Reset();
                    //Llenado del DataSet con resultado de comando ejecutado al .xls
                    MyCommand.Fill(DtSet);
                    //Asignacion del DataSet como origen de datos del DataGridView 
                    dataGridView2.DataSource = DtSet.Tables[1];

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

        /// <summary>
        /// Metodo del boton de migrar que realiza el proceso de ajuste de transacciones
        /// </summary>
        private void button2_Click(object sender, EventArgs e)
        {


            //Abriendo la coneccion con npgsql
            connectionString = @"Host=" + Globals.Host + ";port=" + Globals.port + ";Database=" + Globals.DB + ";User ID=" + Globals.usuario + ";Password=" + Globals.pass + ";";
            NpgsqlConnection conn = new NpgsqlConnection(connectionString);

            conn.Open();
            //NpgsqlTransaction t = conn.BeginTransaction();
            //Recorriendo el Datagridview e insertando cada valor
            count = 0;



            foreach (DataGridViewRow ROW in dataGridView1.Rows)
            {
                try
                {

                    switch (comboBox1.SelectedIndex)
                    {
                        case 0:
                            {

                                callbackInsertCargo(conn, ROW);
                                break;
                            }
                        case 1:
                            {

                                callbackInsertAjustePrecio(conn, ROW);
                                break;
                            }
                    }

                }

                catch (Exception ex)
                {
                    //Mensaje de error en la insercion de datos
                    MessageBox.Show("Hubo un Error en la insercion de datos. Excepcion: " + ex.Message.ToString());
                    // Cambio de color de la fila del DataGridView cuya insercion arrojo una excepcion                       
                    ROW.DefaultCellStyle.BackColor = Color.Red;
                }
            }


            //t.Commit();
            conn.Close();

            MessageBox.Show(count + " Filas se almacenaron correctamente", "Atencion", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Metodo que realiza el cargo
        /// </summary>
        private void callbackInsertCargo(NpgsqlConnection conn, DataGridViewRow ROW)
        {
            double total = 0;
            if (ROW.Cells["org_hijo"].Value != null)
            {
                //insercion de ajuste para el cargo
                sql = @"INSERT INTO admin.int_ajuste_precio(org_hijo,cod_motivo,descri,cod_autoriza,
                    nomb_autoriza,cod_persona, nomb_persona,cantidad_item,total_precio,total_utilidad, 
                    reg_usu_cc, reg_estatus, nro_items, doc_control, migrado) VALUES(@org_hijo,@cod_motivo,@descri,@cod_autoriza,
                    @nomb_autoriza,@cod_persona, @nomb_persona,@cantidad_item,@total_precio,@total_utilidad, 
                    @reg_usu_cc, @reg_estatus, @nro_items, @doc_control, @migrado)";
                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn);
                dbcmd.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
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

                dbcmd.Parameters[0].Value = Globals.org;
                dbcmd.Parameters[1].Value = ROW.Cells["cod_motivo"].Value;
                dbcmd.Parameters[2].Value = "CARGA INICIAL DE INVENTARIO";
                dbcmd.Parameters[3].Value = ROW.Cells["cod_autoriza"].Value;
                dbcmd.Parameters[4].Value = ROW.Cells["nomb_autoriza"].Value;
                dbcmd.Parameters[5].Value = ROW.Cells["cod_persona"].Value;
                dbcmd.Parameters[6].Value = ROW.Cells["nomb_persona"].Value;
                dbcmd.Parameters[7].Value = ROW.Cells["nro_items"].Value;
                dbcmd.Parameters[8].Value = ROW.Cells["total"].Value;
                dbcmd.Parameters[9].Value = "0";
                dbcmd.Parameters[10].Value = "INNOVA";
                dbcmd.Parameters[11].Value = ROW.Cells["reg_estatus"].Value;
                dbcmd.Parameters[12].Value = ROW.Cells["nro_items"].Value;
                dbcmd.Parameters[13].Value = ROW.Cells["doc_control"].Value;
                dbcmd.Parameters[14].Value = true;

                count += dbcmd.ExecuteNonQuery();

                sql = @"SELECT doc from admin.int_ajuste_precio order by fecha_reg desc";
                dbcmd = new NpgsqlCommand(sql, conn);

                string reader = dbcmd.ExecuteScalar().ToString();
                NpgsqlTransaction t = conn.BeginTransaction();
                //Insercion del detalle del ajuste
                foreach (DataGridViewRow ROW2 in dataGridView2.Rows)
                {
                    if (ROW2.Cells["cod_articulo"].Value != null)
                    {

                        sql = @"INSERT INTO admin.int_ajuste_precio_det(org_hijo,doc,cod_alterno,cod_articulo,
                        costo,costo_promedio,fecha,tipo_ajuste,item) VALUES(@org_hijo,@doc,
                        @cod_alterno,@cod_articulo,@costo,@costo_promedio,@fecha,@tipo_ajuste,@item)";
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
                        //dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Integer));

                        dbcmd.Prepare();

                        dbcmd.Parameters[0].Value = ROW.Cells["org_hijo"].Value;
                        dbcmd.Parameters[1].Value = Convert.ToInt64(reader);
                        dbcmd.Parameters[2].Value = ROW2.Cells["cod_alterno"].Value.ToString().Replace(" ", string.Empty);
                        dbcmd.Parameters[3].Value = ROW2.Cells["cod_articulo"].Value.ToString().Replace(" ", string.Empty);
                        dbcmd.Parameters[4].Value = ROW2.Cells["costo"].Value;
                        dbcmd.Parameters[5].Value = ROW2.Cells["costo_promedio"].Value;
                        dbcmd.Parameters[6].Value = ExtractDate(ROW.Cells["fecha"].Value.ToString());
                        dbcmd.Parameters[7].Value = ROW2.Cells["tipo_ajuste"].Value;
                        dbcmd.Parameters[8].Value = ROW2.Cells["item"].Value;
                        //dbcmd.Parameters[9].Value = 1;

                        PRUEBA = ROW2.Cells["cod_alterno"].Value.ToString().Replace(" ", string.Empty);

                        count += dbcmd.ExecuteNonQuery();
                        total += (Convert.ToDouble(ROW2.Cells["costo"].Value) * Convert.ToDouble(ROW2.Cells["cantidad"].Value));
                    }
                }

                t.Commit();


                //insercion del cargo
                sql = @"INSERT INTO admin.int_cargo(org_hijo,cod_terminal,doc_num,tipo_opera,
                        descri,descorta,cod_dep,cod_autoriza,cod_persona,nomb_autoriza,nomb_persona,fecha,
                        cod_motivo, motivo, total, reg_usu_cc, reg_usu_cu, reg_estatus, cod_moneda, factor,
                        observacion,doc_control, nro_items, cod_ajuste_precio, migrado) VALUES(@org_hijo,
                        @cod_terminal,@doc_num,@tipo_opera,@descri,@descorta,@cod_dep,@cod_autoriza,@cod_persona,
                        @nomb_autoriza,@nomb_persona,@fecha,@cod_motivo, @motivo, @total, @reg_usu_cc, @reg_usu_cu, 
                        @reg_estatus, @cod_moneda, @factor,@observacion,@doc_control,@nro_items, @cod_ajuste_precio, @migrado)";
                dbcmd = new NpgsqlCommand(sql, conn);
                dbcmd.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@cod_terminal", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@doc_num", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@tipo_opera", NpgsqlDbType.Integer));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@descorta", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@cod_dep", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@cod_autoriza", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@cod_persona", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@nomb_autoriza", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@nomb_persona", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@fecha", NpgsqlDbType.Date));
                dbcmd.Parameters.Add(new NpgsqlParameter("@cod_motivo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@motivo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@total", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cu", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@reg_estatus", NpgsqlDbType.Integer));
                dbcmd.Parameters.Add(new NpgsqlParameter("@cod_moneda", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@factor", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@observacion", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@doc_control", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@nro_items", NpgsqlDbType.Integer));
                dbcmd.Parameters.Add(new NpgsqlParameter("@cod_ajuste_precio", NpgsqlDbType.Bigint));
                dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));

                dbcmd.Prepare();

                dbcmd.Parameters[0].Value = ROW.Cells["org_hijo"].Value;
                dbcmd.Parameters[1].Value = ROW.Cells["cod_terminal"].Value;
                dbcmd.Parameters[2].Value = ROW.Cells["doc_num"].Value;
                dbcmd.Parameters[3].Value = ROW.Cells["tipo_opera"].Value;
                dbcmd.Parameters[4].Value = ROW.Cells["descri"].Value;
                dbcmd.Parameters[5].Value = ROW.Cells["descorta"].Value;
                dbcmd.Parameters[6].Value = ROW.Cells["cod_dep"].Value;
                dbcmd.Parameters[7].Value = ROW.Cells["cod_autoriza"].Value;
                dbcmd.Parameters[8].Value = ROW.Cells["cod_persona"].Value;
                dbcmd.Parameters[9].Value = ROW.Cells["nomb_autoriza"].Value;
                dbcmd.Parameters[10].Value = ROW.Cells["nomb_persona"].Value;
                dbcmd.Parameters[11].Value = ExtractDate(ROW.Cells["fecha"].Value.ToString());
                dbcmd.Parameters[12].Value = ROW.Cells["cod_motivo"].Value;
                dbcmd.Parameters[13].Value = ROW.Cells["motivo"].Value;
                dbcmd.Parameters[14].Value = total;
                dbcmd.Parameters[15].Value = "INNOVA";
                dbcmd.Parameters[16].Value = "INNOVA";
                dbcmd.Parameters[17].Value = ROW.Cells["reg_estatus"].Value;
                dbcmd.Parameters[18].Value = ROW.Cells["cod_moneda"].Value;
                dbcmd.Parameters[19].Value = ROW.Cells["factor"].Value;
                dbcmd.Parameters[20].Value = ROW.Cells["observacion"].Value;
                dbcmd.Parameters[21].Value = ROW.Cells["doc_control"].Value;
                dbcmd.Parameters[22].Value = ROW.Cells["nro_items"].Value;
                dbcmd.Parameters[23].Value = reader;
                dbcmd.Parameters[24].Value = true;

                count += dbcmd.ExecuteNonQuery();
                sql = @"SELECT doc from admin.int_cargo order by fecha_reg desc";
                dbcmd = new NpgsqlCommand(sql, conn);

                reader = dbcmd.ExecuteScalar().ToString();
                t = conn.BeginTransaction();
                foreach (DataGridViewRow ROW2 in dataGridView2.Rows)
                {

                    if (ROW2.Cells["org_hijo"].Value != null)
                    {
                        sql = @"INSERT INTO admin.int_cargo_det(org_hijo, doc, item, cod_alterno,
                        cod_articulo, cod_innova, descri, cantidad, existencia, existencia_anterior,
                        costo_anterior, costo_promedio_ant, costo_promedio, cod_dep, costo, total, 
                        precio_utilidad, descorta, tipo_opera, reg_estatus, tipo_ajuste, observacion,
                        tipo_documento) VALUES(@org_hijo, @doc, @item, @cod_alterno,
                        @cod_articulo, @cod_innova, @descri, @cantidad, @existencia, @existencia_anterior,
                        @costo_anterior, @costo_promedio_ant, @costo_promedio, @cod_dep, @costo, @total, 
                        @precio_utilidad, @descorta, @tipo_opera, @reg_estatus, @tipo_ajuste, @observacion,
                        @tipo_documento)";
                        dbcmd = new NpgsqlCommand(sql, conn);
                        dbcmd.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
                        dbcmd.Parameters.Add(new NpgsqlParameter("@doc", NpgsqlDbType.Bigint));
                        dbcmd.Parameters.Add(new NpgsqlParameter("@item", NpgsqlDbType.Integer));
                        dbcmd.Parameters.Add(new NpgsqlParameter("@cod_alterno", NpgsqlDbType.Varchar));
                        dbcmd.Parameters.Add(new NpgsqlParameter("@cod_articulo", NpgsqlDbType.Varchar));
                        dbcmd.Parameters.Add(new NpgsqlParameter("@cod_innova", NpgsqlDbType.Varchar));
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
                        dbcmd.Parameters.Add(new NpgsqlParameter("@observacion", NpgsqlDbType.Varchar));
                        dbcmd.Parameters.Add(new NpgsqlParameter("@tipo_documento", NpgsqlDbType.Integer));
                        //dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Integer));

                        dbcmd.Prepare();

                        dbcmd.Parameters[0].Value = ROW2.Cells["org_hijo"].Value;
                        dbcmd.Parameters[1].Value = Convert.ToInt64(reader);
                        dbcmd.Parameters[2].Value = ROW2.Cells["item"].Value;
                        dbcmd.Parameters[3].Value = ROW2.Cells["cod_alterno"].Value.ToString().Replace(" ", string.Empty);
                        dbcmd.Parameters[4].Value = ROW2.Cells["cod_articulo"].Value.ToString().Replace(" ", string.Empty);
                        dbcmd.Parameters[5].Value = ROW2.Cells["cod_innova"].Value;
                        dbcmd.Parameters[6].Value = ROW2.Cells["descri"].Value;
                        dbcmd.Parameters[7].Value = ROW2.Cells["cantidad"].Value;
                        dbcmd.Parameters[8].Value = ROW2.Cells["existencia"].Value;
                        dbcmd.Parameters[9].Value = ROW2.Cells["existencia_anterior"].Value;
                        dbcmd.Parameters[10].Value = ROW2.Cells["costo_anterior"].Value;
                        dbcmd.Parameters[11].Value = ROW2.Cells["costo_promedio_ant"].Value;
                        dbcmd.Parameters[12].Value = ROW2.Cells["costo_promedio"].Value;
                        dbcmd.Parameters[13].Value = ROW2.Cells["cod_dep"].Value;
                        dbcmd.Parameters[14].Value = ROW2.Cells["costo"].Value;
                        dbcmd.Parameters[15].Value = ROW2.Cells["total"].Value;
                        dbcmd.Parameters[16].Value = convertBoolean(ROW2.Cells["precio_utilidad"].Value);
                        dbcmd.Parameters[17].Value = ROW2.Cells["descorta"].Value;
                        dbcmd.Parameters[18].Value = ROW2.Cells["tipo_opera"].Value;
                        dbcmd.Parameters[19].Value = ROW2.Cells["reg_estatus"].Value;
                        dbcmd.Parameters[20].Value = ROW2.Cells["tipo_ajuste"].Value;
                        dbcmd.Parameters[21].Value = ROW2.Cells["observacion"].Value;
                        dbcmd.Parameters[22].Value = ROW2.Cells["tipo_documento"].Value;
                        //dbcmd.Parameters[23].Value = 1;
                        PRUEBA = ROW2.Cells["cod_alterno"].Value.ToString().Replace(" ", string.Empty);
                        count += dbcmd.ExecuteNonQuery();



                    }
                }
                t.Commit();

            }
        }
        /// <summary>
        /// Metodo que realiza el ajuste
        /// </summary>
        private void callbackInsertAjustePrecio(NpgsqlConnection conn, DataGridViewRow ROW)
        {
            if (ROW.Cells["org_hijo"].Value != null)
            {
                //insercion de ajuste para el cargo
                sql = @"INSERT INTO admin.int_ajuste_precio(org_hijo,cod_motivo,descri,cod_autoriza,
                    nomb_autoriza,cod_persona, nomb_persona,cantidad_item,total_precio,total_utilidad, 
                    reg_usu_cc, reg_estatus, nro_items, doc_control, migrado) VALUES(@org_hijo,@cod_motivo,@descri,@cod_autoriza,
                    @nomb_autoriza,@cod_persona, @nomb_persona,@cantidad_item,@total_precio,@total_utilidad, 
                    @reg_usu_cc, @reg_estatus, @nro_items, @doc_control, @migrado)";
                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn);
                dbcmd.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
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
                dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Integer));

                dbcmd.Prepare();

                dbcmd.Parameters[0].Value = "INNOVA01";
                dbcmd.Parameters[1].Value = ROW.Cells["cod_motivo"].Value;
                dbcmd.Parameters[2].Value = ROW.Cells["descri"].Value;
                dbcmd.Parameters[3].Value = ROW.Cells["cod_autoriza"].Value;
                dbcmd.Parameters[4].Value = ROW.Cells["nomb_autoriza"].Value;
                dbcmd.Parameters[5].Value = ROW.Cells["cod_persona"].Value;
                dbcmd.Parameters[6].Value = ROW.Cells["nomb_persona"].Value;
                dbcmd.Parameters[7].Value = ROW.Cells["cantidad_item"].Value;
                dbcmd.Parameters[8].Value = ROW.Cells["total_precio"].Value;
                dbcmd.Parameters[9].Value = ROW.Cells["total_utilidad"].Value;
                dbcmd.Parameters[10].Value = "INNOVA";
                dbcmd.Parameters[11].Value = ROW.Cells["reg_estatus"].Value;
                dbcmd.Parameters[12].Value = ROW.Cells["nro_items"].Value;
                dbcmd.Parameters[13].Value = ROW.Cells["doc_control"].Value;
                dbcmd.Parameters[14].Value = 1;

                count += dbcmd.ExecuteNonQuery();

                sql = @"SELECT doc from admin.int_ajuste_precio order by fecha_reg desc";
                dbcmd = new NpgsqlCommand(sql, conn);

                string reader = dbcmd.ExecuteScalar().ToString();
                NpgsqlTransaction t = conn.BeginTransaction();
                //Insercion del detalle del ajuste
                foreach (DataGridViewRow ROW2 in dataGridView2.Rows)
                {
                    if (ROW2.Cells["cod_articulo"].Value != null)
                    {

                        sql = @"INSERT INTO admin.int_ajuste_precio_det(org_hijo,doc,cod_alterno,cod_articulo,
                        costo,costo_promedio,fecha,tipo_ajuste,item, migrado) VALUES(@org_hijo,@doc,
                        @cod_alterno,@cod_articulo,@costo,@costo_promedio,@fecha,@tipo_ajuste,@item, @migrado)";
                        dbcmd = new NpgsqlCommand(sql, conn, t);
                        dbcmd.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
                        dbcmd.Parameters.Add(new NpgsqlParameter("@doc", NpgsqlDbType.Bigint));
                        dbcmd.Parameters.Add(new NpgsqlParameter("@cod_alterno", NpgsqlDbType.Varchar));
                        dbcmd.Parameters.Add(new NpgsqlParameter("@cod_articulo", NpgsqlDbType.Varchar));
                        dbcmd.Parameters.Add(new NpgsqlParameter("@costo", NpgsqlDbType.Double));
                        dbcmd.Parameters.Add(new NpgsqlParameter("@costo_promedio", NpgsqlDbType.Double));
                        dbcmd.Parameters.Add(new NpgsqlParameter("@fecha", NpgsqlDbType.Date));
                        dbcmd.Parameters.Add(new NpgsqlParameter("@tipo_ajuste", NpgsqlDbType.Integer));
                        dbcmd.Parameters.Add(new NpgsqlParameter("@item", NpgsqlDbType.Integer));
                        dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Integer));

                        dbcmd.Prepare();

                        dbcmd.Parameters[0].Value = "INNOVA01";
                        dbcmd.Parameters[1].Value = Convert.ToInt64(reader);
                        dbcmd.Parameters[2].Value = ROW2.Cells["cod_alterno"].Value.ToString().Replace(" ", string.Empty);
                        dbcmd.Parameters[3].Value = ROW2.Cells["cod_articulo"].Value.ToString().Replace(" ", string.Empty);
                        dbcmd.Parameters[4].Value = ROW2.Cells["costo"].Value;
                        dbcmd.Parameters[5].Value = ROW2.Cells["costo_promedio"].Value;
                        dbcmd.Parameters[6].Value = ExtractDate(ROW.Cells["fecha"].Value.ToString());
                        dbcmd.Parameters[7].Value = ROW2.Cells["tipo_ajuste"].Value;
                        dbcmd.Parameters[8].Value = ROW2.Cells["item"].Value;
                        dbcmd.Parameters[9].Value = 1;

                        count += dbcmd.ExecuteNonQuery();
                    }
                }
                t.Commit();
            }
        }

        /// <summary>
        /// Metodo que regresa un objeto tipo date
        /// </summary>
        public static DateTime ExtractDate(string myDate)
        {

            DateTime dt = DateTime.ParseExact(myDate, "M/d/yyyy", CultureInfo.InvariantCulture);

            return dt.Date;
        }
        /// <summary>
        /// Metodo que toma un string y segun su valor retorna un valor booleano
        /// </summary>
        public bool convertBoolean(object obj)
        {
            string text = Convert.ToString(obj);
            if (text.Equals("t", StringComparison.OrdinalIgnoreCase) || text.Equals("true", StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }
    }
}
