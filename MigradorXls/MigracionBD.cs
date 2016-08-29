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
using System.IO;
using System.Collections;

namespace MigradorXls
{
    public partial class MigracionBD : Form
    {
        #region Declaraciones
        string sql;
        string file;
        string seleccion = "";
        string connectionString = @"Host=" + Globals.Host + ";port=" + Globals.port + ";Database=" + Globals.DB + ";User ID=" + Globals.usuario + ";Password=" + Globals.pass + ";";
        int codInterno = 0;
        int count = 0;
        bool act = true;
        Exception c = new Exception();
        DataSet TAB = new DataSet();
        Dictionary<string, DateTime?> adelantos = new Dictionary<string, DateTime?>();
        DataConvert dt = new DataConvert();
        List<string> RIF = new List<string>();
        public MigracionBD()
        {
            InitializeComponent();

            var db = DBConn.Instance;
            var c = db.Collection<Config>();
            int cn = c.Count();
            Globals.ServidorSaint = c.FindById(cn).descServSaint;
            Globals.NombBDSaint = c.FindById(cn).BDSaint;
            var z = db.Collection<adminA2>();
            List<adminA2> ad = z.Find(Query.All()).ToList();
            comboBox2.DataSource = ad;
            comboBox2.DisplayMember = "desc";
            comboBox2.ValueMember = "id";
        }
        #endregion
        
        #region Migraciones

        #region A2(DBIsam)

        #region Zonas
        private void A2Zonas(OdbcConnection objODBCCon)
        {
            string oString = "Select FZ_CODIGO, FZ_DESCRIPCION, FZ_STATUS from Szonas";
            OdbcDataAdapter comm = new OdbcDataAdapter(oString, objODBCCon);
            comm.Fill(TAB);
            
        }

        private void callbackInsertZona(NpgsqlConnection conn, DataGridViewRow ROW) 
        {
             
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
                dbcmd.Parameters[2].Value = ROW.Cells["FZ_CODIGO"].Value;
                dbcmd.Parameters[3].Value = ROW.Cells["FZ_DESCRIPCION"].Value;
                dbcmd.Parameters[4].Value = ROW.Cells["FZ_DESCRIPCION"].Value;
                dbcmd.Parameters[5].Value = 0;
                dbcmd.Parameters[6].Value = 0;
                dbcmd.Parameters[7].Value = 0;
                dbcmd.Parameters[8].Value = "INNOVA";
                dbcmd.Parameters[9].Value = "INNOVA";
                dbcmd.Parameters[10].Value = 1;
                dbcmd.Parameters[11].Value = ROW.Cells["FZ_STATUS"].Value;
                dbcmd.Parameters[12].Value = true;

                count +=dbcmd.ExecuteNonQuery();
            
        }
        #endregion

        #region Moneda
        private void A2Moneda(OdbcConnection objODBCCon)
        {
            string oString = "Select FM_CODE, FM_DESCRIPCION, FM_STATUS, FM_SIMBOLO, FM_FACTOR from Smoneda";
            OdbcDataAdapter comm = new OdbcDataAdapter(oString, objODBCCon);
            comm.Fill(TAB);

        }


        private void callbackInsertMoneda(NpgsqlConnection conn, DataGridViewRow ROW)
        {
            
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
            dbcmd.Parameters[2].Value = ROW.Cells["FM_CODE"].Value;
            dbcmd.Parameters[3].Value = ROW.Cells["FM_DESCRIPCION"].Value;
            dbcmd.Parameters[4].Value = ROW.Cells["FM_DESCRIPCION"].Value;
            dbcmd.Parameters[5].Value = ROW.Cells["FM_SIMBOLO"].Value;
            dbcmd.Parameters[6].Value = ROW.Cells["FM_FACTOR"].Value;
            dbcmd.Parameters[7].Value = 1;
            dbcmd.Parameters[8].Value = "INNOVA";
            dbcmd.Parameters[9].Value = "INNOVA";
            dbcmd.Parameters[10].Value = 1;
            dbcmd.Parameters[11].Value = ROW.Cells["FM_STATUS"].Value;
            dbcmd.Parameters[12].Value = true;

            count += dbcmd.ExecuteNonQuery();
            

        }
        #endregion

        #region Vendedores
        private void A2Vendedores(OdbcConnection objODBCCon)
        {
            string oString = "Select FV_CODIGO, FV_DESCRIPCION, FV_DESCRIPCIONDETALLADA, FV_DIRECCION, FV_ZONAVENTA, FV_STATUS FROM Svendedores";
            OdbcDataAdapter comm = new OdbcDataAdapter(oString, objODBCCon);
            comm.Fill(TAB);
        }
        private void callbackInsertVendedores(NpgsqlConnection conn, DataGridViewRow ROW)
        {
            sql = @"INSERT INTO admin.org_talento(org_hijo,cod_interno,
                        codigo,cedula,descri,fecha_nac,reg_usu_cc,reg_usu_cu,
                        reg_estatus,disponible,tipo_cont,tipo_pers,cod_zona,migrado, 
                        descorta, direc1, observacion )
                        VALUES(@orgHijo , @codInterno, @codigo, @cedula, @descri, 
                        @fechaNac,@reg_usu_cc , @reg_usu_cu, @regEstatus, @disponible, @tipoCont, @tipoPers, 
                        @codZona , @migrado, @descorta,  @direc1, @observacion)";
            NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn);
            
            dbcmd.Parameters.Add(new NpgsqlParameter("@orghijo", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@codInterno", NpgsqlDbType.Bigint));
            dbcmd.Parameters.Add(new NpgsqlParameter("@codigo", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@cedula", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@fechaNac", NpgsqlDbType.Date));
            dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@reg_usu_cu", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
            dbcmd.Parameters.Add(new NpgsqlParameter("@disponible", NpgsqlDbType.Boolean));
            dbcmd.Parameters.Add(new NpgsqlParameter("@tipoCont", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@tipoPers", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@codZona", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));
            dbcmd.Parameters.Add(new NpgsqlParameter("@descorta", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@direc1", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@observacion", NpgsqlDbType.Varchar));


            dbcmd.Prepare();


            dbcmd.Parameters[0].Value = Globals.org;
            dbcmd.Parameters[1].Value = codInterno;
            dbcmd.Parameters[2].Value = ROW.Cells["FV_CODIGO"].Value.ToString().Replace(" ", string.Empty);
            dbcmd.Parameters[3].Value = ROW.Cells["FV_CODIGO"].Value.ToString().Replace(" ", string.Empty);
            dbcmd.Parameters[4].Value = ROW.Cells["FV_DESCRIPCIONDETALLADA"].Value;
            dbcmd.Parameters[5].Value = dt.ExtractDate("01/01/1991");
            dbcmd.Parameters[6].Value = "INNOVA"; 
            dbcmd.Parameters[7].Value = "INNOVA"; 
            dbcmd.Parameters[8].Value = 1; 
            dbcmd.Parameters[9].Value = ROW.Cells["FV_STATUS"].Value;
            dbcmd.Parameters[10].Value = "02.1";
            dbcmd.Parameters[11].Value = "03.1";
            dbcmd.Parameters[12].Value = ROW.Cells["FV_ZONAVENTA"].Value;
            dbcmd.Parameters[13].Value = true;
            dbcmd.Parameters[14].Value = ROW.Cells["FV_DESCRIPCION"].Value;
            dbcmd.Parameters[15].Value = ROW.Cells["FV_DIRECCION"].Value;
            dbcmd.Parameters[16].Value = "ESTA DATA FUE MIGRADA, POR FAVOR REVISAR TODOS LOS DATOS";
            
            count += dbcmd.ExecuteNonQuery();
        }
        #endregion

        #region Categorias
        private void A2Cat(OdbcConnection objODBCCon)
        {
            string oString = "Select FD_CODIGO, FD_DESCRIPCION, FD_STATUS from Scategoria";
            OdbcDataAdapter comm = new OdbcDataAdapter(oString, objODBCCon);
            comm.Fill(TAB);

        }

        private void callbackInsertCat(NpgsqlConnection conn, DataGridViewRow ROW)
        {

            sql = @"INSERT INTO admin.inv_cat(org_hijo,
                        cod_interno,cat_hijo,descri,reg_usu_cc,reg_usu_cu,
                        reg_estatus,disponible) 
                        VALUES(@org_hijo, @codInterno, @catHijo, @descri,
                        @reg_usu_cc, @reg_usu_cu, @regEstatus, @disponible)";


            NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn);

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
            dbcmd.Parameters[1].Value = codInterno;
            dbcmd.Parameters[2].Value = ROW.Cells["FD_CODIGO"].Value.ToString().Replace(" ", string.Empty);
            dbcmd.Parameters[3].Value = ROW.Cells["FD_DESCRIPCION"].Value;
            dbcmd.Parameters[4].Value = "INNOVA";
            dbcmd.Parameters[5].Value = "INNOVA";
            dbcmd.Parameters[6].Value = 1;
            dbcmd.Parameters[7].Value = ROW.Cells["FD_STATUS"].Value;
            count += dbcmd.ExecuteNonQuery();

        }
        #endregion

        #region Depositos
        private void A2Deposito(OdbcConnection objODBCCon)
        {
            string oString = "Select FDP_CODIGO, FDP_DESCRIPCION, FDP_STATUS from Sdepositos";
            OdbcDataAdapter comm = new OdbcDataAdapter(oString, objODBCCon);
            comm.Fill(TAB);

        }

        private void callbackInsertDeposito(NpgsqlConnection conn, DataGridViewRow ROW)
        {

            sql = @"INSERT INTO admin.inv_dep(org_hijo,
                        cod_interno,codigo,descri,maximo,minimo,espacio_mq,espacio_vol,espacio_uso,
                        reg_usu_cc,reg_usu_cu,reg_estatus,disponible) 
                        VALUES(@org_hijo , @codInterno, @codigo, @descri,@maximo,@minimo,
                        @espaciomq,@espaciovol,@espaciouso,@reg_usu_cc, @reg_usu_cu, 
                        @regEstatus, @disponible)";


            NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn);

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
            dbcmd.Parameters[1].Value = codInterno;
            dbcmd.Parameters[2].Value = ROW.Cells["FDP_CODIGO"].Value.ToString().Replace(" ", string.Empty);
            dbcmd.Parameters[3].Value = ROW.Cells["FDP_DESCRIPCION"].Value;
            dbcmd.Parameters[4].Value = 0;
            dbcmd.Parameters[5].Value = 0;
            dbcmd.Parameters[6].Value = 0;
            dbcmd.Parameters[7].Value = 0;
            dbcmd.Parameters[8].Value = 0;
            dbcmd.Parameters[9].Value = "INNOVA";
            dbcmd.Parameters[10].Value = "INNOVA";
            dbcmd.Parameters[11].Value = 1;
            dbcmd.Parameters[12].Value = ROW.Cells["FDP_STATUS"].Value;
            count += dbcmd.ExecuteNonQuery();

        }
        #endregion

        #region Clientes
        private void A2Clientes(OdbcConnection objODBCCon)
        {
            string oString = "Select FC_CODIGO, FC_DESCRIPCION, FC_DESCRIPCIONDETALLADA,FC_RIF,FC_CONTACTO, FC_TELEFONO, FC_EMAIL, FC_DIRECCION1, FC_RETENCION, FC_SALDO, FC_STATUS FROM Sclientes";
            OdbcDataAdapter comm = new OdbcDataAdapter(oString, objODBCCon);
            comm.Fill(TAB);
            dataGridView1.DataSource = TAB.Tables[0];
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                RIF.Add(row.Cells["FC_RIF"].Value.ToString());
            }
        }
        private void callbackInsertClientes(NpgsqlConnection conn, DataGridViewRow ROW)
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
                        @migrado, @esdatos, @esvip, @espronto, @observacion, @tipo_ret_iva, @telefono, @email,@nombPersona)";


            NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn);
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
            dbcmd.Parameters[1].Value = codInterno;
            dbcmd.Parameters[2].Value = ROW.Cells["FC_CODIGO"].Value.ToString().Replace(" ", string.Empty);
            dbcmd.Parameters[3].Value = ROW.Cells["FC_DESCRIPCION"].Value;
            dbcmd.Parameters[4].Value = "03.1";
            dbcmd.Parameters[5].Value = tipoContribuyente(ROW.Cells["FC_RIF"].Value.ToString());
            dbcmd.Parameters[6].Value = ROW.Cells["FC_RIF"].Value.ToString().Replace(" ", string.Empty);
            dbcmd.Parameters[7].Value = ROW.Cells["FC_DIRECCION1"].Value;
            dbcmd.Parameters[8].Value = 0;
            dbcmd.Parameters[9].Value = false;
            dbcmd.Parameters[10].Value = false;
            dbcmd.Parameters[11].Value = false;
            dbcmd.Parameters[12].Value = false;
            dbcmd.Parameters[13].Value = 0;
            dbcmd.Parameters[14].Value = 0;
            dbcmd.Parameters[15].Value = 0;
            dbcmd.Parameters[16].Value = 0;
            dbcmd.Parameters[17].Value = 0;
            dbcmd.Parameters[18].Value = 0;
            dbcmd.Parameters[19].Value = 0;
            dbcmd.Parameters[20].Value = 0;
            dbcmd.Parameters[21].Value = 0;
            dbcmd.Parameters[22].Value = dt.GetValue<double>(ROW.Cells["FC_SALDO"].Value);
            dbcmd.Parameters[23].Value = "INNOVA";
            dbcmd.Parameters[24].Value = "INNOVA";
            dbcmd.Parameters[25].Value = 1;
            dbcmd.Parameters[26].Value = ROW.Cells["FC_STATUS"].Value;
            dbcmd.Parameters[27].Value = true;
            dbcmd.Parameters[28].Value = dt.GetValue<double>(ROW.Cells["FC_RETENCION"].Value);
            dbcmd.Parameters[29].Value = 0;
            dbcmd.Parameters[30].Value = 0;
            dbcmd.Parameters[31].Value = 0;
            dbcmd.Parameters[32].Value = 0;
            dbcmd.Parameters[33].Value = 0;
            dbcmd.Parameters[34].Value = true;
            dbcmd.Parameters[35].Value = true;
            dbcmd.Parameters[36].Value = true;
            dbcmd.Parameters[37].Value = "ESTA DATA FUE MIGRADA, POR FAVOR VERIFICAR LOS DATOS";
            dbcmd.Parameters[38].Value = retencion(dt.GetValue<double>(ROW.Cells["FC_RETENCION"].Value).ToString());
            dbcmd.Parameters[39].Value = ROW.Cells["FC_TELEFONO"].Value;
            dbcmd.Parameters[40].Value = ROW.Cells["FC_EMAIL"].Value;
            dbcmd.Parameters[41].Value = ROW.Cells["FC_CONTACTO"].Value;
            count += dbcmd.ExecuteNonQuery();
        }
        #endregion

        #region Proveedores
        private void A2Prov(OdbcConnection objODBCCon)
        {
            string oString = "Select FP_CODIGO, FP_DESCRIPCION, FP_DESCRIPCIONDETALLADA,FP_TELEFONO,FP_EMAIL,FP_CONTACTO, FP_RIF, FP_DIRECCION1, FP_RETENCIONIVA, FP_SALDO, FP_STATUS FROM Sproveedor";
            OdbcDataAdapter comm = new OdbcDataAdapter(oString, objODBCCon);
            comm.Fill(TAB);
            dataGridView1.DataSource = TAB.Tables[0];
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                RIF.Add(row.Cells["FP_RIF"].Value.ToString());
            }
        }
        private void callbackInsertProv(NpgsqlConnection conn, DataGridViewRow ROW)
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
                        @migrado, @porcretiva, @Observacion, @tipo_ret_iva,@telefono,@email, @nombPersona)";


                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn);
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
                dbcmd.Parameters[1].Value = codInterno;
                dbcmd.Parameters[2].Value = ROW.Cells["FP_CODIGO"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[3].Value = ROW.Cells["FP_DESCRIPCION"].Value;
                dbcmd.Parameters[4].Value = "03.1";
                dbcmd.Parameters[5].Value = tipoContribuyente(ROW.Cells["FP_RIF"].Value.ToString());
                dbcmd.Parameters[6].Value = ROW.Cells["FP_RIF"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[7].Value = ROW.Cells["FP_DIRECCION1"].Value;
                dbcmd.Parameters[8].Value = 0;
                dbcmd.Parameters[9].Value = false;
                dbcmd.Parameters[10].Value = false;
                dbcmd.Parameters[11].Value = false;
                dbcmd.Parameters[12].Value = false;
                dbcmd.Parameters[13].Value = 0;
                dbcmd.Parameters[14].Value = 0;
                dbcmd.Parameters[15].Value = 0;
                dbcmd.Parameters[16].Value = 0;
                dbcmd.Parameters[17].Value = 0;
                dbcmd.Parameters[18].Value = 0;
                dbcmd.Parameters[19].Value = 0;
                dbcmd.Parameters[20].Value = 0;
                dbcmd.Parameters[21].Value = 0;
                dbcmd.Parameters[22].Value = dt.GetValue<double>(ROW.Cells["FP_SALDO"].Value);
                dbcmd.Parameters[23].Value = "INNOVA";
                dbcmd.Parameters[24].Value = "INNOVA";
                dbcmd.Parameters[25].Value = 1;
                dbcmd.Parameters[26].Value = ROW.Cells["FP_STATUS"].Value;
                dbcmd.Parameters[27].Value = true;
                dbcmd.Parameters[28].Value = dt.GetValue<double>(ROW.Cells["FP_RETENCIONIVA"].Value);
                dbcmd.Parameters[29].Value = "ESTA DATA FUE MIGRADA, POR FAVOR VERIFICAR LOS DATOS";
                dbcmd.Parameters[30].Value = retencion(dt.GetValue<double>(ROW.Cells["FP_RETENCIONIVA"].Value).ToString());
                dbcmd.Parameters[31].Value = ROW.Cells["FP_TELEFONO"].Value;
                dbcmd.Parameters[32].Value = ROW.Cells["FP_EMAIL"].Value;
                dbcmd.Parameters[33].Value = ROW.Cells["FP_CONTACTO"].Value;
            count += dbcmd.ExecuteNonQuery();
            
        }
        #endregion

        #region CXC
        private void A2CxC(OdbcConnection objODBCCon)
        {
            string oString = @"Select FCC_NUMERO , FCC_CODIGO, FCC_DESCRIPCIONMOV,FCC_FECHAEMISION,FCC_FECHAVENCIMIENTO, 
                            FCC_MONTODOCUMENTO, FCC_SALDODOCUMENTO,FCC_IMPUESTO1,FCC_IMPUESTO1PORCENT,FCC_MTOIMPUESTO1,FCC_IMPUESTO2,FCC_IMPUESTO2PORCENT,FCC_MTOIMPUESTO2,FCC_BASEIMPONIBLE,FCC_BASEIMPONIBLE2, FCC_TIPOTRANSACCION,
                            FCC_CONTROL,FCC_FECHARECEPCION, FCC_NROVENDEDOR,FCC_MACHINENAME  FROM Scuentasxcobrar 
                            WHERE (FCC_TIPOTRANSACCION=1 OR FCC_TIPOTRANSACCION=2 OR FCC_TIPOTRANSACCION=5 OR FCC_TIPOTRANSACCION=7 OR FCC_TIPOTRANSACCION=9)";
            OdbcDataAdapter comm = new OdbcDataAdapter(oString, objODBCCon);
            comm.Fill(TAB);
        }
        private void callbackInsertCxC(NpgsqlConnection conn, DataGridViewRow ROW)
        {
            sql = @"INSERT INTO admin.fin_cxc(org_hijo,doc_num,cod_cli,fecha_emi,fecha_ven,factor,
                      total,saldo,saldo_inicial,total_exento,doc_control,reg_usu_cc,
                      reg_estatus,cod_empleado , migrado, cod_impresorafiscal, descri, tipo_opera, debito, credito) 
                      VALUES(@org_hijo,@docNum,@codCli,@fechaEmi, @fechaVen, @factor, @total, @saldo, @saldoInicial, @totalEx, 
                      @doc_control,@reg_usu_cc, @reg_estatus, @codEmpleado, @migrado, @cod_impresorafiscal, @descri, @tipoOpera,
                        @debito, @credito)";
            NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn);
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
            dbcmd.Parameters[1].Value = ROW.Cells["FCC_NUMERO"].Value;
            dbcmd.Parameters[2].Value = ROW.Cells["FCC_CODIGO"].Value.ToString().Replace(" ", string.Empty);
            dbcmd.Parameters[3].Value = ROW.Cells["FCC_FECHAEMISION"].Value;
            dbcmd.Parameters[4].Value = ROW.Cells["FCC_FECHAVENCIMIENTO"].Value;
            dbcmd.Parameters[5].Value = 0;
            dbcmd.Parameters[6].Value = ROW.Cells["FCC_MONTODOCUMENTO"].Value;
            dbcmd.Parameters[7].Value = ROW.Cells["FCC_SALDODOCUMENTO"].Value;
            dbcmd.Parameters[8].Value = ROW.Cells["FCC_MONTODOCUMENTO"].Value;
            dbcmd.Parameters[9].Value = montoEx((double)ROW.Cells["FCC_MONTODOCUMENTO"].Value, ((double)ROW.Cells["FCC_MTOIMPUESTO1"].Value + (double)ROW.Cells["FCC_MTOIMPUESTO2"].Value), ((double)ROW.Cells["FCC_BASEIMPONIBLE"].Value - (double)ROW.Cells["FCC_BASEIMPONIBLE2"].Value));
            if(string.IsNullOrWhiteSpace(ROW.Cells["FCC_CONTROL"].Value.ToString())) dbcmd.Parameters[10].Value = ROW.Cells["FCC_NUMERO"].Value;
            else dbcmd.Parameters[10].Value = ROW.Cells["FCC_CONTROL"].Value;
            dbcmd.Parameters[11].Value = "INNOVA";
            dbcmd.Parameters[12].Value = 1;
            dbcmd.Parameters[13].Value = dt.GetValue<string>(ROW.Cells["FCC_NROVENDEDOR"].Value.ToString().Replace(" ", string.Empty));
            dbcmd.Parameters[14].Value = true;
            dbcmd.Parameters[15].Value = dt.GetValue<string>(ROW.Cells["FCC_MACHINENAME"].Value);
            dbcmd.Parameters[16].Value = ROW.Cells["FCC_DESCRIPCIONMOV"].Value;
            dbcmd.Parameters[17].Value = tipoOperaA2(ROW.Cells["FCC_TIPOTRANSACCION"].Value.ToString(),false);
            switch (ROW.Cells["FCC_TIPOTRANSACCION"].Value.ToString())
            {
                case "1":
                    {
                        dbcmd.Parameters[18].Value = 0;
                        dbcmd.Parameters[19].Value = ROW.Cells["FCC_MONTODOCUMENTO"].Value;
                        break;
                    }
                case "7":
                case "5":
                case "9":
                    {
                        dbcmd.Parameters[18].Value = ROW.Cells["FCC_MONTODOCUMENTO"].Value;
                        dbcmd.Parameters[19].Value = 0;
                        break;
                    }
                case "2":
                case "8":
                    {
                        dbcmd.Parameters[18].Value = 0;
                        dbcmd.Parameters[19].Value = ROW.Cells["FCC_MONTODOCUMENTO"].Value;
                        break;
                    }
                
            }

            count += dbcmd.ExecuteNonQuery();

            if ((double)ROW.Cells["FCC_IMPUESTO1"].Value > 0) callbackInsertcuentasImp(conn, ROW, 1, "FCC", "fin_cxc");
            if ((double)ROW.Cells["FCC_IMPUESTO2"].Value > 0) callbackInsertcuentasImp(conn, ROW, 2, "FCC", "fin_cxc");

        }

        private void callbackInsertcuentasImp(NpgsqlConnection conn, DataGridViewRow ROW, int z, string ad,string tabla)
        {
            sql = @"SELECT MAX(doc) FROM admin."+tabla;
            NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn);

            string reader = dbcmd.ExecuteScalar().ToString();

            sql = @"INSERT INTO admin."+tabla+"_imp(org_hijo,porcentaje,cod_impuesto,base, total, doc,reg_estatus,migrado) VALUES(@orgHijo , @porcentaje,@codImpuesto, @base , @total,@doc, @regEstatus, @migrado)";
            dbcmd = new NpgsqlCommand(sql, conn);
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
            dbcmd.Parameters[1].Value = ROW.Cells[ad+"_IMPUESTO" + z + ""].Value;
            dbcmd.Parameters[2].Value = Impuesto(Convert.ToDouble(ROW.Cells[ad + "_MTOIMPUESTO" + z + ""].Value), conn, Convert.ToDouble(ROW.Cells[ad + "_IMPUESTO"+z].Value),Convert.ToBoolean(ROW.Cells[ad+"_IMPUESTO"+z+"PORCENT"].Value));
            if(z==1)dbcmd.Parameters[3].Value = ROW.Cells[ad+"_BASEIMPONIBLE"].Value;
            else dbcmd.Parameters[3].Value = ROW.Cells[ad+"_BASEIMPONIBLE"+z].Value;
            dbcmd.Parameters[4].Value = ROW.Cells[ad+"_MONTODOCUMENTO"].Value;
            dbcmd.Parameters[5].Value = reader;
            dbcmd.Parameters[6].Value = 1;
            dbcmd.Parameters[7].Value = true;

            count += dbcmd.ExecuteNonQuery();
        }
        #endregion

        #region CxP
        private void A2CxP(OdbcConnection objODBCCon)
        {
            string oString = @"Select FCP_NUMERO , FCP_CODIGO, FCP_DESCRIPCIONMOV,FCP_FECHAEMISION,FCP_FECHAVENCIMIENTO, 
                            FCP_MONTODOCUMENTO, FCP_SALDODOCUMENTO,FCP_IMPUESTO1,FCP_IMPUESTO1PORCENT,FCP_MTOIMPUESTO1,FCP_IMPUESTO2,FCP_IMPUESTO2PORCENT,FCP_MTOIMPUESTO2,FCP_BASEIMPONIBLE,FCP_BASEIMPONIBLE2, 
                            FCP_TIPOTRANSACCION, FCP_CONTROL,FCP_FECHARECEPCION, FCP_NROVENDEDOR, FCP_MACHINENAME  FROM Scuentasxpagar 
                            WHERE (FCP_TIPOTRANSACCION=1 OR FCP_TIPOTRANSACCION=2 OR FCP_TIPOTRANSACCION=5 OR FCP_TIPOTRANSACCION=7 OR FCP_TIPOTRANSACCION=9) ";
            OdbcDataAdapter comm = new OdbcDataAdapter(oString, objODBCCon);
            comm.Fill(TAB);
        }
        private void callbackInsertCxP(NpgsqlConnection conn, DataGridViewRow ROW)
        {
            sql = @"INSERT INTO admin.fin_cxp(org_hijo,doc_num,cod_prov,fecha_emi,fecha_ven,factor,
                      total,saldo,saldo_inicial,total_exento,doc_control,reg_usu_cc,
                      reg_estatus, migrado, cod_impresorafiscal, descri, tipo_opera,debito,credito) 
                      VALUES(@org_hijo,@docNum,@codPro,
                      @fechaEmi, @fechaVen, @factor, @total, @saldo, @saldoInicial, @totalEx, 
                      @doc_control,@reg_usu_cc, @reg_estatus, @migrado, @cod_impresorafiscal, @descri, @tipoOpera, @debito, @credito)";
            NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn);
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
            dbcmd.Parameters[1].Value = ROW.Cells["FCP_NUMERO"].Value;
            dbcmd.Parameters[2].Value = ROW.Cells["FCP_CODIGO"].Value.ToString().Replace(" ", string.Empty);
            dbcmd.Parameters[3].Value = ROW.Cells["FCP_FECHAEMISION"].Value;
            dbcmd.Parameters[4].Value = ROW.Cells["FCP_FECHAVENCIMIENTO"].Value;
            dbcmd.Parameters[5].Value = 0;
            dbcmd.Parameters[6].Value = ROW.Cells["FCP_MONTODOCUMENTO"].Value;
            dbcmd.Parameters[7].Value = ROW.Cells["FCP_SALDODOCUMENTO"].Value;
            dbcmd.Parameters[8].Value = ROW.Cells["FCP_MONTODOCUMENTO"].Value;
            dbcmd.Parameters[9].Value = montoEx((double)ROW.Cells["FCP_MONTODOCUMENTO"].Value, ((double)ROW.Cells["FCP_MTOIMPUESTO1"].Value + (double)ROW.Cells["FCP_MTOIMPUESTO2"].Value), ((double)ROW.Cells["FCP_BASEIMPONIBLE"].Value - (double)ROW.Cells["FCP_BASEIMPONIBLE2"].Value));
            if (string.IsNullOrWhiteSpace(ROW.Cells["FCP_CONTROL"].Value.ToString())) dbcmd.Parameters[10].Value = ROW.Cells["FCP_NUMERO"].Value;
            else dbcmd.Parameters[10].Value = ROW.Cells["FCP_CONTROL"].Value;
            dbcmd.Parameters[11].Value = "INNOVA";
            dbcmd.Parameters[12].Value = 1;
            dbcmd.Parameters[13].Value = true;
            dbcmd.Parameters[14].Value = dt.GetValue<string>(ROW.Cells["FCP_MACHINENAME"].Value);
            dbcmd.Parameters[15].Value = ROW.Cells["FCP_DESCRIPCIONMOV"].Value;
            dbcmd.Parameters[16].Value = tipoOperaA2(ROW.Cells["FCP_TIPOTRANSACCION"].Value.ToString(), false);
            switch (ROW.Cells["FCP_TIPOTRANSACCION"].Value.ToString())
            {
                case "1":
                    {
                        dbcmd.Parameters[17].Value = 0;
                        dbcmd.Parameters[18].Value = ROW.Cells["FCP_MONTODOCUMENTO"].Value;
                        break;
                    }
                case "7":
                case "5":
                    {
                        dbcmd.Parameters[17].Value = ROW.Cells["FCP_MONTODOCUMENTO"].Value;
                        dbcmd.Parameters[18].Value = 0;
                        break;
                    }
                case "2":
                case "8":
                case "9":
                    {
                        dbcmd.Parameters[17].Value = 0;
                        dbcmd.Parameters[18].Value = ROW.Cells["FCP_MONTODOCUMENTO"].Value;
                        break;
                    }

            }

            count += dbcmd.ExecuteNonQuery();

            if ((double)ROW.Cells["FCP_MTOIMPUESTO1"].Value > 0) callbackInsertcuentasImp(conn, ROW, 1, "FCP", "fin_cxp");
            if ((double)ROW.Cells["FCP_MTOIMPUESTO2"].Value > 0) callbackInsertcuentasImp(conn, ROW, 2, "FCP", "fin_cxp");

        }
        #endregion

        #region Adelantos (Clientes)
        private void A2AdelantosCli(OdbcConnection objODBCCon)
        {
            string oString = @"Select  FCC_CODIGO, FCC_DESCRIPCIONMOV,FCC_FECHAEMISION, 
                            FCC_MONTODOCUMENTO  FROM Scuentasxcobrar 
                            WHERE (FCC_TIPOTRANSACCION=6)";
            OdbcDataAdapter comm = new OdbcDataAdapter(oString, objODBCCon);
            comm.Fill(TAB);
        }
        private void callbackInsertAdelantosCli(NpgsqlConnection conn, DataGridViewRow ROW)
        {
            NpgsqlCommand dbcmd = new NpgsqlCommand();
            if (!adelantos.ContainsKey(ROW.Cells["FCC_CODIGO"].Value.ToString().Replace(" ", string.Empty)))
            {
                adelantos[ROW.Cells["FCC_CODIGO"].Value.ToString().Replace(" ", string.Empty)] = dt.ExtractDate(ROW.Cells["FCC_FECHAEMISION"].Value.ToString());

                sql = @"INSERT INTO admin.fin_cli_adelanto(org_hijo,saldo,reg_usu_cc,reg_estatus, cli_hijo,
                     migrado) VALUES(@orgHijo ,
                    @saldo, @regusucc , @regEstatus,@clihijo, @migrado)";
                dbcmd = new NpgsqlCommand(sql, conn);
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
                dbcmd.Parameters[4].Value = ROW.Cells["FCC_CODIGO"].Value.ToString().Replace(" ",string.Empty);
                dbcmd.Parameters[5].Value = true;

                count += dbcmd.ExecuteNonQuery();

            }

            sql = @"INSERT INTO admin.fin_cli_ade_det(org_hijo,monto,observacion, cli_hijo,
                     fecha,migrado) VALUES(@orgHijo , @monto,
                    @observacion , @clihijo,@fecha, @migrado)";
            dbcmd = new NpgsqlCommand(sql, conn);
            dbcmd.Parameters.Add(new NpgsqlParameter("@orgHijo", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@monto", NpgsqlDbType.Double));
            dbcmd.Parameters.Add(new NpgsqlParameter("@observacion", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@clihijo", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@fecha", NpgsqlDbType.Date));
            dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));


            dbcmd.Prepare();

            dbcmd.Parameters[0].Value = Globals.org;
            dbcmd.Parameters[1].Value = ROW.Cells["FCC_MONTODOCUMENTO"].Value;
            dbcmd.Parameters[2].Value = ROW.Cells["FCC_DESCRIPCIONMOV"].Value;
            dbcmd.Parameters[3].Value = ROW.Cells["FCC_CODIGO"].Value.ToString().Replace(" ", string.Empty);
            dbcmd.Parameters[4].Value = dt.ExtractDate(ROW.Cells["FCC_FECHAEMISION"].Value.ToString());
            dbcmd.Parameters[5].Value = true;

            count += dbcmd.ExecuteNonQuery();


        }
        #endregion

        #region Adelantos (Proveedores)
        private void A2AdelantosProv(OdbcConnection objODBCCon)
        {
            string oString = @"Select  FCC_CODIGO, FCC_DESCRIPCIONMOV,FCC_FECHAEMISION, 
                            FCC_MONTODOCUMENTO  FROM Scuentasxpagar 
                            WHERE (FCC_TIPOTRANSACCION=6)";
            OdbcDataAdapter comm = new OdbcDataAdapter(oString, objODBCCon);
            comm.Fill(TAB);
        }
        private void callbackInsertA2AdelantosProv(NpgsqlConnection conn, DataGridViewRow ROW)
        {
            NpgsqlCommand dbcmd = new NpgsqlCommand();
            if (!adelantos.ContainsKey(ROW.Cells["FCC_CODIGO"].Value.ToString().Replace(" ", string.Empty)))
            {
                adelantos[ROW.Cells["FCC_CODIGO"].Value.ToString().Replace(" ", string.Empty)] = dt.ExtractDate(ROW.Cells["FCC_FECHAEMISION"].Value.ToString());

                sql = @"INSERT INTO admin.fin_prov_adelanto(org_hijo,saldo,reg_usu_cc,reg_estatus, prov_hijo,
                     migrado) VALUES(@orgHijo ,
                    @saldo, @regusucc , @regEstatus,@provhijo, @migrado)";
                dbcmd = new NpgsqlCommand(sql, conn);
                dbcmd.Parameters.Add(new NpgsqlParameter("@orgHijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@saldo", NpgsqlDbType.Double));
                dbcmd.Parameters.Add(new NpgsqlParameter("@regusucc", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@regEstatus", NpgsqlDbType.Integer));
                dbcmd.Parameters.Add(new NpgsqlParameter("@provhijo", NpgsqlDbType.Varchar));
                dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));


                dbcmd.Prepare();

                dbcmd.Parameters[0].Value = Globals.org;
                dbcmd.Parameters[1].Value = 0;
                dbcmd.Parameters[2].Value = "INNOVA";
                dbcmd.Parameters[3].Value = 1;
                dbcmd.Parameters[4].Value = ROW.Cells["FCP_CODIGO"].Value.ToString().Replace(" ", string.Empty);
                dbcmd.Parameters[5].Value = true;

                count += dbcmd.ExecuteNonQuery();

            }

            sql = @"INSERT INTO admin.fin_prov_ade_det(org_hijo,monto,observacion, prov_hijo,
                     fecha,migrado) VALUES(@orgHijo , @monto,
                    @observacion , @provhijo,@fecha, @migrado)";
            dbcmd = new NpgsqlCommand(sql, conn);
            dbcmd.Parameters.Add(new NpgsqlParameter("@orgHijo", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@monto", NpgsqlDbType.Double));
            dbcmd.Parameters.Add(new NpgsqlParameter("@observacion", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@provhijo", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@fecha", NpgsqlDbType.Date));
            dbcmd.Parameters.Add(new NpgsqlParameter("@migrado", NpgsqlDbType.Boolean));


            dbcmd.Prepare();

            dbcmd.Parameters[0].Value = Globals.org;
            dbcmd.Parameters[1].Value = ROW.Cells["FCP_MONTODOCUMENTO"].Value;
            dbcmd.Parameters[2].Value = ROW.Cells["FCP_DESCRIPCIONMOV"].Value;
            dbcmd.Parameters[3].Value = ROW.Cells["FCP_CODIGO"].Value.ToString().Replace(" ", string.Empty);
            dbcmd.Parameters[4].Value = dt.ExtractDate(ROW.Cells["FCP_FECHAEMISION"].Value.ToString());
            dbcmd.Parameters[5].Value = true;

            count += dbcmd.ExecuteNonQuery();
        }
        #endregion

        #region Banco
        private void A2Banco(OdbcConnection objODBCCon)
        {
            string oString = "Select FB_CODIGO, FB_DESCRIPCION, FB_DESCRIPCIONDETALLADA, FB_CONTACTO, FZ_STATUS from Szonas";
            OdbcDataAdapter comm = new OdbcDataAdapter(oString, objODBCCon);
            comm.Fill(TAB);
        }

        private void callbackInsertBanco(NpgsqlConnection conn, DataGridViewRow ROW)
        {

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
            dbcmd.Parameters[2].Value = ROW.Cells["FZ_CODIGO"].Value;
            dbcmd.Parameters[3].Value = ROW.Cells["FZ_DESCRIPCION"].Value;
            dbcmd.Parameters[4].Value = ROW.Cells["FZ_DESCRIPCION"].Value;
            dbcmd.Parameters[5].Value = 0;
            dbcmd.Parameters[6].Value = 0;
            dbcmd.Parameters[7].Value = 0;
            dbcmd.Parameters[8].Value = "INNOVA";
            dbcmd.Parameters[9].Value = "INNOVA";
            dbcmd.Parameters[10].Value = 1;
            dbcmd.Parameters[11].Value = ROW.Cells["FZ_STATUS"].Value;
            dbcmd.Parameters[12].Value = true;

            count += dbcmd.ExecuteNonQuery();

        }
        #endregion

        #endregion

        #region Saint(SQLServer)
        private void SaZonas(SqlConnection objODBCCon)
        {
            string oString = "Select FZ_CODIGO, FZ_DESCRIPCION, FZ_STATUS from Szonas";
            SqlDataAdapter comm = new SqlDataAdapter(oString, objODBCCon);
            comm.Fill(TAB);

        }

        #region Moneda
        private void SaMoneda(SqlConnection objODBCCon)
        {
            string oString = "Select CodMone, Descripcion, CodMone as Simbolo, Factor from SBMONE";
            SqlDataAdapter comm = new SqlDataAdapter(oString, objODBCCon);
            comm.Fill(TAB);
            dataGridView1.DataSource = TAB.Tables[0];

        }
        private void callbackInsertMonedaSaint(NpgsqlConnection conn, DataGridViewRow ROW)
        {
            
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
            dbcmd.Parameters[2].Value = ROW.Cells["CodMone"].Value;
            dbcmd.Parameters[3].Value = ROW.Cells["Descripcion"].Value;
            dbcmd.Parameters[4].Value = ROW.Cells["Descripcion"].Value;
            dbcmd.Parameters[5].Value = ROW.Cells["Simbolo"].Value;
            dbcmd.Parameters[6].Value = ROW.Cells["Factor"].Value;
            dbcmd.Parameters[7].Value = 1;
            dbcmd.Parameters[8].Value = "INNOVA";
            dbcmd.Parameters[9].Value = "INNOVA";
            dbcmd.Parameters[10].Value = 1;
            dbcmd.Parameters[11].Value = true;
            dbcmd.Parameters[12].Value = true;

            count += dbcmd.ExecuteNonQuery();
        }
        #endregion

        #region Clientes
        private void SaClientes(SqlConnection objODBCCon)
        {
            string oString = @"Select CodClie, Descrip, DescripExt,ID3,Represent, Telef, Email, Direc1, Descto, 
                                Saldo, FechaUV,MontoMax,MtoMaxCred,MontoUV,NumeroUV,FechaUP,MontoUP,NumeroUP,PagosA,RetenIva FROM SACLIE";
            SqlDataAdapter comm = new SqlDataAdapter(oString, objODBCCon);
            comm.Fill(TAB);
            dataGridView1.DataSource = TAB.Tables[0];
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                RIF.Add(row.Cells["ID3"].Value.ToString());
            }
        }
        private void callbackInsertClientesSaint(NpgsqlConnection conn, DataGridViewRow ROW)
        {
            sql = @"INSERT INTO admin.ven_cli(org_hijo,cod_interno,cli_hijo,descri,
                        tipo_cont,tipo_pers,porc_ret_iva,rif,direc1,monto_descuento,es_descuento,
                        es_exento,es_retencion,es_monto,monto_min,monto_max,monto_cred_max,monto_acum,
                        pri_vmonto,ult_vmonto,ult_vfecha,ult_vdoc,ult_pmonto,ult_pfecha,ult_pdoc,
                        pago_max,pago_adel,longitud, latitud, altitud,
                        pago_prom,monto_cred_min,saldo, reg_usu_cc,reg_usu_cu,reg_estatus,disponible, 
                        migrado,es_datos,es_vip,es_pronto, observacion, tipo_ret_iva, telefono, email, nomb_persona) 
                        VALUES(@org_hijo,@cod_interno,@cli_hijo,@descri,@tipo_cont,@tipo_pers,
                        @PorcRetIva,@rif,@direc1,@descuento,@es_descuento,
                        @es_exento,@es_retencion,@es_monto,@monto_min,@monto_max,@monto_cred_max,@monto_acum,
                        @pri_vmonto,@ult_vmonto,@ult_vfecha,@ult_vdoc,@ult_pmonto,@ult_pfecha,@ult_pdoc,
                        @pago_max,@pago_adel,@longitud,@latitud,@altitud,
                        @pago_prom,@montoCredMin,@saldo, @reg_usu_cc,@reg_usu_cu,@reg_estatus,@disponible, 
                        @migrado, @esdatos, @esvip, @espronto, @observacion, @tipo_ret_iva, @telefono, @email,@nombPersona)";


            NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn);
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
            dbcmd.Parameters.Add(new NpgsqlParameter("@ult_vfecha", NpgsqlDbType.Timestamp));
            dbcmd.Parameters.Add(new NpgsqlParameter("@ult_vdoc", NpgsqlDbType.Bigint));
            dbcmd.Parameters.Add(new NpgsqlParameter("@ult_pmonto", NpgsqlDbType.Double));
            dbcmd.Parameters.Add(new NpgsqlParameter("@ult_pfecha", NpgsqlDbType.Timestamp));
            dbcmd.Parameters.Add(new NpgsqlParameter("@ult_pdoc", NpgsqlDbType.Bigint));
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
            dbcmd.Parameters[1].Value = codInterno;
            dbcmd.Parameters[2].Value = ROW.Cells["CodClie"].Value.ToString().Replace(" ", string.Empty);
            dbcmd.Parameters[3].Value = ROW.Cells["Descrip"].Value;
            dbcmd.Parameters[4].Value = "03.1";
            dbcmd.Parameters[5].Value = "02.1";
            dbcmd.Parameters[6].Value = ROW.Cells["ID3"].Value.ToString().Replace(" ", string.Empty);
            dbcmd.Parameters[7].Value = ROW.Cells["Direc1"].Value;
            dbcmd.Parameters[8].Value = ROW.Cells["Descto"].Value;
            dbcmd.Parameters[9].Value = false;
            dbcmd.Parameters[10].Value = false;
            dbcmd.Parameters[11].Value = false;
            dbcmd.Parameters[12].Value = false;
            dbcmd.Parameters[13].Value = 0;
            dbcmd.Parameters[14].Value = ROW.Cells["MontoMax"].Value.ToString().Replace(".",",");
            dbcmd.Parameters[15].Value = ROW.Cells["MtoMaxCred"].Value.ToString().Replace(".", ",");
            dbcmd.Parameters[16].Value = 0;
            dbcmd.Parameters[17].Value = ROW.Cells["MontoUV"].Value.ToString().Replace(".", ",");
            dbcmd.Parameters[18].Value = dt.ExtractDate(ROW.Cells["FechaUV"].Value.ToString());
            dbcmd.Parameters[19].Value = ROW.Cells["NumeroUV"].Value.FromStringToInt();
            dbcmd.Parameters[20].Value = ROW.Cells["MontoUP"].Value.ToString().Replace(".", ",");
            dbcmd.Parameters[21].Value = dt.ExtractDate(ROW.Cells["FechaUP"].Value.ToString());
            dbcmd.Parameters[22].Value = ROW.Cells["NumeroUP"].Value.FromStringToInt();
            dbcmd.Parameters[23].Value = 0;
            dbcmd.Parameters[24].Value = ROW.Cells["PagosA"].Value.ToString().Replace(".", ",");
            dbcmd.Parameters[25].Value = 0;
            dbcmd.Parameters[26].Value = ROW.Cells["Saldo"].Value.ToString().Replace(".", ",");
            dbcmd.Parameters[27].Value = "INNOVA";
            dbcmd.Parameters[28].Value = "INNOVA";
            dbcmd.Parameters[29].Value = 1;
            dbcmd.Parameters[30].Value = true;
            dbcmd.Parameters[31].Value = true;
            dbcmd.Parameters[32].Value = ROW.Cells["RetenIva"].Value.ToString().Replace(".", ",");
            dbcmd.Parameters[33].Value = 0;
            dbcmd.Parameters[34].Value = 0;
            dbcmd.Parameters[35].Value = 0;
            dbcmd.Parameters[36].Value = 0;
            dbcmd.Parameters[37].Value = 0;
            dbcmd.Parameters[38].Value = true;
            dbcmd.Parameters[39].Value = true;
            dbcmd.Parameters[40].Value = true;
            dbcmd.Parameters[41].Value = "ESTA DATA FUE MIGRADA, POR FAVOR VERIFICAR LOS DATOS";
            dbcmd.Parameters[42].Value = retencion(ROW.Cells["RetenIva"].Value.ToString());
            dbcmd.Parameters[43].Value = ROW.Cells["Telef"].Value;
            dbcmd.Parameters[44].Value = "";//ROW.Cells["Email"].Value;
            dbcmd.Parameters[45].Value = ROW.Cells["Represent"].Value;
            count += dbcmd.ExecuteNonQuery();
        }
        #endregion

        #region Proveedores
        private void SaProv(SqlConnection objODBCCon)
        {
            string oString = @"Select CodProv, Descrip,Telef, Email, Represent, ID3, Direc1, 
            FechaUC, NumeroUC, MontoUC, FechaUP, NumeroUP, MontoUP, MontoMax, PagosA, PromPago,
            RetenIVA, Saldo FROM SAPROV";
            SqlDataAdapter comm = new SqlDataAdapter(oString, objODBCCon);
            comm.Fill(TAB);
            dataGridView1.DataSource = TAB.Tables[0];
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                RIF.Add(row.Cells["ID3"].Value.ToString());
            }
        }
        private void callbackInsertProvSaint(NpgsqlConnection conn, DataGridViewRow ROW)
        {
            sql = @"INSERT INTO admin.com_prov(org_hijo,cod_interno,prov_hijo,descri,
                        tipo_cont,tipo_pers,rif,direc1,descuento,es_descuento,
                        es_exento,es_retencion,es_monto,monto_min,monto_max,monto_cred,
                        pri_monto,ult_fecha,ult_doc,ult_monto,rec_fecha,rec_doc,rect_monto,pago_max,pago_ade,
                        pago_prom,saldo, reg_usu_cc,reg_usu_cu,reg_estatus,disponible, migrado,
                        porc_ret_iva, observacion, tipo_ret_iva, telefono, email, nomb_persona) 
                        VALUES(@org_hijo,@cod_interno,@prov_hijo,@descri,
                        @tipo_cont,@tipo_pers,@rif,@direc1,@descuento,@es_descuento,
                        @es_exento,@es_retencion,@es_monto,@monto_min,@monto_max,@monto_cred,
                        @pri_monto,@ult_fecha,@ult_doc,@ult_monto,@rect_fecha,@rect_doc,@rect_monto,@pago_max,@pago_ade,
                        @pago_prom,@saldo, @reg_usu_cc,@reg_usu_cu,@reg_estatus,@disponible, 
                        @migrado, @porcretiva, @Observacion, @tipo_ret_iva,@telefono,@email, @nombPersona)";


            NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn);
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
            dbcmd.Parameters.Add(new NpgsqlParameter("@ult_fecha", NpgsqlDbType.Timestamp));
            dbcmd.Parameters.Add(new NpgsqlParameter("@ult_doc", NpgsqlDbType.Varchar));
            dbcmd.Parameters.Add(new NpgsqlParameter("@ult_monto", NpgsqlDbType.Double));
            dbcmd.Parameters.Add(new NpgsqlParameter("@rect_fecha", NpgsqlDbType.Timestamp));
            dbcmd.Parameters.Add(new NpgsqlParameter("@rect_doc", NpgsqlDbType.Varchar));
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
            dbcmd.Parameters[1].Value = codInterno;
            dbcmd.Parameters[2].Value = ROW.Cells["CodProv"].Value.ToString().Replace(" ", string.Empty);
            dbcmd.Parameters[3].Value = ROW.Cells["Descrip"].Value;
            dbcmd.Parameters[4].Value = "03.1";
            dbcmd.Parameters[5].Value = "02.1";
            dbcmd.Parameters[6].Value = ROW.Cells["ID3"].Value.ToString().Replace(" ", string.Empty);
            dbcmd.Parameters[7].Value = ROW.Cells["Direc1"].Value;
            dbcmd.Parameters[8].Value = 0;
            dbcmd.Parameters[9].Value = false;
            dbcmd.Parameters[10].Value = false;
            dbcmd.Parameters[11].Value = false;
            dbcmd.Parameters[12].Value = false;
            dbcmd.Parameters[13].Value = 0;
            dbcmd.Parameters[14].Value = 0;
            dbcmd.Parameters[15].Value = 0;
            dbcmd.Parameters[16].Value = 0;
            dbcmd.Parameters[17].Value = ROW.Cells["FechaUP"].Value; 
            dbcmd.Parameters[18].Value = ROW.Cells["NumeroUP"].Value; 
            dbcmd.Parameters[19].Value = ROW.Cells["MontoUP"].Value; 
            dbcmd.Parameters[20].Value = ROW.Cells["FechaUC"].Value; 
            dbcmd.Parameters[21].Value = ROW.Cells["NumeroUC"].Value; 
            dbcmd.Parameters[22].Value = ROW.Cells["MontoUC"].Value; 
            dbcmd.Parameters[23].Value = 0;
            dbcmd.Parameters[24].Value = 0;
            dbcmd.Parameters[25].Value = 0;
            dbcmd.Parameters[26].Value = ROW.Cells["Saldo"].Value;
            dbcmd.Parameters[27].Value = "INNOVA";
            dbcmd.Parameters[28].Value = "INNOVA";
            dbcmd.Parameters[29].Value = 1;
            dbcmd.Parameters[30].Value = true;
            dbcmd.Parameters[31].Value = true;
            dbcmd.Parameters[32].Value = ROW.Cells["RetenIVA"].Value;
            dbcmd.Parameters[33].Value = "ESTA DATA FUE MIGRADA, POR FAVOR VERIFICAR LOS DATOS";
            dbcmd.Parameters[34].Value = retencion(ROW.Cells["RetenIVA"].Value.ToString());
            dbcmd.Parameters[35].Value = ROW.Cells["Telef"].Value;
            dbcmd.Parameters[36].Value = "";
            dbcmd.Parameters[37].Value = ROW.Cells["Represent"].Value;
            count += dbcmd.ExecuteNonQuery();

        }
        #endregion

        #endregion

        #endregion

        #region Metodos y Eventos

        private void button2_Click(object sender, EventArgs e)
        {
            MigrarData();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            comboBox2.DataSource = null;
            if (radioButton1.Checked)
            {
                var db = DBConn.Instance;
                var c = db.Collection<adminA2>();
                List<adminA2> ad = c.Find(Query.All()).ToList();
                comboBox2.DataSource = ad;
                comboBox2.DisplayMember = "desc";
                comboBox2.ValueMember = "id";
            }
            else
            {
                var db = DBConn.Instance;
                var c = db.Collection<adminSaint>();
                List<adminSaint> ad = c.Find(Query.All()).ToList();
                comboBox2.DataSource = ad;
                comboBox2.DisplayMember = "desc";
                comboBox2.ValueMember = "id";
            }
        }

        private void MigrarData()
        {
            count = 0;
            codInterno = 0;
            NpgsqlConnection conn = new NpgsqlConnection(connectionString);
            conn.Open();
            NpgsqlTransaction t = conn.BeginTransaction();
            foreach (DataGridViewRow ROW in dataGridView1.Rows)
            {
                if (Convert.ToBoolean(ROW.Cells["Migrar"].Value) == true)
                {
                    codInterno++;
                    try
                    {
                        Cursor.Current = Cursors.WaitCursor;
                        splitContainer1.Enabled = false;
                        if (radioButton1.Checked)
                        {
                            switch (comboBox2.Text)
                            {
                                case "Zonas":
                                    {
                                        callbackInsertZona(conn, ROW);
                                        break;
                                    }
                                case "Moneda":
                                    {
                                        callbackInsertMoneda(conn, ROW);
                                        break;
                                    }
                                case "Talento":
                                    {
                                        callbackInsertVendedores(conn, ROW);
                                        break;
                                    }
                                case "Clientes":
                                    {
                                        callbackInsertClientes(conn, ROW);
                                        break;
                                    }
                                case "Proveedores":
                                    {
                                        callbackInsertProv(conn, ROW);
                                        break;
                                    }
                                case "Categorias":
                                    {
                                        callbackInsertCat(conn, ROW);
                                        break;
                                    }
                                case "Depositos":
                                    {
                                        callbackInsertDeposito(conn, ROW);
                                        break;
                                    }
                                case "Cuentas por Cobrar":
                                    {
                                        callbackInsertCxC(conn, ROW);
                                        break;
                                    }
                                case "Cuentas por Pagar":
                                    {
                                        callbackInsertCxP(conn, ROW);
                                        break;
                                    }
                            }
                        }
                        else if (radioButton2.Checked)
                        {
                            switch (comboBox2.Text)
                            {
                                case "Zonas":
                                    {
                                        callbackInsertZona(conn, ROW);
                                        break;
                                    }
                                case "Moneda":
                                    {
                                        callbackInsertMonedaSaint(conn, ROW);
                                        break;
                                    }
                                case "Talento":
                                    {
                                        callbackInsertVendedores(conn, ROW);
                                        break;
                                    }
                                case "Clientes":
                                    {
                                        callbackInsertClientesSaint(conn, ROW);
                                        break;
                                    }
                                case "Proveedores":
                                    {
                                        callbackInsertProvSaint(conn, ROW);
                                        break;
                                    }
                            }
                        }
                    }
                    catch (NpgsqlException ex)
                    {
                        count = 0;
                        logWrite(ex, "");
                        var db = DBConn.Instance;
                        var col = db.Collection<Errores>();
                        try
                        {
                            ROW.Cells["Error"].Value = col.Find(x => x.codigo == ex.Code.ToString()).FirstOrDefault().Desc;
                        }
                        catch (Exception)
                        {
                            ROW.Cells["Error"].Value = "Error no identificado";
                        }
                        ROW.DefaultCellStyle.BackColor = Color.Red;
                        MessageBox.Show(ex.Message);
                        break;
                    }
                }
            }
            splitContainer1.Enabled = true;
            Cursor.Current = Cursors.Default;
            MessageBox.Show("Se migraron exitosamente " + count + " registros", "Atencion", MessageBoxButtons.OK, MessageBoxIcon.Information);
            t.Commit();
            conn.Close();
        }

        private void ExtractA2Table()
        {
            try
            {
                
                TAB.Reset();
                string txtConStr = "DSN=conDBisam";
                OdbcConnection objODBCCon = new OdbcConnection(txtConStr);
                objODBCCon.Open();
                dataGridView1.Columns["Error"].Visible = true;
                dataGridView1.Columns["Migrar"].Visible = true;
                dataGridView1.Columns["numero"].Visible = true;
                switch (comboBox2.Text)
                {
                    case "Zonas":
                        {
                            A2Zonas(objODBCCon);
                            dataGridView1.DataSource = TAB.Tables[0];
                            
                            break;
                        }
                    case "Moneda":
                        {
                            A2Moneda(objODBCCon);
                            dataGridView1.DataSource = TAB.Tables[0];
                            break;
                        }
                    case "Talento":
                        {
                            A2Vendedores(objODBCCon);
                            dataGridView1.DataSource = TAB.Tables[0];
                            break;
                        }
                    case "Clientes":
                        {
                            A2Clientes(objODBCCon);
                            seleccion="FC";
                            rifs(seleccion);
                            break;
                        }
                    case "Proveedores":
                        {
                            A2Prov(objODBCCon);
                            seleccion = "FP";
                            rifs(seleccion);
                            break;
                        }
                    case "Categorias":
                        {
                            A2Cat(objODBCCon);
                            dataGridView1.DataSource = TAB.Tables[0];
                            seleccion = "FD";
                            break;
                        }
                    case "Depositos":
                        {
                            A2Deposito(objODBCCon);
                            dataGridView1.DataSource = TAB.Tables[0];
                            seleccion = "FDP";
                            break;
                        }
                    case "Cuentas por Cobrar":
                        {
                            A2CxC(objODBCCon);
                            dataGridView1.DataSource = TAB.Tables[0];
                            break;
                        }
                    case "Cuentas por Pagar":
                        {
                            A2CxP(objODBCCon);
                            dataGridView1.DataSource = TAB.Tables[0];
                            break;
                        }
                    case "Adelantos (Clientes)":
                        {
                            A2AdelantosCli(objODBCCon);
                            dataGridView1.DataSource = TAB.Tables[0];
                            break;
                        }
                    case "Adelantos (Proveedores)":
                        {
                            A2AdelantosProv(objODBCCon);
                            dataGridView1.DataSource = TAB.Tables[0];
                            break;
                        }
                }
                
                loopdatagrid(seleccion , true, act);
                objODBCCon.Close();
            }
            catch (Exception E)
            {
                MessageBox.Show(E.Message.ToString());
            }

        }

        private void ExtractSaintTable()
        {
            try
            {
                seleccion = "";
                TAB.Reset();
                SqlConnection myConnection = new SqlConnection("server=" + Globals.ServidorSaint + ";" +
                                       "Trusted_Connection=yes;" +
                                       "database=" + Globals.NombBDSaint + "; " +
                                       "connection timeout=30");
                myConnection.Open();
                dataGridView1.Columns["Error"].Visible = true;
                dataGridView1.Columns["Migrar"].Visible = true;
                dataGridView1.Columns["numero"].Visible = true;
                switch (comboBox2.Text)
                {
                    case "Moneda":
                        {
                            SaMoneda(myConnection);
                            break;
                        }
                    case "Clientes":
                        {
                            SaClientes(myConnection);
                            seleccion = "CodClie";
                            break;
                        }
                    case "Proveedores":
                        {
                            SaProv(myConnection);
                            seleccion = "CodProv";
                            break;
                        }
                }
                
                loopdatagridSa(true, seleccion, act);
                myConnection.Close();
            }
            catch(Exception x)
            {
                MessageBox.Show("Hubo un error\n"+x.Message);
            }
            
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
            if (radioButton1.Checked == true)
                ExtractA2Table();
            else if (radioButton2.Checked == true)
                ExtractSaintTable();
            
            button2.Enabled = true;
        }

        private void loopdatagrid(string tabla, bool mod, bool select)
        {
            if (mod)
            {
                int cont = 0;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    cont++;
                    if (!(row.DefaultCellStyle.BackColor == Color.DimGray))
                    {
                        row.Cells["Migrar"].Value = true;
                        row.Cells["numero"].Value = cont;
                        if ((tabla == "FP" || tabla == "FC") && string.IsNullOrWhiteSpace(dt.GetValue<string>(row.Cells[tabla + "_RIF"].Value)) && mod)
                        {
                            row.Cells[tabla + "_RIF"].Value = row.Cells[tabla + "_CODIGO"].Value;
                            row.Cells["Error"].Value = "Rif es un valor provisional";
                            row.DefaultCellStyle.BackColor = Color.Yellow;
                            row.Cells["Migrar"].Value = false;
                        }
                    }
                }
            }else
            {
                
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (select) row.Cells["Migrar"].Value = true;
                    else row.Cells["Migrar"].Value = false;
                }
            }
            
                
        }

        private void loopdatagridSa(bool mod,string tabla,bool select)
        {
            if (mod)
            {
                int cont = 0;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    cont++;
                    if (!(row.DefaultCellStyle.BackColor == Color.DimGray))
                    {
                        row.Cells["Migrar"].Value = true;
                        row.Cells["numero"].Value = cont;
                        if (string.IsNullOrWhiteSpace(dt.GetValue<string>(row.Cells["ID3"].Value)) && tabla!="")
                        {
                            row.Cells["ID3"].Value = row.Cells[tabla].Value;
                            row.Cells["Error"].Value = "Rif es un valor provisional";
                            row.DefaultCellStyle.BackColor = Color.Yellow;
                            row.Cells["Migrar"].Value = false;
                        }
                    }
                }
            }
            else
            {

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (select) row.Cells["Migrar"].Value = true;
                    else row.Cells["Migrar"].Value = false;
                }
            }


        }

        private string tipoContribuyente(string rif)
        {
            if (rif.Replace(" ", string.Empty).StartsWith("j", StringComparison.InvariantCultureIgnoreCase)) return "02.3";
            else return "02.1";
        }

        private string retencion(string ret)
        {
            switch (ret.Replace(" ", string.Empty))
            {
                case "100": return "15.1";
                    
                case "75":  return "15.2";

                case "0": return "15.3";

                default: return "15.3";  
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            loopdatagrid("", false, act);
            act = !act;
        }

        private string tipoOperaA2(string opera, bool tabla)
        {

            switch (opera)
            {
                case "1":
                    {
                        return "20";
                    }
                case "2":
                case "8":
                    {
                        return "31";
                    }
                case "5":
                case "7":
                    {
                        return "32";
                    }
                case "9":
                    {
                        if (tabla) return "31";
                        else return "32";
                    }
                default: return null;
            }

        }

        private double montoEx(double monto, double iva, double baseimp)
        {
            return (monto - iva) - baseimp;
        }

        public static string Impuesto(double monto, NpgsqlConnection conn, double porc, bool tipo)
        {
            double cod = 0;
            string reader;
            string sql="";
            
            sql = @"select codigo from admin.gen_tributo where porcentaje="+porc;
            NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn);
            try
            {
                reader = dbcmd.ExecuteScalar().ToString();
                return reader;
            }
            catch (Exception)
            {
                
                    sql = @"SELECT COUNT(*) FROM admin.gen_tributo";
                    dbcmd = new NpgsqlCommand(sql, conn);
                    reader = dbcmd.ExecuteScalar().ToString();
                    cod = Convert.ToDouble(reader) + 1;
                    sql = @"INSERT INTO admin.gen_tributo(org_hijo,
                        cod_interno,codigo,descri,porcentaje,monto,reg_usu_cc,reg_usu_cu,
                        reg_estatus,disponible) 
                        VALUES(@org_hijo, @codInterno, @codigo, @descri,@porcentaje,@monto,
                        @reg_usu_cc, @reg_usu_cu, @regEstatus, @disponible)";
                
                    dbcmd = new NpgsqlCommand(sql, conn);
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
                    dbcmd.Parameters[1].Value = 1;
                    dbcmd.Parameters[2].Value = "00" + cod;
                    dbcmd.Parameters[3].Value = "Impuesto al" + porc + " %";
                    if (tipo)
                    {
                    dbcmd.Parameters[4].Value = porc;
                    dbcmd.Parameters[5].Value = 0;
                    }else
                    {
                    dbcmd.Parameters[4].Value = 0;
                    dbcmd.Parameters[5].Value = porc;
                    }
                   
                    dbcmd.Parameters[6].Value = "INNOVA";
                    dbcmd.Parameters[7].Value = "INNOVA";
                    dbcmd.Parameters[8].Value = 1;
                    dbcmd.Parameters[9].Value = true;

                    dbcmd.ExecuteNonQuery();

                    reader = "00" + cod;
                return reader;
            }          
        }
        
        private void rifs (string tabla)
        {
            foreach(var z in RIF)
            {
                bool i = false;
                foreach (DataGridViewRow row in (dataGridView1.Rows))
                {
                    if (!String.IsNullOrWhiteSpace(row.Cells[tabla + "_RIF"].Value.ToString()) && row.Cells[tabla+"_RIF"].Value.ToString() == z && i == false) i = true;
                    else if(row.Cells[tabla + "_RIF"].Value.ToString() == z && i == true)
                    {
                        Exception c= new Exception();
                        row.Cells["Error"].Value = "Ya existe un registro con este rif";
                        row.DefaultCellStyle.BackColor = Color.DimGray;
                        row.Cells["Migrar"].Value = false;
                        if (tabla == "FP") logWrite(c, "En proveedores el RIF " + row.Cells[tabla + "_RIF"].Value.ToString() + " esta repetido");
                        else if (tabla == "FC") logWrite(c, "En Clientes el RIF " + row.Cells[tabla + "_RIF"].Value.ToString() + " esta repetido");


                    }
                }
            }
            
        }

        public static void logWrite(Exception x, string codigo)
        {
            if (string.IsNullOrWhiteSpace(codigo))
            {
                File.AppendAllText("Log.txt", DateTime.Now.ToString() + " :" + x.Message + Environment.NewLine);
                
            }else
            {
                File.AppendAllText("Log.txt", DateTime.Now.ToString() + " Error:" + codigo + Environment.NewLine);
            }
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.DefaultCellStyle.BackColor == Color.Red)
                    dataGridView1.CurrentCell = dataGridView1.Rows[row.Index].Cells[0];
            }
        }

        #endregion


    }
}
