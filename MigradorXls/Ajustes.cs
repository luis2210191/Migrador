using Npgsql;
using NpgsqlTypes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MigradorXls
{
    public partial class Ajustes : Form
    {
        string connectionstring;
        string sql;
        double valorIni;
        public double[] Porcutil { get; set; }
        public bool status { get; set; }
        DataTable inventario = new DataTable();
        DataTable Codigo1 = new DataTable();
        DataTable Codigo2 = new DataTable();
        /// <summary>
        /// Constructor
        /// </summary>
        public Ajustes()
        {
            InitializeComponent();
        }
        /// <summary>
        /// Evento que se activa durante la carga del form Ajustes
        /// </summary>
        private void Ajustes_Load(object sender, EventArgs e)
        {
            try
            {
                
                DataTable categoria = new DataTable();
                connectionstring = @"Host=" + Globals.Host + ";port=" + Globals.port + ";Database=" + Globals.DB + ";User ID=" + Globals.usuario + ";Password=" + Globals.pass + ";";
                NpgsqlConnection conn = new NpgsqlConnection(connectionstring);
                conn.Open();
                sql = "select cat_hijo,descri from admin.inv_cat";
                NpgsqlCommand com = new NpgsqlCommand(sql, conn);
                NpgsqlDataAdapter ad = new NpgsqlDataAdapter(com);
                ad.Fill(categoria);
                sql = @"SELECT a.codigo, a.descri FROM admin.inv_art AS a JOIN admin.inv_cat_art as b on a.codigo=b.cod_articulo where b.cat_hijo='"+comboBox1.SelectedValue+"' ";
                com = new NpgsqlCommand(sql, conn);
                ad = new NpgsqlDataAdapter(com);
                ad.Fill(Codigo1);
                ad.Fill(Codigo2);
                conn.Close();
                DataRow workRow = categoria.NewRow();
                workRow["descri"] = "Ninguna";
                workRow["cat_hijo"] = "Ninguna";
                categoria.Rows.Add(workRow);
                comboBox1.DataSource = categoria;
                comboBox1.DisplayMember = "descri";
                comboBox1.ValueMember = "cat_hijo";
                

                comboBox2.DataSource = Codigo1;
                comboBox2.DisplayMember ="codigo";
                comboBox2.ValueMember = "descri";

                comboBox3.DataSource = Codigo2;
                comboBox3.DisplayMember = "codigo";
                comboBox3.ValueMember = "descri";
                sql = @"SELECT org_hijo from admin.cfg_org";
                //Obteniendo organizacion
                conn = new NpgsqlConnection(connectionstring);
                NpgsqlCommand dbcmd = new NpgsqlCommand(sql, conn);
                conn.Open();
                Globals.org = dbcmd.ExecuteScalar().ToString();

                sql = "select tipo_porc_utilidad from admin.cfg_preferencia";
                dbcmd = new NpgsqlCommand(sql, conn);
                Globals.pref = (int)dbcmd.ExecuteScalar();
                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }
        /// <summary>
        /// Evento click del boton que realiza la busqueda en el inventario
        /// </summary>
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string condicion = "";
                NpgsqlConnection conn = new NpgsqlConnection(connectionstring);
                conn.Open();
                if (radioButton1.Checked == true)
                {
                    if (checkBox1.Checked == false)
                    {
                        condicion = @"WHERE b.cat_hijo ='" + comboBox1.SelectedValue + "'";
                    }
                    else if (checkBox1.Checked == true && comboBox1.SelectedText == "Ninguna")
                    {
                        condicion = @"Where a.codigo BETWEEN '" + comboBox2.SelectedValue + "' AND '" + comboBox3.SelectedValue + "'";
                    }
                    else
                    {
                        condicion = @"WHERE b.cat_hijo ='" + comboBox1.SelectedValue + "' AND a.codigo BETWEEN '" + comboBox2.Text + "' AND '" + comboBox3.Text + "'";
                    }

                }
                else if (radioButton2.Checked == true)
                {
                    condicion = @"WHERE a.codigo ='" + textBox1.Text + "'";
                }

                sql = @"select distinct a.codigo, a.costo, a.descri, d.utilidad1 as porc_util1,
                        d.utilidad2 as porc_util2, d.utilidad3 as porc_util3, 
                        d.utilidad4 as porc_util4,e.utilidad1, e.utilidad2, e.utilidad3, 
                        e.utilidad4, c.precio1, c.precio2, c.precio3, c.precio4 
                        from admin.inv_art as a Join admin.inv_cat_art as b on a.org_hijo = b.org_hijo 
                        and a.codigo = b.cod_articulo JOIN admin.tvinv003_p as c on a.codigo = c.codigo_art
                        JOIN admin.tvinv003_u as d on a.codigo = d.codigo_art
                        JOIN admin.tvinv003_um as e on a.codigo = e.codigo_art " + condicion + "ORDER BY a.codigo";

                NpgsqlCommand com = new NpgsqlCommand(sql, conn);
                NpgsqlDataAdapter ad = new NpgsqlDataAdapter(com);
                inventario.Clear();
                ad.Fill(inventario);
                dataGridView1.AutoGenerateColumns = false;
                dataGridView1.DataSource = null;
                dataGridView1.Rows.Clear();
                dataGridView1.DataSource = inventario;
                conn.Close();
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if(cell.OwningColumn.Name != "codigo" && cell.OwningColumn.Name != "descripcion" && cell.OwningColumn.Name != "costo" && cell.OwningColumn.Name != "Restaurar" && cell.OwningColumn.Name != "Modificado")
                        {
                            cell.ReadOnly = false;
                        }                        
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Hubo un problema al cargar la informacion", "Error",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
            
        }
        /// <summary>
        /// Evento de cambio de valor del radioButton 1 (categoria)
        /// </summary>
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {                  
                comboBox1.Enabled = true;
                textBox1.Enabled = false;
                checkBox1.Enabled = true;
            }           
        }
        /// <summary>
        /// Evento de cambio de valor del radioButton 2 (Codigo)
        /// </summary>
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked == true)
            {
                comboBox2.Enabled = false;
                comboBox3.Enabled = false;
                comboBox1.Enabled = false;
                textBox1.Enabled = true;
                checkBox1.Enabled = false;
            }
        }
        /// <summary>
        /// Evento de cambio de valor del radioButton 3 (Todos)
        /// </summary>
        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked == true)
            {
                comboBox2.Enabled = false;
                comboBox3.Enabled = false;
                comboBox1.Enabled = false;
                textBox1.Enabled = false;
                checkBox1.Enabled = false;
            }
            
        }
        /// <summary>
        /// Evento de cambio del checkBox que maneja si se usara un rango de valores en la busqueda del inventario
        /// </summary>
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true && radioButton1.Checked == true)
            {
                comboBox2.Enabled = true;
                comboBox3.Enabled = true;
            }
            else
            {
                comboBox2.Enabled = false;
                comboBox3.Enabled = false;
            }
        }
        /// <summary>
        /// Evento de cambio de valor del combobox1 (Categoria)
        /// </summary>
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(comboBox1.SelectedValue.ToString()=="Ninguna")
            {
                sql = @"SELECT a.codigo, a.descri FROM admin.inv_art AS a JOIN admin.inv_cat_art as b on a.codigo=b.cod_articulo";

            }
            else
            {
                sql = @"SELECT a.codigo, a.descri FROM admin.inv_art AS a JOIN admin.inv_cat_art as b on a.codigo=b.cod_articulo where b.cat_hijo='" + comboBox1.SelectedValue + "' ";

            }
            NpgsqlConnection conn = new NpgsqlConnection(connectionstring);
            conn.Open();
            Codigo1.Clear();
            Codigo2.Clear();
            NpgsqlCommand  com = new NpgsqlCommand(sql, conn);
            NpgsqlDataAdapter ads = new NpgsqlDataAdapter(com);
            ads.Fill(Codigo1);
            ads.Fill(Codigo2);
            conn.Close();
           
            comboBox2.DataSource = Codigo1;
            comboBox2.DisplayMember = "codigo";
            comboBox2.ValueMember = "descri";

            comboBox3.DataSource = Codigo2;
            comboBox3.DisplayMember = "codigo";
            comboBox3.ValueMember = "descri";
        }
        /// <summary>
        /// Evento de finalizacion de edicion de celda en datagrid
        /// </summary>
        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            
            if (valorIni != Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value))
            {
                int index = e.ColumnIndex;

                string columnName = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].OwningColumn.Name;

                int size = columnName.Length;

                string lastChar = columnName.Substring(size - 1);

                if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].OwningColumn.Name.Contains("porc_util"))
                {
                    dataGridView1.Rows[e.RowIndex].Cells["tipoAjuste"].Value = 1;
                    dataGridView1.Rows[e.RowIndex].Cells["precio" + lastChar + ""].Value = calculoPrecio(Globals.pref, (double)dataGridView1.Rows[e.RowIndex].Cells["costo"].Value, (double)dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value);
                    dataGridView1.Rows[e.RowIndex].Cells["utilidad" + lastChar + ""].Value = calculoUtilidad((double)dataGridView1.Rows[e.RowIndex].Cells["costo"].Value, (double)dataGridView1.Rows[e.RowIndex].Cells["precio" + lastChar + ""].Value);

                }
                else if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].OwningColumn.Name.Contains("utilidad"))
                {
                    dataGridView1.Rows[e.RowIndex].Cells["tipoAjuste"].Value = 1;
                    dataGridView1.Rows[e.RowIndex].Cells["precio" + lastChar + ""].Value = (double)dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value + (double)dataGridView1.Rows[e.RowIndex].Cells["costo"].Value;
                    dataGridView1.Rows[e.RowIndex].Cells["porc_util" + lastChar + ""].Value = calculoPorcentaje(Globals.pref,(double)dataGridView1.Rows[e.RowIndex].Cells["precio" + lastChar + ""].Value, (double)dataGridView1.Rows[e.RowIndex].Cells["costo"].Value);
                }
                
                else
                {
                    dataGridView1.Rows[e.RowIndex].Cells["tipoAjuste"].Value = 0;
                    dataGridView1.Rows[e.RowIndex].Cells["porc_util" + lastChar + ""].Value = calculoPorcentaje(Globals.pref, (double)dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value, (double)dataGridView1.Rows[e.RowIndex].Cells["costo"].Value);
                    dataGridView1.Rows[e.RowIndex].Cells["utilidad" + lastChar + ""].Value = calculoUtilidad((double)dataGridView1.Rows[e.RowIndex].Cells["costo"].Value, (double)dataGridView1.Rows[e.RowIndex].Cells["precio" + lastChar + ""].Value);
                }
                

                foreach (DataGridViewCell cell in dataGridView1.Rows[e.RowIndex].Cells)
                {
                    if (cell.OwningColumn.Name != "codigo" && cell.OwningColumn.Name != "descripcion" && cell.OwningColumn.Name != "costo" && cell.OwningColumn.Name != "Restaurar" && cell.OwningColumn.Name != "Modificado")
                    {
                        if (!cell.ReadOnly && cell.OwningColumn.Name.Contains(lastChar) && cell.ColumnIndex != index)
                        {
                            cell.ReadOnly = true;
                            cell.Style.BackColor = Color.Gray;
                        }
                    }
                }
                dataGridView1.Rows[e.RowIndex].Cells[0].Value = true;
            }   
        }
        /// <summary>
        /// Evento de inicio de edicion de celda del datagrid
        /// </summary>
        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            valorIni = Convert.ToDouble(dataGridView1.SelectedCells[0].Value);
        }
        /// <summary>
        /// Evento click del boton para aplicar un porcentaje de utilidad a todos los registros en el grid
        /// </summary>
        private void button2_Click(object sender, EventArgs e)
        {
            using (var f = new PorcUtil() { Owner = this })
            {
                f.ShowDialog();
                if (f.DialogResult == DialogResult.OK)
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    { 
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                             switch (cell.OwningColumn.Name)
                                {
                                    case "porc_util1":
                                        {
                                            cell.Value = Porcutil[0];
                                            break;
                                        }
                                    case "porc_util2":
                                        {
                                            cell.Value = Porcutil[1];
                                            break;
                                        }
                                    case "porc_util3":
                                        {
                                            cell.Value = Porcutil[2];
                                            break;
                                        }
                                    case "porc_util4":
                                        {
                                            cell.Value = Porcutil[3];
                                            break;
                                        }
                                }
                                                     
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Evento click para el boton que realiza el ajuste de precios
        /// </summary>
        private void button3_Click(object sender, EventArgs e)
        {
            string codigos;
            int COUNT = 1;
            double TOTALP=0;
            int cantidad=0;
            bool STATUS=false;
            try
            {
                NpgsqlConnection conn = new NpgsqlConnection(connectionstring);
                conn.Open();
                NpgsqlTransaction t = conn.BeginTransaction();
                NpgsqlCommand com = new NpgsqlCommand();
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {

                    if (Convert.ToBoolean(row.Cells[0].Value) == true)
                    {
                        STATUS = true;
                        TOTALP += Convert.ToDouble(row.Cells["Costo"].Value);
                        codigos = row.Cells[3].Value.ToString();
                        for (COUNT = 1;COUNT<=4;COUNT++)
                        {
                            sql = @"UPDATE admin.inv_art_precio SET porc_utilidad=@porcutilidad" + COUNT + ",utilidad=@utilidad" + COUNT + ",precio =@precio" + COUNT + " WHERE cod_articulo = @codigo AND cod_precio='0" + COUNT + "'";
                            com = new NpgsqlCommand(sql, conn);
                            com.Parameters.Add(new NpgsqlParameter("@porcutilidad" + COUNT + "", NpgsqlDbType.Double));
                            com.Parameters.Add(new NpgsqlParameter("@utilidad" + COUNT + "", NpgsqlDbType.Double));
                            com.Parameters.Add(new NpgsqlParameter("@precio" + COUNT + "", NpgsqlDbType.Double));
                            com.Parameters.Add(new NpgsqlParameter("@codigo", NpgsqlDbType.Varchar));
                            com.Prepare();
                            com.Parameters[0].Value = row.Cells["porc_util" + COUNT + ""].Value;
                            com.Parameters[1].Value = row.Cells["utilidad" + COUNT + ""].Value;
                            com.Parameters[2].Value = row.Cells["precio" + COUNT + ""].Value;
                            com.Parameters[3].Value = row.Cells["codigo"].Value;
                            com.ExecuteNonQuery();
                        }
                        
                        cantidad++;
                    }
                }
               

                if (STATUS==true)
                {
                    
                    sql = @"INSERT INTO admin.int_ajuste_precio(org_hijo,descri,total_precio,total_utilidad, 
                    reg_usu_cc, reg_estatus, nro_items) VALUES(@org_hijo,@descri,@total_precio,@total_utilidad, 
                    @reg_usu_cc, @reg_estatus, @nro_items)";
                    com = new NpgsqlCommand(sql, conn);
                    com.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
                    com.Parameters.Add(new NpgsqlParameter("@descri", NpgsqlDbType.Varchar));
                    com.Parameters.Add(new NpgsqlParameter("@total_precio", NpgsqlDbType.Double));
                    com.Parameters.Add(new NpgsqlParameter("@total_utilidad", NpgsqlDbType.Double));
                    com.Parameters.Add(new NpgsqlParameter("@reg_usu_cc", NpgsqlDbType.Varchar));
                    com.Parameters.Add(new NpgsqlParameter("@reg_estatus", NpgsqlDbType.Integer));
                    com.Parameters.Add(new NpgsqlParameter("@nro_items", NpgsqlDbType.Integer));

                    com.Prepare();

                    com.Parameters[0].Value = Globals.org;
                    com.Parameters[1].Value = "AJUSTE DE PRECIOS";
                    com.Parameters[2].Value = TOTALP;
                    com.Parameters[3].Value = "0";
                    com.Parameters[4].Value = "INNOVA";
                    com.Parameters[5].Value = 1;
                    com.Parameters[6].Value = cantidad;

                    com.ExecuteNonQuery();

                    sql = @"SELECT doc from admin.int_ajuste_precio order by fecha_reg desc";
                    com = new NpgsqlCommand(sql, conn);
                    string reader = com.ExecuteScalar().ToString();
                    //Insercion del detalle del ajuste
                    int Item = 1;
                    foreach (DataGridViewRow ROW2 in dataGridView1.Rows)
                    {
                        if (Convert.ToBoolean(ROW2.Cells[0].Value)==true)
                        {
                            
                            sql = @"INSERT INTO admin.int_ajuste_precio_det(org_hijo,doc,cod_alterno,cod_articulo,
                        costo,costo_promedio,fecha,tipo_ajuste,item) VALUES(@org_hijo,@doc,
                        @cod_alterno,@cod_articulo,@costo,@costo_promedio,@fecha,@tipo_ajuste,@item)";
                            com = new NpgsqlCommand(sql, conn);
                            com.Parameters.Add(new NpgsqlParameter("@org_hijo", NpgsqlDbType.Varchar));
                            com.Parameters.Add(new NpgsqlParameter("@doc", NpgsqlDbType.Bigint));
                            com.Parameters.Add(new NpgsqlParameter("@cod_alterno", NpgsqlDbType.Varchar));
                            com.Parameters.Add(new NpgsqlParameter("@cod_articulo", NpgsqlDbType.Varchar));
                            com.Parameters.Add(new NpgsqlParameter("@costo", NpgsqlDbType.Double));
                            com.Parameters.Add(new NpgsqlParameter("@costo_promedio", NpgsqlDbType.Double));
                            com.Parameters.Add(new NpgsqlParameter("@fecha", NpgsqlDbType.Date));
                            com.Parameters.Add(new NpgsqlParameter("@tipo_ajuste", NpgsqlDbType.Integer));
                            com.Parameters.Add(new NpgsqlParameter("@item", NpgsqlDbType.Integer));

                            com.Prepare();

                            com.Parameters[0].Value = Globals.org;
                            com.Parameters[1].Value = Convert.ToInt64(reader);
                            com.Parameters[2].Value = ROW2.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);
                            com.Parameters[3].Value = ROW2.Cells["codigo"].Value.ToString().Replace(" ", string.Empty);
                            com.Parameters[4].Value = ROW2.Cells["costo"].Value;
                            com.Parameters[5].Value = ROW2.Cells["costo"].Value;
                            com.Parameters[6].Value = DateTime.Now;
                            com.Parameters[7].Value = ROW2.Cells["tipoAjuste"].Value;
                            com.Parameters[8].Value = Item;
                            com.ExecuteNonQuery();
                            Item++;
                        }
                    }
                }
                    
               t.Commit();
                conn.Close();
                MessageBox.Show("Ajuste realizado con exito");
            }
            catch (Exception EX)
            {
                MessageBox.Show("Excepcion : "+EX.Message);
            }
        }
        /// <summary>
        /// Evento click del contenido de una celda en el datagrid
        /// </summary>
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

            try
            {
                var senderGrid = (DataGridView)sender;
                DataTable art = new DataTable();

                if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn &&
                    e.RowIndex >= 0)
                {
                    string codigo;
                    NpgsqlConnection conn = new NpgsqlConnection(connectionstring);
                    conn.Open();
                    codigo = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                    sql = @"select distinct a.codigo, a.costo, a.descri, d.utilidad1 as porc_util1,
                        d.utilidad2 as porc_util2, d.utilidad3 as porc_util3, 
                        d.utilidad4 as porc_util4,e.utilidad1, e.utilidad2, e.utilidad3, 
                        e.utilidad4, c.precio1, c.precio2, c.precio3, c.precio4 
                        from admin.inv_art as a Join admin.inv_cat_art as b on a.org_hijo = b.org_hijo 
                        and a.codigo = b.cod_articulo JOIN admin.tvinv003_p as c on a.codigo = c.codigo_art
                        JOIN admin.tvinv003_u as d on a.codigo = d.codigo_art
                        JOIN admin.tvinv003_um as e on a.codigo = e.codigo_art WHERE a.codigo='" + codigo + "'";
                    NpgsqlCommand com = new NpgsqlCommand(sql, conn);
                    NpgsqlDataAdapter ads = new NpgsqlDataAdapter(com);
                    ads.Fill(art);
                    conn.Close();

                    //Asignar valores originales de la fila
                    dataGridView1.Rows[e.RowIndex].Cells[5].Value = art.Rows[0][1];
                    dataGridView1.Rows[e.RowIndex].Cells[6].Value = art.Rows[0][3];
                    dataGridView1.Rows[e.RowIndex].Cells[7].Value = art.Rows[0][7];
                    dataGridView1.Rows[e.RowIndex].Cells[8].Value = art.Rows[0][11];
                    dataGridView1.Rows[e.RowIndex].Cells[9].Value = art.Rows[0][4];
                    dataGridView1.Rows[e.RowIndex].Cells[10].Value = art.Rows[0][8];
                    dataGridView1.Rows[e.RowIndex].Cells[11].Value = art.Rows[0][12];
                    dataGridView1.Rows[e.RowIndex].Cells[12].Value = art.Rows[0][5];
                    dataGridView1.Rows[e.RowIndex].Cells[13].Value = art.Rows[0][9];
                    dataGridView1.Rows[e.RowIndex].Cells[14].Value = art.Rows[0][13];
                    dataGridView1.Rows[e.RowIndex].Cells[15].Value = art.Rows[0][6];
                    dataGridView1.Rows[e.RowIndex].Cells[16].Value = art.Rows[0][10];
                    dataGridView1.Rows[e.RowIndex].Cells[17].Value = art.Rows[0][14];
                    dataGridView1.Rows[e.RowIndex].Cells[0].Value = false;
                    //Volver los campos recien restaurados editables
                    foreach (DataGridViewCell cell in dataGridView1.Rows[e.RowIndex].Cells)
                    {
                        if (cell.ColumnIndex >= 5)
                        {
                            cell.ReadOnly = false;
                            cell.Style.BackColor = Color.White;
                        }
                    }

                }
            }
            catch(Exception x)
            {
                MessageBox.Show("Se produjo una excepcion del tipo : "+x);
            }            
        }
        /// <summary>
        /// Metodo para el calculo del procentaje de utilidad teniendo el precio, el costo y el tipo del calculo(lineal o financiero))
        /// </summary>
        private double calculoPorcentaje(int tipo, double precio, double costo)
        {
            double result = 0;
            if (tipo == 1)
            {
                result = (precio - costo) * 100 / costo;//lineal
            }else
            {
                result = (precio - costo) * 100 / precio;//financiero
            }
            return result;
        }
        /// <summary>
        /// Metodo para el calculo del precio teniendo el porcentaje de utilidad, el costo y el tipo del calculo(lineal o financiero)
        /// </summary>
        private double calculoPrecio(int tipo,double costo, double porc)
        {
            double result=0;
            if (tipo == 1)
                {
                    result = costo + (costo * porc / 100);//lineal
                }
                else
                {
                    result = costo / ((100 - porc)/100);//Financiero
                }

                return result;
        }
        /// <summary>
        /// Metodo para el calculo del la utilidad teniendo el precio y el costo
        /// </summary>   
        private double calculoUtilidad(double costo, double precio)
        {
               double result = precio - costo;            
            return result;
        }
        
        
    }
}
