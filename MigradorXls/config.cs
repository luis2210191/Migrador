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
    public partial class config : Form
    {
        int count;
        Config con = new Config();
        /// <summary>
        /// Constructor
        /// </summary>
        public config()
        {
            InitializeComponent();
            var db = DBConn.Instance;
            var col = db.Collection<Config>();
            count = col.Count();
            try
            {
                textBox1.Text = col.Find(x => x.id == count).FirstOrDefault().descServSaint;
                textBox2.Text = col.Find(x => x.id == count).FirstOrDefault().BDSaint;
                textBox3.Text = col.Find(x => x.id == count).FirstOrDefault().nombConexA2;
            }
            catch
            {

            }

        }

        /// <summary>
        /// Evento click del boton para guardar la configuracion
        /// </summary>
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                var db = DBConn.Instance;
                var col = db.Collection<Config>();
                count = col.Count();
                con.id = count + 1;
                con.descServSaint = textBox1.Text;
                con.BDSaint = textBox2.Text;
                con.nombConexA2 = textBox3.Text;
                col.Insert(con);
                MessageBox.Show("Informacion guardada con exito", "Atencion", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception)
            {
                MessageBox.Show("Hubo un error al tratar de guardar la configuracion", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
    }
    public class Config
    {
        public string descServSaint { get; set;}
        public string BDSaint { get; set;}
        public string nombConexA2 { get; set;}
        public int id { get; set;}

        /// <summary>
        /// Constructor
        /// </summary>
        public Config()
        {

        }
        /// <summary>
        /// Sobrecarga del constructor
        /// </summary>
        public Config(string ServSaint, string bdSaint, string ConexA2, int Id)
        {
            this.descServSaint = ServSaint;
            this.BDSaint = bdSaint;
            this.nombConexA2 = ConexA2;
            this.id = Id;
        }
    }
}
