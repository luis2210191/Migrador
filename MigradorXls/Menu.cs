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
    public partial class Menu : Form
    {


        //Main main = new Main();
        /// <summary>
        /// Constructor
        /// </summary>
        public Menu()
        {
            Login fLogin = new Login();
            if (fLogin.ShowDialog() == DialogResult.Cancel)
            {
                Environment.Exit(-1);
            }
            InitializeComponent();

            label3.Text += Application.ProductVersion;

        }

        /// <summary>
        /// Evento click del boton que abre la ventana de migracion por formato .xls
        /// </summary>
        private void button1_Click(object sender, EventArgs e)
        {
            if (this.ActiveMdiChild != null) this.ActiveMdiChild.Close();
            var myForm = new Main();
            myForm.MdiParent = this;
            myForm.ControlBox = false;
            myForm.MaximizeBox = false;
            myForm.MinimizeBox = false;
            myForm.ShowIcon = false;
            myForm.Text = "";
            myForm.Dock = DockStyle.Fill;
            myForm.Show();

        }

        /// <summary>
        /// Evento click del boton que cierra el programa
        /// </summary>
        private void button4_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("¿Esta seguro que desea salir?", "Atencion", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
            {
                Application.Exit();
            }
        }

        /// <summary>
        /// Evento click del boton que abre la ventana de migracion desde BD
        /// </summary>
        private void button2_Click(object sender, EventArgs e)
        {
            if (this.ActiveMdiChild != null) this.ActiveMdiChild.Close();
            var myForm = new MigracionBD();
            myForm.MdiParent = this;
            myForm.ControlBox = false;
            myForm.MaximizeBox = false;
            myForm.MinimizeBox = false;
            myForm.ShowIcon = false;
            myForm.Text = "";
            myForm.Dock = DockStyle.Fill;
            myForm.Show();
        }

        /// <summary>
        /// Evento click que abre la seccion de configuracion
        /// </summary>
        private void button5_Click(object sender, EventArgs e)
        {
            config c = new config();
            c.ShowDialog();
        }
    }


}
