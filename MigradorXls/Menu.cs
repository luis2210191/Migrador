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

        private void button1_Click(object sender, EventArgs e)
        {
            Main main = new Main();
            main.ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Ajustes aj = new Ajustes();
            aj.ShowDialog();
        }

        private void button4_Click(object sender, EventArgs e)
        {
           if( MessageBox.Show("¿Esta seguro que desea salir?", "Atencion", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
            {
                Application.Exit();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MigracionBD bd = new MigracionBD();
            bd.Show();
        }
    }


}
