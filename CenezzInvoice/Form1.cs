using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CenezzInvoice
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void ribbonButton1_Click(object sender, EventArgs e)
        {
            add add = new add();
            add.MdiParent = this;
            add.ControlBox = false;
            add.MaximizeBox = false;
            add.MinimizeBox = false;
            add.WindowState = FormWindowState.Maximized;
            add.Show();
        }

        private void ribbonButton2_Click(object sender, EventArgs e)
        {
            list lista = new list();
            lista.MdiParent = this;
            lista.ControlBox = false;
            lista.MaximizeBox = false;
            lista.MinimizeBox = false;
            lista.WindowState = FormWindowState.Maximized;
            lista.Show();
        }

        private void ribbonOrbMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

            try
            {
                config.tr.Close();
                config.tr.Dispose();
            }
            catch (System.Exception ex)
            {
                System.ArgumentException argEx = new System.ArgumentException("" + ex);
                //throw argEx;
                MessageBox.Show("" + argEx);
            }
            try
            {
                config.conn.Open();
                //MessageBox.Show("Conexion exitosa", "Atencion", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //config.conn.Close();
            }
            catch (Exception ex)
            {
                config.conn.Close();
                MessageBox.Show("Error al interner abrir la conexion ( " + ex.Message + " )", "Atencion", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
                Environment.Exit(0);
                Application.Exit();
            }
            //MessageBox.Show("TEMP: " + System.IO.Path.GetTempPath());
            System.IO.Path.GetTempPath();
        }

        private void ribbonButton3_Click(object sender, EventArgs e)
        {
            arts arts = new arts();
            arts.MdiParent = this;
            arts.ControlBox = false;
            arts.MaximizeBox = false;
            arts.MinimizeBox = false;
            arts.WindowState = FormWindowState.Maximized;
            arts.Show();
        }

        private void ribbonButton4_Click(object sender, EventArgs e)
        {
            clients clients = new clients();
            clients.MdiParent = this;
            clients.ControlBox = false;
            clients.MaximizeBox = false;
            clients.MinimizeBox = false;
            clients.WindowState = FormWindowState.Maximized;
            clients.Show();
        }

        private void ribbonButton5_Click(object sender, EventArgs e)
        {
            emiters emiters = new emiters();
            emiters.MdiParent = this;
            emiters.ControlBox = false;
            emiters.MaximizeBox = false;
            emiters.MinimizeBox = false;
            emiters.WindowState = FormWindowState.Maximized;
            emiters.Show();
        }

        private void ribbonButton7_Click(object sender, EventArgs e)
        {
            gastos gastos = new gastos();
            gastos.MdiParent = this;
            gastos.ControlBox = false;
            gastos.MaximizeBox = false;
            gastos.MinimizeBox = false;
            gastos.WindowState = FormWindowState.Maximized;
            gastos.Show();
        }

        private void ribbonButton8_Click(object sender, EventArgs e)
        {
            configui configui = new configui();
            configui.MdiParent = this;
            configui.ControlBox = false;
            configui.MaximizeBox = false;
            configui.MinimizeBox = false;
            configui.WindowState = FormWindowState.Maximized;
            configui.Show();
        }
    }
}
