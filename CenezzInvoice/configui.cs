using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CenezzInvoice
{
    public partial class configui : Form
    {
        public configui()
        {
            InitializeComponent();
        }

        private void configui_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }

        private void configui_Deactivate(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (uuid.Text != "")
            {
                if (DialogResult.Yes == MessageBox.Show("¿Desea ACTUALIZAR?\r\nlos parametros de configuración.", "--Actualizar parametros--                ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
                {
                    SqlConnection con = new SqlConnection(config.cade);
                    con.Open();
                    string qu = "UPDATE configs SET tarima='" + palletkgs.Text + "', obs1='" + obs1.Text + "', obs2='" + obs2.Text + "', obs3='" + obs3.Text + "', obs4='" + obs4.Text + "', obs5='" + obs5.Text + "'  WHERE id=" + uuid.Text + ";";
                    SqlCommand myCo = new SqlCommand(qu, config.conn);
                    myCo.ExecuteNonQuery();
                    myCo.Dispose();
                    con.Close();
                }

            }

        }

        private void containsadd_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void cveadd_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void configui_Load(object sender, EventArgs e)
        {

            ip.Text = "" + config.srv;
            dbs.Text = "" + config.dbb;
            puerto.Text = "" + config.porto;
            user.Text = "" + config.usr;
            pass.Text = "" + config.pss;
            prefix.Text = "" +config.prefix;
            sufix.Text = "" + config.numemp;


            SqlConnection con = new SqlConnection("" + config.cade);
            con.Open();
            string query = "SELECT year FROM folios ORDER BY year DESC;";
            SqlCommand cm = new SqlCommand(query, con);
            SqlDataAdapter da = new SqlDataAdapter(cm);
            DataTable dt = new DataTable();
            da.Fill(dt);
            int cuenta = dt.Rows.Count;
            if (cuenta > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    ejer.Items.Add("" + row["year"]);
                }
            }
            dt.Dispose(); cm.Dispose(); da.Dispose();


            query = "SELECT TOP(1)* FROM configs ORDER BY id ASC;";
            cm = new SqlCommand(query, con);
            da = new SqlDataAdapter(cm);
            dt = new DataTable();
            da.Fill(dt);
            cuenta = dt.Rows.Count;
            if (cuenta > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    uuid.Text = "" + row["id"];
                    palletkgs.Text = "" + row["tarima"];
                    obs1.Text = "" + row["obs1"];
                    obs2.Text = "" + row["obs2"];
                    obs3.Text = "" + row["obs3"];
                    obs4.Text = "" + row["obs4"];
                    obs5.Text = "" + row["obs5"];
                }
            }
            dt.Dispose();cm.Dispose(); da.Dispose();
            con.Close();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string ruta = @"" + Path.GetDirectoryName(Application.ExecutablePath) + '\\' + "config.ini";
            try
            {
                using (StreamWriter file = new StreamWriter(@"" + ruta))
                {
                    file.Flush();
                    file.Close();
                }
                
            }
            catch
            {
                MessageBox.Show("No existe el archivo de configuración.");
            }

            try
            {
                using (StreamWriter file = new StreamWriter(@"" + ruta + "", true))
                {
                    // AQUI PONER EL UPDATE DEL CONFIG
                    /*
                    0       127.0.0.1
                    1       sa
                    2       momonet
                    3       cenezzimports
                    4       1433
                    5
                    6       dbo
                    */

                    file.WriteLine("" + ip.Text);
                    file.WriteLine("" + user.Text);
                    file.WriteLine("" + pass.Text);
                    file.WriteLine("" + dbs.Text);
                    file.WriteLine("" + puerto.Text);
                    file.WriteLine("" + sufix.Text);
                    file.WriteLine("" + prefix.Text);
                    file.Close();
                }
            }
            catch
            {
                MessageBox.Show("No se pudo actualizar el archivo de configuración.");
            }

            MessageBox.Show("LOS PARAMETROS HAN SIDO ACTUALIZADOS,\r\nREINICIE EL PROGRAMA PARA ACTIVARLOS.", "  --- ACTUALIZANDO CONFIGURACIÓN DE CONEXIÓN ---  ", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

        private void ejer_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ejer.Text != "")
            {
                SqlConnection con = new SqlConnection("" + config.cade);
                con.Open();
                string query = "SELECT folio FROM folios WHERE year= '" + ejer.Text + "';";
                SqlCommand cm = new SqlCommand(query, con);
                SqlDataAdapter da = new SqlDataAdapter(cm);
                DataTable dt = new DataTable();
                da.Fill(dt);
                int cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        folio.Text = "" + row["folio"];
                    }
                }
                dt.Dispose(); cm.Dispose(); da.Dispose();
            }
        }

        private void folio_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) )
            {
                e.Handled = true;
            }
        }

        private void folio_TextChanged(object sender, EventArgs e)
        {
            if (folio.Text != "")
            {
                SqlConnection con = new SqlConnection("" + config.cade);
                con.Open();
                string query = "SELECT number FROM invoicespl WHERE number= '"+ folio.Text  + "-"+ ejer.Text + "';";
                SqlCommand cm = new SqlCommand(query, con);
                SqlDataAdapter da = new SqlDataAdapter(cm);
                DataTable dt = new DataTable();
                da.Fill(dt);

                //MessageBox.Show("" + query);
                int cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    savefolio.Enabled = false;
                }
                else
                {
                    savefolio.Enabled = true;
                }
                dt.Dispose(); cm.Dispose(); da.Dispose();
            }
            else
            {
                savefolio.Enabled = false;
            }
        }

        private void savefolio_Click(object sender, EventArgs e)
        {
            if (ejer.Text != "" && folio.Text != "")
            {
                if (DialogResult.Yes == MessageBox.Show("¿Desea ACTUALIZAR?\r\nel folio del ejercicio: " + ejer.Text + ".", "--Actualizar folios--                ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
                {
                    SqlConnection con = new SqlConnection(config.cade);
                    con.Open();
                    string qu = "UPDATE folios SET folio='" + folio.Text + "' WHERE year='" + ejer.Text + "';";
                    SqlCommand myCo = new SqlCommand(qu, config.conn);
                    myCo.ExecuteNonQuery();
                    myCo.Dispose();
                    con.Close();
                }

            }
        }
    }
}
