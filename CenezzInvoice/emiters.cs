using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CenezzInvoice
{
    public partial class emiters : Form
    {
        public emiters()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        internal static string invoice_query = "SELECT  id AS Clave, nom AS Nombre, nif AS NIF, pais AS Pais FROM empresasipl ORDER BY Nombre ASC ";
        internal static string invoice_querys = "SELECT  id AS Clave, nom AS Nombre, nif AS NIF, pais AS Pais FROM empresasipl ";

        private void button2_Click(object sender, EventArgs e)
        {
            if (cve.Text != "")
            {
                if (DialogResult.Yes == MessageBox.Show("¿Desea ELIMINAR?", "--Eliminar Emisor --                ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
                {
                    string cellValueid = cve.Text;
                    SqlConnection con = new SqlConnection(config.cade);
                    con.Open();
                    string qu = "DELETE FROM empresasipl WHERE id=" + cellValueid + ";";
                    SqlCommand myCo = new SqlCommand(qu, config.conn);
                    myCo.ExecuteNonQuery();
                    myCo.Dispose();

                    string range = "";
                    //range = " WHERE CAST(fecha AS DATE) >= CAST('" + inicial + "'  AS DATE) AND CAST(fecha AS DATE) <= CAST('" + final + "' AS DATE)";
                    SqlDataAdapter DA = new SqlDataAdapter();
                    string sqlSelectAll = invoice_query + " " + range + ";";
                    DA.SelectCommand = new SqlCommand(sqlSelectAll, con);

                    DataTable table = new DataTable();
                    DA.Fill(table);

                    BindingSource bSource = new BindingSource();
                    bSource.DataSource = table;
                    lister.DataSource = bSource;

                    con.Close();

                }
            }
            else { MessageBox.Show("Debes seleccionar un emisor para eliminarlo."); }
        }

        private void emiters_Load(object sender, EventArgs e)
        {


            SqlConnection con = new
            SqlConnection("" + config.cade);
            con.Open();

            string range = "";
            //range = " WHERE CAST(fecha AS DATE) >= CAST('" + inicial + "'  AS DATE) AND CAST(fecha AS DATE) <= CAST('" + final + "' AS DATE)";
            SqlDataAdapter DA = new SqlDataAdapter();
            string sqlSelectAll = invoice_query + " " + range + ";";
            DA.SelectCommand = new SqlCommand(sqlSelectAll, con);

            DataTable table = new DataTable();
            DA.Fill(table);

            BindingSource bSource = new BindingSource();
            bSource.DataSource = table;
            lister.DataSource = bSource;

            paiss.Items.Add("");
            string query = "SELECT pais FROM empresasipl GROUP BY pais ORDER BY pais ASC;";
            SqlCommand cm = new SqlCommand(query, con);
            SqlDataAdapter da = new SqlDataAdapter(cm);
            DataTable dt = new DataTable();
            da.Fill(dt);
            int cuenta = dt.Rows.Count;
            if (cuenta > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    paiss.Items.Add("" + row["pais"]);
                }
            }
            da.Dispose(); cm.Dispose(); dt.Dispose();


            pais.Items.Add("");
            paise.Items.Add("");
            query = "SELECT pais FROM empresasipl GROUP BY pais ORDER BY pais ASC;";
            cm = new SqlCommand(query, con);
            da = new SqlDataAdapter(cm);
            dt = new DataTable();
            da.Fill(dt);
            cuenta = dt.Rows.Count;
            if (cuenta > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    pais.Items.Add("" + row["pais"]);
                    paise.Items.Add("" + row["pais"]);
                }
            }
            da.Dispose(); cm.Dispose(); dt.Dispose();
            resizegrid();
            con.Close();
        }

        private void resizegrid()
        {
            lister.Columns[0].Width = 40;
            lister.Columns[1].Width = 370;
            lister.Columns[2].Width = 150;
            lister.Columns[3].Width = 100;
            //lister.Columns[4].Width = 100;
            //lister.Columns[5].Width = 50;
            //lister.Columns[5].Visible = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnection con = new
            SqlConnection("" + config.cade);    
            con.Open();

            if (nifs.Text != "" || noms.Text != "" || paiss.Text != "")
            {


                string range = "";
                string nifss = "", nomss = "", paisss = "";
                if (nifs.Text != "")
                {
                    nifss = " AND nif LIKE '%" + nifs.Text + "%' ";
                }
                if (noms.Text != "")
                {
                    nomss = " AND nom  LIKE '%" + noms.Text + "%' ";
                }
                if (paiss.Text != "")
                {
                    paisss = " AND pais='" + paiss.Text + "' ";
                }

                range = "" + nifss + "" + nomss + "" + paisss;

                int largo = range.Length;
                if (largo >= 4)
                {
                    range = range.Substring(4);
                }

                range = " WHERE " + range;
                //range = " WHERE CAST(fecha AS DATE) >= CAST('" + inicial + "'  AS DATE) AND CAST(fecha AS DATE) <= CAST('" + final + "' AS DATE)";
                SqlDataAdapter DA = new SqlDataAdapter();
                string sqlSelectAll = invoice_querys + " " + range + "  ORDER BY nom ASC ;";
                DA.SelectCommand = new SqlCommand(sqlSelectAll, con);

                DataTable table = new DataTable();
                DA.Fill(table);

                BindingSource bSource = new BindingSource();
                bSource.DataSource = table;
                lister.DataSource = bSource;

            }
            else
            {

                string range = "";
                //range = " WHERE CAST(fecha AS DATE) >= CAST('" + inicial + "'  AS DATE) AND CAST(fecha AS DATE) <= CAST('" + final + "' AS DATE)";
                SqlDataAdapter DA = new SqlDataAdapter();
                string sqlSelectAll = invoice_query + " " + range + ";";
                DA.SelectCommand = new SqlCommand(sqlSelectAll, con);

                DataTable table = new DataTable();
                DA.Fill(table);

                BindingSource bSource = new BindingSource();
                bSource.DataSource = table;
                lister.DataSource = bSource;
            }
            con.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (cvee.Text != "" && cve.Text != "")
            {

                if (DialogResult.Yes == MessageBox.Show("¿Desea guardar?", " -- Almacenar cambios del emisor --                ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
                {


                    string idl = "";
                    SqlConnection con = new SqlConnection(config.cade);
                    con.Open();


                    string qu = "UPDATE empresasipl  SET  nom='" + nome.Text + "', nif='" + nife.Text + "', calle='" + callee.Text + "', num='" + numee.Text + "', numi='" + numie.Text + "', col='" + cole.Text + "', mun='" + mune.Text + "', cd='" + cde.Text + "', estado='" + edoe.Text + "', pais='" + paise.Text + "', cp='" + cpe.Text + "' ";
                    qu = qu + "WHERE id=" + cvee.Text + ";";
                    SqlCommand myCo = new SqlCommand(qu, config.conn);
                    myCo.ExecuteNonQuery();
                    myCo.Dispose();


                    string range = " WHERE id =" + cvee.Text + ";";
                    //range = " WHERE CAST(fecha AS DATE) >= CAST('" + inicial + "'  AS DATE) AND CAST(fecha AS DATE) <= CAST('" + final + "' AS DATE)";
                    SqlDataAdapter DA = new SqlDataAdapter();
                    //string sqlSelectAll = invoice_query + " " + range + "";
                    string sqlSelectAll = invoice_querys +" " + range;
                    DA.SelectCommand = new SqlCommand(sqlSelectAll, con);

                    DataTable table = new DataTable();
                    DA.Fill(table);

                    BindingSource bSource = new BindingSource();
                    bSource.DataSource = table;
                    lister.DataSource = bSource;
                    con.Close();

                    cvee.Text = "";
                    nome.Text = "";
                    nife.Text = "";
                    callee.Text = "";
                    numee.Text = "";
                    numie.Text = "";
                    cole.Text = "";
                    mune.Text = "";
                    cde.Text = "";
                    edoe.Text = "";
                    paise.Text = "";
                }

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show("¿Desea guardar?", " -- Almacenar nuevo cliente --                ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
            {
                SqlConnection con = new SqlConnection(config.cade);
                con.Open();


                string qu = "INSERT INTO  empresasipl (nom, nif, calle, num, numi, col, mun, cd, estado, pais, cp) ";
                qu = qu + "VALUES('" + nom.Text + "','" + nif.Text + "','" + calle.Text + "','" + nume.Text + "','" + numi.Text + "','" + col.Text + "','" + mun.Text + "','" + cd.Text + "','" + edo.Text + "','" + pais.Text + "','" + cp.Text + "');SELECT SCOPE_IDENTITY();";
                SqlCommand myCo = new SqlCommand(qu, config.conn);
                string uuid = "" + myCo.ExecuteScalar().ToString();
                myCo.Dispose();

                cve.Text = "";
                nom.Text = "";
                nif.Text = "";
                calle.Text = "";
                nume.Text = "";
                numi.Text = "";
                col.Text = "";
                mun.Text = "";
                cd.Text = "";
                edo.Text = "";
                pais.Text = "";


                string range = " WHERE id =" + uuid + ";";
                //range = " WHERE CAST(fecha AS DATE) >= CAST('" + inicial + "'  AS DATE) AND CAST(fecha AS DATE) <= CAST('" + final + "' AS DATE)";
                SqlDataAdapter DA = new SqlDataAdapter();
                //string sqlSelectAll = invoice_query + " " + range + "";
                string sqlSelectAll = invoice_querys +" " + range;
                DA.SelectCommand = new SqlCommand(sqlSelectAll, con);

                DataTable table = new DataTable();
                DA.Fill(table);

                BindingSource bSource = new BindingSource();
                bSource.DataSource = table;
                lister.DataSource = bSource;
                cve.Text = "" + uuid;
                con.Close();
            }
        }

        private void lister_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            cvee.Text = "";
            nome.Text = "";
            nife.Text = "";
            callee.Text = "";
            numee.Text = "";
            numie.Text = "";
            cole.Text = "";
            mune.Text = "";
            cde.Text = "";
            edoe.Text = "";
            paise.Text = "";

            if (e.RowIndex != -1)
            {

                var dataIndexNo = lister.Rows[e.RowIndex].Index.ToString();
                string cellValueid = lister.Rows[e.RowIndex].Cells[0].Value.ToString();
                //string cellValueid = lister.Rows[e.RowIndex].Cells[5].Value.ToString();

                cve.Text = "" + cellValueid;
                SqlConnection con = new
                SqlConnection("" + config.cade);
                con.Open();
                string query = "SELECT  * FROM empresasipl WHERE id=" + cellValueid + ";";
                SqlCommand cm = new SqlCommand(query, con);
                SqlDataAdapter da = new SqlDataAdapter(cm);
                DataTable dt = new DataTable();
                da.Fill(dt);
                int cuenta = dt.Rows.Count;
                if (cuenta > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        cvee.Text = "" + row["id"];
                        nome.Text = "" + row["nom"];
                        nife.Text = "" + row["nif"];
                        callee.Text = "" + row["calle"];
                        numee.Text = "" + row["num"];
                        numie.Text = "" + row["numi"];
                        cole.Text = "" + row["col"];
                        mune.Text = "" + row["mun"];
                        cde.Text = "" + row["cd"];
                        edoe.Text = "" + row["estado"];
                        paise.Text = "" + row["pais"];
                        cpe.Text = "" + row["cp"];

                        
                    }
                }
                da.Dispose(); cm.Dispose(); dt.Dispose();
                con.Close();
            }

        }

        private void emiters_KeyUp(object sender, KeyEventArgs e)
        {

            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }

        private void emiters_Deactivate(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cp_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
            /*
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
            */
        }

        private void cpe_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
            /*
            // If you want, you can allow decimal (float) numbers
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
            */
        }
    }
}
